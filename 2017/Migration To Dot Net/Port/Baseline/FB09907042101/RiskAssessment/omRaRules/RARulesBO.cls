VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "RARulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'PROG   Date        Description
'LH     26/10/2006  Phase 2: Created
'MCh    08/11/2006  EP2_34 - Rules added: 3565, 3575, 3580, 3585, 3600,
'                                         3610, 3615, 3620, 3625, 3630, 3640
'LH     08/11/2005  EP2_34 - Rules added: 3675, 3690, 3700, 3705, 3710
'MCh    08/11/2006  EP2_34 - Rules added: 3550, 3660
'LH     09/11/2006  EP2_34 - Rules added: 3520, 3525, 3535, 3550, 3555, 3605, 3670
'MCh    13/11/2006  EP2_34 - Rules added: 3590, 3595, 6500, 6525
'LH     15/11/2006  EP2_78 - Calls valuation rules
'LH     16/11/2006  EP2_34 - Added Mortgage Type Rule Triggers facility
'LH     17/11/2006  EP2_34 - Added GetGlobalParameter() and called it from rules where appropriate.
'LH     17/11/2006  EP2_34 - Fixed GetGlobalParameter()
'MCh    23/11/2006  EP2_34 - Added DateDiffInYearsRoundedUp and updated rules 3520, 3525, 3535 to us it as spec'd.
'DS     24/11/2006  EP2_34 - Added Rule 3665
'OS     24/11/2006  EP2_34 - Added Rule 3515
'DS     25/11/2006  EP2_34 - Added Rule 3695
'DS     27/11/2006  EP2_34 - Modified Rules 3665 and 3695. Rule 3516 modified by OS
'OS     28/11/2006  EP2_34 - Modified Rules 3570 and 3575. Added Rule 3645
'DS     28/11/2006  EP2_34 - Added DIP Rule 3655
'OS     29/11/2006  EP2_34 - Added Rule 3650
'DS     29/11/2006  EP2_34 - Added App Rule 6810
'DS     30/11/2006  EP2_34 - Added App Rule 6805
'OS     01/12/2006  EP2_34 - Added Rules 3730, 6785, 6790, 6795, 6800
'LH     05/12/2006  EP2_34 - Added rule 3925
'DS     02/12/2006  EP2_34 - Added App Rule 6555
'DS     05/12/2006  EP2_34 - Added App Rule 6625
'OS     05/12/2006  EP2_34 - Added App Rules 6530, 6535 and 6780
'LH     05/12/2006  EP2_34 - Replaced global param code with GetGlobalParameter() in rules where appropriate.
'LH     06/12/2006  EP2_34 - Errors now correctly handled from ValuationRulesBO component rule errors
'LH     06/12/2006  EP2_34 - Changed experian based rule checks from LP_RULE to APP_RULE
'OS     06/12/2006  EP2_34 - Added rule values to 6530, 6535
'LH     06/12/2006  EP2_331 - Changed values for Enum RiskAssessmentDecision
'LH     07/12/2006  EP2_331 - Finished off rule 3925 and added rule 3935.
'LH     07/12/2006  EP2_331 - Added rule 3930 and rule 6775.
'LH     08/12/2006  EP2_331 - Changed call to rules 3925 & 3935 to check web ingested/paper case
'                             only after checking if the rule should be run.
'LH     08/12/2006  EP2_331 - Changed rule 3580 to include additional XML node check.
'LH     08/12/2006  EP2_331 - Changed raiseerror() calls in rules to a fail and writing a line to the log file
'LH     08/12/2006  EP2_331 - Added extra XML node checks to 6525 and added extra check in truedatediff()
'LH     12/12/2006  EP2_432 - Added rule 6765
'LH     13/12/2006  EP2_432 - Added rule 3755, 3830, additional XML node checking to 3565
'LH     13/12/2006  EP2_477 - Changed rules 3600, 3610, 3615, 3620 to include "EMPT" according to spec change.
'LH     14/12/2006  EP2_490 - Added rules 6635, 6820, 6825, 6830, 6835, 6840, 6845, 6850, 6855, 6860, 6865,
'                             6870, 6875, 6880, 6595, 6755, 6640
'LH     14/12/2006  EP2_521 - Fixed rule 3670, 3675 to return the response after a failure has been found and not before
'LH     15/12/2006  EP2_508 - Fixed rule 3600, 3610, 3615, 3620, 6555, 6785, 6795, 6800
'LH     15/12/2006  EP2_496 - Fixed rule 3730
'LH     15/12/2006  EP2_524 - Changed rules 6785, 6795, 6800, 6805, 6810 according to spec change
'LH     15/12/2006  EP2_536 - Fixed rule 6505
'LH     18/12/2006  EP2_536 - Added rule 6545, fixed rule 6560
'LH     19/12/2006  EP2_578 - Fixed rule 6765
'LH     19/12/2006  EP2_590 - Fixed rule 3590
'LH     19/12/2006  EP2_594 - Fixed rules 3570, 3575, 6525 & 3730
'LH     19/12/2006  EP2_597 - Fixed rule 3590 - additional XML checks required
'LH     28/12/2006  EP2_646 - Modified rules 3595, 3920, 6745 - additional XML checks required
'LH     15/01/2007  EP2_766 - Fixed GetMortgageSubQuoteNumber() and fixed DateDiffInYearsRoundedUp()
'LH     15/01/2007  EP2_830 - Fixed rule 3535
'LH     15/01/2007  EP2_856 - Fixed rules where references PORTEDLOAN - should be MANUALPORTEDLOANIND.
'LH     15/01/2007  EP2_781 - Changed rules 3615, 3620, 6800, 6805 - now correctly returns <= 12 months employment history.
'                                                                    i.e. Employment started 12 months 1 day ago is no longer included.
'LH     15/01/2007  EP2_814 - Changed rule 6765 in accordance with spec change.
'LH     16/01/2007  EP2_874 - Fixed rule 6510.
'LH     16/01/2007  EP2_877 - Fixed rules 6500, 6510 & 6515.
'IK     18/01/2007  EP2_876 - fix DIP product cascade calls (no accepted quote at DIP), remove Public qualifiers
'LH     16/01/2007  EP2_879 - Fixed rules 6530, 6535, 3515, 3516, 3645 to correctly count number of applicants
'LH     22/01/2007  EP2_924 - Fixed rules 3590, 3595
'LH     22/01/2007  EP2_879 - Fixed rule 6530
'LH     22/01/2007  EP2_877 - Fixed rule 6515
'LH     23/01/2007  EP2_879 - Fixed rule 6530
'LH     23/01/2007  EP2_881 - Fixed rule 6530
'LH     24/01/2007  EP2_775 - Fixed DateDiffInMonthsRoundedDown(), DateDiffInMonthsRoundedUp()
'                             Changed name of  CalculateAgeInFull() to ActualDateDiff()
'LH     24/01/2007  EP2_965 - Fixed rules 6605, 6580, 6615
'LH     25/01/2007  EP2_963 - Fixed rule 6580, 6615
'LH     25/01/2007  EP2_961 - Fixed rule 6555
'LH     26/01/2007  EP2_1009 - Fixed GetDateTodayFromIngestion()
'LH     29/01/2007  EP2_1035 - Fixed rule 6770, also fixed APP_RULE number lookup
'LH     29/01/2007  EP2_1061 - Changed rules 6530, 6535 according to spec change
'LH     30/01/2007  EP2_1045 - Changed rule 6615 according to spec change
'LH     30/01/2007  EP2_1111 - Fixed rule 6560
'LH     31/01/2007  EP2_1001 - Fixed rule 6795, DateDiffInMonthsRoundedDown(), DateDiffInMonthsRoundedUp()
'LH     31/01/2007  EP2_1137 - Fixed rule 6810
'LH     01/02/2007  EP2_1153 - Fixed rule 3935, 6755
'LH     01/02/2007  EP2_766 - Fixed rule 3520
'LH     01/02/2007  EP2_985 - Fixed rule 6785
'LH     01/02/2007  EP2_1132 - Fixed rule 6805
'LH     01/02/2007  EP2_1176 - Fixed rule 3695
'LH     02/02/2007  EP2_1137 - Fixed rule 6810
'LH     02/02/2007  EP2_1191 - Fixed rule 3710
'LH     02/02/2007  EP2_820 - Fixed method TrueDateDiff() for weeks calc - should round up
'LH     05/02/2007  EP2_1220 - Fixed rules 3810, 3815, 3820, 3825, 6865, 6870, 6875, 6880
'LH     05/02/2007  EP2_1229 - Fixed rules 3515, 3516, 3520, 3555, 3790, 3795, 3800, 3805, 6510, 6520, 6845, 6850, 6855, 6860
'LH     07/02/2007  EP2_1133 - Changed rule number 6675 to 6700
'LH     08/02/2007  EP2_1282 - Changed rule number 6815 according to spec change
'LH     10/02/2007  EP2_1311 - Changed rule number 3500 according to spec change
'LH     14/02/2007  EP2_1371 - Changed rules according to spec change
'LH     16/02/2007  EP2_1463 - Fixed rule 3595
'LH     20/02/2007  EP2_1458 - Fixed rule 3516
'OS     26/02/2007  EP2_1634 - Fixed rules 3590, 3595, 6530 and 6535
'OS     27/02/2007  EP2_1685 - Fixed rules 3650 and 3655
'OS     02/03/2007  EP2_1706 - Changed rule 6545 according to updated specs
'LDM    05/03/2007  EP2_1648 - If App is Further Advance, Credit Limit Increase, Product Switch, Transfer of Equity must have a validation type other than "other" to run rule.
'LDM    05/03/2007  EP2_1650 - changed rule 3595 to deal with missing employment records
'LDM    07/03/2007  EP2_1648 - Use the stagesequence no from the stage table to know how far the case has progressed
'LDM    08/03/2007  EP2_1734 - Deal with if the DOB is missing for a applicant/guarantor
'LDM    13/03/2007  EP2_1706 - Fixed rule 6545
'JD     21/03/2007  EP2_1282 - Fixed rule 6815 - correct search string
'LDM    21/03/2007  EP2_1852 - Fixed rule 6530, 3590
'JD     29/03/2007  EP2_2129 Don't run rules 3925,3935 if the app has no quote
'IK     05/04/2007  EP2_2268  - rule 3655 failure is refer (not decline)
'AShaw  06/04/2007  EP2_2228 - Rules 3650 3655. Add New High and Low Global Param checks.
'AW     10/04/2007  EP2_2328   Amended Rule 3590 to cater for emloyed contract
'OS     16/04/2007  EP2_2427   Removed rules 3515, 3516 and 3520, amended 3570, 3640, 3670, 3675 and created 3940
'AShaw  20/04/2007  EP2_2506   Rule3940 altered to fail on Age error.
'AShaw  20/04/2007  EP2_2451   EP2_2328 - Fixed Rule3590 but not Rule3595.
'                              Add same fix for Rule3595.DIP assessment
'                              Also - Rule6535 failing for similar reasons. App Assessment.
'-----------------------------------------------------------------------------------

Option Explicit

Private gstrRuleNumber As String
Private gxmlRequest As MSXML2.DOMDocument40
Private gbErrFlag As Boolean
Private gintErrNumber As Integer
Private gstrErrDescription As String

Private gintTermAsMonths As Integer
Private gintUnderWriterCounter As Integer
Private gintProcessorCounter As Integer
Private gintFraudCounter As Integer

Private gblnSeniorUnderWriter As Boolean

Private gblnReferralOveride As Boolean

Private Enum RiskAssessmentReferralType
    Processor = 20
    Underwriter = 40
End Enum

'If a rule passes then return Level = 1.
'If a "Refer" rule other than #3925 fails then return Level = 3.
'If rule #3925 fails then return Level = 2.
'If a "Decline" rule fails then return Level = 4.
Private Enum RiskAssessmentDecision
   Accept = 1
   AcceptCascade = 2
   Referral = 3
   DeclinePolicy = 4
End Enum

Private Enum RuleSet
    DIPSet = 1
    AppSet = 2
    ValnSet = 3
End Enum

Private Enum GlobalParameterType
    Amount = 1
    Percentage = 2
    BooleanType = 3
    StringType = 4
End Enum

Private Enum QuoteType
   Accepted = 1
   Active = 2
   All = 3
End Enum

Private Enum enmLogType
   LogError = 1&
   LogWarning = 2&
   LogInfo = 4&
End Enum

Private Enum enmErrLevel
   lInfo = &H60000000
   lWarning = &HA0000000
   lError = &HE0000000
End Enum

Private Declare Function RegisterEventSource _
   Lib "advapi32" Alias "RegisterEventSourceA" _
   (ByVal lpUNCServerName As String, _
    ByVal lpSourceName As String) As Long

Private Declare Function DeregisterEventSource _
   Lib "advapi32" _
   (ByVal hEventLog As Long) As Long

Private Declare Function ReportEvent _
   Lib "advapi32" Alias "ReportEventA" _
   (ByVal hEventLog As Long, _
    ByVal wType As Long, _
    ByVal wCategory As Long, _
    ByVal dwEventID As Long, _
    ByVal lpUserSid As Long, _
    ByVal wNumStrings As Long, _
    ByVal dwDataSize As Long, _
    lpStrings As Any, _
    lpRawData As Any) As Long

Public Function RiskAssessmentRules(ByVal vstrgxmlRequest As String) As String

    On Error GoTo OmRequestExit
    
    Dim xmlResponse As MSXML2.DOMDocument40, _
        xmlNode As IXMLDOMNode, _
        xmlResponseNode As IXMLDOMNode, _
        xmlApplicationNode As IXMLDOMNode

    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String
        
    Dim bParseOK As Boolean
    
    gintUnderWriterCounter = 0
    gintProcessorCounter = 0
    gblnReferralOveride = False
    
    Set xmlResponse = New MSXML2.DOMDocument40
    xmlResponse.async = False
    
    Set xmlNode = xmlResponse.createElement("RESPONSE")
    Set xmlResponseNode = xmlResponse.appendChild(xmlNode)
    
    Set gxmlRequest = New MSXML2.DOMDocument40
    gxmlRequest.async = False
    
    bParseOK = gxmlRequest.loadXML(vstrgxmlRequest)
    
    Call Main
    
    LogDetails 1, "========================================================================="
    LogDetails 1, "START: RiskAssessmentRules()", gxmlRequest, "omRARequest"
    
    If Not bParseOK Then
    
        RaiseError xmlParseErr(gxmlRequest)
            
    Else
        
        Set xmlApplicationNode = gxmlRequest.selectSingleNode("REQUEST/APPLICATION")
    
        If xmlApplicationNode Is Nothing Then
            RaiseError "missing REQUEST/APPLICATION node: " & vstrgxmlRequest
        End If
        
        RunRules xmlApplicationNode, xmlResponseNode
        
        Set xmlApplicationNode = Nothing
        
    End If
    
OmRequestExit:
    If Err.number <> 0 Or gbErrFlag Then
        RiskAssessmentRules = CreateErrorResponse
        LogDetails 3, "Err.Number=" & Err.number & " Desc=" & CreateErrorResponse(xmlResponse), xmlResponse, "omRAResponse"
        LogErrorToEventViewer "Err.Number=" & Err.number & " Desc=" & CreateErrorResponse(xmlResponse), vbLogEventTypeError
    Else
        RiskAssessmentRules = xmlResponse.xml
        'xmlResponse.Save "C:\RARules\RAResponse" & Format(Now(), "ddmmyyyyhhmmss") & ".xml"
        LogDetails 1, "END: RiskAssessmentRules()", xmlResponse, "omRAResponse"
    End If
    
    'Close the file
    LogDetails 10, "", xmlResponse
    
    Set xmlNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlApplicationNode = Nothing
    Set gxmlRequest = Nothing
    Set xmlResponse = Nothing

End Function
Private Sub RunRules( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode)

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlList As IXMLDOMNodeList
    
    Dim intFailLevel As Integer, _
        intFailCount As Integer, _
        intStageSeqNo As Integer, _
        intDIPCaseStageMax As Integer
    
    Dim blVex As Boolean 'has omRARules been called from omVex?
        
    Dim strContext As String
    
    
    'log details to log file
    LogDetails 1, Chr(9) & "START: RunRules()"
    
    'Determine whether this component is being called from omVex
    'Do not run full case assessment rules if called from omVex
    Set xmlList = gxmlRequest.selectNodes("REQUEST[@CONTEXT]")
    
    If xmlList.length > 0 Then
       If Not xmlList.Item(0).Attributes.getNamedItem("CONTEXT") Is Nothing Then
            strContext = xmlList.Item(0).Attributes.getNamedItem("CONTEXT").nodeValue
       End If
    End If
    
    If UCase(strContext) = "VEX" Then
        blVex = True
    End If
    
    intFailLevel = RiskAssessmentDecision.Accept
    
    'LDM 07/03/2007 EP2_1648
    intStageSeqNo = GetCaseStageSeqNo
    
    intDIPCaseStageMax = CSng(GetGlobalParameter(xmlApplicationNode, "DIPCaseStageMax", Amount))
    
    If intStageSeqNo > 0 And intStageSeqNo <= intDIPCaseStageMax Then
        'Run case assessment decision in process rules
        intFailLevel = RunRulesDIP(xmlApplicationNode, xmlResponseNode)
    ElseIf intStageSeqNo > intDIPCaseStageMax Or intStageSeqNo = -99 Then
        'Run case assessment post-app rules
        intFailLevel = RunRulesApp(xmlApplicationNode, xmlResponseNode, blVex)
    End If
    
    Set xmlNode = xmlResponseNode.ownerDocument.createAttribute("TYPE")
    xmlNode.Text = "SUCCESS"
    xmlResponseNode.Attributes.setNamedItem xmlNode
    
    Set xmlNode = xmlResponseNode.ownerDocument.createAttribute("FAILLEVEL")
    xmlNode.Text = intFailLevel
    xmlResponseNode.Attributes.setNamedItem xmlNode
    
    Set xmlNode = xmlResponseNode.ownerDocument.createAttribute("UNDERWRITERCOUNT")
    xmlNode.Text = gintUnderWriterCounter
    xmlResponseNode.Attributes.setNamedItem xmlNode
    
    Set xmlNode = xmlResponseNode.ownerDocument.createAttribute("PROCESSORCOUNT")
    xmlNode.Text = gintProcessorCounter
    xmlResponseNode.Attributes.setNamedItem xmlNode
    
    Set xmlNode = xmlResponseNode.ownerDocument.createAttribute("FRAUDCOUNT")
    xmlNode.Text = gintFraudCounter
    xmlResponseNode.Attributes.setNamedItem xmlNode
    
    'if this component is being called from omVex, add a VALUATIONSATISFIED attribute to the RESPONSE
    If blVex Then
        Set xmlNode = xmlResponseNode.ownerDocument.createAttribute("VALUATIONSATISFIED")
        If intFailLevel > RiskAssessmentDecision.Accept Then 'if at least 1 rule has failed
            xmlNode.Text = False
        Else
            xmlNode.Text = True
        End If
        xmlResponseNode.Attributes.setNamedItem xmlNode
    End If
    
    Set xmlNode = Nothing
    
    'log details to log file
    LogDetails 1, Chr(9) & "END: RunRules()"
    
End Sub

Private Function RunRulesDIP( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByRef xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intFailLevel As Integer, _
        intResult As Integer
        
    Dim bRunRule3005and3340 As Boolean, _
        bRunRule3500 As Boolean, _
        bRunRule3505 As Boolean, _
        bRunRule3510 As Boolean, _
        bRunRule3640 As Boolean, _
        bRunRule3695 As Boolean
        
    Dim bCreditCheckAvailable As Boolean
    Dim strApplicationNumber As String
    Dim intMortgageSubquoteNumber As Integer
    
    bRunRule3005and3340 = True
    bRunRule3500 = True
    bRunRule3505 = True
    bRunRule3510 = True
    bRunRule3640 = True
    bRunRule3695 = True
    bCreditCheckAvailable = True
     
    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "START: RunRulesDIP()"
    
    If xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK") Is Nothing Then
        bCreditCheckAvailable = False
    End If
    
    Set xmlNode = xmlApplicationNode.selectSingleNode("/REQUEST/APPLICATION")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("APPLICATIONNUMBER") Is Nothing Then
            strApplicationNumber = xmlNode.Attributes.getNamedItem("APPLICATIONNUMBER").nodeValue
            LogDetails 1, String(2, Chr(9)) & "Application number: " & strApplicationNumber
        End If
    End If
    
    intFailLevel = RiskAssessmentDecision.Accept
    
    'Experian: Experian failure or time-out
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3500) Then
        intResult = Rule3500(xmlApplicationNode, xmlResponseNode)
        'if rule 3500 was unsuccessful
        If intResult <> RiskAssessmentDecision.Accept Then
            bRunRule3500 = False
        End If
        intFailLevel = EpsomMax(intFailLevel, intResult)
    End If
    
    'Experian: Experian cannot assign product scheme
    If bCreditCheckAvailable Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3505) Then
            intResult = Rule3505(xmlApplicationNode, xmlResponseNode)
            'if rule 3505 was unsuccessful
            If intResult <> RiskAssessmentDecision.Accept Then
                'bRunRule3005and3340 = False
                 bRunRule3505 = False
            End If
            intFailLevel = EpsomMax(intFailLevel, intResult)
        End If
    Else
        bRunRule3005and3340 = False
    End If
    
    'Experian: Invalid product scheme returned from Experian
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3510) Then
        intResult = Rule3510(xmlApplicationNode, xmlResponseNode)
        'if rule 3510 was unsuccessful
        If intResult <> RiskAssessmentDecision.Accept Then
            'bRunRule3005and3340 = False
            bRunRule3510 = False
        End If
        intFailLevel = EpsomMax(intFailLevel, intResult)
    End If
    
    '====OS 16/04/2007 (EP2_2427)====
    'OS 24/11/2006
    'Age at end of term beyond residential lower limit
    'If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3515) Then
    '    intFailLevel = EpsomMax(intFailLevel, Rule3515(xmlApplicationNode, xmlResponseNode))
    'End If
    
    'OS 24/11/2006
    'Age at end of term beyond residential lower limit
    'If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3516) Then
    '   intFailLevel = EpsomMax(intFailLevel, Rule3516(xmlApplicationNode, xmlResponseNode))
    'End If
    
    'Age at end of term beyond residential maximum
    'If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3520) Then
    '    intFailLevel = EpsomMax(intFailLevel, Rule3520(xmlApplicationNode, xmlResponseNode))
    'End If
    '====End EP2_2427=====
    
    'Age now less than minimum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3525) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3525(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Age now less than minimum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3535) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3535(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 31/10/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3550) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3550(xmlApplicationNode, xmlResponseNode))
    End If
    
    'RULE 3555: Term out of range
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3555) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3555(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 31/10/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3565) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3565(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 31/10/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3570) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3570(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 31/10/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3575) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3575(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 01/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3580) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3580(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 01/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3585) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3585(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 08/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3590) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3590(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 10/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3595) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3595(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 02/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3600) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3600(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Rule 3605: Non-permanent employment
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3605) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3605(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 03/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3610) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3610(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 03/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3615) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3615(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 06/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3620) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3620(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 07/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3625) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3625(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 07/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3630) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3630(xmlApplicationNode, xmlResponseNode))
    End If
    
    'MCh 07/11/2006
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3640) Then
        intResult = Rule3640(xmlApplicationNode, xmlResponseNode)
        'If Rule 3640 was unsuccessful
        If intResult <> RiskAssessmentDecision.Accept Then
            bRunRule3640 = False
        End If
        intFailLevel = EpsomMax(intFailLevel, intResult)
    End If
    
    'OS 27/11/2006
    'One employed applicant cannot afford repayment
    'IF Rule 3640 = [FAIL] then do not run the rule
    If (bRunRule3640) Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3645) Then
           intFailLevel = EpsomMax(intFailLevel, Rule3645(xmlApplicationNode, xmlResponseNode))
        End If
    End If
    
    'OS 28/11/2006
    'Total card balance not equal CAIS within tolerance
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3650) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3650(xmlApplicationNode, xmlResponseNode))
    End If
    
    
    'DS 28/11/2006
    'Total number of existing loans held and total monthly repayment does not match CAIS data.
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3655) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3655(xmlApplicationNode, xmlResponseNode))
    End If
    
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3660) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3660(xmlApplicationNode, xmlResponseNode))
    End If
    
    'DS 24/11/2006
    'Loan greater than LTV maximum for scheme
    'IF Rule 3500 = [FAIL] OR Rule 3505 = [FAIL] OR RULE 3510 = [FAIL] then do not run this rule.
    If (bRunRule3500 And bRunRule3505 And bRunRule3510) Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3665) Then
            intFailLevel = EpsomMax(intFailLevel, Rule3665(xmlApplicationNode, xmlResponseNode))
        End If
        
    End If
    
    'Property value / PP less than minimum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3670) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3670(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Property value / PP less than upper limit
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3675) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3675(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Loan less than minimu
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3690) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3690(xmlApplicationNode, xmlResponseNode))
    End If
    
    'DS 25/11/2006
    'LTV > maximum allowed for product scheme.
    'IF the Quotation exists(This check is done inside the Rule code)
    'OR Rule 3500 = [FAIL]OR Rule 3505 = [FAIL] OR RULE 3510 = [FAIL] then do not run this rule
    If (bRunRule3500 And bRunRule3505 And bRunRule3510) Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3695) Then
            intResult = Rule3695(xmlApplicationNode, xmlResponseNode)
            'if rule 3695 was unsuccessful
            If intResult <> RiskAssessmentDecision.Accept Then
                 bRunRule3695 = False
            End If
            intFailLevel = EpsomMax(intFailLevel, intResult)
        End If
    End If
        
    'LTV maximum exceeded for BTL remortgage
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3700) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3700(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Remortgage where purpose is "Other"
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3705) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3705(xmlApplicationNode, xmlResponseNode))
    End If
    
    'BTL remortgage where purpose not property-related
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3710) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3710(xmlApplicationNode, xmlResponseNode))
    End If

    'Invalid repayment vehicle for IO residential
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3720) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3720(xmlApplicationNode, xmlResponseNode))
    End If
        
    'Invalid repayment vehicle for IO BTL
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3725) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3725(xmlApplicationNode, xmlResponseNode))
    End If
    
    'OS - 29/11/2006
    '"Additional Information" present
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3730) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3730(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Gifted deposit or unknown source
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3735) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3735(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Builder gifted deposit greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3740) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3740(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Deposit includes Builders Incentive and max LTV
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3745) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3745(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Guarantor present on case
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3750) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3750(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Product switch with right-to-buy
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3755) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3755(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Total number of components greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3760) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3760(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Total number of components greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3765) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3765(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of C&I components greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3770) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3770(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of IO components greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3775) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3775(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of C&I components greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3780) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3780(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of IO components greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3785) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3785(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of C&I terms greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3790) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3790(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of IO terms greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3795) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3795(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of C&I terms greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3800) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3800(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of IO terms greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3805) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3805(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of C&I products greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3810) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3810(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of IO products greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3815) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3815(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of C&I products greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3820) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3820(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Number of IO products greater than maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3825) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3825(xmlApplicationNode, xmlResponseNode))
    End If
    
    'FA/CLI with outstanding retention
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3830) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3830(xmlApplicationNode, xmlResponseNode))
    End If
    
    'Experian: AuthenticatePlus "High Risk" code
    'IF rule 3500 = [FAIL] then do not run this rule.
    If bRunRule3500 Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3840) Then
            intFailLevel = EpsomMax(intFailLevel, Rule3840(xmlApplicationNode, xmlResponseNode))
        End If
    End If
    
    'Experian: Applicant authentication documentation required
    'IF rule 3500 = [FAIL] then do not run this rule.
    
    If bRunRule3500 Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3870) Then
            intFailLevel = EpsomMax(intFailLevel, Rule3870(xmlApplicationNode, xmlResponseNode))
        End If
    End If

    'Unacceptable strategy manager code
    'IF rule 3500 = [FAIL] then do not run this rule.
    If bRunRule3500 Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3875) Then
            intFailLevel = EpsomMax(intFailLevel, Rule3875(xmlApplicationNode, xmlResponseNode))
        End If
    End If

    'Affordability Index below minimum
    'IF rule 3500 = [FAIL] then do not run this rule.
    If bRunRule3500 Then
        If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3920) Then
            intFailLevel = EpsomMax(intFailLevel, Rule3920(xmlApplicationNode, xmlResponseNode))
        End If
    End If
    
    'Total Allowable Income < Total Approximate Monthly Cost
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3930) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3930(xmlApplicationNode, xmlResponseNode))
    End If
    
    'EP2_876
    'product cascade
    'EP2_2129 Don't run these rules if the app has no quote
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    'Locate any LoanComponent record where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
    Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
    If Not xmlNode Is Nothing Then
        If IsWebIngestedCase(xmlApplicationNode) Then
            If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3925) Then
                If (bRunRule3505 And bRunRule3510 And bRunRule3695) Then
                    intFailLevel = EpsomMax(intFailLevel, Rule3925(xmlApplicationNode, xmlResponseNode))
                End If
            End If
        ElseIf IsPaperCase(xmlApplicationNode) Then
            If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3935) Then
                If (bRunRule3505 And bRunRule3510 And bRunRule3695) Then
                    intFailLevel = EpsomMax(intFailLevel, Rule3935(xmlApplicationNode, xmlResponseNode))
                End If
            End If
        End If
    End If
    'EP2_876_ends
    
    'OS 16/04/2006 - EP2_2427
    'Age at end of term > residential / BTL (Income) maximum
    If RunRule(RuleSet.DIPSet, xmlApplicationNode, 3940) Then
        intFailLevel = EpsomMax(intFailLevel, Rule3940(xmlApplicationNode, xmlResponseNode))
    End If
    
    Set xmlNode = Nothing
    
    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "END: RunRulesDIP() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intFailLevel
    
    RunRulesDIP = intFailLevel
    
End Function

Private Function RunRulesApp( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByRef xmlResponseNode As IXMLDOMNode, _
    ByVal blVex As Boolean) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    
    Dim intFailLevel As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intResult As Integer
        
    Dim strApplicationNumber As String
    
    Dim objValuationRulesBO As ValuationRulesBO
    
    Dim bValuationReportExists As Boolean
    
    Dim bRunRule6645 As Boolean, _
        bRunRule6650 As Boolean, _
        bRunRule6660 As Boolean
    
    bValuationReportExists = True
    bRunRule6645 = True
    bRunRule6650 = True
    bRunRule6660 = True
    
    intFailLevel = RiskAssessmentDecision.Accept
    
    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "START: RunRulesApp()"
    
    'only run these rules if this component is not called from omVex
    If Not blVex Then
            
        Set xmlNode = xmlApplicationNode.selectSingleNode("/REQUEST/APPLICATION")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("APPLICATIONNUMBER") Is Nothing Then
                strApplicationNumber = xmlNode.Attributes.getNamedItem("APPLICATIONNUMBER").nodeValue
                LogDetails 1, String(2, Chr(9)) & "Application number: " & strApplicationNumber
            End If
        End If
        
        'Age beyond maximum limit
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6500) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6500(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Age now less than minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6505) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6505(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Age beyond maximum limit
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6510) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6510(xmlApplicationNode, xmlResponseNode))
        End If
    
        'Age beyond BTL maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6515) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6515(xmlApplicationNode, xmlResponseNode))
        End If
    
        'Term out of range
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6520) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6520(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Supplied address history less than minimum years.
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6525) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6525(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 02/12/2006
        'Single app; income below minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6530) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6530(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 04/12/2006
        'Joint app; income below minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6535) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6535(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Confirmed income does not match declared
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6545) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6545(xmlApplicationNode, xmlResponseNode))
        End If
        
        'All applicants unemployed
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6550) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6550(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Contract less than minimum or not renewable
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6555) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6555(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Self-employed time trading less than minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6560) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6560(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Accountant's qualification
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6565) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6565(xmlApplicationNode, xmlResponseNode))
        End If
        
        'With current lender less than 6 months
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6570) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6570(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Existing borrowings from DB
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6575) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6575(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Property value / PP less than minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6580) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6580(xmlApplicationNode, xmlResponseNode))
        End If
        
        'LTV maximum exceeded for BTL remortgage
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6590) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6590(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Loan greater than LTV maximum for scheme
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6595) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6595(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Loan Amount greater than auto accept limit
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6600) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6600(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Under-valued or gift
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6605) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6605(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Deposit includes Builders Incentive and max LTV
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6610) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6610(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Builder gifted deposit greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6615) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6615(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Gifted deposit
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6620) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6620(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Remortgage where purpose is "Other"
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6625) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6625(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Total number of components greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6635) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6635(xmlApplicationNode, xmlResponseNode))
        End If
    
        'RULE 6640: FA/CLI with outstanding retention
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6640) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6640(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Experian unable to allocate product scheme
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6645) Then
            intResult = Rule6645(xmlApplicationNode, xmlResponseNode)
            'if rule 6645 was unsuccessful
            If intResult <> RiskAssessmentDecision.Accept Then
                bRunRule6645 = False
            End If
            intFailLevel = EpsomMax(intFailLevel, intResult)
        End If
        
        'Invalid product scheme returned from Experian
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6650) Then
            intResult = Rule6650(xmlApplicationNode, xmlResponseNode)
            'if rule 6650 was unsuccessful
            If intResult <> RiskAssessmentDecision.Accept Then
                bRunRule6650 = False
            End If
            intFailLevel = EpsomMax(intFailLevel, intResult)
        End If
        
        'Experian failure or time-out
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6660) Then
            intResult = Rule6660(xmlApplicationNode, xmlResponseNode)
            'if rule 6660 was unsuccessful
            If intResult <> RiskAssessmentDecision.Accept Then
                bRunRule6660 = False
            End If
            intFailLevel = EpsomMax(intFailLevel, intResult)
        End If
        
        'Authenticate "High Risk" code
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6665) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6665(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Applicant authentication documentation required
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6695) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6695(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Unacceptable Strategy Manager Code
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6700) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6700(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Affordability Index below minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6745) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6745(xmlApplicationNode, xmlResponseNode))
        End If
        
        'product scheme cascade
        'IF Rule 6645 = [FAIL] OR Rule 6650 = [FAIL] OR RULE 6660 = [FAIL] then do not run the rule.
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6755) Then
            If (bRunRule6645 And bRunRule6650 And bRunRule6660) Then
                intFailLevel = EpsomMax(intFailLevel, Rule6755(xmlApplicationNode, xmlResponseNode))
            End If
        End If
    
        'The property being mortgaged is self-built.
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6760) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6760(xmlApplicationNode, xmlResponseNode))
        End If
        
        'For BTL-R % rental income cover < min LTV allowed
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6765) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6765(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Unencumbered property
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6770) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6770(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Total Allowable Income < Total Approximate Monthly Cost
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6775) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6775(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 01/12/2006
        'Grey-listed broker
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6780) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6780(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 30/11/2006
        'Employment history < minimum period
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6785) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6785(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 30/11/2006
        'Non-permanent employment
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6790) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6790(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 30/11/2006
        'Time in current employment less than minimum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6795) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6795(xmlApplicationNode, xmlResponseNode))
        End If
        
        'OS - 30/11/2006
        'Number or employers greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6800) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6800(xmlApplicationNode, xmlResponseNode))
        End If
        
        'DS 30/11/2006
        'Total number of existing loans held and total monthly repayment does not match CAIS data.
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6805) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6805(xmlApplicationNode, xmlResponseNode))
        End If
        
        'DS 29/11/2006
        'Total number of existing loans held and total monthly repayment does not match CAIS data.
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6810) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6810(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Free Legals Product Selected
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6815) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6815(xmlApplicationNode, xmlResponseNode))
        End If
    
        
        'Total number of components greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6820) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6820(xmlApplicationNode, xmlResponseNode))
        End If
         
        'Number of C&I components greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6825) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6825(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of IO components greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6830) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6830(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of C&I components greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6835) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6835(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of IO components greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6840) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6840(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of C&I terms greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6845) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6845(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of IO terms greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6850) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6850(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of C&I terms greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6855) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6855(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of IO terms greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6860) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6860(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of C&I products greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6865) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6865(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of IO products greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6870) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6870(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of C&I products greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6875) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6875(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Number of IO products greater than maximum
        If RunRule(RuleSet.AppSet, xmlApplicationNode, 6880) Then
            intFailLevel = EpsomMax(intFailLevel, Rule6880(xmlApplicationNode, xmlResponseNode))
        End If
    End If
            
    'Call valuation rules
    Set objValuationRulesBO = New ValuationRulesBO
    objValuationRulesBO.RunRulesValuation gxmlRequest, xmlResponseNode, intFailLevel, gintUnderWriterCounter, gintProcessorCounter, gstrRuleNumber, gbErrFlag, gintErrNumber, gstrErrDescription
    
    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "END: RunRulesApp() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intFailLevel
    
    RunRulesApp = intFailLevel
    
End Function
' IK EP1008 - avoid duplicate keys
Private Sub AddResponse(xmlResponseNode As IXMLDOMNode, number, textOne, textTwo, score, ScoreCardInd)
    
    Dim thisElem As IXMLDOMElement
    Set thisElem = xmlResponseNode.ownerDocument.createElement("RULE")
    
    thisElem.setAttribute "NUMBER", number
    thisElem.setAttribute "TEXTONE", textOne
    thisElem.setAttribute "TEXTTWO", textTwo
    thisElem.setAttribute "SCORE", score
    thisElem.setAttribute "SCORECARDIND", ScoreCardInd
    
    ' IK EP1008 - avoid duplicate keys
    If xmlResponseNode.selectSingleNode("RULE[@NUMBER='" & number & "']") Is Nothing Then
        xmlResponseNode.appendChild thisElem
    End If
    ' IK EP1008 - ends
    
    Set thisElem = Nothing

End Sub

Private Function xmlParseErr(XMLdoc)
    Dim sError As String
    sError = "xml Parser error" & vbCrLf
    sError = sError & "reason : " & XMLdoc.parseError.reason & vbCrLf
    sError = sError & "line No. : " & XMLdoc.parseError.Line & vbCrLf
    sError = sError & "character : " & XMLdoc.parseError.linepos & vbCrLf
    sError = sError & "source text : " & XMLdoc.parseError.srcText & vbCrLf
    xmlParseErr = sError
End Function

Private Sub RaiseError(ByVal vstrDescription)

    Err.Raise vbObjectError + 512 + 999, "omRARulesBO", vstrDescription

End Sub

Private Function CreateErrorResponse(Optional xmlResponse As DOMDocument40) As String
    
    Dim xmlResponseDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = True
    xmlResponseDoc.async = True
    xmlResponseDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
    If Err.number > vbObjectError Then
        xmlElem.setAttribute "TYPE", "APPERR"
    Else
        xmlElem.setAttribute "TYPE", "SYSERR"
    End If
    Set xmlNode = xmlResponseDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlResponseDoc.createElement("ERROR")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlResponseDoc.createElement("NUMBER")
    If gintErrNumber = 0 Then
        xmlElem.Text = Err.number
    Else
        xmlElem.Text = gintErrNumber
    End If
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("SOURCE")
    xmlElem.Text = "omRARulesBO"
    
    If Len(gstrRuleNumber) > 0 Then
        xmlElem.Text = xmlElem.Text & ".Rule" & gstrRuleNumber
    End If
    
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("VERSION")
    If Len(App.Comments) > 0 Then
        xmlElem.Text = App.Comments
    Else
        xmlElem.Text = App.Major & "." & App.Major & "." & App.Revision
    End If
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("DESCRIPTION")
    If Len(gstrErrDescription) = 0 Then
        xmlElem.Text = Err.Description
    Else
        xmlElem.Text = gstrErrDescription
    End If
    xmlNode.appendChild xmlElem
    
    If Not xmlResponse Is Nothing Then
        Set xmlResponse = xmlResponseDoc
    End If
    
    CreateErrorResponse = xmlResponseDoc.xml
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDoc = Nothing

End Function

Private Function GetAttributeAsText(xmlNode, strAttribName)

    Dim xmlAttrib As IXMLDOMNode

    Dim strResult As String
    
    Set xmlAttrib = xmlNode.Attributes.getNamedItem(strAttribName)
    
    If Not xmlAttrib Is Nothing Then
        strResult = xmlAttrib.Text
    End If
    
    Set xmlAttrib = Nothing
    
    GetAttributeAsText = strResult

End Function

Private Function IsWebIngestedCase(ByVal xmlApplicationNode As IXMLDOMNode) As Boolean

Dim intWebChannelId As Integer
Dim intChannelID As Integer
Dim xmlNode As IXMLDOMNode

    intWebChannelId = CInt(GetGlobalParameter(xmlApplicationNode, "WebChannelId", Amount))
    Set xmlNode = xmlApplicationNode.selectSingleNode("./APPLICATIONDATA")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("CHANNELID") Is Nothing Then
            intChannelID = CInt(xmlNode.Attributes.getNamedItem("CHANNELID").Text)
        End If
    End If
    
    If intChannelID = intWebChannelId Then
        IsWebIngestedCase = True
    End If
    
End Function

Private Function IsPaperCase(ByVal xmlApplicationNode As IXMLDOMNode) As Boolean

Dim intPaperChannelId As Integer
Dim intChannelID As Integer
Dim xmlNode As IXMLDOMNode

    intPaperChannelId = CInt(GetGlobalParameter(xmlApplicationNode, "PaperChannelId", Amount))
    Set xmlNode = xmlApplicationNode.selectSingleNode("./APPLICATIONDATA")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("CHANNELID") Is Nothing Then
            intChannelID = CInt(xmlNode.Attributes.getNamedItem("CHANNELID").Text)
        End If
    End If
    
    If intChannelID = intPaperChannelId Then
        IsPaperCase = True
    End If
    
End Function

Private Function GetDateTodayFromIngestion(ByVal xmlApplicationNode As MSXML2.IXMLDOMNode) As Date

    Dim attChannelID As MSXML2.IXMLDOMAttribute
    Dim attReceivedDate As MSXML2.IXMLDOMAttribute
    Dim dtmToday As Date
    Dim xmlNode As IXMLDOMNode
    
    dtmToday = Now()
    
    If Not (xmlApplicationNode Is Nothing) Then
        
        'use APPLICATIONFACTFIND/@APPLICATIONRECEIVEDDATE, else use today
        Set attReceivedDate = xmlApplicationNode.Attributes.getNamedItem("APPLICATIONRECEIVEDDATE")
        If Not (attReceivedDate Is Nothing) Then
            dtmToday = CDate(attReceivedDate.Text)
        End If
        
        GetDateTodayFromIngestion = dtmToday
    Else
        RaiseError ("Error in GetDateTodayFromIngestion(): APPLICATION node IS NOTHING")
    End If

End Function
Private Function EpsomMax(CurrentLevel As Integer, NewLevel As Integer) As Integer
    
    'returns highest level of RiskAssessmentDecision
    If NewLevel > CurrentLevel Then
        EpsomMax = NewLevel
    Else
        EpsomMax = CurrentLevel
    End If
        
End Function

Private Function GetMortgageSubQuoteNumber(eQuoteType As QuoteType) As Integer
    
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intAcceptedQuoteNumber As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intQuoteNumber As Integer, _
        intActiveQuoteNumber As Integer
    
    LogDetails 1, String(2, Chr(9)) & "START: GetMortgageSubQuoteNumber(eQuoteType=" & eQuoteType & ") 1=accepted quote, 2=active, 3=either"
  
    If eQuoteType = Accepted Or eQuoteType = All Then
        'Get Accepted Quote Number
        Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION[@ACCEPTEDQUOTENUMBER]")
        If xmlNodeList.length > 0 Then
            intAcceptedQuoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("ACCEPTEDQUOTENUMBER").Text
            intQuoteNumber = intAcceptedQuoteNumber
        End If
    End If
    
    If eQuoteType = Active Or eQuoteType = All Then
        'if quote type is 'all' and accepted quote number was not found then find the active quote number
        If intAcceptedQuoteNumber = 0 Then
            'Get Active Quote Number
            Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION[@ACTIVEQUOTENUMBER]")
            If xmlNodeList.length > 0 Then
                intActiveQuoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("ACTIVEQUOTENUMBER").Text
                intQuoteNumber = intActiveQuoteNumber
            End If
        End If
    End If
    
    'Get Mortgage Subquote Number
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/QUOTATION[@QUOTATIONNUMBER=" & intQuoteNumber & "]")
    If xmlNodeList.length > 0 Then
        intMortgageSubquoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("MORTGAGESUBQUOTENUMBER").Text
    End If
    
    LogDetails 1, String(2, Chr(9)) & "END: GetMortgageSubQuoteNumber() returned: " & intMortgageSubquoteNumber

    GetMortgageSubQuoteNumber = intMortgageSubquoteNumber

End Function

Private Function GetAcceptedQuoteNumberOrActiveQuoteNumber() As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intQuoteNumber As Integer, _
        intMortgageSubquoteNumber As Integer
    
    
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION")
    If xmlNodeList.length > 0 Then
        'Get Accepted Quote Number
        If Not xmlNodeList.Item(0).Attributes.getNamedItem("ACCEPTEDQUOTENUMBER") Is Nothing Then
            intQuoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("ACCEPTEDQUOTENUMBER").nodeValue
        Else
            'Get Active Quote Number
            If Not xmlNodeList.Item(0).Attributes.getNamedItem("ACTIVEQUOTENUMBER") Is Nothing Then
                intQuoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("ACTIVEQUOTENUMBER").nodeValue
            End If
        End If
    End If
    
    GetAcceptedQuoteNumberOrActiveQuoteNumber = intQuoteNumber

End Function

'-----------------------------------------------------------------------
'Function:  GetComboValueIDForValidationType
'Author  :  MChivers - 03/11/2006
'Desc    :  Returns the COMBOVALUE for the COMBOGROUP GROUPNAME Specified (as string)
'------------------------------------------------------------------------
Private Function GetComboValueIDForValidationType(ByVal xmlApplicationNode As MSXML2.IXMLDOMNode, _
                                                  ByVal sGroupName As String, _
                                                  ByVal sValidationType As String) As String
    
    Dim ndComboValue As MSXML2.IXMLDOMNode
  
    Set ndComboValue = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='" & sGroupName & "']/COMBOVALUE[COMBOVALIDATION/@VALIDATIONTYPE='" & sValidationType & "']")

    If Not (ndComboValue Is Nothing) Then
        GetComboValueIDForValidationType = ndComboValue.Attributes.getNamedItem("VALUEID").Text
        Set ndComboValue = Nothing
    End If

End Function

Private Function GetCaseStage() As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList
    Dim intCaseStage As Integer
        
    LogDetails 1, String(2, Chr(9)) & "START: GetCaseStage()"
    
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/CASEACTIVITY/CASESTAGE")
    If xmlNodeList.length > 0 Then
        'Get case activity ID
        If Not xmlNodeList.Item(0).Attributes.getNamedItem("STAGEID") Is Nothing Then
            intCaseStage = xmlNodeList.Item(0).Attributes.getNamedItem("STAGEID").nodeValue
        End If
    End If
    
    LogDetails 1, String(2, Chr(9)) & "END: GetCaseStage() returned: " & intCaseStage
    
    GetCaseStage = intCaseStage

End Function
'LDM 7/3/7 EP2_1648
'return the stage sequence no from the stage table associated with the stageid, that the case is currently at.
'return  -99 if STAGESEQUENCENO is null, -1 if cant find the stage element
Private Function GetCaseStageSeqNo() As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList
    Dim intCaseStageSeqNo As Integer
        
    LogDetails 1, String(2, Chr(9)) & "START: GetCaseStageSeqNo()"
    
    intCaseStageSeqNo = -1
    
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/CASEACTIVITY/CASESTAGE/STAGE")
    
    If xmlNodeList.length > 0 Then
        'Get case activity ID
        If Not xmlNodeList.Item(0).Attributes.getNamedItem("STAGESEQUENCENO") Is Nothing Then
            intCaseStageSeqNo = CInt(xmlNodeList.Item(0).Attributes.getNamedItem("STAGESEQUENCENO").nodeValue)
        Else
            intCaseStageSeqNo = -99
        End If
    End If
    
    LogDetails 1, String(2, Chr(9)) & "END: GetCaseStageSeqNo() returned: " & intCaseStageSeqNo
    
    GetCaseStageSeqNo = intCaseStageSeqNo

End Function

Private Function RunRule(ByVal eRuleSet As RuleSet, _
                        ByVal xmlApplicationNode As IXMLDOMNode, _
                        ByVal intRuleNumber As Integer) As Boolean
                            
'checks whether rule is in XML - if not in XML then don't run it
'this is very handy when testing - enables rules to be switched on/off from database
'if a rule is available to run (ie the rule number exists in combo RiskAssessmentOverrideScore)
'then run ValidateRuleMortgageTypes as a further check to see whether the rule should be run

Dim blnSuccess As Boolean
Dim xmlNode As IXMLDOMNode
    
    Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='RiskAssessmentOverrideScore']/COMBOVALUE[@VALUEID='" & intRuleNumber & "']")
    If Not xmlNode Is Nothing Then
        
        blnSuccess = ValidateRuleMortgageTypes(xmlApplicationNode, intRuleNumber, eRuleSet)
    End If
    
    RunRule = blnSuccess
    
    Set xmlNode = Nothing
End Function

Private Function GetRuleValidationType(ByVal xmlApplicationNode As IXMLDOMNode, ByVal intRuleNumber As Integer) As Integer
'Retrieves rule validation type from XML

Dim intRuleValidationType As Integer
Dim xmlNode As IXMLDOMNode

    Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='RiskAssessmentOverrideScore']/COMBOVALUE[@VALUEID='" & intRuleNumber & "']/COMBOVALIDATION")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("VALIDATIONTYPE") Is Nothing Then
            intRuleValidationType = CInt(xmlNode.Attributes.getNamedItem("VALIDATIONTYPE").nodeValue)
        End If
    Else
        RaiseError ("Validation type not found for rulenumber = " & intRuleNumber & " (Combo 'RiskAssessmentOverrideScore')")
    End If

    GetRuleValidationType = intRuleValidationType

End Function

Function GetRuleDescription(ByVal xmlApplicationNode As IXMLDOMNode, ByVal intRuleNumber As Integer) As String
'Retrieves rule description from XML

Dim strRuleDescription As String
Dim xmlNode As IXMLDOMNode

    Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='RiskAssessmentOverrideScore']/COMBOVALUE[@VALUEID='" & intRuleNumber & "']")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("VALUENAME") Is Nothing Then
            strRuleDescription = xmlNode.Attributes.getNamedItem("VALUENAME").nodeValue
        End If
    Else
        RaiseError ("Description not found for rulenumber = " & intRuleNumber & " (Combo 'RiskAssessmentOverrideScore')")
    End If

    GetRuleDescription = strRuleDescription
    
End Function

' -----------------------------------------------------------------------
' Function:     ValidateRuleMortgageTypes
' Author:       Lee Hudson
' Date:         16/11/2006
' Parameters:   xmlApplicationNode - the application XML passed into this component.
'               intRuleNumber - the rule number to check.
'               eRuleSet - the DIP, App or Valuation set that the rule belongs to.
' Returns:      Boolean (True or False).
'               True = run the rule, False = don't run the film.
'
' Purpose:  Decides whether a rule should be run based upon:
'           a) The mortgage type.
'           b) Whether the rule has a validation type that matches
'              the mortgage type
'
'           3 rule sets are stored as combos in the database, these 3 combo's are:
'               RiskAssessmentDIPSet (DIP rules)
'               RiskAssessmentAppSet (Application rules)
'               RiskAssessmentValSet (Valuation rules, called from ValuationRulesBO)
'
'           Each combo will list all the rules in that particular set with:
'               ValueId = Rule Id
'               Value = Very brief description
'               ValidationType(s) = Type of mortgage codes to which this rule applies. (Note: Every rule contains "OTHER".)
'
'               Mortgage codes in the ValidationType combo field will have one of the following values:
'               FACLI = Further Advance or Credit Limit Increase
'               TOEA = Transfer of Equity (Add)
'               TOER = Transfer of equity (Remove)
'               PSW = Product Switch
'               OTHER = None of the above
'
'           Example for RiskAssessmentDIPSet
'               3515, "Age at end of term > 75", "TOEA, OTHER"
'               3520, "Age at end of term > 76", "TOEA, OTHER"
'               3525, "Age at app < min", "TOEA, "OTHER"
'
'           In the case of a mortgage type of TransferOfEquity (TOE),
'           GetTransferOfEquityType() is called to determine whether the
'           mortgage type is Transfer Of Equity (Add) or Transfer of Equity (Remove).
'
'           In the case where a rule number is NOT added to a RiskAssessment Set
'           (ie RiskAssessmentDIPSet, RiskAssessmentAppSet, RiskAssessmentValSet)
'           this function will ALWAYS return TRUE (ie the rule will pass and bypass
'           any further checks within this function).
' -----------------------------------------------------------------------

Private Function ValidateRuleMortgageTypes(ByVal xmlApplicationNode As IXMLDOMNode, _
                                        ByVal intRuleNumber As Integer, _
                                        ByVal eRuleSet As RuleSet) As Boolean

Dim strRuleValidationType As String
Dim blSuccess As Boolean
Dim xmlRuleAppType As IXMLDOMNode
Dim xmlRuleAppTypes As IXMLDOMNodeList
Dim strGroupName As String  'combo groupname of rule set
Dim strTransferOfEquityType As String

    'For all other mortgage type scenarios than:
    'Further Advance, Credit Limit Increase, Product Switch, Transfer of Equity
    'the rule is run; subject to any other limiting conditions (after this function has been called).
     If Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_F") Is Nothing Or _
        Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_CLI") Is Nothing Or _
        Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PSW") Is Nothing Or _
        Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_TOE") Is Nothing Then

        Select Case eRuleSet
            Case DIPSet
                strGroupName = "RiskAssessmentDIPSet"
            Case AppSet
                strGroupName = "RiskAssessmentAppSet"
            Case ValnSet
                strGroupName = "RiskAssessmentValSet"
        End Select
        
        Set xmlRuleAppTypes = xmlApplicationNode.selectNodes("./COMBOGROUP[@GROUPNAME='" & strGroupName & "']/COMBOVALUE[@VALUEID='" & intRuleNumber & "']/COMBOVALIDATION")
        'If a rule has not been added to the rule set, then run the rule; subject to any other limiting conditions.
        If xmlRuleAppTypes.length = 0 Then
            blSuccess = True
        
        'the rule has been added to the rule set
        Else
            For Each xmlRuleAppType In xmlRuleAppTypes
                If Not xmlRuleAppType Is Nothing Then
                    If Not xmlRuleAppType.Attributes.getNamedItem("VALIDATIONTYPE") Is Nothing Then
                        strRuleValidationType = xmlRuleAppType.Attributes.getNamedItem("VALIDATIONTYPE").nodeValue
                        
                        'LDM 05/03/2007 EP2_1648
                        'For type of application (Further Advance, Credit Limit Increase, Product Switch, Transfer of Equity);
                        'the rule will only be run if the validation type has corresponding FACLI, PSW, TOEA, TOER
                        'If validation type is 'OTHER' and there are no other validation types for the rule
                        'then we exit the loop and indicate that the rule should not be
                        'attempted to be run;
                        If strRuleValidationType = "OTHER" And xmlRuleAppTypes.length = 1 Then
                            blSuccess = False
                            Exit For
                        End If
                        
                        ' The mortgage is a Further Advance (aka Additional Borrowing) or
                        ' The mortgage is a Credit Limit Increase
                        If Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_F") Is Nothing Or _
                        Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_CLI") Is Nothing Then
                            If strRuleValidationType = "FACLI" Then
                                blSuccess = True
                                Exit For
                            End If
                        
                        ' The mortgage is a Product Switch.
                        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PSW") Is Nothing Then
                            If strRuleValidationType = "PSW" Then
                                blSuccess = True
                                Exit For
                            End If
                        
                        ' The mortgage is a Transfer of Equity.
                        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_TOE") Is Nothing Then
                            'Call GetTransferOfEquityType to determine the whether
                            'the mortgage is a Transfer of Equity (Add) or
                            'the mortgage is a Transfer of Equity (Remove).
                            strTransferOfEquityType = GetTransferOfEquityType(xmlApplicationNode)
                            If strTransferOfEquityType = "TOEA" And strRuleValidationType = "TOEA" Then
                                blSuccess = True
                                Exit For
                            End If
                            
                            If strTransferOfEquityType = "TOER" And strRuleValidationType = "TOER" Then
                                blSuccess = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next xmlRuleAppType
        End If
    Else
        'Application type is not one of:
        'Further Advance, Credit Limit Increase, Product Switch, Transfer of Equity
        'the rule is run; subject to any other limiting conditions (after this function has been called).
        blSuccess = True
    End If
    
    ValidateRuleMortgageTypes = blSuccess

End Function

'-----------------------------------------------------------------------
'Function:  GetTransferOfEquityType
'Author  :  L.Hudson - 16/11/2006
'Desc    :  Returns the mortgage equity type
'------------------------------------------------------------------------
Private Function GetTransferOfEquityType(ByVal xmlApplicationNode As IXMLDOMNode) As String

Dim xmlTOE_List As IXMLDOMNodeList 'Transfer of Equity List
Dim strTransferOfEquityType As String

    Set xmlTOE_List = xmlApplicationNode.selectNodes("./REMOVEDTOECUSTOMER[@TYPE='A']")
    'IF at least one record is read with RemovedToECustomer.Type = "A" ("Added") then:
    If xmlTOE_List.length > 0 Then
        'The mortgage is a Transfer of Equity (Add).
        strTransferOfEquityType = "TOEA"
    Else
        'ELSE IF no records are read with RemovedToECustomer.Type = "A"
        'AND at least one record is read with RemovedToECustomer.Type = "R" ("Removed") then:
        Set xmlTOE_List = xmlApplicationNode.selectNodes("./REMOVEDTOECUSTOMER[@TYPE='R']")
        If xmlTOE_List.length > 0 Then
            'The mortgage is a Transfer of Equity (Remove).
            strTransferOfEquityType = "TOER"
        End If
    End If

    GetTransferOfEquityType = strTransferOfEquityType
End Function


Private Sub CountReferral(ByVal intScore As Integer)
'decides the referral type and adds to the appropriate counter

If intScore = RiskAssessmentReferralType.Processor Then
    gintProcessorCounter = gintProcessorCounter + 1
ElseIf intScore = RiskAssessmentReferralType.Underwriter Then
    gintUnderWriterCounter = gintUnderWriterCounter + 1
End If

End Sub

Private Function GetGlobalParameter(ByVal xmlApplicationNode As IXMLDOMNode, _
                                    ByVal sName As String, _
                                    ByVal eGlobalParameterType As GlobalParameterType) As String
   
    Dim xmlNode As IXMLDOMNode
    Dim sParamValue As String
    Dim sGlobalParameterType As String
    
'    Private Enum GlobalParameterType
'        Amount = 1
'        Percentage = 2
'        BooleanType = 3
'        StringType = 4
'    End Enum

    Select Case eGlobalParameterType
        Case Amount
            sGlobalParameterType = "AMOUNT"
            
        Case Percentage
            sGlobalParameterType = "PERCENTAGE"
            
        Case BooleanType
            sGlobalParameterType = "BOOLEAN"
            
        Case StringType
            sGlobalParameterType = "STRING"
    End Select
    
    Set xmlNode = xmlApplicationNode.selectSingleNode("./GLOBALPARAMETER[@NAME='" & sName & "']")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem(sGlobalParameterType) Is Nothing Then
            sParamValue = xmlNode.Attributes.getNamedItem(sGlobalParameterType).nodeValue
        End If
    Else
        RaiseError "Globalparameter '" & sName & "' not found"
    End If
        
    GetGlobalParameter = sParamValue
    
End Function

Function ActualDateDiff(ByVal dtStartDate As Date, ByVal dtEndDate As Date, _
ByRef intYears As Integer, ByRef intMonths As Integer, ByRef intDays As Integer) As Boolean
    
    Dim intYear As Integer
    Dim intMonth  As Integer
    Dim intDay As Integer
    Dim dtTemp As Date
    
    ActualDateDiff = True
    
    dtTemp = DateSerial(Year(dtEndDate), Month(dtStartDate), Day(dtStartDate))
    intYear = Year(dtEndDate) - Year(dtStartDate) + (dtTemp > dtEndDate)
    intMonth = Month(dtEndDate) - Month(dtStartDate) - (12 * (dtTemp > dtEndDate))
    intDay = Day(dtEndDate) - Day(dtStartDate)
    If intDay < 0 Then
        intMonth = intMonth - 1
        intDay = Day(DateSerial(Year(dtEndDate), Month(dtEndDate) + 1, 0)) + intDay + 1
    End If
    intMonths = intMonth
    intYears = intYear
    intDays = intDay
    
End Function

Private Function DateDiffInMonthsRoundedUp(ByVal dtStartDate As Date, ByVal dtEndDate As Date) As Integer
    'Rounds up in months difference between start date and current date.
    'This is to address rounding problems with DateDiff().
    'Datediff only subtracts month numeric from only month numeric,
    'hence in case: DateDiff ("m", "28/03/2006", "25/04/2006") DateDiff
    'would subtract 4-3 which would return 1.
    
    Dim intYears As Integer
    Dim intMonths As Integer
    Dim intDays As Integer
    Dim intResult As Integer
        
    If IsDate(dtStartDate) Then
        
        ActualDateDiff dtStartDate, dtEndDate, intYears, intMonths, intDays
        If intDays > 0 Then
            intMonths = (intYears * 12) + intMonths + 1
        End If
        
        intResult = intMonths
        
    End If
    
    DateDiffInMonthsRoundedUp = intResult
    
    Exit Function

End Function

Private Function DateDiffInMonthsRoundedDown(ByVal dtStartDate As Date, _
                                     ByVal dtEndDate As Date) As Integer
    
    'Rounds down in months difference between start date and end date.
    Dim intYears As Integer
    Dim intMonths As Integer
    Dim intDays As Integer
    
    If IsDate(dtStartDate) And IsDate(dtEndDate) And dtEndDate <> "00:00:00" Then
    
        ActualDateDiff dtStartDate, dtEndDate, intYears, intMonths, intDays
                                
        DateDiffInMonthsRoundedDown = (intYears * 12) + intMonths

    Else
        RaiseError "Error Invalid date passed into function DateDiffInMonthsRoundedDown"
    End If
    
End Function


Private Function LogErrorToEventViewer(sErrMsg As String, _
eEventType As LogEventTypeConstants) As Boolean
    Dim lEventLogHwnd As Long
    Dim LogType As enmLogType
    Dim lEventID As Long
    Dim lCategory As Long
    Dim sServerName As String
    Dim lRet As Long
   
    LogErrorToEventViewer = True
    Const sSourceName = "omRARules"
    lCategory = 0
    sServerName = vbNullString
            
    If eEventType = vbLogEventTypeError Then
        LogType = LogError
        lEventID = 3& Or enmErrLevel.lError
    ElseIf eEventType = vbLogEventTypeInformation Then
        LogType = LogInfo
        lEventID = 1& Or enmErrLevel.lInfo
    ElseIf eEventType = vbLogEventTypeWarning Then
        LogType = LogWarning
        lEventID = 2& Or enmErrLevel.lWarning
    End If
    
    lEventLogHwnd = RegisterEventSource(lpUNCServerName:=sServerName, lpSourceName:=sSourceName)
    
    If lEventLogHwnd = 0 Then
        LogErrorToEventViewer = False
        Exit Function
    End If
    
    lRet = ReportEvent(hEventLog:=lEventLogHwnd, _
                       wType:=LogType, _
                       wCategory:=lCategory, _
                       dwEventID:=lEventID, _
                       lpUserSid:=0, _
                       wNumStrings:=1, _
                       dwDataSize:=0, _
                       lpStrings:=sErrMsg, _
                       lpRawData:=0)
                       
    If lRet = False Then
        LogErrorToEventViewer = False
    End If
                       
    DeregisterEventSource lEventLogHwnd
End Function

Private Function DateDiffInYearsRoundedDown(dtStartDate As Date, dtEndDate As Date) As Integer
    'Returns difference between two dates in years (rounded down)
    'Ie if date diff is 17 yrs 11 months, return 17

    Dim intTotalYears As Integer
    Dim intMonths As Integer
    Dim intDays As Integer
    
    ActualDateDiff dtStartDate, dtEndDate, intTotalYears, intMonths, intDays
                
    DateDiffInYearsRoundedDown = intTotalYears
    
End Function

Private Function DateDiffInYearsRoundedUp(dtStartDate As Date, dtEndDate As Date) As Integer
    'Returns difference between two dates in years (rounded up)
    'DateDiffInYearsRoundedUp(01/01/2005, 01/06/2005) returns 1
    'DateDiffInYearsRoundedUp(01/01/2005, 01/06/2005) returns 1

    Dim intTotalYears As Integer
    Dim intMonths As Integer
    Dim intDays As Integer

    ActualDateDiff dtStartDate, dtEndDate, intTotalYears, intMonths, intDays
    If intMonths > 0 Or intDays > 0 Then
        intTotalYears = intTotalYears + 1
    End If

    DateDiffInYearsRoundedUp = intTotalYears

End Function

'-----------------------------------------------------------------------
'Function:  TrueDateDiff
'Author  :  MChivers - 03/11/2006
'Desc    :  Returns literal difference between two dates
'Examples:
'           YEARS
'           TrueDateDiff("yyyy", 01/11/2005, 01/11/2006) = 1'
'           TrueDateDiff("yyyy", 02/11/2005, 01/11/2006) = 0
'           TrueDateDiff("yyyy", 01/11/2006, 01/11/2005) = -1
'
'           MONTHS
'           TrueDateDiff("m", 02/11/2005, 01/11/2006) = 11
'           TrueDateDiff("m", 01/11/2005, 01/11/2006) = 12
'           TrueDateDiff("m", 01/12/2006, 01/11/2006) = -1
'
'           WEEKS
'           TrueDateDiff("w", 01/11/2006, 07/11/2006) = 1
'           TrueDateDiff("w", 01/11/2006, 06/11/2006) = 0
'           TrueDateDiff("w", 07/11/2006, 01/11/2006) = -1
'------------------------------------------------------------------------
Private Function TrueDateDiff(ByVal sInterval As String, _
                              ByVal dtmStart As Date, _
                              ByVal dtmEnd As Date) As Integer

    Dim iNumOfYears As Integer
    Dim iNumOfMonths As Integer
    Dim iNumOfWeeks As Integer
    Dim iNumOfDays As Integer
    
    Dim iResult As Integer
    
    If dtmStart > 0 And dtmEnd > 0 Then
    
        iNumOfDays = DateDiff("d", dtmStart, dtmEnd)
        iNumOfDays = IIf(iNumOfDays < 0, iNumOfDays + -1, iNumOfDays + 1)
        
        Select Case sInterval
            Case "d"
                
                iResult = iNumOfDays
                
            Case "w"
                
                iResult = Fix(iNumOfDays / 7)
                
                'Round the number of weeks up
                If iNumOfDays - (iResult * 7) > 0 Then
                    iResult = iResult + 1
                End If
                
            Case "m"
                        
                iNumOfMonths = ((Year(dtmEnd) - Year(dtmStart)) * 12)
                iNumOfMonths = iNumOfMonths + (Month(dtmEnd) - Month(dtmStart))
                
                If iNumOfMonths > 0 Then
                    iNumOfMonths = iNumOfMonths + IIf((Day(dtmEnd) - Day(dtmStart)) >= 0, 0, -1)
                End If
                
                iResult = iNumOfMonths
            
            Case "yyyy"
                
                iNumOfYears = Year(dtmEnd) - Year(dtmStart)
                iNumOfYears = iNumOfYears + IIf((Month(dtmEnd) - Month(dtmStart)) >= 0, 0, -1)
                
                If Month(dtmEnd) = Month(dtmStart) Then
                    iNumOfYears = iNumOfYears + IIf((Day(dtmEnd) - Day(dtmStart)) >= 0, 0, -1)
                End If
                
                iResult = iNumOfYears
                            
        End Select
    End If
    
    TrueDateDiff = iResult

End Function


'AW EP923   30/06/2006
Private Function HasExistingBorrowings(ByVal strCompanyName As String) As Boolean

    Dim blnIsDB As Boolean
    
    If (Len(strCompanyName) = 2 And (InStr(1, strCompanyName, "DB") > 0)) Or _
                    InStr(1, strCompanyName, "DB ") > 0 Or _
                    InStrRev(strCompanyName, " DB") > 0 Or _
                    InStr(strCompanyName, " DB ") > 0 Or _
                    InStr(strCompanyName, "DEUTSCHE") > 0 Then
                            
        blnIsDB = True

    End If
    
    HasExistingBorrowings = blnIsDB
    
    Exit Function

End Function

Private Function Rule3500(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim objApplicant As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bFail As Boolean
    Dim objNode As IXMLDOMNode
        
    'Rule 3500: Experian failure or time-out
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3500"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
  
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' IF the ApplicationCreditCheck record is not found
    '    OR the KnowYourCustomerCheck record for any applicant is not found then [FAIL].
    
    If xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK") Is Nothing Then
        bFail = True
    End If
    
    For Each objNode In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
        If objNode.selectSingleNode("./KNOWYOURCUSTOMERCHECK") Is Nothing Then
            bFail = True
        End If
    Next objNode
    
    If bFail Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3500 = intResult
End Function
Private Function Rule3505(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode

    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim strEligInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3505"

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    ' Experian unable to allocate product scheme
    ' IF BespokeDecision.EligibilityIndicator = "NB" OR = "NR" then [FAIL].

    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If

    If strEligInd = "NB" Or strEligInd = "NR" Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Product Scheme: " & strEligInd
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3505 = intResult

End Function

Private Function Rule3510(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim objApplicant As IXMLDOMNode
    
    Dim xmlNode As IXMLDOMNode
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strEligInd As String
        
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bFail As Boolean
    Dim objNode As IXMLDOMNode
        
    'Invalid product scheme returned from Experian
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3510"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule3510() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'Locate the entry in combo "ProductScheme" where ValidationType = BespokeDecision.EligibilityIndicator.
    'IF the combo entry is not found then [FAIL]
    Set xmlNode = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("ELIGIBILITYINDICATOR") Is Nothing Then
            strEligInd = xmlNode.Attributes.getNamedItem("ELIGIBILITYINDICATOR").nodeValue
        End If
        
        
        Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='ProductScheme']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE='" & strEligInd & "']")
        If xmlNode Is Nothing Then
            bFail = True
        End If
    End If
    
    If bFail Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Product Scheme: " & strEligInd
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3510 = intResult
End Function

'========================================================
'OS 23/11/2006
'========================================================

Private Function Rule3515( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlApplicants As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAgeEndTermLow As Integer, _
        intMaxAgeEndTermHigh As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intNumberOfApplicants As Integer, _
        intApplicantsAge As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3515"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3515: Age at end of term beyond residential lower limit
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    Set xmlApplicants = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    intNumberOfApplicants = xmlApplicants.length
    
    
    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType = "BR") OR the Quotation is not found
    'or number of applicants is more than one then do not run the rule
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
        intMortgageSubquoteNumber = 0 Or intNumberOfApplicants > 1 Then
        
        'Do nothing
    Else
        
        'Get lower limit for loan term
        intMaxAgeEndTermLow = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBLo", Amount)
        
         'Get upper limit for loan term
        intMaxAgeEndTermHigh = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBHi", Amount)
        
        'Locate all LoanComponent records where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        'Calculate and temporarily store the biggest term (in years) across all new loan components.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
            End If
        Next xmlNode

        
        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                'Get customer number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                    strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
                
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    'LDM 08/03/2006 EP2_1734
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'IF this is a web-ingested case then calculate age as a system date (in whole years rounded up) using CustomerVersion.DateOfBirth
                'else calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded up) using CustomerVersion.DateOfBirth
                dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
                
                'IF age (at end of term) > global parameter "LPMaxAgeEndTermResLTBLo"
                'AND <= global parameter "LPMaxAgeEndTermResLTBHi" then [FAIL]
                If (intApplicantsAge + intMaxTermInYears) > intMaxAgeEndTermLow _
                And (intApplicantsAge + intMaxTermInYears) <= intMaxAgeEndTermHigh Then
                    blnSuccess = False
                    Exit For
                End If
        
            End If
        Next xmlApplicant
        
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
        End If
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        Else
            strRuleValue = "Age: " & intApplicantsAge + intMaxTermInYears
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3515 = intResult
    
    Set xmlNode = Nothing
        
End Function


'========================================================
'OS 24/11/2006
'========================================================

Private Function Rule3516( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicants As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlIncome As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAgeEndTermLow As Integer, _
        intMaxAgeEndTermHigh As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intApplicantsAge As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intNumberOfApplicants As Integer, _
        intApplicantCount As Integer, _
        intApplicantFailCount As Integer, _
        intApplicantFailAge As Integer
        
    Dim dblAllowableIncome As Double, _
        dblTotalApproximateMonthlyCost As Double
    
    
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess As Boolean, _
        blnOneApplicantFailed, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3516"
    strScoreCardInd = 0 'False
    'Set default value for Allowable income
    dblAllowableIncome = 0
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3516: Age at end of term beyond residential lower limit
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    blnOneApplicantFailed = False
    bDOBExists = True
     
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    Set xmlApplicants = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    intNumberOfApplicants = xmlApplicants.length
    
    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType = "BR") OR the Quotation is not found
    'or number of applicants is more than one then do not run the rule
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
        intMortgageSubquoteNumber = 0 Or intNumberOfApplicants = 1 Then
        
        'Do nothing
        
    Else
        'Get TotalApproximateMonthlyCost
        If Not xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST") Is Nothing Then
            dblTotalApproximateMonthlyCost = CDbl(xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST").Text)
        End If
        
        'Get lower limit for loan term
        intMaxAgeEndTermLow = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBLo", Amount)
        
         'Get upper limit for loan term
        intMaxAgeEndTermHigh = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBHi", Amount)
        
        'Locate all LoanComponent records where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        'Calculate and temporarily store the biggest term (in years) across all new loan components.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
            End If
        Next xmlNode

        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                'Get customer number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                    strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
                
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    'LDM 08/03/2006 EP2_1734
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'IF this is a web-ingested case then calculate age as a system date (in whole years rounded down) using CustomerVersion.DateOfBirth
                'else calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded down) using CustomerVersion.DateOfBirth
                dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
                    
                intApplicantCount = intApplicantCount + 1
                
                'IF age (at end of term) > global parameter "LPMaxAgeEndTermResLTBLo"
                'AND <= global parameter "LPMaxAgeEndTermResLTBHi" then [FAIL]
                If (intApplicantsAge + intMaxTermInYears) > intMaxAgeEndTermLow _
                And (intApplicantsAge + intMaxTermInYears) <= intMaxAgeEndTermHigh Then
                    If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERROLETYPE_TYPE_G") Is Nothing Then
                        blnSuccess = False
                        Exit For
                    ElseIf intApplicantFailCount = 1 Then 'both applicants failed, fail and exit loop
                        intApplicantFailCount = intApplicantFailCount + 1
                        blnSuccess = False
                        Exit For
                    Else
                        'One of the applicants has met the conditions of the age check
                        blnOneApplicantFailed = True
                        intApplicantFailAge = intApplicantsAge + intMaxTermInYears
                        intApplicantFailCount = intApplicantFailCount + 1
                    End If
                End If
                
                'if one applicant failed and we've processed all applicants and/or guarantor
                If blnOneApplicantFailed And intApplicantCount = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']").length Then
                    'An applicant has not met the conditions of the age check
                    'Store Allowable income for this applicant
                    If Not xmlNode Is Nothing _
                    And Not xmlApplicant.Attributes.getNamedItem("CUSTOMERROLETYPE_TYPE_A") Is Nothing Then
                        'Calculate income for applicant
                        dblAllowableIncome = 0
                        
                        'Get income from NETCONFIRMEDALLOWABLEINCOME attribute
                        Set xmlIncome = xmlNode.selectSingleNode("./INCOMESUMMARY[@NETCONFIRMEDALLOWABLEINCOME]")
                        If Not xmlIncome Is Nothing Then
                            dblAllowableIncome = CDbl(xmlIncome.Attributes.getNamedItem("NETCONFIRMEDALLOWABLEINCOME").Text)
                        End If
                        
                        'Get income from NETALLOWABLEANNUALINCOME attribute
                        If dblAllowableIncome = 0 Then
                            Set xmlIncome = xmlNode.selectSingleNode("./INCOMESUMMARY[@NETALLOWABLEANNUALINCOME]")
                            If Not xmlIncome Is Nothing Then
                                dblAllowableIncome = CDbl(xmlIncome.Attributes.getNamedItem("NETALLOWABLEANNUALINCOME").Text)
                            End If
                        End If
                    End If
                End If
        
            End If
        Next xmlApplicant
        
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        ElseIf blnOneApplicantFailed Then
            'If both applicants failed the check, then Allowable Income would be zero
            If dblTotalApproximateMonthlyCost > dblAllowableIncome Then
                blnSuccess = False
                strRuleValue = "Age: " & intApplicantFailAge
            Else
                strRuleValue = "Age: " & intApplicantsAge + intMaxTermInYears
            End If
            
        Else
            strRuleValue = "Age: " & intApplicantsAge + intMaxTermInYears
        End If
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
        End If
        
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3516 = intResult
    
    Set xmlNode = Nothing
        
End Function

'========================================================


Private Function Rule3520( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAgeEndTerm As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intApplicantsAge As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean
        
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3520"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3520: Age at end of term beyond residential maximum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType = "BR") OR the Quotation is not found then do not run the rule
    If xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
        intMortgageSubquoteNumber > 0 Then
        
        
        'Get max age for loan term
        intMaxAgeEndTerm = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBHi", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        'Calculate and temporarily store the biggest term (in years) across all new loan components.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
            End If
        Next xmlNode

        
        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']")
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    'LDM 8/3/7 EP2_1734
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'IF this is a web-ingested case then calculate age as a system date (in whole years rounded down) using CustomerVersion.DateOfBirth
                'else calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded down) using CustomerVersion.DateOfBirth
                dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
                    
                'IF age (at end of term) >= global parameter "LPMaxAgeEndTermResLTBHi" then [FAIL]
                If (intApplicantsAge + intMaxTermInYears) > intMaxAgeEndTerm Then
                    blnSuccess = False
                    Exit For
                End If
        
            End If
        Next xmlApplicant
        
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        Else
            strRuleValue = "Age: " & intApplicantsAge + intMaxTermInYears
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3520 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3525( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAge As Integer, _
        intApplicantsAge As Integer, _
        intMinimumAgeRes As Integer, _
        intMinimumAgeBTL As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3525"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3525: Age now less than minimum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True
        
    'Get Minimum Age Residential
    intMinimumAgeRes = CSng(GetGlobalParameter(xmlApplicationNode, "MinimumAgeRes", Amount))
    
    'Get Minimum Age Buy-To-Let
    intMinimumAgeBTL = CSng(GetGlobalParameter(xmlApplicationNode, "MinimumAgeBTL", Amount))

    'IF this is a web-ingested case then calculate age as at system date (in whole years rounded down) using CustomerVersion.DateOfBirth.
    'ELSE calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded down) using CustomerVersion.DateOfBirth.
    dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
    
    'Read each applicant/guarantor for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If Not xmlApplicant Is Nothing Then
            'get customer version number
            If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
            End If
            
            Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']")
            
            blnSuccess = False
            If Not xmlNode Is Nothing Then
                'get customers date of birth
                'LDM 08/03/2006 EP2_1734
                If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                    If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                        dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                        blnSuccess = True
                    End If
                End If
            End If
            
            If Not blnSuccess Then
                bDOBExists = False
                Exit For
            End If
            
            'Calculate applicant's age at application (rounded down)
            intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
            
            'IF ApplicationFactFind.NatureOfLoan = "Residential" OR = "Let-to-Buy" (combo "NatureOfLoan" ValidationType = "RS" OR = "LT") AND any customer's age < global parameter "MinimumAgeRes" then [FAIL].
            If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Or _
            Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing Then
                If intApplicantsAge < intMinimumAgeRes Then
                    blnSuccess = False
                    Exit For
                End If
            'ELSE IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let" (combo "NatureOfLoan" ValidationType = "BI" OR = "BR") AND any customer's age < global parameter "MinimumAgeBTL" then [FAIL].
            ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
            Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
                If intApplicantsAge < intMinimumAgeBTL Then
                    blnSuccess = False
                    Exit For
                End If
            End If
        End If
    Next xmlApplicant
    

    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.DeclinePolicy
    End If
    
    If Not bDOBExists Then
        strRuleValue = "No DOB"
    Else
        strRuleValue = "Age: " & intApplicantsAge
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3525 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3535( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intApplicantsAge As Integer, _
        intMaxAgeBTL As Integer, _
        intRuleValidationType As Integer, _
        intDays As Integer, _
        intMonths As Integer, _
        intYears As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3535"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3535: Age now less than minimum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True
        
    'IF ApplicationFactFind.NatureOfLoan not = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType not = "BR") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
    
        'Get Minimum Age Residential
        intMaxAgeBTL = CSng(GetGlobalParameter(xmlApplicationNode, "LPMaxAgeBTL", Amount))
                       
        'IF this is a web-ingested case then calculate age as at system date (in whole years rounded up) using CustomerVersion.DateOfBirth.
        'ELSE calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded up) using CustomerVersion.DateOfBirth.
        dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
       
        'Read each applicant/guarantor for the application
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
            If Not xmlApplicant Is Nothing Then
        
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']")
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    'LDM 08/03/2006 EP2_1734
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'Calculate applicant's age at application (rounded up)
                'MCh 23/11/2006 spec updated. Now uses new function
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
                                               
                If intApplicantsAge > intMaxAgeBTL Then
                    blnSuccess = False
                    Exit For
                End If
                
            End If
        Next xmlApplicant
        
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        Else
            strRuleValue = "Age: " & intApplicantsAge
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3535 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3550(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim dtmToday As Date
    Dim dtmNameChange As Date
    Dim intNameChangeYears As Integer
    Dim ndAlias As MSXML2.IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3550"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    intNameChangeYears = CInt(GetGlobalParameter(xmlApplicationNode, "NameChangeYears", Amount))
    
    dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
    Set ndAlias = xmlApplicationNode.selectSingleNode("//ALIAS[position()=1]")
    
    If Not (ndAlias Is Nothing) Then
    
        dtmNameChange = CDate(ndAlias.Attributes.getNamedItem("DATEOFCHANGE").Text)
            
        'Iterate all ALIAS nodes of all CUSTOMERS to get most recent name change date
        For Each ndAlias In xmlApplicationNode.selectNodes("//ALIAS")
            If DateDiff("d", dtmNameChange, CDate(ndAlias.Attributes.getNamedItem("DATEOFCHANGE").Text)) > 0 Then
                dtmNameChange = CDate(ndAlias.Attributes.getNamedItem("DATEOFCHANGE").Text)
            End If
        Next
            
        'If change was in last 3 years
        If TrueDateDiff("yyyy", dtmNameChange, dtmToday) < intNameChangeYears Then
            'rule failed
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
            strRuleValue = "Change Date: " & CStr(dtmNameChange)
        End If
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3550 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndAlias = Nothing
    
End Function

Private Function Rule3555( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMinTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMinimumTerm As Integer, _
        intMaximumTerm As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3555"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3555: Term out of range
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF the Quotation is not found then do not run the rule
    If intMortgageSubquoteNumber > 0 Then
        
        intMinTermInYears = 99
        
        'Get max term
        intMaximumTerm = GetGlobalParameter(xmlApplicationNode, "MaximumTerm", Amount)
        
        'Get min term
        intMinimumTerm = GetGlobalParameter(xmlApplicationNode, "MinimumTerm", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
                'store the minimum term found so far...
                If intTempTermInYears < intMinTermInYears Then
                    intMinTermInYears = intTempTermInYears
                End If
                
            End If
        Next xmlNode
                  
        'When all loan components have been processed, check the smallest and biggest term.
        'IF MinTerm < global parameter "MinimumTerm" then set TempTerm = MinTerm and [FAIL]. ELSE IF MaxTerm > global parameter
        If intMinTermInYears < intMinimumTerm Then
            intTempTermInYears = intMinTermInYears
            blnSuccess = False
        ElseIf intMaxTermInYears > intMaximumTerm Then
            intTempTermInYears = intMaxTermInYears
            blnSuccess = False
        End If
    
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
        
        strRuleValue = "Term (whole years): " & intTempTermInYears
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3555 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3565(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim intCountryID As Integer
    Dim ndAddress As MSXML2.IXMLDOMNode
    Dim ndXML As MSXML2.IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3565"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'Iterate all customer Current addresses and see if any are not UK
    For Each ndAddress In xmlApplicationNode.selectNodes("./CUSTOMERROLE/CUSTOMERVERSION/CUSTOMERADDRESS/ADDRESS")
        If (ndAddress.Attributes.getNamedItem("COUNTRY_TYPE_UK") Is Nothing) Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
                       
            If Not ndAddress.Attributes.getNamedItem("COUNTRY") Is Nothing Then
                'Get Address Country for current address and exit
                intCountryID = ndAddress.Attributes.getNamedItem("COUNTRY").Text
                If Not xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='Country']/COMBOVALUE") Is Nothing Then
                    Set ndXML = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='Country']/COMBOVALUE[@VALUEID='" & intCountryID & "']")
                    strRuleValue = "Country: " & ndXML.Attributes.getNamedItem("VALUENAME").Text
                End If
            End If
            Exit For
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3565 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndAddress = Nothing
    Set ndXML = Nothing
    
End Function
Private Function Rule3570(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim lstCustomerAddress As MSXML2.IXMLDOMNodeList
    Dim attMovedOut As MSXML2.IXMLDOMAttribute
    Dim attMovedIn As MSXML2.IXMLDOMAttribute
    Dim iCustomerAddress As Integer
    
    Dim iAddressValidationYears As Integer
    Dim iMaxAddressGap As Integer
    Dim dtmMovedOut As Date
    Dim dtmMovedIn As Date
    Dim dtmToday As Date
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3570"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iAddressValidationYears = CInt(GetGlobalParameter(xmlApplicationNode, "AddressValidationYears", Amount))
    iMaxAddressGap = CInt(GetGlobalParameter(xmlApplicationNode, "MaxAddressGap", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        Set lstCustomerAddress = ndCustomerRole.selectNodes("./CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H='true' or @ADDRESSTYPE_TYPE_P='true']")
    
        If Not (lstCustomerAddress Is Nothing) Then
            
            'Get [date moved out] for the most recent CUSTOMERADDRESS
            If Not lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                dtmMovedIn = CDate(lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDIN").Text)
            Else
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATEMOVEDIN missing for Customer EMPLOYMENT, return a fail"
                Exit For
            End If
            
            If Not lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDOUT") Is Nothing Then
                dtmMovedOut = CDate(lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDOUT").Text)
            End If
            
            'Is there < AddressValidationYears of history between today and earliest movedin date?
            If TrueDateDiff("yyyy", dtmMovedIn, dtmToday) < iAddressValidationYears Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                Exit For
            Else
                'Iterate Previous and Current CUSTOMERADDRESS starting from last but one node
                iCustomerAddress = IIf((lstCustomerAddress.length > 1), (lstCustomerAddress.length - 2), 0)
                
                Do While iCustomerAddress >= 0
                    
                    If Not lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                        dtmMovedIn = CDate(lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDIN").Text)
                    Else
                        intScore = intRuleValidationType
                        intResult = RiskAssessmentDecision.Referral
                        CountReferral (intScore)
                        LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATEMOVEDIN missing for Customer EMPLOYMENT, return a fail"
                        Exit For
                    End If
                    
                    If dtmMovedOut > 0 Then 'ie if dtmMovedOut has been found (won't exist with a current address)
                        'Modified by OS - 28/11/2006
                        'Global parameter MaxAddressGap now specified in days instead of months
                        If TrueDateDiff("d", dtmMovedOut, dtmMovedIn) > iMaxAddressGap Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            Exit For
                        End If
                    End If
                    
                    Set attMovedOut = lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDOUT")
                    If Not (attMovedOut Is Nothing) Then
                        dtmMovedOut = CDate(attMovedOut.Text)
                    Else
                        If iCustomerAddress <> 0 Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            LogDetails 1, String(4, Chr(9)) & "Required attribute missing: DateMovedOut, return a fail"
                        End If
                    End If
                    
                    iCustomerAddress = (iCustomerAddress - 1)
                Loop
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3570 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    Set lstCustomerAddress = Nothing
    
End Function

Private Function Rule3575(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim lstCustomerAddress As MSXML2.IXMLDOMNodeList
    Dim attMovedOut As MSXML2.IXMLDOMAttribute
    Dim attMovedIn As MSXML2.IXMLDOMAttribute
    Dim iCustomerAddress As Integer
    
    Dim iAddressValidationUKYears As Integer
    Dim iMaxAddressGap As Integer
    Dim dtmMovedOut As Date
    Dim dtmMovedIn As Date
    Dim dtmToday As Date
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3575"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iAddressValidationUKYears = CInt(GetGlobalParameter(xmlApplicationNode, "AddressValidationUKYears", Amount))
    iMaxAddressGap = CInt(GetGlobalParameter(xmlApplicationNode, "MaxAddressGap", Amount))

    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        Set lstCustomerAddress = ndCustomerRole.selectNodes("./CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H='true' or @ADDRESSTYPE_TYPE_P='true' and ADDRESS/@COUNTRY_TYPE_UK='true']")
    
        If Not (lstCustomerAddress Is Nothing) Then
            
            'Get [date moved out] for the most recent CUSTOMERADDRESS
            If Not lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                dtmMovedIn = CDate(lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDIN").Text)
            Else
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATEMOVEDIN missing for Customer EMPLOYMENT, return a fail"
                Exit For
            End If
            
            If Not lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDOUT") Is Nothing Then
                dtmMovedOut = CDate(lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDOUT").Text)
            End If
            
            'Is there < AddressValidationYears of history between today and earliest movedin date?
            If TrueDateDiff("yyyy", dtmMovedIn, dtmToday) < iAddressValidationUKYears Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                Exit For
            Else
                'Iterate Previous and Current CUSTOMERADDRESS starting from last but one node to first
                iCustomerAddress = IIf(lstCustomerAddress.length > 1, (lstCustomerAddress.length - 2), 0)
                
                Do While iCustomerAddress >= 0
                    If Not lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                        dtmMovedIn = CDate(lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDIN").Text)
                    Else
                        intScore = intRuleValidationType
                        intResult = RiskAssessmentDecision.Referral
                        CountReferral (intScore)
                        LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATEMOVEDIN missing for Customer EMPLOYMENT, return a fail"
                        Exit For
                    End If
                    
                    If dtmMovedOut > 0 Then 'ie if dtmMovedOut has been found (won't exist with a current address)
                        'Modified by OS - 28/11/2006
                        'Global parameter MaxAddressGap now specified in days instead of months
                        If TrueDateDiff("d", dtmMovedOut, dtmMovedIn) > iMaxAddressGap Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            Exit For
                        End If
                    End If
                    
                    Set attMovedOut = lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDOUT")
                    If Not (attMovedOut Is Nothing) Then
                        dtmMovedOut = CDate(attMovedOut.Text)
                    Else
                        If iCustomerAddress <> 0 Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            LogDetails 1, String(4, Chr(9)) & "Required attribute missing: DateMovedOut, return a fail"
                        End If
                    End If
                    
                    iCustomerAddress = (iCustomerAddress - 1)
                Loop
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3575 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    Set lstCustomerAddress = Nothing
    
End Function

Private Function Rule3580(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim iNonEEAValueID As Integer
    Dim ndCustomerVersion As MSXML2.IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3580"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iNonEEAValueID = GetComboValueIDForValidationType(xmlApplicationNode, "Nationality", "NON")

    'Iterate each CUSTOMERVERSION
    For Each ndCustomerVersion In xmlApplicationNode.selectNodes("./CUSTOMERROLE/CUSTOMERVERSION")
        If Not ndCustomerVersion.Attributes.getNamedItem("NATIONALITY") Is Nothing Then
            'If any customer has a NonEEA Nationality fail rule and exit
            If ndCustomerVersion.Attributes.getNamedItem("NATIONALITY").Text = iNonEEAValueID Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                Exit For
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3580 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerVersion = Nothing
    
End Function

Private Function Rule3585(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerAddress As MSXML2.IXMLDOMNode
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3585"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'Iterate each current CUSTOMERADDRESS
    For Each ndCustomerAddress In xmlApplicationNode.selectNodes(".//CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H='true']")
    
        If Not (ndCustomerAddress.Attributes.getNamedItem("NATUREOFOCCUPANCY_TYPE_X") Is Nothing) Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
            Exit For
        End If
                        
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3585 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerAddress = Nothing
    
End Function

Private Function Rule3590(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndAllowableIncomeFactors As MSXML2.IXMLDOMNode
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim ndEarnedIncome As MSXML2.IXMLDOMNode
    Dim ndNetProfit As MSXML2.IXMLDOMNode
    
    Dim iEmploymentStatus_EMP As Integer
    Dim iEmploymentStatus_SELF As Integer
    Dim iSelfEmployedIncomeType_ANP  As Integer
    Dim iIncomeGroup_ED As Integer
    Dim iIncomeGroup_SD As Integer
    Dim iMinIncomeSingle As Double
    
    Dim attAllowableIncomeFactor As MSXML2.IXMLDOMNode
    Dim iEarnedIncomeAmount As Double
    Dim iEarnedIncomeFrequency As Integer
    Dim iEarnedIncomeType As Integer
    Dim iAllowableIncomeFactor As Integer
    Dim iTotalIncome As Double
    Dim iYear1Profits As Double
            
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3590"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'If there's just one applicant - take a deep breath and continue
    If xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']").length = 1 Then
      
        Set ndAllowableIncomeFactors = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS")
        iMinIncomeSingle = CDbl(GetGlobalParameter(xmlApplicationNode, "MinIncomeSingle", Amount))
        
        iEmploymentStatus_EMP = GetComboValueIDForValidationType(xmlApplicationNode, "EmploymentStatus", "EMP")
        iIncomeGroup_ED = GetComboValueIDForValidationType(xmlApplicationNode, "IncomeGroupType", "ED")
        iIncomeGroup_SD = GetComboValueIDForValidationType(xmlApplicationNode, "IncomeGroupType", "SD")
        iEmploymentStatus_SELF = GetComboValueIDForValidationType(xmlApplicationNode, "EmploymentStatus", "SELF")
        iSelfEmployedIncomeType_ANP = GetComboValueIDForValidationType(xmlApplicationNode, "SelfEmpDeclaredIncomeType", "ANP")
        
        'Iterate all customers
        For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        
            iTotalIncome = 0
        
            'For all 'Employed' Employment records
            Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1']")
            
            If Not ndEmployment Is Nothing Then

                'LDM 21/03/07 EP2_1852  if employed then look at earned income records
                If Not ndEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_EMP='true']") Is Nothing Then
                'Employed
                
                    'Calculate totalincome from all earnedIncome records
                    For Each ndEarnedIncome In ndEmployment.selectNodes("./EARNEDINCOME")
                            
                        If Not ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT") Is Nothing Then
                            iEarnedIncomeAmount = ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT").Text
                        End If
                        
                        If Not ndEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE") Is Nothing Then
                            iEarnedIncomeFrequency = ndEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE").Text
                        End If
                        
                        If Not ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE") Is Nothing Then
                            iEarnedIncomeType = ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE").Text
                        End If
                                            
                        If iEarnedIncomeAmount <> 0 Then
                            Set attAllowableIncomeFactor = _
                                       ndAllowableIncomeFactors.selectSingleNode("./ALLOWABLEINCOMEFACTOR[@INCOMEGROUP='" & iIncomeGroup_ED & "' and @EMPLOYMENTSTATUS='" & iEmploymentStatus_EMP & "' and @TYPE='" & iEarnedIncomeType & "']")
                                                                    
                            If Not (attAllowableIncomeFactor Is Nothing) Then
                                If Not attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                    iAllowableIncomeFactor = attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text
                                    iTotalIncome = iTotalIncome + ((iEarnedIncomeAmount * iEarnedIncomeFrequency) * (iAllowableIncomeFactor / 100))
                                End If
                            
                            End If
                        
                        End If
                        
                    Next
                    
                'LDM 21/03/07 EP2_1852 if self employed then look at net profit records
                'AW 10/04/07 EP2_2328
                ElseIf Not ndEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_SELF='true' or @EMPLOYMENTSTATUS_TYPE_CON='true']") Is Nothing Then
                'Self employed
                    Dim ProfitExists As Boolean
                    ProfitExists = True
                
                    Set ndNetProfit = ndEmployment.selectSingleNode("./NETPROFIT")
                    If ndEmployment.selectSingleNode("./NETPROFIT[@YEAR1AMOUNT]") Is Nothing Then
                        ProfitExists = False
                    End If
                            
                    If Not (ndNetProfit Is Nothing) And ProfitExists Then
                    
                        iYear1Profits = ndNetProfit.Attributes.getNamedItem("YEAR1AMOUNT").Text
                        
                        Set attAllowableIncomeFactor = _
                                   ndAllowableIncomeFactors.selectSingleNode("./ALLOWABLEINCOMEFACTOR[@INCOMEGROUP='" & iIncomeGroup_SD & "' and @EMPLOYMENTSTATUS='" & iEmploymentStatus_SELF & "' and @TYPE='" & iSelfEmployedIncomeType_ANP & "']")
                                                                
                        If Not (attAllowableIncomeFactor Is Nothing) Then
                            If Not attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                iAllowableIncomeFactor = attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text
                                iTotalIncome = iYear1Profits * (iAllowableIncomeFactor / 100)
                            End If
    
                        End If
                    Else
                        If ndNetProfit Is Nothing Then
                            LogDetails 1, String(4, Chr(9)) & "Mandatory element missing: NETPROFIT, return a fail"
                        End If
                        If ProfitExists = False Then
                            LogDetails 1, String(4, Chr(9)) & "Mandatory attibute missing: NETPROFIT@YEAR1AMOUNT, return a fail"
                        End If
                        
                        intScore = intRuleValidationType
                        intResult = RiskAssessmentDecision.DeclinePolicy
                        CountReferral (intScore)
                    End If
                    
                End If
            End If
            
            
            'If any Customer has < iMinIncomeSingle gross income rules fails.
            If iTotalIncome < iMinIncomeSingle Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.DeclinePolicy
                CountReferral (intScore)
                strRuleValue = "Income: " & CStr(iTotalIncome)
                Exit For
            End If
            
        Next
    
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If

    Rule3590 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function

'AShaw 20/04/2007 - EP2_2451
Private Function Rule3595(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode, _
                 Optional ByVal blnRunForGuarantors As Boolean) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim lstCustomerRoles As MSXML2.IXMLDOMNodeList
    Dim ndAllowableIncomeFactors As MSXML2.IXMLDOMNode
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim ndEarnedIncome As MSXML2.IXMLDOMNode
    Dim ndNetProfit As MSXML2.IXMLDOMNode
    
    Dim iEmploymentStatus_EMP As Integer
    Dim iEmploymentStatus_SELF As Integer
    Dim iEmploymentStatus_CON As Integer
    Dim iSelfEmployedIncomeType_ANP  As Integer
    Dim iIncomeGroup_ED As Integer
    Dim iIncomeGroup_SD As Integer
    Dim iMinIncomeSingle As Double
    Dim iMinIncomeJoint As Double
    
    Dim attAllowableIncomeFactor As MSXML2.IXMLDOMNode
    Dim iEarnedIncomeAmount As Double
    Dim iEarnedIncomeFrequency As Integer
    Dim iEarnedIncomeType As Integer
    Dim iAllowableIncomeFactor As Integer
    Dim iTotalIncome As Double
    Dim iYear1Profits As Double
    Dim bContinue As Boolean
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3595"
    bContinue = True
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    If blnRunForGuarantors Then
        'Get Gaurantor nodes
        Set lstCustomerRoles = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_G='true']")
    Else
        'Get applicant nodes
        Set lstCustomerRoles = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    End If
    
    If blnRunForGuarantors Or (lstCustomerRoles.length > 1) Then
        
        Set ndAllowableIncomeFactors = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS")
       
        iMinIncomeSingle = CDbl(GetGlobalParameter(xmlApplicationNode, "MinIncomeSingle", Amount))
        iMinIncomeJoint = CDbl(GetGlobalParameter(xmlApplicationNode, "MinIncomeJoint", Amount))

        iEmploymentStatus_EMP = GetComboValueIDForValidationType(xmlApplicationNode, "EmploymentStatus", "EMP")
        iIncomeGroup_ED = GetComboValueIDForValidationType(xmlApplicationNode, "IncomeGroupType", "ED")
        iIncomeGroup_SD = GetComboValueIDForValidationType(xmlApplicationNode, "IncomeGroupType", "SD")
        iEmploymentStatus_SELF = GetComboValueIDForValidationType(xmlApplicationNode, "EmploymentStatus", "SELF")
        ' EP2_2451
        iEmploymentStatus_CON = GetComboValueIDForValidationType(xmlApplicationNode, "EmploymentStatus", "CON")
        iSelfEmployedIncomeType_ANP = GetComboValueIDForValidationType(xmlApplicationNode, "SelfEmpDeclaredIncomeType", "ANP")
        
        iTotalIncome = 0
        
        'Iterate all customers
        For Each ndCustomerRole In lstCustomerRoles
        
            'For all 'Employed' Employment records
            Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1']")
            
            'LDM 05/03/2007 EP2_1650
            'if dont have a valid employment record with a main status of 1 then fail the rule
            If ndEmployment Is Nothing Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.DeclinePolicy
                CountReferral (intScore)
                strRuleValue = "Employment missing"
                LogDetails 1, String(4, Chr(9)) & "Mandatory element missing: EMPLOYMENT, return a fail"
                bContinue = False
                Exit For
            End If
            
            
            If Not ndEmployment Is Nothing Then

                'LDM 21/03/07 EP2_1852  if employed then look at earned income records
                If Not ndEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_EMP='true']") Is Nothing Then
                'Employed
                
                    'Calculate totalincome from all earnedIncome records
                    For Each ndEarnedIncome In ndEmployment.selectNodes("./EARNEDINCOME")
                            
                        If Not ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT") Is Nothing Then
                            iEarnedIncomeAmount = ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT").Text
                        End If
                        
                        If Not ndEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE") Is Nothing Then
                            iEarnedIncomeFrequency = ndEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE").Text
                        End If
                        
                        If Not ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE") Is Nothing Then
                            iEarnedIncomeType = ndEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE").Text
                        End If
                                            
                        If iEarnedIncomeAmount <> 0 Then
                            Set attAllowableIncomeFactor = _
                                       ndAllowableIncomeFactors.selectSingleNode("./ALLOWABLEINCOMEFACTOR[@INCOMEGROUP='" & iIncomeGroup_ED & "' and @EMPLOYMENTSTATUS='" & iEmploymentStatus_EMP & "' and @TYPE='" & iEarnedIncomeType & "']")
                                                                    
                            If Not (attAllowableIncomeFactor Is Nothing) Then
                                If Not attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                    iAllowableIncomeFactor = attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text
                                    iTotalIncome = iTotalIncome + ((iEarnedIncomeAmount * iEarnedIncomeFrequency) * (iAllowableIncomeFactor / 100))
                                End If
                            
                            End If
                        
                        End If
                        
                    Next
                    
                'LDM 21/03/07 EP2_1852 if self employed then look at net profit records
                'AW 10/04/07 EP2_2328
                ElseIf Not ndEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_SELF='true' or @EMPLOYMENTSTATUS_TYPE_CON='true']") Is Nothing Then
                'Self employed
                    Dim ProfitExists As Boolean
                    ProfitExists = True
                
                    Set ndNetProfit = ndEmployment.selectSingleNode("./NETPROFIT")
                    If ndEmployment.selectSingleNode("./NETPROFIT[@YEAR1AMOUNT]") Is Nothing Then
                        ProfitExists = False
                    End If
                            
                    If Not (ndNetProfit Is Nothing) And ProfitExists Then
                    
                        iYear1Profits = ndNetProfit.Attributes.getNamedItem("YEAR1AMOUNT").Text
                        
                        Set attAllowableIncomeFactor = _
                                   ndAllowableIncomeFactors.selectSingleNode("./ALLOWABLEINCOMEFACTOR[@INCOMEGROUP='" & iIncomeGroup_SD & "' and @EMPLOYMENTSTATUS='" & iEmploymentStatus_SELF & "' and @TYPE='" & iSelfEmployedIncomeType_ANP & "']")
                                                                
                        If Not (attAllowableIncomeFactor Is Nothing) Then
                            If Not attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                iAllowableIncomeFactor = attAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text
                                iTotalIncome = iYear1Profits * (iAllowableIncomeFactor / 100)
                            End If
    
                        End If
                    Else
                        If ndNetProfit Is Nothing Then
                            LogDetails 1, String(4, Chr(9)) & "Mandatory element missing: NETPROFIT, return a fail"
                        End If
                        If ProfitExists = False Then
                            LogDetails 1, String(4, Chr(9)) & "Mandatory attibute missing: NETPROFIT@YEAR1AMOUNT, return a fail"
                        End If
                        
                        intScore = intRuleValidationType
                        intResult = RiskAssessmentDecision.DeclinePolicy
                        CountReferral (intScore)
                    End If
                    
                End If
            End If
            
        Next
           
            
        If bContinue Then
        
            If blnRunForGuarantors Then
                'Totals for guarantor(s)
                If iTotalIncome < iMinIncomeSingle Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.DeclinePolicy
                    CountReferral (intScore)
                    strRuleValue = "Income: " & CStr(iTotalIncome)
                End If
            Else
                            
                If iTotalIncome < iMinIncomeJoint Then
                    'totalIncome < iMinIncomeJoint and no guarantor so fail
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.DeclinePolicy
                    CountReferral (intScore)
                    strRuleValue = "Income: " & CStr(iTotalIncome)
                Else
                    If xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_G='true']").length > 0 Then
                        'Run rule again for Guarantor(s)
                        Rule3595 xmlApplicationNode, xmlResponseNode, True
                    End If
                End If
                
            End If
        
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    End If

    Rule3595 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function
Private Function Rule3600(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim iMinEmployHistMonths As Integer
    Dim blnEmploymentExists As Boolean
    
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    
    Dim dtmStartedOrEst As Date
    Dim dtmLeftOrCeased As Date
    Dim dtmToday As Date
        
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3600"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iMinEmployHistMonths = CInt(GetGlobalParameter(xmlApplicationNode, "MinEmploymentHist", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        blnEmploymentExists = False
    
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        If Not ndEmployment Is Nothing Then

            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
            'Loop all Emploment records. if the customer was employed iMinEmployHistMonths ago then rule passes
            For Each ndEmployment In ndCustomerRole.selectNodes("./CUSTOMERVERSION/EMPLOYMENT[not(@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true')]")
                Set attStartedOrEst = ndEmployment.Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
                Set attLeftOrCeased = ndEmployment.Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                
                If Not (attStartedOrEst Is Nothing) Then
                    dtmStartedOrEst = attStartedOrEst.Text
                    
                    'Is the dtmStartedOrEst over iMinEmployHistMonths (12 months) ago
                    If TrueDateDiff("m", dtmStartedOrEst, dtmToday) >= iMinEmployHistMonths Then
                        If Not (attLeftOrCeased Is Nothing) Then
                            dtmLeftOrCeased = attLeftOrCeased.Text
                                
                            'Is the dtmLeftOrCeased more recent than iMinEmployHistMonths (12 months) ago
                            If TrueDateDiff("m", dtmLeftOrCeased, dtmToday) < iMinEmployHistMonths Then
                                blnEmploymentExists = True
                                Exit For
                            End If
                        Else
                            'No DATELEFTORCEASEDTRADING found so still employed
                            blnEmploymentExists = True
                            Exit For
                        End If
                    End If
                Else
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED missing for Customer EMPLOYMENT, return a fail"
                End If
            Next
                    
            If Not (blnEmploymentExists) Then
                'Valid Employment record for Customer didn't exist so fail rule and exit loop
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                Exit For
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3600 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function

Private Function Rule3605(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
            
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3605"
    
    'Rule 3605: Non-permanent employment
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        'Locate the Employment record where Employment.MainStatus = True AND Employment.EmploymentStatus = "Employed - Temporary" (combo "EmploymentStatus" ValidationType = "EMPT").
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        'IF the record is found then [FAIL].
        If Not ndEmployment Is Nothing Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
            Exit For
        End If

    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3605 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function

Private Function Rule3610(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim iMinTimeEmployed As Integer
    
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    
    Dim dtmStartedOrEst As Date
    Dim dtmLeftOrCeased As Date
    Dim dtmToday As Date
        
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3610"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iMinTimeEmployed = CInt(GetGlobalParameter(xmlApplicationNode, "MinTimeEmployed", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        If Not ndEmployment Is Nothing Then
    
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
            
            Set attStartedOrEst = ndEmployment.Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
            Set attLeftOrCeased = ndEmployment.Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                
            If Not (attStartedOrEst Is Nothing) Then
                If DateDiffInMonthsRoundedDown(attStartedOrEst.Text, dtmToday) < iMinTimeEmployed Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore) 'add to appropriate referral count
                    strRuleValue = "Time (months): " & CStr(DateDiffInMonthsRoundedDown(attStartedOrEst.Text, dtmToday))
                    Exit For
                End If
            Else
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED missing for Customer EMPLOYMENT, return a fail"
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3610 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function

Private Function Rule3615(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim iMinEmployHistMonths As Integer
    Dim iMaxNumEmployments As Integer
    
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    
    Dim dtmStartedOrEst As Date
    Dim dtmLeftOrCeased As Date
    Dim dtmToday As Date
    Dim iEmploymentCount As Integer
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3615"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    iMinEmployHistMonths = CInt(GetGlobalParameter(xmlApplicationNode, "MinEmploymentHist", Amount))
    iMaxNumEmployments = CInt(GetGlobalParameter(xmlApplicationNode, "MaxNumberEmployments", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        
        'default for each customer
        iEmploymentCount = 0
    
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        If Not ndEmployment Is Nothing Then
    
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
            
            'Loop all Emploment records. if the customer was employed iMinEmployHistMonths ago then rule passes
            For Each ndEmployment In ndCustomerRole.selectNodes("./CUSTOMERVERSION/EMPLOYMENT[not(@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true')]")
                Set attLeftOrCeased = ndEmployment.Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                
                If Not (attLeftOrCeased Is Nothing) Then
                    dtmLeftOrCeased = attLeftOrCeased.Text
                    
                    'Is the dtmLeftOrCeased more recent than (or in the future) iMinEmployHistMonths (12 months) ago
                    If DateDiffInMonthsRoundedUp(dtmLeftOrCeased, dtmToday) <= iMinEmployHistMonths Then
                        iEmploymentCount = iEmploymentCount + 1
                    End If
                Else
                    'Current employment wont have an end date
                    iEmploymentCount = iEmploymentCount + 1
                End If
            Next
            
            'If theres more than iMaxNumEmployments in last iMinEmployHistMonths
            'for any customer rule fails. Exit loop
            If iEmploymentCount > iMaxNumEmployments Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                Exit For
            End If
        End If
    Next
       
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3615 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function

Private Function Rule3620(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer

    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim lstEmployment  As MSXML2.IXMLDOMNodeList
    Dim iEmployment  As Integer
    Dim iMinEmployHistMonths As Integer
    Dim iMaxEmploymentGapWeeks As Integer
    Dim iMinMonthsCount As Integer
    
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim dtmLeftOrCeased As Date
    Dim dtmStartedOrEst As Date
    Dim dtmToday As Date

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3620"
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iMinEmployHistMonths = CInt(GetGlobalParameter(xmlApplicationNode, "MinEmploymentHist", Amount))
    iMaxEmploymentGapWeeks = CInt(GetGlobalParameter(xmlApplicationNode, "MaxEmploymentGap", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")

        'Check there's a main employment record else fail
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")

        If Not ndEmployment Is Nothing Then
            
            '---------------------------------------------------
            'continue
    
            iMinMonthsCount = 0
            iEmployment = 0
            
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
            Set lstEmployment = ndCustomerRole.selectNodes("./CUSTOMERVERSION/EMPLOYMENT[not(@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true')]")
            
            'Check gap between today and dtmLeftOrCeased of most recent employment
            Set attLeftOrCeased = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
            If Not (attLeftOrCeased Is Nothing) Then
                dtmLeftOrCeased = CDate(attLeftOrCeased.Text)
                
                'Is Gap between Today and last dtmLeftOrCeased more then 4 weeks?
                If TrueDateDiff("w", dtmLeftOrCeased, dtmToday) > iMaxEmploymentGapWeeks Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore) 'add to appropriate referral count
                    Exit For
                End If
            Else
                'Current employment so there's no DateLeftOrCeasedTrading
                dtmLeftOrCeased = dtmToday
            End If
            
            'Get dtmStartedOrEst for most recent employment
            Set attStartedOrEst = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
            If Not (attStartedOrEst Is Nothing) Then
                dtmStartedOrEst = CDate(attStartedOrEst.Text)
                iEmployment = iEmployment + 1
                
                'Continue if we're still processing records within the last MinEmploymentHist months else exit
                Do While (iEmployment <= (lstEmployment.length - 1)) And (DateDiffInMonthsRoundedUp(dtmLeftOrCeased, dtmToday) <= iMinEmployHistMonths)
                    
                    Set attLeftOrCeased = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                    If Not (attLeftOrCeased Is Nothing) Then
                        dtmLeftOrCeased = CDate(attLeftOrCeased.Text)
                        
                        'Is Gap between the previous Employment dtmStartedOrEst and the current
                        'Employment dtmLeftOrCeased > iMaxEmploymentGapWeeks (weeks)
                        If TrueDateDiff("w", dtmLeftOrCeased, dtmStartedOrEst) > iMaxEmploymentGapWeeks Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            Exit For
                        End If
                    
                    End If
                    
                    'Get dtmStartedOrEst of current Employment record
                    Set attStartedOrEst = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
                    If Not (attStartedOrEst Is Nothing) Then
                        dtmStartedOrEst = CDate(attStartedOrEst.Text)
                    End If
                    
                    iEmployment = iEmployment + 1
                Loop
            
            Else
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED missing for Customer EMPLOYMENT, return a fail"
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
            End If
        End If
    Next

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3620 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Set ndCustomerRole = Nothing
    Set lstEmployment = Nothing

End Function

Private Function Rule3625(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer

    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    
    Dim iMinTimeSelfEmployed As Integer
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim dtmStartedOrEst As Date
    Dim dtmToday As Date

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3625"
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iMinTimeSelfEmployed = CInt(GetGlobalParameter(xmlApplicationNode, "MinTimeSelfEmployed", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")

        'Check there's a main employment record else fail
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_SELF='true']")

        If Not (ndEmployment Is Nothing) Then

            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
            Set attStartedOrEst = ndEmployment.Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
            If Not (attStartedOrEst Is Nothing) Then
                dtmStartedOrEst = CDate(attStartedOrEst.Text)
                
                If DateDiffInMonthsRoundedDown(dtmStartedOrEst, dtmToday) < iMinTimeSelfEmployed Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    strRuleValue = "Start Date: " & dtmStartedOrEst
                    Exit For
                End If
            Else
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED missing for Customer EMPLOYMENT, return a fail"
            End If
        End If
        
    Next

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3625 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Set ndCustomerRole = Nothing

End Function

Private Function Rule3630(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer

    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim ndContractDetails As MSXML2.IXMLDOMNode
    
    Dim attLengthOfContractMonths As MSXML2.IXMLDOMAttribute
    Dim attLengthOfContractYears As MSXML2.IXMLDOMAttribute
    Dim attLikelyToBeRenewed As MSXML2.IXMLDOMNode
    
    Dim iTotalContractMonths As Integer
    Dim iMinTimeContract As Integer
    Dim blnLikelyToBeRenewed As Boolean
    Dim dtmToday As Integer
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3630"
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iMinTimeContract = CInt(GetGlobalParameter(xmlApplicationNode, "MinTimeContract", Amount))
    
    'Iterate each CUSTOMERROLE
    
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")

        'Check there's a main employment record else fail
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_CON='true']")

        If Not (ndEmployment Is Nothing) Then
            
            Set ndContractDetails = ndEmployment.selectSingleNode("./CONTRACTDETAILS")
            
            If Not (ndContractDetails Is Nothing) Then
            
                Set attLengthOfContractMonths = ndContractDetails.Attributes.getNamedItem("LENGTHOFCONTRACTMONTHS")
                Set attLengthOfContractYears = ndContractDetails.Attributes.getNamedItem("LENGTHOFCONTRACTYEARS")
                Set attLikelyToBeRenewed = ndContractDetails.Attributes.getNamedItem("LIKELYTOBERENEWED")
                
                If Not (attLengthOfContractYears Is Nothing) Then
                   iTotalContractMonths = CInt(attLengthOfContractYears.Text) * 12
                End If
                
                If Not (attLengthOfContractMonths Is Nothing) Then
                    iTotalContractMonths = iTotalContractMonths + CInt(attLengthOfContractMonths.Text)
                End If
                
                If Not (attLikelyToBeRenewed Is Nothing) Then
                    blnLikelyToBeRenewed = attLikelyToBeRenewed.Text
                End If
                
                If (iTotalContractMonths < iMinTimeContract) Or Not (blnLikelyToBeRenewed) Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    strRuleValue = "Length (months): " & iTotalContractMonths
                    Exit For
                End If
            End If
        End If
        
    Next

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule3630 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Set ndCustomerRole = Nothing

End Function

Private Function Rule3640(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer

    Dim lstCustomerRole As MSXML2.IXMLDOMNodeList
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    
    Dim iCustomersUnemployed  As Integer
    Dim iTotalCustomers As Integer

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3640"
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'OS 16/04/2007 - EP2_2427
    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType = "BR") then do not run the rule.
    If xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
    'Get count of customers
        Set lstCustomerRole = xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        iTotalCustomers = lstCustomerRole.length
        
        For Each ndCustomerRole In lstCustomerRole
    
            'Get the main unemployed employment record
            Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true']")
    
            'if no employment record exit
            If ndEmployment Is Nothing Then
                Exit For
            End If
        
            iCustomersUnemployed = iCustomersUnemployed + 1
        
        Next
    
        'If all customers are unemployed rule fails
        If (iCustomersUnemployed = iTotalCustomers) Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
    
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    End If
    
    Rule3640 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Set ndCustomerRole = Nothing

End Function

'=======================================
'OS 27/11/2006
'=======================================
Private Function Rule3645(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlApplicants As IXMLDOMNodeList
    Dim xmlIncome As IXMLDOMNode
    Dim xmlEmploymentRecord As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intNumberOfApplicants As Integer
        
    Dim dblAllowableIncome As Double, _
        dblTotalApproximateMonthlyCost As Double
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String

        
    Dim blnSuccess As Boolean, _
        blnOneApplicantFailed As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3645"
    strScoreCardInd = 0 'False
    'Set default value for Allowable income
    dblAllowableIncome = 0
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3645: One employed applicant cannot afford repayment
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    blnOneApplicantFailed = False
    
    Set xmlApplicants = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    intNumberOfApplicants = xmlApplicants.length
    
    'IF number of applicants is one then do not run the rule
    If intNumberOfApplicants > 1 Then
        
        'Get TotalApproximateMonthlyCost
        If Not xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST") Is Nothing Then
            dblTotalApproximateMonthlyCost = CDbl(xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST").Text)
        End If
        
        'Read each applicant for the application (applicants)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
            If Not xmlApplicant Is Nothing Then
                'Get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                'Get customer number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                    strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
                
                If Not xmlNode Is Nothing Then
                    Set xmlEmploymentRecord = xmlNode.selectSingleNode("./EMPLOYMENT[@MAINSTATUS='1']")
                    
                    If Not xmlEmploymentRecord Is Nothing Then
                        If Not xmlEmploymentRecord.Attributes.getNamedItem("EMPLOYMENTSTATUS_TYPE_NOEMPLOYER") Is Nothing Then
                            blnOneApplicantFailed = True
                        Else
                            'Calculate income for applicant
                            dblAllowableIncome = 0
                            
                            'Get income from NETCONFIRMEDALLOWABLEINCOME attribute
                            Set xmlIncome = xmlNode.selectSingleNode("./INCOMESUMMARY[@NETCONFIRMEDALLOWABLEINCOME]")
                            If Not xmlIncome Is Nothing Then
                                dblAllowableIncome = CDbl(xmlIncome.Attributes.getNamedItem("NETCONFIRMEDALLOWABLEINCOME").Text)
                            End If
                            
                            'Get income from NETALLOWABLEANNUALINCOME attribute
                            If dblAllowableIncome = 0 Then
                                Set xmlIncome = xmlNode.selectSingleNode("./INCOMESUMMARY[@NETALLOWABLEANNUALINCOME]")
                                If Not xmlIncome Is Nothing Then
                                    dblAllowableIncome = CDbl(xmlIncome.Attributes.getNamedItem("NETALLOWABLEANNUALINCOME").Text)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next xmlApplicant
        
        If blnOneApplicantFailed Then
            
            'If both applicants failed the check, then Allowable Income would be zero
            If dblTotalApproximateMonthlyCost > dblAllowableIncome Then
                blnSuccess = False
            End If
        End If
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
        End If
        
        'strRuleValue = "Age: " & intApplicantsAge + intMaxTermInYears
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3645 = intResult
    
    Set xmlNode = Nothing
                        
End Function

'=======================================
'OS 28/11/2006
'=======================================

Private Function Rule3650(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlBureauData As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
    
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String

    Dim dblTotalBalance As Double, _
        dblLowerLimit As Double, _
        dblUpperLimit As Double, _
        dblCAISCardVarianceLow As Double, _
        dblCAISCardVarianceHigh As Double, _
        dblCurrentBalanceActiveCards As Double
    
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3650"
    strScoreCardInd = 0 'False
    dblTotalBalance = 0
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3650: Total card balance not equal CAIS within tolerance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True

    ' EP2_2228
    dblCAISCardVarianceLow = CDbl(GetGlobalParameter(xmlApplicationNode, "CAISCardVarianceLow", Percentage))
    dblCAISCardVarianceHigh = CDbl(GetGlobalParameter(xmlApplicationNode, "CAISCardVarianceHigh", Percentage))
   
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
        If Not xmlApplicant Is Nothing Then
            'Get customer version number
            If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
            End If
            
            'Get customer number
            If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
            End If
            
            Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
            
            'Iterate through each LoansLiabilities record
            For Each xmlNode In xmlApplicant.selectNodes("./CUSTOMERVERSION/ACCOUNTRELATIONSHIP[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']/ACCOUNT/LOANSLIABILITIES[@AGREEMENTTYPE_TYPE_CC='true' or @AGREEMENTTYPE_TYPE_CHC='true' or @AGREEMENTTYPE_TYPE_CSC='true']")
                If Not xmlNode.Attributes.getNamedItem("TOTALOUTSTANDINGBALANCE") Is Nothing Then
                    'Calculate Total Balance
                    dblTotalBalance = dblTotalBalance + CDbl(xmlNode.Attributes.getNamedItem("TOTALOUTSTANDINGBALANCE").Text)
                End If
            Next xmlNode
        End If
    Next xmlApplicant
    
    'Get Bespoke Bureau Data
    Set xmlBureauData = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK[position()=1]/CREDITCHECK/BESPOKEBUREAUDATA")
            
    'The collected total must be within the Experian CAIS data limits
    ' EP2_2228 - Now uses different High and low variances.
    If Not xmlBureauData Is Nothing Then
        If Not xmlBureauData.Attributes.getNamedItem("CURRENTBALANCEACTIVECARDS") Is Nothing Then
            dblCurrentBalanceActiveCards = CDbl(xmlBureauData.Attributes.getNamedItem("CURRENTBALANCEACTIVECARDS").Text)
            dblLowerLimit = Round(dblCurrentBalanceActiveCards * dblCAISCardVarianceLow / 100, 0)
            dblUpperLimit = Round(dblCurrentBalanceActiveCards * dblCAISCardVarianceHigh / 100, 0)
            
            If dblTotalBalance < dblLowerLimit Or dblTotalBalance > dblUpperLimit Then
                blnSuccess = False
            End If
        End If
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
    End If
    
    strRuleValue = "CAIS: " & dblCurrentBalanceActiveCards & "; Entered: " & dblTotalBalance
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
  
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3650 = intResult
    
    Set xmlNode = Nothing
        
End Function

'======================================================================================
'Added By:Dhira Singh
'DIP Rule : 3655
'======================================================================================
Private Function Rule3655(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlBureauData As IXMLDOMNode
    Dim intResult As Integer, _
    intScore As Integer, _
    intRuleValidationType As Integer, _
    intMortgageSubquoteNumber As Integer
    
    Dim strRuleValue As String, _
    strCustomerVersionNumber As String, _
    strCustomerNumber As String, _
    strRuleDescription As String
    
    Dim dblTotalPayment As Double, _
    dblLowerLimit As Double, _
    dblUpperLimit As Double, _
    dblCAISLoanVarianceLow As Double, _
    dblCAISLoanVarianceHigh As Double, _
    dblTotalRepayLoansMore12m As Double
    
    Dim blnSuccess As Boolean
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3655"
    strScoreCardInd = 0 'False
    dblTotalPayment = 0
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3655: Total number of existing loans held and total monthly repayment does not match CAIS data
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    ' EP2_2228
    dblCAISLoanVarianceLow = CDbl(GetGlobalParameter(xmlApplicationNode, "CAISLoanVarianceLow", Percentage))
    dblCAISLoanVarianceHigh = CDbl(GetGlobalParameter(xmlApplicationNode, "CAISLoanVarianceHigh", Percentage))
    
    'FOR each LoansLiabilities record where LoansLiabilities.AgreementType = "Personal Loan"
    'OR = "Secured Loan" OR = "Car Loan" OR = "Business Loan" OR = "Other Loan"
    '(combo "AgreementType" ValidationType = "P" OR = "S" OR = "C" OR = "BL" OR = "OL")
    'set totalpayment = totalpayment + LoansLiabilities.MonthlyRepayment.
    
    'Specific Database logic
    'Read the ApplicationCreditCheck record with the highest value of DateTime
    'for the ApplicationFactFind. Read the CreditCheck record for the ApplicationCreditCheck
    'and the BespokeBureauData record for the CreditCheck.
    'FOR each CustomerRole record with CustomerRoleType = "Applicant"
    'read all AccountRelationship records for the CustomerNumber AND CustomerversionNumber.
    'FOR each AccountRelationship record read the LoansLiabilities record for the AccountGuid - Not yet implemented
    
   
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
        If Not xmlApplicant Is Nothing Then
            
            'get customer version number
            If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
            End If
            
            'get customer number
            If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
            End If
            
            For Each xmlNode In xmlApplicant.selectNodes("./CUSTOMERVERSION/ACCOUNTRELATIONSHIP[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER ='" & strCustomerNumber & "']/ACCOUNT/LOANSLIABILITIES[@AGREEMENTTYPE_TYPE_P='true' or @AGREEMENTTYPE_TYPE_S='true' or @AGREEMENTTYPE_TYPE_C = 'true' or @AGREEMENTTYPE_TYPE_BL = 'true' or @AGREEMENTTYPE_TYPE_OL = 'true']")
                If Not xmlNode.Attributes.getNamedItem("MONTHLYREPAYMENT") Is Nothing Then
                    dblTotalPayment = dblTotalPayment + CDbl(xmlNode.Attributes.getNamedItem("MONTHLYREPAYMENT").Text)
                End If
            Next xmlNode
        End If
    Next xmlApplicant
    
    
    'The collected total must match Experian CAIS data within 10%.
    'Set lowerlimit = BespokeBureauData.TotalRepayLoansMore12m * (100 - global parameter "CAISLoanVariance") / 100 rounded to nearest integer.
    'Set upperlimit = BespokeBureauData.TotalRepayLoansMore12m * (100 + global parameter "CAISLoanVariance") / 100 rounded to nearest integer.
    
    'get most recent application credit check record
    Set xmlBureauData = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK[1]/CREDITCHECK/BESPOKEBUREAUDATA")
    
    If Not xmlBureauData Is Nothing Then
        If Not xmlBureauData.Attributes.getNamedItem("TOTALREPAYLOANSMORE12M") Is Nothing Then
            
            dblTotalRepayLoansMore12m = CDbl(xmlBureauData.Attributes.getNamedItem("TOTALREPAYLOANSMORE12M").Text)
            dblLowerLimit = Round(dblTotalRepayLoansMore12m * dblCAISLoanVarianceLow / 100, 0)
            dblUpperLimit = Round(dblTotalRepayLoansMore12m * dblCAISLoanVarianceHigh / 100, 0)
            
            'IF totalpayment < lowerlimit OR > upperlimit then [FAIL].
            If dblTotalPayment < dblLowerLimit Or dblTotalPayment > dblUpperLimit Then
                blnSuccess = False
            End If
            
        End If
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        ' EP2_2268
        intResult = RiskAssessmentDecision.Referral
    End If
    
    strRuleValue = "CAIS: " & dblTotalRepayLoansMore12m & "; Entered  : " & dblTotalPayment
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    Rule3655 = intResult
    Set xmlNode = Nothing
    
End Function
'======================================================================================
'End :DIP Rule 3655
'======================================================================================

Private Function Rule3660(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim dblAmountRequested As Double, _
        dblMaxAutoAcceptLimit As Double
        
    Dim strScoreCardInd As String
    
    Dim ndQuotation As MSXML2.IXMLDOMNode
    
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3660"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
        
    dblMaxAutoAcceptLimit = CDbl(GetGlobalParameter(xmlApplicationNode, "MaxAutoAcceptLimit", Amount))
    
    Set ndQuotation = xmlApplicationNode.selectSingleNode("./QUOTATION")
    
    If ndQuotation Is Nothing Then
        'get loan amount for the application from NEWLOAN
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
    Else
        'get loan amount for the application from MORTGAGESUBQUOTE
        If Not ndQuotation.selectSingleNode("./MORTGAGESUBQUOTE/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(ndQuotation.selectSingleNode("./MORTGAGESUBQUOTE/@AMOUNTREQUESTED").Text)
        End If
    End If
    
    'If NewLoan.AmountRequested > global parameter MaxAutoAcceptLimit (£250,000) then [FAIL]
    If dblAmountRequested > dblMaxAutoAcceptLimit Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        strRuleValue = "Loan: £" & dblAmountRequested
        CountReferral (intScore)
    End If
        
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3660 = intResult
        
End Function

'======================================================================================
'Added By:Dhira Singh
'DIP Rule : 3665
'======================================================================================
 Private Function Rule3665(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intProdScheme As Integer
    Dim intRuleValidationType As Integer
    
    Dim dblBandAmount As Double
    Dim dblLTV As Double
    Dim dblAmountRequested As Double
    Dim dblPurchPriceEstVal As Double
    
    Dim bRes As Boolean
    Dim bLTB As Boolean
    Dim bBTL As Boolean
    
    Dim strEligInd As String
    Dim strSeek As String
    
    Dim xmlNode As IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3665"
    
    'Rule 3665 : LTV or loan amount > max limits according to product category and scheme.
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    
    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If
    
    If strEligInd <> "" Then
    
        If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Then
            strSeek = "MaxLoanByLTVRes"
        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing Then
            strSeek = "MaxLoanByLTVLTB"
        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
            strSeek = "MaxLoanByLTVBTL"
        End If
            
        
        Select Case strEligInd
                Case "ST":  strSeek = strSeek & "PR"
                Case "UL":  strSeek = strSeek & "NP"
                Case Else:  strSeek = strSeek & strEligInd
        End Select
             
             
        
        If Not xmlApplicationNode.selectSingleNode("QUOTATION") Is Nothing Then
            If Not xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@LTV") Is Nothing Then
                dblLTV = xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@LTV").Text
            End If
            
            If Not xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED") Is Nothing Then
            
              dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED").Text)
            End If
            
        Else
            If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
                dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
            End If
        
            If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
                dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
            End If
            
            dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
        End If
        
   
      ' Locate the global parameter band where High Band = smallest value >= LTV.
        If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='" & strSeek & "'][@HIGHBAND>='" & CStr(dblLTV) & "']") Is Nothing Then

            Set xmlNode = xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='" & strSeek & "'][@HIGHBAND>='" & CStr(dblLTV) & "']")
            dblBandAmount = CDbl(xmlNode.Attributes.getNamedItem("AMOUNT").Text)

        End If
     
        'IF Loan > band Amount then [FAIL]
        If dblAmountRequested > dblBandAmount Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

    End If
    
    
    strRuleValue = "Loan: " & dblAmountRequested
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Rule3665 = intResult
End Function
 
'======================================================================================
'End :DIP Rule 3665
'======================================================================================

Private Function Rule3670(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer, _
        lngMinPurchPriceRes As Long, _
        lngMinPurchPriceBTL As Long
    
    Dim dblValuation As Double
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3670"
    
    'Rule 3670: Property value / PP less than minimum
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    lngMinPurchPriceRes = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceRes", Amount)
    lngMinPurchPriceBTL = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceBTL", Amount)

    'Set valuation = ApplicationFactFind.PurchasePriceOrEstimatedValue.
    If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
        dblValuation = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
    End If
            
    'IF ApplicationFactFind.NatureOfLoan = "Residential" OR = "Let-to-Buy" (combo "NatureOfLoan" ValidationType = "RS" OR = "LT")
    'AND valuation < global parameter "MinPurchPriceRes" then [FAIL].
    If (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing) And _
    (dblValuation < lngMinPurchPriceRes) Then
        blnSuccess = False
    'ELSE IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let (Income)" OR = "Buy-to-Let (Rental)" (combo "NatureOfLoan" ValidationType = "BI" OR = "BR")
    'AND valuation < global parameter "MinPurchPriceBTL" then [FAIL].
    ElseIf (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing) And _
    (dblValuation < lngMinPurchPriceBTL) Then
        blnSuccess = False
    End If
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.DeclinePolicy
    End If
        
    strRuleValue = "Value: " & dblValuation
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3670 = intResult
End Function

Private Function Rule3675(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer, _
        lngMinPurchPriceRes As Long, _
        lngMinPurchPriceResHigh As Long, _
        lngMinPurchPriceBTL As Long, _
        lngMinPurchPriceBTLHigh As Long
    
    Dim dblValuation As Double
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3675"
    
    'Rule 3675: Property value / PP less than upper limit
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    lngMinPurchPriceRes = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceRes", Amount)
    lngMinPurchPriceResHigh = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceResHigh", Amount)
    lngMinPurchPriceBTL = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceBTL", Amount)
    lngMinPurchPriceBTLHigh = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceBTLHigh", Amount)

    'Set valuation = ApplicationFactFind.PurchasePriceOrEstimatedValue.
    If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
        dblValuation = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
    End If
            
    'IF ApplicationFactFind.NatureOfLoan = "Residential" OR = "Let-to-Buy" (combo "NatureOfLoan" ValidationType = "RS" OR = "LT")
    'AND (valuation >= global parameter "MinPurchPriceRes" AND < global parameter "MinPurchPriceResHigh) then [FAIL].
    If (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing) And _
    (dblValuation >= lngMinPurchPriceRes And dblValuation < lngMinPurchPriceResHigh) Then
        blnSuccess = False
    'ELSE IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let (Income)" OR = "Buy-to-Let (Rental)"
    '(combo "NatureOfLoan" ValidationType = "BI" OR = "BR") AND (valuation >= global parameter "MinPurchPriceBTL" AND < global parameter "MinPurchPriceBTLHigh) then [FAIL].
    ElseIf (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing) And _
    (dblValuation >= lngMinPurchPriceBTL And dblValuation < lngMinPurchPriceBTLHigh) Then
        blnSuccess = False
    End If
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
    End If
        
    strRuleValue = "Value: " & dblValuation
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3675 = intResult
End Function

Private Function Rule3690(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer
    
    Dim dblMinLoan As Double
    Dim dblAmountRequested As Double
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3690"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'rule 3690: Loan less than minimum
        
    dblMinLoan = CDbl(GetGlobalParameter(xmlApplicationNode, "MinimumLoan", Amount))
    
    'IF the Quotation is not found then set loan = NewLoan.AmountRequested ELSE set loan = MortgageSubQuote.AmountRequested.
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    If intMortgageSubquoteNumber = 0 Then
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
    Else
        Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("AMOUNTREQUESTED") Is Nothing Then
                dblAmountRequested = xmlNode.Attributes.getNamedItem("AMOUNTREQUESTED").nodeValue
            End If
        End If
    End If
    
    If dblAmountRequested < dblMinLoan Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.DeclinePolicy
    End If
    
    strRuleValue = "Loan: " & dblAmountRequested
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3690 = intResult
End Function

'======================================================================================
'Added By:Dhira Singh
'DIP Rule : 3695
'======================================================================================
Private Function Rule3695(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim dblBandAmount As Double
    Dim dblAmountRequested As Double
    Dim dblPurchPriceEstVal As Double
    Dim dblLTV As Double
    
    Dim xmlNode As IXMLDOMNode
    Dim iNatureOfLoan As Integer
    Dim strEligInd As String, strProdScheme As String
    Dim strTypeOfBuyer As String
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3695"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

'    // This rule is only run for a "Scheme Finder" DIP; i.e. where the mortgage product is not known and, consequently, a quotation has not been created.
'
'    IF the Quotation exists OR Rule 3500 = [FAIL] OR Rule 3505 = [FAIL] OR RULE 3510 = [FAIL] then do not run this rule.
'
'    IF ApplicationFactFind.NatureOfLoan = "Residential" (combo "NatureOfLoan" ValidationType = "RS") then
'    IF ApplicationFactFind.SelfCertInd = False use global banded parameter "MaxLTVResProdSchemeFS"
'    ELSE use global banded parameter "MaxLTVResProdSchemeSC".
'
'    IF ApplicationFactFind.NatureOfLoan = "Let-to-Buy" (combo "NatureOfLoan" ValidationType = "LT") then
'    IF ApplicationFactFind.SelfCertInd = False use global banded parameter "MaxLTVLTBProdSchemeFS"
'    ELSE use global banded parameter "MaxLTVLTBProdSchemeSC".
'
'    IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let" (combo "NatureOfLoan" ValidationType = "BI" OR = "BR") then
'    IF ApplicationFactFind.TypeOfApplication not = "Remortgage" (combo "TypeOfMortgage" ValidationType = "R") use global banded parameter "MaxLTVBTLProdScheme"
'    ELSE use global banded parameter "MaxLTVBTLProdSchemeRmtg".
'
'    Calculate LTV = NewLoan.AmountRequested /
'    ApplicationFactFind.PurchasePriceOrEstimatedValue * 100 rounded to 2dp.
'
'    Locate the entry in combo "ProductScheme" where ValidationType =
'    BespokeDecision.EligibilityIndicator. Locate the global parameter band where
'    High Band = ValueId from combo. IF LTV > band Amount then [FAIL].

    
   If xmlApplicationNode.selectSingleNode("QUOTATION") Is Nothing Then
        If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
            strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
        End If
        
        If strEligInd <> "" Then
            Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='ProductScheme']/COMBOVALUE[./COMBOVALIDATION/@VALIDATIONTYPE='" & strEligInd & "']")
            If Not xmlNode Is Nothing Then
                strProdScheme = xmlNode.Attributes.getNamedItem("VALUEID").nodeValue
            End If
        End If

        If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Then
            If Not xmlApplicationNode.Attributes.getNamedItem("APPLICATIONINCOMESTATUS_TYPE_FS") Is Nothing Then
                If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVResProdSchemeFS'][@HIGHBAND='" & strProdScheme & "']") Is Nothing Then
                    dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVResProdSchemeFS'][@HIGHBAND='" & strProdScheme & "']").Attributes.getNamedItem("PERCENTAGE").Text)
                End If
            Else
                If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVResProdSchemeSC'][@HIGHBAND='" & strProdScheme & "']") Is Nothing Then
                    dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVResProdSchemeSC'][@HIGHBAND='" & strProdScheme & "']").Attributes.getNamedItem("PERCENTAGE").Text)
                End If
            End If
        
        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing Then
            If Not xmlApplicationNode.Attributes.getNamedItem("APPLICATIONINCOMESTATUS_TYPE_FS") Is Nothing Then
                If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVLTBProdSchemeFS'][@HIGHBAND='" & strProdScheme & "']") Is Nothing Then
                    dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVLTBProdSchemeFS'][@HIGHBAND='" & strProdScheme & "']").Attributes.getNamedItem("PERCENTAGE").Text)
                End If
            Else
                If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVLTBProdSchemeSC'][@HIGHBAND='" & strProdScheme & "']") Is Nothing Then
                    dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVLTBProdSchemeSC'][@HIGHBAND='" & strProdScheme & "']").Attributes.getNamedItem("PERCENTAGE").Text)
                End If
            End If
            
        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
        
            Set xmlNode = xmlApplicationNode.selectSingleNode("./APPLICATIONDATA")

            If Not xmlNode.Attributes.getNamedItem("TYPEOFBUYER") Is Nothing Then
                strTypeOfBuyer = xmlNode.Attributes.getNamedItem("TYPEOFBUYER").nodeValue
                'Is it Remortgage
                Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='TypeOfBuyerNewLoan']/COMBOVALUE[@VALUEID='" & strTypeOfBuyer & "']/COMBOVALIDATION[@VALIDATIONTYPE='R']")
                
                If Not xmlNode Is Nothing Then
                    If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVBTLProdSchemeRmtg'][@HIGHBAND='" & strProdScheme & "']") Is Nothing Then
                        dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVBTLProdSchemeRmtg'][@HIGHBAND='" & strProdScheme & "']").Attributes.getNamedItem("PERCENTAGE").Text)
                    End If
                Else
                    If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVBTLProdScheme'][@HIGHBAND='" & strProdScheme & "']") Is Nothing Then
                        dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='MaxLTVBTLProdScheme'][@HIGHBAND='" & strProdScheme & "']").Attributes.getNamedItem("PERCENTAGE").Text)
                    End If
                End If
            End If
        End If
        
        'Calc LTV
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
    
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
        End If
    
        dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
        
        If dblLTV > dblBandAmount Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
   
    
        strRuleValue = "LTV: " & dblLTV
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
   End If
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Rule3695 = intResult
End Function
'======================================================================================
'End :DIP Rule 3695
'======================================================================================

Private Function Rule3700(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer
    
    Dim dblLTV As Double
    Dim dblAmountRequested As Double
    Dim dblPurchPriceEstVal As Double
    Dim dblMaxBTLRemortgageLTV As Double
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3700"
    
    'Rule 3700: LTV maximum exceeded for BTL remortgage
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF (ApplicationFactFind.NatureOfLoan not = "Buy-to-Let (Income)" AND not = "Buy-to-Let (Rental)" (combo "NatureOfLoan" ValidationType not = "BI" AND not = "BR")) OR ApplicationFactFind.TypeOfApplication not = "Remortgage" (combo "TypeOfMortgage" ValidationType = "R") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing Then
            
        dblMaxBTLRemortgageLTV = GetGlobalParameter(xmlApplicationNode, "MaxBTLRemortgageLTV", Percentage)
            
        'IF the Quotation is not found then calculate LTV = NewLoan.AmountRequested / ApplicationFactFind.PurchasePriceOrEstimatedValue * 100 (rounded to 2dp)
        intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
        If intMortgageSubquoteNumber = 0 Then
            If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
                dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
            End If
                
            If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
                dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
            End If
            
            
            dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
            
        'ELSE set LTV = MortgageSubQuote.LTV.
        Else
            Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("LTV") Is Nothing Then
                    dblLTV = xmlNode.Attributes.getNamedItem("LTV").nodeValue
                End If
            End If
        End If
        
        If dblLTV > dblMaxBTLRemortgageLTV Then
            blnSuccess = False
        End If
        
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
    
        strRuleValue = "LTV: " & dblLTV
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3700 = intResult
End Function

Private Function Rule3705( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String, _
        strPurpose As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3705"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3705: Remortgage where purpose is "Other"
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF ApplicationFactFind.TypeOfApplication not = "Remortgage" (combo "TypeOfMortgage" ValidationType not = "R") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing Then
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            'IF any LoanComponent has LoanComponent.SubPurposeOfLoan = "Other" (combo "SubPurposeOfLoan" ValidationType = "O") then [FAIL].
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("SUBPURPOSEOFLOAN_TYPE_O") Is Nothing Then
                    blnSuccess = False
                End If
            End If
        Next xmlNode
                
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3705 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3710( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String, _
        strPurpose As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3710"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3710: BTL remortgage where purpose not property-related
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF (ApplicationFactFind.NatureOfLoan not = "Buy-to-Let (Rental)" AND not = "Buy-to-Let (Income)" (combo "NatureOfLoan" ValidationType not = "BR" AND not = "BI")) OR ApplicationFactFind.TypeOfApplication not = "Remortgage" (combo "TypeOfMortgage" ValidationType not = "R") then do not run the rule.
    If (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing) And _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing Then
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            'IF any LoanComponent has LoanComponent.SubPurposeOfLoan not = a property-related purpose
            '(combo "SubPurposeOfLoan" ValidationType not = "PROP") then set purpose = combo "SubPurposeOfLoan" Value for the offending LoanComponent and [FAIL].
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("SUBPURPOSEOFLOAN") Is Nothing And _
                xmlNode.Attributes.getNamedItem("SUBPURPOSEOFLOAN_TYPE_PROP") Is Nothing Then
                    blnSuccess = False
                    strPurpose = xmlNode.Attributes.getNamedItem("SUBPURPOSEOFLOAN_TEXT").nodeValue
                    strRuleValue = "Purpose: " & strPurpose
                End If
            End If
        Next xmlNode
                
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3710 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3720( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String, _
        strMethod As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3720"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3720: Invalid repayment vehicle for IO residential
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF ApplicationFactFind.NatureOfLoan not = "Residential" AND not = "Let-to-Buy" (combo "NatureOfLoan" ValidationType not = "RS" AND not = "LT") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing Then
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I").
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0' and @REPAYMENTMETHOD_TYPE_I='true']")
            'IF any LoanComponent has LoanComponent.RepaymentVehicle = "Sale of Residential Property" OR = "Inheritance" OR = "Sale of Business"
            '(combo "RepaymentVehicle" ValidationType = "SRP" OR = "INH" OR = "SOB") then set method = combo "RepaymentVehicle" Value for the offending LoanComponent and [FAIL].
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_SRP") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_INH") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_SOB") Is Nothing Then
                    blnSuccess = False
                    strMethod = xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TEXT").nodeValue
                    strRuleValue = "Method: " & strMethod
                End If
            End If
        Next xmlNode
                
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3720 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3725( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String, _
        strMethod As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3725"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3725: Invalid repayment vehicle for IO BTL
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF ApplicationFactFind.NatureOfLoan not = "Buy-to-Let (Rental)" AND not = "Buy-to-Let (Income)" (combo "NatureOfLoan" ValidationType not = "BR" AND not = "BI") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Then
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I").
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0' and @REPAYMENTMETHOD_TYPE_I='true']")
            'IF any LoanComponent has LoanComponent.RepaymentVehicle = "Sale of Residential Property" OR = "Inheritance" OR = "Sale of Business" OR = "ISA" OR = "Pension" OR = "Endowment" OR = "Other"
            '(combo "RepaymentVehicle" ValidationType = "SRP" OR = "INH" OR = "SOB" OR = "I" OR = "P" OR = "E" OR = "N") then set method = combo "RepaymentVehicle" Value for the offending LoanComponent and [FAIL].
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_SRP") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_INH") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_SOB") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_I") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_P") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_E") Is Nothing Or _
                Not xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TYPE_N") Is Nothing Then
                    blnSuccess = False
                    strMethod = xmlNode.Attributes.getNamedItem("REPAYMENTVEHICLE_TEXT").nodeValue
                    strRuleValue = "Method: " & strMethod
                End If
            End If
        Next xmlNode
                
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3725 = intResult
    
    Set xmlNode = Nothing
        
End Function

'========================================================
'OS 29/11/2006
'========================================================

Private Function Rule3730( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlMemoPad As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String, _
        strMethod As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3730"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3730: "Additional Information" present
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    'Select the MemoPad nodeRead the MemoPad record where MemoPad.EntryType = "DIP Additional Info" (combo "MemoPadEntryType" ValidationType = "DAI").
    Set xmlMemoPad = xmlApplicationNode.selectSingleNode("./MEMOPAD[@ENTRYTYPE_TYPE_DAI = 'true']")
    
    'If "Additional Information" free-format notes field not = empty, then [FAIL]
    If Not xmlMemoPad Is Nothing Then
        If Not xmlMemoPad.Attributes.getNamedItem("MEMOENTRY") Is Nothing Then
            If xmlMemoPad.Attributes.getNamedItem("MEMOENTRY").Text <> "" Then
                blnSuccess = False
            End If
        End If
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3730 = intResult
    
    Set xmlNode = Nothing
        
End Function


Private Function Rule3735( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intPropertyDepositCount As Integer, _
        intPropertyDepositCount_GFT_OTH As Integer, _
        intRuleValidationType As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3735"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3735: Gifted deposit or unknown source
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
        
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT")
    intPropertyDepositCount = xmlNodeList.length
    
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT[@SOURCEOFFUNDING_TYPE_GFT = 'true' or @SOURCEOFFUNDING_TYPE_OTH = 'true']")
    intPropertyDepositCount_GFT_OTH = xmlNodeList.length
        
    'IF any NewPropertyDeposit record has NewPropertyDeposit.SourceOfFunding = "Gifted Deposit" OR = "Family Gift" OR = "Other" (combo "DepositSource" ValidationType = "GFT" OR = "OTH") then [FAIL].
    If intPropertyDepositCount > 0 And intPropertyDepositCount_GFT_OTH > 0 Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3735 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3740( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMaxBISDeposit As Integer
    
    Dim lngBISDepositTotalAmount As Long
        
    Dim dblThreshold As Double, _
        dblPurchasePrice As Double
    
    Dim strRuleValue As String, _
        strRuleDescription As String
       
    Dim blnSuccess As Boolean
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3740"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3740: Builder gifted deposit greater than maximum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True

    'Global parameter "MaxBISDeposit" = 5% that represents the maximum percentage of the valuation that can be from a builder's incentive scheme.
    intMaxBISDeposit = GetGlobalParameter(xmlApplicationNode, "MaxBISDeposit", Amount)
    
    If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
        dblPurchasePrice = xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").nodeValue
    End If
    
    dblThreshold = Round(dblPurchasePrice * intMaxBISDeposit / 100, 2)
    
    'Read all Builders Incentive Scheme deposit records for the case (see database notes)
    'Set BISDeposit = Sum NewPropertyDeposit.Amount for the records read
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT[@SOURCEOFFUNDING_TYPE_BIS = 'true']")
    For Each xmlNode In xmlNodeList
        If Not xmlNode.Attributes.getNamedItem("AMOUNT") Is Nothing Then
            lngBISDepositTotalAmount = lngBISDepositTotalAmount + xmlNode.Attributes.getNamedItem("AMOUNT").nodeValue
        End If
    Next xmlNode
    
    'IF BISDeposit > Threshold then [FAIL].
    If lngBISDepositTotalAmount > dblThreshold Then
        blnSuccess = False
    End If
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Deposit: £" & lngBISDepositTotalAmount & "; Valuation: £" & dblPurchasePrice
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3740 = intResult
        
End Function

Private Function Rule3745( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        dblMaxLTVAllowed As Double, _
        intPropertyDepositCount_BIS As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer

    Dim dblAmountRequested As Double, _
        dblPurchPriceEstVal As Double, _
        dblLTV As Double
        
    Dim strRuleValue As String, _
        strRuleDescription As String

        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3745"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3745: Deposit includes Builders Incentive and max LTV
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
  
    'Get max LTV allowed
    dblMaxLTVAllowed = GetGlobalParameter(xmlApplicationNode, "MaximumLTVAllowed", Percentage)
    
    'IF the Quotation is not found then calculate LTV = NewLoan.AmountRequested / ApplicationFactFind.PurchasePriceOrEstimatedValue * 100 (rounded to 2dp)
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    If intMortgageSubquoteNumber = 0 Then
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
            
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
        End If
        
        
        dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
        
    'ELSE set LTV = MortgageSubQuote.LTV.
    Else
        Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("LTV") Is Nothing Then
                dblLTV = xmlNode.Attributes.getNamedItem("LTV").nodeValue
            End If
        End If
    End If
        
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT[@SOURCEOFFUNDING_TYPE_BIS = 'true']")
    intPropertyDepositCount_BIS = xmlNodeList.length
        
    'IF LTV = global parameter "MaximumLTVAllowed" AND at least one "Builders Incentive Scheme" record is found then [FAIL].
    If intPropertyDepositCount_BIS > 0 And dblLTV = dblMaxLTVAllowed Then
        blnSuccess = False
    End If
    
    strRuleValue = "LTV: " & dblLTV
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3745 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule3750( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3750"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3750: Guarantor present on case
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    'IF the "Guarantor" CustomerRole record is found then [FAIL]
    Set xmlApplicant = xmlApplicationNode.selectSingleNode("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_G='true']")
    If Not xmlApplicant Is Nothing Then
        blnSuccess = False
    End If
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
        
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
          
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3750 = intResult
    
    Set xmlApplicant = Nothing
        
End Function

Private Function Rule3755(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlApplicant As IXMLDOMNode, _
        xmlAccountRelationships As IXMLDOMNode, _
        xmlMortgageAccounts As IXMLDOMNode, _
        xmlMortgageLoans As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
    
    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim dtPreEmptionEndDate As Date
    
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3755"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3755: Product switch with right-to-buy
    'This rule is only run for product switch; this restriction is handled by the calling component.
    'An existing right-to-buy that is now being switched can be identified by having a pre-emption end date.

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
   
    'Read each applicant for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        'FOR each MortgageAccount record where Account.IMPORTEDINDICATOR = true, read all related MortgageAccount records.
        For Each xmlAccountRelationships In xmlApplicant.selectNodes("CUSTOMERVERSION/ACCOUNTRELATIONSHIP")
            For Each xmlMortgageAccounts In xmlAccountRelationships.selectNodes("ACCOUNT[@IMPORTEDINDICATOR='1']")
                
                'IF the Account record is found then read all MortgageAccount records for the AccountGuid.
                For Each xmlMortgageLoans In xmlAccountRelationships.selectNodes("./MORTGAGEACCOUNT")

                    'IF MortgageAccount.PreEmptionEndDate not = null then [FAIL].
                    If Not xmlMortgageLoans.Attributes.getNamedItem("PREEMPTIONENDDATE") Is Nothing Then
                        blnSuccess = False
                        Set xmlApplicant = Nothing  'set xml to nothing so we exit loops
                        Set xmlMortgageLoans = Nothing
                        Set xmlMortgageAccounts = Nothing
                        Set xmlAccountRelationships = Nothing
                    End If
                Next xmlMortgageLoans
            Next xmlMortgageAccounts
        Next xmlAccountRelationships
    Next xmlApplicant
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
  
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3755 = intResult
    
    Set xmlApplicant = Nothing
    Set xmlMortgageLoans = Nothing
    Set xmlMortgageAccounts = Nothing
    Set xmlAccountRelationships = Nothing
        
End Function

Private Function Rule3760( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponentsNew As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3760"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3760: Total number of components greater than maximum
    'Multi-component loan and invalid number of components for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") then do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponentsNew = GetGlobalParameter(xmlApplicationNode, "MaxComponentsNew", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsNew" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponentsNew Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3760 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3765( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponentsFA As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3765"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3765: Total number of components greater than maximum
    'Multi-component loan and invalid number of components for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") then do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponentsFA = GetGlobalParameter(xmlApplicationNode, "MaxComponentsFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsFA" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponentsFA Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3765 = intResult

    Set xmlNode = Nothing

End Function
Private Function Rule3770( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponentsNew As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3770"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3770: Number of C&I components greater than maximum
    'Multi-component loan and number of C&I components greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponentsNew = GetGlobalParameter(xmlApplicationNode, "MaxComponentsCINew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsCINew" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponentsNew Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3770 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3775( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponents As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3775"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3775: Number of IO components greater than maximum

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponents = GetGlobalParameter(xmlApplicationNode, "MaxComponentsIONew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsIONew" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponents Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3775 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3780( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponents As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3780"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3780: Number of C&I components greater than maximum
    'Multi-component loan and number of C&I components greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponents = GetGlobalParameter(xmlApplicationNode, "MaxComponentsCIFA", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsCIFA" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponents Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3780 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3785( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponents As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3785"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3785: Number of IO components greater than maximum
    'Multi-component loan and number of IO components greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponents = GetGlobalParameter(xmlApplicationNode, "MaxComponentsIOFA", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsIOFA" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponents Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3785 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3790( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3790"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3790: Number of C&I terms greater than maximum
    'Multi-component loan and number of C&I terms greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsCINew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsCINew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3790 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3795( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3795"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3795: Number of IO terms greater than maximum
    'Multi-component loan and number of IO terms greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsIONew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsIONew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3795 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3800( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3800"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3800: Number of C&I terms greater than maximum
    'Multi-component loan and number of C&I terms greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsCIFA", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsCIFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3800 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3805( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3805"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3805: Number of IO terms greater than maximum
    'Multi-component loan and number of IO terms greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsIOFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsIOFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3805 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3810( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3810"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3810: Number of C&I products greater than maximum
    'Multi-component loan and number of C&I products greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsCINew", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsCINew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3810 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3815( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3815"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3815: Number of IO products greater than maximum
    'Multi-component loan and number of IO products greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsIONew", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsIONew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3815 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3820( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3820"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3820: Number of C&I products greater than maximum
    'Multi-component loan and number of C&I products greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsCIFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsCIFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3820 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3825( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3825"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3825: Number of IO products greater than maximum
    'Multi-component loan and number of IO products greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsIOFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsIOFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3825 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3830(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlApplicant As IXMLDOMNode, _
        xmlAccountRelationships As IXMLDOMNode, _
        xmlMortgageAccounts As IXMLDOMNode, _
        xmlMortgageLoans As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
    
    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim dblDisbursedAmount As Double
    Dim lngOriginalLoanAmount As Long
    
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3830"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3830: FA/CLI with outstanding retention
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
   
    'Read each applicant for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        'FOR each MortgageAccount record where Account.IMPORTEDINDICATOR = true, read all related MortgageLoan records.
        For Each xmlAccountRelationships In xmlApplicant.selectNodes("CUSTOMERVERSION/ACCOUNTRELATIONSHIP")
            For Each xmlMortgageAccounts In xmlAccountRelationships.selectNodes("ACCOUNT[@IMPORTEDINDICATOR='1']")
                
                'IF the Account record is found then read all MortgageLoan records for the AccountGuid.
                For Each xmlMortgageLoans In xmlAccountRelationships.selectNodes("./MORTGAGELOAN")
                    If Not xmlMortgageLoans.Attributes.getNamedItem("DISBURSEDAMOUNT") Is Nothing Then
                        dblDisbursedAmount = xmlMortgageLoans.Attributes.getNamedItem("DISBURSEDAMOUNT").nodeValue
                    End If
                    
                    If Not xmlMortgageLoans.Attributes.getNamedItem("ORIGINALLOANAMOUNT") Is Nothing Then
                        lngOriginalLoanAmount = xmlMortgageLoans.Attributes.getNamedItem("ORIGINALLOANAMOUNT").nodeValue
                    End If
                    
                    'IF any MortgageLoan has DisbursedAmount < OriginalLoanAmount then [FAIL].
                    If dblDisbursedAmount < lngOriginalLoanAmount Then
                        blnSuccess = False
                        Set xmlApplicant = Nothing  'set xml to nothing so we exit loops
                        Set xmlMortgageLoans = Nothing
                        Set xmlMortgageAccounts = Nothing
                        Set xmlAccountRelationships = Nothing
                    End If
                Next xmlMortgageLoans
            Next xmlMortgageAccounts
        Next xmlAccountRelationships
    Next xmlApplicant
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
    End If
    
    strRuleValue = "Disbursed: " & dblDisbursedAmount & "; Loan: " & lngOriginalLoanAmount
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
  
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3830 = intResult
    
    Set xmlApplicant = Nothing
    Set xmlMortgageLoans = Nothing
    Set xmlMortgageAccounts = Nothing
    Set xmlAccountRelationships = Nothing
        
End Function

Private Function Rule3840(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strHighRiskPolicyRule As String, _
        strHighRiskCodeDesc As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bHighRisk As Boolean, _
        bHighRiskFound As Boolean
    
    Dim xmlCustomer As IXMLDOMNode, _
        xmlComboHighPolicy As IXMLDOMNode, _
        xmlHighPolicy As IXMLDOMNode, _
        xmlNode As IXMLDOMNode, _
        xmlHighRiskCode As IXMLDOMNode
        
    Dim xmlHighPolicies As IXMLDOMNodeList, _
        xmlHighRiskCodes As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    
    'RULE 3840: Authenticate "High Risk" code
    gstrRuleNumber = "3840"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'apply this rule to each applicant
    For Each xmlCustomer In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If Not xmlCustomer Is Nothing Then
            ' Read all KnowYourCustomerHighRiskPolicyRules records for the
            ' KnowYourCustomerAuthenticateResults record.
            ' FOR each record read:
            ' IF HighRiskPolicyRule = ValidationType from any entry in combo "AuthHighRiskCodes" then [FAIL]
            
            'make sure we are searching against the latest knowyourcustomercheck record - use last()
            Set xmlHighPolicies = xmlCustomer.selectNodes("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS/KNOWYOURCUSTOMERHIGHRISKPOLICYRULES")
            For Each xmlHighPolicy In xmlHighPolicies
                If Not xmlHighPolicy Is Nothing Then
                    If Not xmlHighPolicy.Attributes.getNamedItem("HIGHRISKPOLICYRULE") Is Nothing Then
                        strHighRiskPolicyRule = xmlHighPolicy.Attributes.getNamedItem("HIGHRISKPOLICYRULE").nodeValue
                        Set xmlComboHighPolicy = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='AuthHighRiskCodes']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE='" & strHighRiskPolicyRule & "']")
                        If Not xmlComboHighPolicy Is Nothing Then
                            Set xmlNode = xmlComboHighPolicy.selectSingleNode("../@VALUENAME")
                            If Not xmlNode Is Nothing Then
                                'get SM value name
                                strHighRiskCodeDesc = xmlNode.Text
                            End If
                            
                            'search for rule validation type
                            Set xmlHighRiskCodes = xmlApplicationNode.selectNodes("./COMBOGROUP[@GROUPNAME='AuthHighRiskCodes']/COMBOVALUE[@VALUENAME='" & strHighRiskCodeDesc & "']/COMBOVALIDATION/@VALIDATIONTYPE")
                            For Each xmlHighRiskCode In xmlHighRiskCodes
                                'if DIP rule found
                                If InStr(1, xmlHighRiskCode.Text, "DIP_RULE_") Then
                                    'get DIP rule number only
                                    gstrRuleNumber = Mid(xmlHighRiskCode.Text, 10)
                                    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
                                    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
                                    bHighRiskFound = True
                                    
                                    'report fail for each high risk policy code found
                                    strRuleValue = gstrRuleNumber
                                    intScore = intRuleValidationType
                                    intResult = RiskAssessmentDecision.Referral
                                    CountReferral (intScore) 'add to appropriate referral count
                                    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
                                End If
                            Next xmlHighRiskCode
                                            
                        End If
                    End If
                End If
            Next xmlHighPolicy
        End If
    
    Next xmlCustomer
    
    'if a high risk has not been found
    If Not bHighRiskFound Then
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3840 = intResult
End Function

Private Function Rule3870(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim objApplicant As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bFail As Boolean
    Dim objNode As IXMLDOMNode
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3870"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' rule 3870
    ' Applicant authentication documentation required
    ' Run the rule for each applicant on the case.
    ' // A fail does not affect the overall DIP but is for information to the user.
    ' IF KnowYourCustomerAuthenticateResults.AuthDecision = "NA00" then [FAIL].
    
    For Each objNode In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If Not objNode.selectSingleNode("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS/@AUTHDECISION") Is Nothing Then
            If objNode.selectSingleNode("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS/@AUTHDECISION").Text = "NA00" Then
                bFail = True
                Exit For
            End If
        End If
    Next objNode
    
    If bFail Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3870 = intResult
End Function

Private Function Rule3875(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim objApplicant As IXMLDOMNode
    
    Dim xmlNode As IXMLDOMNode, _
        xmlSMReasonCode As IXMLDOMNode
        
    Dim xmlCreditCheckReasonCodes As IXMLDOMNodeList, _
        xmlSMReasonCodes As IXMLDOMNodeList
    
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strReasonCode As String, _
        strSMValueName As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intRuleNumberCount As Integer
    Dim intRuleCount As Integer
    
    Dim bSMReasonCodeFound As Boolean
    Dim bSMRuleFound As Boolean
    Dim bReturnRuleResult As Boolean
    
    Dim strSMRuleNumbers() As String
    
    'RULE 3875: unacceptable strategy manager code
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3875"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
        
    'get most recent application credit check record
    Set xmlCreditCheckReasonCodes = xmlApplicationNode.selectNodes("APPLICATIONCREDITCHECK[1]/CREDITCHECKREASONCODE")
    For Each xmlNode In xmlCreditCheckReasonCodes
        'get reason code
        If Not xmlNode.Attributes.getNamedItem("REASONCODE") Is Nothing Then
            strReasonCode = xmlNode.Attributes.getNamedItem("REASONCODE").nodeValue
        End If
        
        'IF ReasonCode = ValidationType from any entry in combo "SMReasonCode"
        'AND combo entry also has ValidationType = "Refer" then
        Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='SMReasonCode']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE='" & strReasonCode & "']")
        If Not xmlNode Is Nothing Then
            Set xmlNode = xmlNode.selectSingleNode("../COMBOVALIDATION[@VALIDATIONTYPE='" & "Refer" & "']")
            If Not xmlNode Is Nothing Then
                Set xmlNode = xmlNode.selectSingleNode("../@VALUENAME")
                If Not xmlNode Is Nothing Then
                    'get SM value name
                    strSMValueName = xmlNode.Text
                End If
                
                'search for rule validation type
                Set xmlSMReasonCodes = xmlApplicationNode.selectNodes("./COMBOGROUP[@GROUPNAME='SMReasonCode']/COMBOVALUE[@VALUENAME='" & strSMValueName & "']/COMBOVALIDATION/@VALIDATIONTYPE")
                For Each xmlSMReasonCode In xmlSMReasonCodes
                    'return the rule result in the XML as default
                    bReturnRuleResult = True
                    
                    'if DIP rule found
                    If InStr(1, xmlSMReasonCode.Text, "DIP_RULE_") Then
                        'get DIP rule number only
                        gstrRuleNumber = Mid(xmlSMReasonCode.Text, 10)
                        
                        If intRuleNumberCount > 0 Then
                            For intRuleCount = 1 To UBound(strSMRuleNumbers)
                                If strSMRuleNumbers(intRuleCount) = gstrRuleNumber Then
                                    bReturnRuleResult = False
                                End If
                            Next intRuleCount
                        End If
            
                        'only return the rule in the XML if it hasn't already been returned
                        If bReturnRuleResult Then
                            
                            strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
                            intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
                            bSMReasonCodeFound = True
                            
                            'report fail for each sm reason code found
                            strRuleValue = gstrRuleNumber
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
                            
                            'increment rule number array counter
                            intRuleNumberCount = intRuleNumberCount + 1
                            
                            ReDim Preserve strSMRuleNumbers(1 To intRuleNumberCount) As String
                            'add the rule number to the array.  This indicates the rule has been returned in the XML.
                            strSMRuleNumbers(intRuleNumberCount) = gstrRuleNumber
                            
                        End If
                    End If
                Next xmlSMReasonCode
                
            End If
        End If
    Next xmlNode
    
    If Not bSMReasonCodeFound Then
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3875 = intResult
End Function

Private Function Rule3920(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim dCurrentAffordabilityIndex As Double
    Dim dMinAffordabilityIndex As Double
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3920"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' Rule 3920: Affordability Index below minimum
    ' IF BXApplicationBureau.RRP2_CurrentAffordabilityIndex < global parameter "MinAffordabilityIndex" then
    '  [FAIL]
    
    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BXAPPLICATIONBUREAU/@RRP2_CURRENTAFFORDABILITYINDEX") Is Nothing Then
        dMinAffordabilityIndex = CDbl(GetGlobalParameter(xmlApplicationNode, "MinAffordabilityIndex", Amount))
        If Not xmlApplicationNode.selectSingleNode("./GLOBALPARAMETER[@NAME=""MinAffordabilityIndex""]") Is Nothing Then
        
            If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BXAPPLICATIONBUREAU/@RRP2_CURRENTAFFORDABILITYINDEX") Is Nothing Then
                dCurrentAffordabilityIndex = CDbl(xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BXAPPLICATIONBUREAU/@RRP2_CURRENTAFFORDABILITYINDEX").Text)
            End If
            
            If dCurrentAffordabilityIndex < dMinAffordabilityIndex Then
                intScore = intRuleValidationType
                strRuleValue = "AI: " & dCurrentAffordabilityIndex
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
            End If
        End If
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3920 = intResult
End Function

Private Function Rule3925(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim blnSuccess As Boolean
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer
    Dim intGroupSequenceNumber As Integer
    
    Dim strValidationType As String
    Dim strEligInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3925"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' Rule 3925: Product cascade
    
    'EP2_876
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'Locate any LoanComponent record where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
    Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
    If Not xmlNode Is Nothing Then
        'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber.
        Set xmlNode = xmlNode.selectSingleNode("SPECIALGROUP")
        If Not xmlNode Is Nothing Then
            'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber
            If Not xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER") Is Nothing Then
                intGroupSequenceNumber = xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER").nodeValue
                Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='SpecialGroup']/COMBOVALUE[@VALUEID='" & intGroupSequenceNumber & "']")
                If Not xmlNode Is Nothing Then
                    Set xmlNode = xmlNode.selectSingleNode("COMBOVALIDATION")
                    If Not xmlNode.Attributes.getNamedItem("VALIDATIONTYPE") Is Nothing Then
                        strValidationType = xmlNode.Attributes.getNamedItem("VALIDATIONTYPE").nodeValue
                    End If
                End If
            End If
        End If
    End If
    
    'IF ValidationType not = BespokeDecision.EligibilityIndicator then [FAIL].
    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If
    
    'If the special group does not match the value determined by Strategy Manager then the rule fails. Check to see if the product can potentially be cascaded.
    If strValidationType <> strEligInd Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
            
        'If other rules have failed
        If gintProcessorCounter > 0 Or gintUnderWriterCounter > 0 Then
            'Another rule has failed; therefore, cascading is not appropriate.
            'exit the cascade process
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        'If no other rules have failed
        Else
            'continue with cascade
            intResult = RiskAssessmentDecision.AcceptCascade
        End If
        
        intScore = intRuleValidationType

    End If
    
    strRuleValue = "Requested: " & strValidationType & ", Strat Man: " & strEligInd
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3925 = intResult
End Function

Private Function Rule3930( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim dblTotalAllowableIncome As Double, _
        dblTotalApproximateMonthlyCost As Double
    
    Dim blnSuccess As Boolean
        
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3930"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 3930: Total Allowable Income < Total Approximate Monthly Cost
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
        
    'Please note: totalallowableincome & total approximatemonthlycost are calculated fields called from omRA
    If Not xmlApplicationNode.Attributes.getNamedItem("TOTALALLOWABLEINCOME") Is Nothing Then
        dblTotalAllowableIncome = xmlApplicationNode.Attributes.getNamedItem("TOTALALLOWABLEINCOME").nodeValue
    End If
    
    If Not xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST") Is Nothing Then
        dblTotalApproximateMonthlyCost = xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST").nodeValue
    End If
    
    'IF Passed in "TotalAllowableIncome" < "TotalApproximateMonthlyCost" THEN Fail
    If dblTotalAllowableIncome < dblTotalApproximateMonthlyCost Then
        blnSuccess = False
    End If

    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule3930 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule3935(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim blnSuccess As Boolean
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer
    Dim intGroupSequenceNumber As Integer
    
    Dim strValidationType As String
    Dim strEligInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3935"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' Rule 3935: Product cascade
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'Locate any LoanComponent record where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
    Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
    If Not xmlNode Is Nothing Then
        'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber.
        Set xmlNode = xmlNode.selectSingleNode("SPECIALGROUP")
        If Not xmlNode Is Nothing Then
            'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber
            If Not xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER") Is Nothing Then
                intGroupSequenceNumber = xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER").nodeValue
                Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='SpecialGroup']/COMBOVALUE[@VALUEID='" & intGroupSequenceNumber & "']")
                If Not xmlNode Is Nothing Then
                    Set xmlNode = xmlNode.selectSingleNode("COMBOVALIDATION")
                    If Not xmlNode.Attributes.getNamedItem("VALIDATIONTYPE") Is Nothing Then
                        strValidationType = xmlNode.Attributes.getNamedItem("VALIDATIONTYPE").nodeValue
                    End If
                End If
            End If
        End If
    End If
    
    'IF ValidationType not = BespokeDecision.EligibilityIndicator then [FAIL].
    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If
    
    'If the special group does not match the value determined by Strategy Manager then the rule fails.
    If strValidationType <> strEligInd Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
        intScore = intRuleValidationType
    End If
    
    strRuleValue = "Requested: " & strValidationType & ", Strat Man: " & strEligInd
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3935 = intResult
End Function

'OS 16/04/2007 - EP2_2427
'AShaw 20/04/2007 - EP2_1506

Private Function Rule3940( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicants As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlIncome As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAgeEndTermLow As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intApplicantsAge As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intNumberOfApplicants As Integer, _
        intApplicantCount As Integer, _
        intApplicantFailCount As Integer, _
        intCurrentApplicantCalcAge As Integer, _
        intApplicantMaxAge As Integer
        
    Dim dblAllowableIncome As Double, _
        dblTotalApproximateMonthlyCost As Double
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess As Boolean, _
        blnOneApplicantFailed As Boolean, _
        bDOBExists As Boolean
        
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "3940"
    strScoreCardInd = 0 'False
    'Set default value for Allowable income
    dblAllowableIncome = 0
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 3940: Age at end of term beyond residential maximum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    blnOneApplicantFailed = False
    bDOBExists = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    Set xmlApplicants = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    intNumberOfApplicants = xmlApplicants.length
    
    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType = "BR") OR the Quotation is not found then do not run the rule
    If xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing And _
        intMortgageSubquoteNumber > 0 Then
        
        'Get TotalApproximateMonthlyCost
        If Not xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST") Is Nothing Then
            dblTotalApproximateMonthlyCost = CDbl(xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST").Text)
        End If
        
        'Get lower limit for loan term
        intMaxAgeEndTermLow = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBLo", Amount)
        
        
        'Locate all LoanComponent records where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        'Calculate and temporarily store the biggest term (in years) across all new loan components.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
            End If
        Next xmlNode

        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                'Get customer number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                    strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
                
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'IF this is a web-ingested case then calculate age as a system date (in whole years rounded down) using CustomerVersion.DateOfBirth
                'else calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded down) using CustomerVersion.DateOfBirth
                dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
                    
                intApplicantCount = intApplicantCount + 1
                
                ' EP2_2056 - Always use Max Age in Rule text.
                intCurrentApplicantCalcAge = intApplicantsAge + intMaxTermInYears
                If intCurrentApplicantCalcAge > intApplicantMaxAge Then
                    intApplicantMaxAge = intCurrentApplicantCalcAge
                End If
                
                
                'IF age (at end of term) > global parameter "LPMaxAgeEndTermResLTBLo" then [FAIL]
                If (intCurrentApplicantCalcAge) > intMaxAgeEndTermLow Then
                    If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERROLETYPE_TYPE_G") Is Nothing Then
                        blnSuccess = False
                        Exit For
                    ElseIf intApplicantFailCount = 1 Then 'both applicants failed, fail and exit loop
                        intApplicantFailCount = intApplicantFailCount + 1
                        blnSuccess = False
                        Exit For
                    Else
                        'One of the applicants has met the conditions of the age check
                        blnOneApplicantFailed = True
                        intApplicantFailCount = intApplicantFailCount + 1
                    End If
                End If
                
                'if one applicant failed and we've processed all applicants and/or guarantor
                If blnOneApplicantFailed And intApplicantCount = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']").length Then
                    'An applicant has not met the conditions of the age check
                    'Store Allowable income for this applicant
                    If Not xmlNode Is Nothing _
                    And Not xmlApplicant.Attributes.getNamedItem("CUSTOMERROLETYPE_TYPE_A") Is Nothing Then
                        'Calculate income for applicant
                        dblAllowableIncome = 0
                        
                        'Get income from NETCONFIRMEDALLOWABLEINCOME attribute
                        Set xmlIncome = xmlNode.selectSingleNode("./INCOMESUMMARY[@NETCONFIRMEDALLOWABLEINCOME]")
                        If Not xmlIncome Is Nothing Then
                            dblAllowableIncome = CDbl(xmlIncome.Attributes.getNamedItem("NETCONFIRMEDALLOWABLEINCOME").Text)
                        End If
                        
                        'Get income from NETALLOWABLEANNUALINCOME attribute
                        If dblAllowableIncome = 0 Then
                            Set xmlIncome = xmlNode.selectSingleNode("./INCOMESUMMARY[@NETALLOWABLEANNUALINCOME]")
                            If Not xmlIncome Is Nothing Then
                                dblAllowableIncome = CDbl(xmlIncome.Attributes.getNamedItem("NETALLOWABLEANNUALINCOME").Text)
                            End If
                        End If
                    End If
                End If
        
            End If
        Next xmlApplicant
        
        ' EP2_2506 - Always use Maximum age if DOB entered (can be overridden below if ANY DOB missing.
        strRuleValue = "Age: " & intApplicantMaxAge
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        ElseIf blnOneApplicantFailed Then                 ' EP2_2506
            blnSuccess = False
        End If
        
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
        End If
        
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule3940 = intResult
    
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlApplicants = Nothing
    Set xmlApplicant = Nothing
    Set xmlIncome = Nothing
        
End Function


Private Function Rule6500(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
       
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim attDOB As MSXML2.IXMLDOMAttribute
    Dim dtmDOB As Date
    Dim dtmToday As Date
    Dim iMaxAgeAtApplication As Integer
    Dim iAge As Integer
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6500"
    'strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
        
    If Not (xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing) Or _
       Not (xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing) Or _
       Not (xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing) Then
        
        iMaxAgeAtApplication = CInt(GetGlobalParameter(xmlApplicationNode, "MaxAgeAtApplication", Amount))
        
        dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
        
        For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
            Set attDOB = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION").Attributes.getNamedItem("DATEOFBIRTH")
        
            If Not (attDOB Is Nothing) Then
            
                dtmDOB = attDOB.Text
                iAge = DateDiffInYearsRoundedUp(dtmDOB, dtmToday)
                
                If iAge >= iMaxAgeAtApplication Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    strRuleValue = "Age: " & CStr(iAge)
                    Exit For 'failed so stop processing
                End If
            
            Else
                LogDetails 1, String(4, Chr(9)) & "Missing attribute DATEOFBIRTH on CUSTOMERVERSION element, return a fail"
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
            End If
        
        Next
     
    End If
     
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6500 = intResult
    
End Function
Private Function Rule6505( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAge As Integer, _
        intApplicantsAge As Integer, _
        intMinimumAgeRes As Integer, _
        intMinimumAgeBTL As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strRuleDescription As String

    Dim dtApplicationDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6505"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6505: Age now less than minimum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True

    'Get Minimum Age Residential
    intMinimumAgeRes = CSng(GetGlobalParameter(xmlApplicationNode, "MinimumAgeRes", Amount))
    
    'Get Minimum Age Buy-To-Let
    intMinimumAgeBTL = CSng(GetGlobalParameter(xmlApplicationNode, "MinimumAgeBTL", Amount))
    
    'Get application received date
    If Not xmlApplicationNode.Attributes.getNamedItem("APPLICATIONRECEIVEDDATE") Is Nothing Then
        dtApplicationDate = xmlApplicationNode.Attributes.getNamedItem("APPLICATIONRECEIVEDDATE").nodeValue
    'IF ApplicationFactFind.ApplicationReceivedDate is null, use System Date to calculate age.
    Else
        dtApplicationDate = Now
    End If
                   
    'Read each applicant/guarantor for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If Not xmlApplicant Is Nothing Then
            'get customer version number
            If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
            End If
            
            Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']")
            blnSuccess = False
            If Not xmlNode Is Nothing Then
                'get customers date of birth
                'LDM 08/03/2006 EP2_1734
                If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                    If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                        dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                        blnSuccess = True
                    End If
                End If
            End If
            
            If Not blnSuccess Then
                bDOBExists = False
                Exit For
            End If
            
            'Calculate applicant's age at application (rounded down)
            intApplicantsAge = DateDiffInYearsRoundedDown(dtDateOfBirth, dtApplicationDate)
            
            'IF ApplicationFactFind.NatureOfLoan = "Residential" OR = "Let-to-Buy" (combo "NatureOfLoan" ValidationType = "RS" OR = "LT") AND any customer's age < global parameter "MinimumAgeRes" then [FAIL].
            If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Or _
            Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing Then
                If intApplicantsAge < intMinimumAgeRes Then
                    blnSuccess = False
                    Exit For
                End If
            'ELSE IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let" (combo "NatureOfLoan" ValidationType = "BI" OR = "BR") AND any customer's age < global parameter "MinimumAgeBTL" then [FAIL].
            ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
            Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
                If intApplicantsAge < intMinimumAgeBTL Then
                    blnSuccess = False
                    Exit For
                End If
            End If
        End If
    Next xmlApplicant
    

    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.DeclinePolicy
    End If
    
    If Not bDOBExists Then
        strRuleValue = "No DOB"
    Else
        strRuleValue = "Age: " & intApplicantsAge
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6505 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6510( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAgeEndTerm As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intApplicantsAge As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strRuleDescription As String

    Dim dtDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6510"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6510: Age beyond maximum limit
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True
   
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let - Rental" (combo "NatureOfLoan" ValidationType = "BR") OR the Quotation is not found then do not run the rule
    If xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing And _
        intMortgageSubquoteNumber > 0 Then
        
        'Get max age for loan term
        intMaxAgeEndTerm = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeEndTermResLTBHi", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
            End If
        Next xmlNode

        
        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']")
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    'LDM 08/03/2006 EP2_1734
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'IF this is a web-ingested case then calculate age as a system date (in whole years rounded down) using CustomerVersion.DateOfBirth
                'else calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole years rounded down) using CustomerVersion.DateOfBirth
                dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtDate)
                    
                'IF age (at end of term) >= global parameter "LPMaxAgeEndTermResLTBHi" then [FAIL]
                If (intApplicantsAge + intMaxTermInYears) >= intMaxAgeEndTerm Then
                    blnSuccess = False
                    Exit For
                End If
        
            End If
        Next xmlApplicant
        
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        Else
            strRuleValue = "Age: " & intApplicantsAge + intMaxTermInYears
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6510 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6515( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intMaxAge As Integer, _
        intApplicantsAge As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strRuleDescription As String

    Dim dtApplicationDate As Date, _
        dtDateOfBirth As Date
        
    Dim blnSuccess, _
        bDOBExists As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6515"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6515: Age beyond BTL maximum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    bDOBExists = True
  
    'if not BTL rental based then don't run this rule
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
        
        'Get max age for loan term
        intMaxAge = GetGlobalParameter(xmlApplicationNode, "LPMaxAgeBTL", Amount)
        
        'Get application received date
        If Not xmlApplicationNode.Attributes.getNamedItem("APPLICATIONRECEIVEDDATE") Is Nothing Then
            dtApplicationDate = xmlApplicationNode.Attributes.getNamedItem("APPLICATIONRECEIVEDDATE").nodeValue
        End If
                       
        'Read each applicant/guarantor for the application
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']")
                blnSuccess = False
                If Not xmlNode Is Nothing Then
                    'get customers date of birth
                    'LDM 08/03/2006 EP2_1734
                    If Not xmlNode.Attributes.getNamedItem("DATEOFBIRTH") Is Nothing Then
                        If Len(xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue) > 0 Then
                            dtDateOfBirth = xmlNode.Attributes.getNamedItem("DATEOFBIRTH").nodeValue
                            blnSuccess = True
                        End If
                    End If
                End If
                
                If Not blnSuccess Then
                    bDOBExists = False
                    Exit For
                End If
                
                'Calculate applicant's age at application (rounded up)
                'DateDiffInYears rounds down, so added 1 to the result to round up in years
                intApplicantsAge = DateDiffInYearsRoundedUp(dtDateOfBirth, dtApplicationDate)
                
                'IF age at application > global parameter "LPMaxAge" then [FAIL]
                If intApplicantsAge > intMaxAge Then
                    blnSuccess = False
                    Exit For
                End If
            End If
        Next xmlApplicant
        
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        If Not bDOBExists Then
            strRuleValue = "No DOB"
        Else
            strRuleValue = "Age: " & intApplicantsAge
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    End If
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6515 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6520( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intLoanTermInYears As Integer, _
        intLoanTermInMonths As Integer, _
        intRuleValidationType As Integer, _
        intTempTermInYears As Integer, _
        intMaxTermInYears As Integer, _
        intMinTermInYears As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMinimumTerm As Integer, _
        intMaximumTerm As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6520"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6520: Term out of range
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF the Quotation is not found then do not run the rule
    If intMortgageSubquoteNumber > 0 Then
        
        intMinTermInYears = 99
        
        'Get max term
        intMaximumTerm = GetGlobalParameter(xmlApplicationNode, "MaximumTerm", Amount)
        
        'Get min term
        intMinimumTerm = GetGlobalParameter(xmlApplicationNode, "MinimumTerm", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        For Each xmlNode In xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                    intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
                End If
            
                If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                    intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
                End If
                
                intTempTermInYears = intLoanTermInYears
                
                If intLoanTermInMonths > 0 Then
                    intTempTermInYears = intTempTermInYears + 1
                End If
                
                'store the maximum term found so far...
                If intTempTermInYears > intMaxTermInYears Then
                    intMaxTermInYears = intTempTermInYears
                End If
            
                'store the minimum term found so far...
                If intTempTermInYears < intMinTermInYears Then
                    intMinTermInYears = intTempTermInYears
                End If
                
            End If
        Next xmlNode
                  
        'When all loan components have been processed, check the smallest and biggest term.
        'IF MinTerm < global parameter "MinimumTerm" then set TempTerm = MinTerm and [FAIL]. ELSE IF MaxTerm > global parameter
        If intMinTermInYears < intMinimumTerm Then
            intTempTermInYears = intMinTermInYears
            blnSuccess = False
        ElseIf intMaxTermInYears > intMaximumTerm Then
            intTempTermInYears = intMaxTermInYears
            blnSuccess = False
        End If
    
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
        
        strRuleValue = "Term (whole years): " & intTempTermInYears
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6520 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6525(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim lstCustomerAddress As MSXML2.IXMLDOMNodeList
    Dim attMovedOut As MSXML2.IXMLDOMAttribute
    Dim attMovedIn As MSXML2.IXMLDOMAttribute
    Dim iCustomerAddress As Integer
    
    Dim iAddressValidationYears As Integer
    Dim iMaxAddressGap As Integer
    Dim dtmMovedOut As Date
    Dim dtmMovedIn As Date
    Dim dtmToday As Date
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6525"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iAddressValidationYears = CInt(GetGlobalParameter(xmlApplicationNode, "AddressValidationYears", Amount))
    iMaxAddressGap = CInt(GetGlobalParameter(xmlApplicationNode, "MaxAddressGap", Amount))

    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        Set lstCustomerAddress = ndCustomerRole.selectNodes("./CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H='true' or @ADDRESSTYPE_TYPE_P='true']")
    
        If Not (lstCustomerAddress Is Nothing) Then
            
            'Get [date moved out] for the most recent CUSTOMERADDRESS
            
            If Not lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                dtmMovedIn = CDate(lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDIN").Text)
            Else
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATEMOVEDIN missing for Customer EMPLOYMENT, return a fail"
                Exit For
            End If
            
            If Not lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDOUT") Is Nothing Then
                dtmMovedOut = CDate(lstCustomerAddress.Item(lstCustomerAddress.length - 1).Attributes.getNamedItem("DATEMOVEDOUT").Text)
            End If
            
            'Is there < AddressValidationYears of history between today and earliest movedin date?
            If TrueDateDiff("yyyy", dtmMovedIn, dtmToday) < iAddressValidationYears Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                Exit For
            Else
                'Iterate Previous and Current CUSTOMERADDRESS starting from last but one node
                iCustomerAddress = IIf((lstCustomerAddress.length > 1), (lstCustomerAddress.length - 2), 0)
                
                Do While iCustomerAddress >= 0
                
                    If Not lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                        dtmMovedIn = CDate(lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDIN").Text)
                    Else
                        intScore = intRuleValidationType
                        intResult = RiskAssessmentDecision.Referral
                        CountReferral (intScore)
                        LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATEMOVEDIN missing for Customer EMPLOYMENT, return a fail"
                        Exit For
                    End If
                    
                    If dtmMovedOut > 0 Then 'ie if dtmMovedOut has been found (won't exist with a current address)
                        If TrueDateDiff("m", dtmMovedOut, dtmMovedIn) > iMaxAddressGap Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            Exit For
                        End If
                    End If
                    
                    Set attMovedOut = lstCustomerAddress.Item(iCustomerAddress).Attributes.getNamedItem("DATEMOVEDOUT")
                    If Not (attMovedOut Is Nothing) Then
                        dtmMovedOut = CDate(attMovedOut.Text)
                    Else
                        If iCustomerAddress <> 0 Then
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            LogDetails 1, String(4, Chr(9)) & "Required attribute missing: DateMovedOut, return a fail"
                        End If
                    End If
                    
                    iCustomerAddress = (iCustomerAddress - 1)
                Loop
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6525 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Set ndCustomerRole = Nothing
    Set lstCustomerAddress = Nothing
    
End Function


'========================================================
'OS 02/12/2006
'========================================================

Private Function Rule6530( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlApplicants As IXMLDOMNodeList
    
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlEmployment As IXMLDOMNode
    Dim xmlEarnedIncome As IXMLDOMNode
    Dim xmlAllowableIncomeFactor As IXMLDOMNode
    Dim xmlNetProfit As IXMLDOMNode
    Dim xmlAccountantsReference As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intNumberOfApplicants As Integer
        
    Dim dblTotalIncome As Double, _
        dblThisIncome As Double, _
        dblFactor As Double, _
        dblMinIncomeSingle As Double, _
        dblYearEnd1 As Double
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String, _
        strEmploymentStatus As String, _
        strIncomeGroup As String, _
        strEarnedIncomeType As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6530"
    strScoreCardInd = 0 'False
    dblTotalIncome = 0
    dblThisIncome = 0
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6530: Single app; income below minimum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    Set xmlApplicants = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    intNumberOfApplicants = xmlApplicants.length
    
    'If Number of "applicants" is not one, then do not run the rule
    If intNumberOfApplicants = 1 Then
        'Get value for global parameter MinIncomeSingle
        dblMinIncomeSingle = CDbl(GetGlobalParameter(xmlApplicationNode, "MinIncomeSingle", Amount))
        
        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                
                dblTotalIncome = 0
                
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                'Get customer number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                    strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
                End If
                
                Set xmlEmployment = xmlApplicant.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
                
                If Not xmlEmployment Is Nothing Then
                    
                    'LDM 21/03/07 EP2_1852  if employed or employed-temp then look at confirmed income first then earned income records
                    If Not xmlEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']") Is Nothing Then
                    
                        If xmlEmployment.selectNodes("./CONFIRMEDEARNEDINCOME").length <> 0 Then
                            'Iterate through each CONFIRMEDEARNEDINCOME Record
                            For Each xmlEarnedIncome In xmlEmployment.selectNodes("./CONFIRMEDEARNEDINCOME")
                                'Combo values
                                strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_EMP"
                                strIncomeGroup = "INCOMEGROUP_TYPE_EC"
                                If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE") Is Nothing Then
                                    strEarnedIncomeType = xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE").Text
                                End If
                                
                                'Get ALLOWABLEINCOMEFACTOR record according to specified criteria
                                Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @TYPE='" & strEarnedIncomeType & "']")
                                
                                If Not xmlAllowableIncomeFactor Is Nothing Then
                                    If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT") Is Nothing _
                                    And Not xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE") Is Nothing _
                                    And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                        'Calculate Total Income
                                        dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                        dblThisIncome = CDbl(xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT").Text) * _
                                                        CDbl(xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE").Text)
                                        dblTotalIncome = dblTotalIncome + (dblThisIncome * dblFactor / 100)
                                    End If
                                End If
                            Next xmlEarnedIncome
                            
                        ElseIf xmlEmployment.selectNodes("./EARNEDINCOME").length <> 0 Then
                            
                            'Iterate through each EARNEDINCOME Record
                            For Each xmlEarnedIncome In xmlEmployment.selectNodes("./EARNEDINCOME")
                                'Combo values
                                strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_EMP"
                                strIncomeGroup = "INCOMEGROUP_TYPE_ED"
                                If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE") Is Nothing Then
                                    strEarnedIncomeType = xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE").Text
                                End If
                                
                                'Get ALLOWABLEINCOMEFACTOR record according to specified criteria
                                Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @TYPE='" & strEarnedIncomeType & "']")
                                
                                If Not xmlAllowableIncomeFactor Is Nothing Then
                                    If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT") Is Nothing _
                                    And Not xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE") Is Nothing _
                                    And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                        'Calculate Total Income
                                        dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                        dblThisIncome = CDbl(xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT").Text) * _
                                                        CDbl(xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE").Text)
                                        dblTotalIncome = dblTotalIncome + (dblThisIncome * dblFactor / 100)
                                    End If
                                End If
        
                            Next xmlEarnedIncome
                        
                        End If
                    
                    'LDM 21/03/07 EP2_1852 if self employed or contractor then look at accountants reference first then net profit records
                    ElseIf Not xmlEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_SELF='true' or @EMPLOYMENTSTATUS_TYPE_CON='true']") Is Nothing Then

                        If xmlEmployment.selectNodes("./ACCOUNTANTSREFERENCE").length > 0 Then
                            
                            Set xmlAccountantsReference = xmlEmployment.selectSingleNode("./ACCOUNTANTSREFERENCE")
                            
                            'Combo values
                            strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_SELF"
                            strIncomeGroup = "INCOMEGROUP_TYPE_SC"
                            strEarnedIncomeType = "TYPE_TYPE_ANP"
                            
                            'Select a suitable ALLOWABLEINCOMEFACTOR record
                            Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @" & strEarnedIncomeType & "]")
                            
                            If Not xmlAllowableIncomeFactor Is Nothing Then
                                If Not xmlAccountantsReference.Attributes.getNamedItem("NETPROFITENTITLEMENT1") Is Nothing _
                                And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                    'Calculate Total Income
                                    dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                    dblYearEnd1 = CDbl(xmlAccountantsReference.Attributes.getNamedItem("NETPROFITENTITLEMENT1").Text)
                                    dblTotalIncome = dblYearEnd1 * dblFactor / 100
                                End If
                            End If
                                                    
                        Else
                            'No EarnedIncome records found
                            'Combo values
                            strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_SELF"
                            strIncomeGroup = "INCOMEGROUP_TYPE_SD"
                            strEarnedIncomeType = "TYPE_TYPE_ANP"
                            
                            'Select a suitable ALLOWABLEINCOMEFACTOR record
                            Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @" & strEarnedIncomeType & "]")
                                            
                            'Get the NETPROFIT record
                            Set xmlNetProfit = xmlEmployment.selectSingleNode("NETPROFIT")
                            
                            If Not xmlAllowableIncomeFactor Is Nothing _
                            And Not xmlNetProfit Is Nothing Then
                                If Not xmlNetProfit.Attributes.getNamedItem("YEAR1AMOUNT") Is Nothing _
                                And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                    'Calculate Total Income
                                    dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                    dblTotalIncome = CDbl(xmlNetProfit.Attributes.getNamedItem("YEAR1AMOUNT").Text) * dblFactor / 100
                                End If
                                
                            End If
                        End If
                        
                    End If
                
                End If
            End If
        
            'Rule fails if Total Income is less than global parameter MinIncomeSingle
            If dblTotalIncome < dblMinIncomeSingle Then
                blnSuccess = False
                Exit For
            End If
        
        Next xmlApplicant
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
        
        strRuleValue = "Income: " & dblTotalIncome
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6530 = intResult
        
End Function

'========================================================
'OS 04/12/2006
'AShaw 20/04/2007 - EP2_2451
'========================================================

Private Function Rule6535( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlApplicants As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlEmployment As IXMLDOMNode
    Dim xmlEarnedIncome As IXMLDOMNode
    Dim xmlAllowableIncomeFactor As IXMLDOMNode
    Dim xmlNetProfit As IXMLDOMNode
    Dim xmlAccountantsReference As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intNumberOfApplicants As Integer
        
    Dim dblTotalIncomeApplicant As Double, _
        dblThisIncome As Double, _
        dblTotalIncomeGuarantor As Double, _
        dblFactor As Double, _
        dblMinIncomeSingle As Double, _
        dblMinIncomeJoint As Double, _
        dblYearEnd1 As Double
        
    Dim strRuleValue As String, _
        strCustomerVersionNumber As String, _
        strCustomerNumber As String, _
        strRuleDescription As String, _
        strEmploymentStatus As String, _
        strIncomeGroup As String, _
        strEarnedIncomeType As String
        
    Dim blnSuccess As Boolean, _
        blIsApplicant As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6535"
    strScoreCardInd = 0 'False
    dblTotalIncomeApplicant = 0
    'Gaurantor's total income initialized to -1
    'If a Gaurantor exists, then this will be set to zero or more
    dblTotalIncomeGuarantor = -1
    
    dblThisIncome = 0
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6535: Joint app; income below minimum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    Set xmlApplicants = xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true']")
    intNumberOfApplicants = xmlApplicants.length
    
    
    'If Number of applicants is one, then do not run the rule
    If intNumberOfApplicants > 1 Then
    
        'Get value for global parameter MinIncomeSingle
        dblMinIncomeSingle = CDbl(GetGlobalParameter(xmlApplicationNode, "MinIncomeSingle", Amount))
        
        'Get value for global parameter MinIncomeJoint
        dblMinIncomeJoint = CDbl(GetGlobalParameter(xmlApplicationNode, "MinIncomeJoint", Amount))
        
        'Read each applicant for the application (applicants or guarantors)
        For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            If Not xmlApplicant Is Nothing Then
                'get customer version number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER") Is Nothing Then
                    strCustomerVersionNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERVERSIONNUMBER").nodeValue
                End If
                
                'Get customer number
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                    strCustomerNumber = xmlApplicant.Attributes.getNamedItem("CUSTOMERNUMBER").nodeValue
                End If
                
                Set xmlEmployment = xmlApplicant.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "' and @CUSTOMERNUMBER='" & strCustomerNumber & "']")
                
                If Not xmlApplicant.Attributes.getNamedItem("CUSTOMERROLETYPE_TYPE_A") Is Nothing Then
                    blIsApplicant = True
                Else
                    blIsApplicant = False 'guarantor
                    dblTotalIncomeGuarantor = 0
                End If
                
                If Not xmlEmployment Is Nothing Then
                    
                    'LDM 21/03/07 EP2_1852  if employed or employed-temp then look at confirmed income first then earned income records
                    If Not xmlEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']") Is Nothing Then
                    
                        If xmlEmployment.selectNodes("./CONFIRMEDEARNEDINCOME").length <> 0 Then
                            'Iterate through each CONFIRMEDEARNEDINCOME Record
                            For Each xmlEarnedIncome In xmlEmployment.selectNodes("./CONFIRMEDEARNEDINCOME")
                                'Combo values
                                strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_EMP"
                                strIncomeGroup = "INCOMEGROUP_TYPE_EC"
                                If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE") Is Nothing Then
                                    strEarnedIncomeType = xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE").Text
                                End If
                                
                                'Get ALLOWABLEINCOMEFACTOR record according to specified criteria
                                Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @TYPE='" & strEarnedIncomeType & "']")
                                
                                If Not xmlAllowableIncomeFactor Is Nothing Then
                                    If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT") Is Nothing _
                                    And Not xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE") Is Nothing _
                                    And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                        'Calculate Total Income
                                        dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                        dblThisIncome = CDbl(xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT").Text) * _
                                                        CDbl(xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE").Text)
                                        If blIsApplicant Then
                                            dblTotalIncomeApplicant = dblTotalIncomeApplicant + (dblThisIncome * dblFactor / 100)
                                        Else
                                            dblTotalIncomeGuarantor = dblTotalIncomeGuarantor + (dblThisIncome * dblFactor / 100)
                                        End If
                                    End If
                                End If
                            Next xmlEarnedIncome
                            
                        ElseIf xmlEmployment.selectNodes("./EARNEDINCOME").length <> 0 Then
                            
                            'Iterate through each EARNEDINCOME Record
                            For Each xmlEarnedIncome In xmlEmployment.selectNodes("./EARNEDINCOME")
                                'Combo values
                                strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_EMP"
                                strIncomeGroup = "INCOMEGROUP_TYPE_ED"
                                If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE") Is Nothing Then
                                    strEarnedIncomeType = xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMETYPE").Text
                                End If
                                
                                'Get ALLOWABLEINCOMEFACTOR record according to specified criteria
                                Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @TYPE='" & strEarnedIncomeType & "']")
                                
                                If Not xmlAllowableIncomeFactor Is Nothing Then
                                    If Not xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT") Is Nothing _
                                    And Not xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE") Is Nothing _
                                    And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                        'Calculate Total Income
                                        dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                        dblThisIncome = CDbl(xmlEarnedIncome.Attributes.getNamedItem("EARNEDINCOMEAMOUNT").Text) * _
                                                        CDbl(xmlEarnedIncome.Attributes.getNamedItem("PAYMENTFREQUENCYTYPE").Text)
                                        If blIsApplicant Then
                                            dblTotalIncomeApplicant = dblTotalIncomeApplicant + (dblThisIncome * dblFactor / 100)
                                        Else
                                            dblTotalIncomeGuarantor = dblTotalIncomeGuarantor + (dblThisIncome * dblFactor / 100)
                                        End If
                                    End If
                                End If
        
                            Next xmlEarnedIncome
                        
                        End If
                    
                    'LDM 21/03/07 EP2_1852 if self employed or contractor then look at accountants reference first then net profit records
                    ElseIf Not xmlEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_SELF='true' or @EMPLOYMENTSTATUS_TYPE_CON='true']") Is Nothing Then

                        If xmlEmployment.selectNodes("./ACCOUNTANTSREFERENCE").length > 0 Then
                            
                            Set xmlAccountantsReference = xmlEmployment.selectSingleNode("./ACCOUNTANTSREFERENCE")
                            
                            'Combo values
                            strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_SELF"
                            strIncomeGroup = "INCOMEGROUP_TYPE_SC"
                            strEarnedIncomeType = "TYPE_TYPE_ANP"
                            
                            'Select a suitable ALLOWABLEINCOMEFACTOR record
                            Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @" & strEarnedIncomeType & "]")
                            
                            If Not xmlAllowableIncomeFactor Is Nothing Then
                                If Not xmlAccountantsReference.Attributes.getNamedItem("NETPROFITENTITLEMENT1") Is Nothing _
                                And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                    'Calculate Total Income
                                    dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                    dblYearEnd1 = CDbl(xmlAccountantsReference.Attributes.getNamedItem("NETPROFITENTITLEMENT1").Text)
                                    If blIsApplicant Then
                                        dblTotalIncomeApplicant = dblYearEnd1 * dblFactor / 100
                                    Else
                                        dblTotalIncomeGuarantor = dblYearEnd1 * dblFactor / 100
                                    End If
                                
                                End If
                            End If
                                                    
                        Else
                            'No EarnedIncome records found
                            'Combo values
                            strEmploymentStatus = "EMPLOYMENTSTATUS_TYPE_SELF"
                            strIncomeGroup = "INCOMEGROUP_TYPE_SD"
                            strEarnedIncomeType = "TYPE_TYPE_ANP"
                            
                            'Select a suitable ALLOWABLEINCOMEFACTOR record
                            Set xmlAllowableIncomeFactor = xmlApplicationNode.selectSingleNode("./ALLOWABLEINCOMEFACTORS/ALLOWABLEINCOMEFACTOR[@" & _
                                strEmploymentStatus & " and @" & strIncomeGroup & " and @" & strEarnedIncomeType & "]")
                                            
                            'Get the NETPROFIT record
                            Set xmlNetProfit = xmlEmployment.selectSingleNode("NETPROFIT")
                            
                            If Not xmlAllowableIncomeFactor Is Nothing _
                            And Not xmlNetProfit Is Nothing Then
                                If Not xmlNetProfit.Attributes.getNamedItem("YEAR1AMOUNT") Is Nothing _
                                And Not xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR") Is Nothing Then
                                    'Calculate Total Income
                                    dblFactor = CDbl(xmlAllowableIncomeFactor.Attributes.getNamedItem("FACTOR").Text)
                                    If blIsApplicant Then
                                        dblTotalIncomeApplicant = CDbl(xmlNetProfit.Attributes.getNamedItem("YEAR1AMOUNT").Text) * dblFactor / 100
                                    Else
                                        dblTotalIncomeGuarantor = CDbl(xmlNetProfit.Attributes.getNamedItem("YEAR1AMOUNT").Text) * dblFactor / 100
                                    End If
                                End If
                                
                            End If
                        End If
                    End If
                End If
            End If
        Next xmlApplicant

                
                
        
        'Rule fails if Total Income for applicants is less than global parameter MinIncomeJoint
        If dblTotalIncomeApplicant < dblMinIncomeJoint Then
            blnSuccess = False
            strRuleValue = "Income: " & dblTotalIncomeApplicant
        Else
            'Rule fails if Total Income for guarantor is less than global parameter MinIncomeSingle
            If dblTotalIncomeGuarantor >= 0 And dblTotalIncomeGuarantor < dblMinIncomeSingle Then
                blnSuccess = False
                strRuleValue = "Income: " & dblTotalIncomeGuarantor
            End If
        End If
    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.DeclinePolicy
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6535 = intResult
        
End Function

Private Function Rule6545( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim objApplicant, _
        objEmployment As IXMLDOMNode
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode, _
        xmlEmploymentReference As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim lngAllowableIncome As Long, _
        lngGrossConfirmedAllowableIncome As Long
    
    Dim strRuleValue As String, _
        strRuleDescription As String
    
    Dim blnSuccess, _
        blnValidEmpType, _
        bContinue As Boolean
        
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6545"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6545: Confirmed income does not match declared
    'If confirmed income not = declared income then refer
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    blnValidEmpType = False
    
    'The rule is run separately for each applicant and guarantor on the application.
    'The rule is not run for self-cert cases nor for customers who are neither employed nor self-employed.
    'Do not run the rule if the application income status is not ='full status'
    If Not xmlApplicationNode.Attributes.getNamedItem("APPLICATIONINCOMESTATUS_TYPE_FS") Is Nothing Then
    
        'For each Customer linked to the case
        For Each objApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A='true' or @CUSTOMERROLETYPE_TYPE_G='true']")
            
            'Locate the Employment record where Employment.MainStatus = True.
            Set objEmployment = objApplicant.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1']")
            
            If Not objEmployment Is Nothing Then
                
                'IF Employment.EmploymentStatus not = "Employed" AND not = "Employed - Temporary" AND not = "Self Employed" AND not = "Contractor" (combo "EmploymentStatus" ValidationType not = "EMP" AND not = "EMPT" AND not = "SELF" AND not = "CON") then do not run the rule.
                 If Not objEmployment.selectSingleNode("./self::node()[@EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true' or @EMPLOYMENTSTATUS_TYPE_SELF='true' or @EMPLOYMENTSTATUS_TYPE_CON='true']") Is Nothing Then
                
                    bContinue = True
                    
                    ' if (employed or employed temp) and do not have a current employers ref then do not run the rule
                    'IF Employment.EmploymentStatus = "Employed" OR = "Employed - Temporary" (combo "EmploymentStatus" ValidationType = "EMP" OR = "EMPT") AND CurrentEmployersRef record not found then dont run the rule for this app/gar.
                    'LDM 13/3/07 EP2_1706
                    If Not objEmployment.selectSingleNode( _
                    "./self::node()[@EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']") Is Nothing Then
                        If objEmployment.selectSingleNode("./CURRENTEMPLOYERSREF") Is Nothing Then
                            bContinue = False
                        End If
                    End If
                    
                    ' if (self employed or contractor) and do not have a accountants reference do not run the rule
                    'ELSE IF Employment.EmploymentStatus = "Self Employed" OR = "Contractor" (combo "EmploymentStatus" ValidationType = "SELF" OR = "CON") AND AccountantsReference record not found then dont run the rule for this app/gar.
                    'LDM 13/3/07 EP2_1706
                    If bContinue = True Then
                        If Not objEmployment.selectSingleNode( _
                        "./self::node()[@EMPLOYMENTSTATUS_TYPE_SELF='true' or @EMPLOYMENTSTATUS_TYPE_CON='true' ]") Is Nothing Then
                            If objEmployment.selectSingleNode("./ACCOUNTANTSREFERENCE") Is Nothing Then
                                bContinue = False
                            End If
                        End If
                    End If
                        
                    'ELSE IF IncomeSummary.AllowableAnnualIncome not = IncomeSummary.GrossConfirmedAllowableIncome then [FAIL].
                    If bContinue = True Then
                        
                        blnValidEmpType = True 'return a result as a valid employment type has been found
                        Set xmlNode = objApplicant.selectSingleNode("./CUSTOMERVERSION/INCOMESUMMARY")
                        If xmlNode Is Nothing Then
                            blnSuccess = False 'this should not happen
                            Exit For
                        Else
                            'get allowable income
                            If Not xmlNode.Attributes.getNamedItem("ALLOWABLEANNUALINCOME") Is Nothing Then
                                lngAllowableIncome = CLng(xmlNode.Attributes.getNamedItem("ALLOWABLEANNUALINCOME").nodeValue)
                            End If
                            
                            'get gross confirmed allowable income
                            If Not xmlNode.Attributes.getNamedItem("GROSSCONFIRMEDALLOWABLEINCOME") Is Nothing Then
                                lngGrossConfirmedAllowableIncome = CLng(xmlNode.Attributes.getNamedItem("GROSSCONFIRMEDALLOWABLEINCOME").nodeValue)
                            End If
                            
                            'IF IncomeSummary.GrossConfirmedAllowableIncome not = IncomeSummary.AllowableAnnualIncome THEN [FAIL]
                            If lngGrossConfirmedAllowableIncome <> lngAllowableIncome Then
                                blnSuccess = False
                                Exit For 'if any applicants fail the rule then fail the rule and exit
                            End If
                            
                        End If
                    End If
                End If
            End If
        Next objApplicant
        
        If blnValidEmpType = True Then
            If Not blnSuccess Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
            End If
        End If
        
        If blnValidEmpType = True Then
            strRuleValue = "Confirmed: " & lngGrossConfirmedAllowableIncome & "; Declared: " & lngAllowableIncome
            AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        End If
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Rule6545 = intResult
    
    Set xmlNodeList = Nothing
    Set objApplicant = Nothing
    Set objEmployment = Nothing
            
End Function

Private Function Rule6550(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intEmploymentCount As Integer
    Dim intUnemployedCount As Integer
        
    Dim intEmploymentStatus
    Dim intRuleValidationType As Integer
    
    Dim objApplicant As IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6550"
    
    'RULE 6550: All applicants unemployed
    'IF all customers have Employment.EmploymentStatus = an "unemployed" value (combo "EmploymentStatus" ValidationType = "NOEMPLOYER") then [FAIL]
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
       
    intEmploymentCount = 0
    intUnemployedCount = 0
    
    For Each objApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        If Not objApplicant.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1']") Is Nothing Then
            With objApplicant.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT").Attributes
            
                intEmploymentCount = intEmploymentCount + 1
                
                If Not .getNamedItem("EMPLOYMENTSTATUS_TYPE_NOEMPLOYER") Is Nothing Then
                    intUnemployedCount = intUnemployedCount + 1
                End If
                
            End With
        End If
    Next objApplicant
    
    'Only fail if all unemployed
    If (intEmploymentCount > 0 And (intEmploymentCount = intUnemployedCount)) Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.DeclinePolicy
    End If
        
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6550 = intResult
End Function


'======================================================================================
'Added By:Dhira Singh
'Application Rule : 6555
'Date : 2nd Dec 2006
'======================================================================================

Private Function Rule6555(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim ndCustomerRole As IXMLDOMNode
    Dim ndContract As IXMLDOMNode
    Dim ndContractMonth As IXMLDOMNode
    Dim ndEmployment As IXMLDOMNode
    
    Dim intResult As Integer
    Dim intScore As Integer
        
    Dim dtDateStarted As Date
    Dim iTimeInMonths As Integer
    Dim iMinTimeEmployed As Integer
    Dim iMinTimeSelfEmployed As Integer
    Dim intLengthOfContractMonths As Integer
    Dim intLengthOfContractYears As Integer
    Dim intLikelyToBeRenewed As Integer
    Dim intLengthOfContract As Integer
    Dim intRuleValidationType As Integer
    
    Dim bEmployed As Boolean
    Dim bSelfEmployed As Boolean
    Dim blnSuccess As Boolean
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6555"
    
    blnSuccess = True
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'Rule 6555 : Contract less than minimum or not renewable
   
    iMinTimeSelfEmployed = CDbl(GetGlobalParameter(xmlApplicationNode, "MinTimeSelfEmployed", Amount))
    
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        
        bEmployed = False
        bSelfEmployed = False
        'Locate the Employment record where Employment.MainStatus = True
        'AND Employment.EmploymentStatus = "Contractor" (combo "EmploymentStatus" ValidationType = "CON").
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_CON='true']")
        
        If Not ndEmployment Is Nothing Then
            
            'Check there's a main employment record
            Set ndContract = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT/CONTRACTDETAILS")
            
            If Not ndContract Is Nothing Then
                
                If Not ndContract.Attributes.getNamedItem("LENGTHOFCONTRACTMONTHS") Is Nothing Then
                    intLengthOfContractMonths = CInt(ndContract.Attributes.getNamedItem("LENGTHOFCONTRACTMONTHS").Text)
                End If
                
                If Not ndContract.Attributes.getNamedItem("LENGTHOFCONTRACTYEARS") Is Nothing Then
                      intLengthOfContractYears = CInt(ndContract.Attributes.getNamedItem("LENGTHOFCONTRACTYEARS").Text)
                End If
                            
                 If Not ndContract.Attributes.getNamedItem("LIKELYTOBERENEWED") Is Nothing Then
                      intLikelyToBeRenewed = CInt(ndContract.Attributes.getNamedItem("LIKELYTOBERENEWED").Text)
                End If
            
            End If
                
               
            'Calculate length of contract in months
            intLengthOfContract = (intLengthOfContractYears * 12) + intLengthOfContractMonths
                  
            'IF any contractor customer has length of contract < global parameter "MinTimeSelfEmployed"
            'OR ContractDetails.LikelyToBeRenewed = False then [FAIL].
            If (intLikelyToBeRenewed = 0) Or (intLengthOfContract < iMinTimeSelfEmployed) Then
                blnSuccess = False
            End If
            
        End If
        
    Next ndCustomerRole
    
    If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Length (months): " & intLengthOfContract
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Rule6555 = intResult
End Function
'======================================================================================
'End: App Rule 6555
'======================================================================================


Private Function Rule6560(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    'Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlCustomerRole As IXMLDOMNode
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMinTimeSelfEmployed As Integer
    Dim intTimeInMonths As Integer
    
    Dim dtDateStarted As Date, _
        dtDate As Date
        
    Dim objApplicant As IXMLDOMNode
    Dim blnSuccess As Boolean

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6560"
    
    'RULE 6560: Self-employed time trading less than minimum
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
        
    'IF ApplicationFactFind.ApplicationReceivedDate is null, use System Date to calculate age.
    dtDate = GetDateTodayFromIngestion(xmlApplicationNode)
     
    For Each xmlCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        'Locate the Employment record where Employment.MainStatus = True AND Employment.EmploymentStatus = "Self Employed" (combo "EmploymentStatus" ValidationType = "SELF").
        If Not xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS_TYPE_SELF='true']") Is Nothing Then
            
            'Calculate age as at ApplicationFactFind.ApplicationReceivedDate (in whole months rounded down) using Employment.DateStartedOrEstablished.
            If Not xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT/@DATESTARTEDORESTABLISHED") Is Nothing Then
                dtDateStarted = xmlCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT/@DATESTARTEDORESTABLISHED").nodeValue
                intTimeInMonths = DateDiff("m", dtDateStarted, dtDate)
            End If
            
            'Get MinTimeSelfEmployed
            intMinTimeSelfEmployed = CDbl(GetGlobalParameter(xmlApplicationNode, "MinTimeSelfEmployed", Amount))
            
            'IF any self-employed applicant has time trading < global parameter "MinTimeSelfEmployed" then [FAIL].
            If intTimeInMonths < intMinTimeSelfEmployed Then
                blnSuccess = False
                Exit For
            End If
        End If
    Next
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
        
    'don't want to display the date if it's not been found
    If dtDateStarted > 0 Then
        strRuleValue = "Start Date: " & dtDateStarted
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6560 = intResult
End Function

Private Function Rule6565( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlApplicant As IXMLDOMNode
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean
      
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6565"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6565: Accountant's qualification
        
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    'Read each applicant for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        
        'check the accountant record
        Set xmlNode = xmlApplicant.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS_TYPE_SELF='true']/ACCOUNTANT")
        If Not xmlNode Is Nothing Then
            
            'if accountant qualifications is null then fail
            If xmlNode.Attributes.getNamedItem("QUALIFICATIONS") Is Nothing Then
                blnSuccess = False
            End If
            
            'if accountant qualifications = "other" then fail
            If Not xmlNode.Attributes.getNamedItem("QUALIFICATIONS_TYPE_O") Is Nothing Then
                blnSuccess = False
            End If
        End If
        
    Next xmlApplicant
    
            
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6565 = intResult
    
    Set xmlNodeList = Nothing
        
End Function


Private Function Rule6570( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlApplicant As IXMLDOMNode, _
        xmlAccountRelationships As IXMLDOMNode, _
        xmlMortgageAccounts As IXMLDOMNode, _
        xmlMortgageLoans As IXMLDOMNode
    
    Dim xmlNodeList As IXMLDOMNodeList
        
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intLoanTerm As Integer, _
        intMinExistLoanTerm As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim dtMortgageLoanStartDate As Date
    
    Dim blnSuccess As Boolean
      
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6570"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6570: With current lender less than 6 months
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    'For each non-DBM mortage account, check if the loan started < 6 months ago.
    'Read each applicant for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        'FOR each MortgageAccount record where MortgageAccount.BMidsAccount = False, read all related MortgageLoan records.
        For Each xmlAccountRelationships In xmlApplicant.selectNodes("CUSTOMERVERSION/ACCOUNTRELATIONSHIP")
                For Each xmlMortgageAccounts In xmlAccountRelationships.selectNodes("MORTGAGEACCOUNT[@BMIDSACCOUNT = '0']")
                    'read all related MortgageLoan records
                    For Each xmlMortgageLoans In xmlAccountRelationships.selectNodes("./MORTGAGELOAN")
                        'get loan start date
                        If Not xmlMortgageLoans.Attributes.getNamedItem("STARTDATE") Is Nothing Then
                            dtMortgageLoanStartDate = CDate(xmlMortgageLoans.Attributes.getNamedItem("STARTDATE").nodeTypedValue)
                            'calculate loan term
                            intLoanTerm = DateDiffInMonthsRoundedDown(dtMortgageLoanStartDate, xmlApplicationNode.Attributes.getNamedItem("APPLICATIONRECEIVEDDATE").Text)
                        
                            'get minimum existing loan term from global param
                            intMinExistLoanTerm = CInt(GetGlobalParameter(xmlApplicationNode, "LPMinExistLoanTerm", Amount))
                            
                            'IF loan term < global parameter "LPMinExistLoanTerm" then [FAIL]
                            If intLoanTerm < intMinExistLoanTerm Then
                                blnSuccess = False
                                Exit For
                            End If
                        End If
                        
                    Next xmlMortgageLoans
                Next xmlMortgageAccounts
        Next xmlAccountRelationships
    Next xmlApplicant
    
            
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Loan term: " & intLoanTerm
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6570 = intResult
    
    Set xmlNodeList = Nothing
        
End Function

Private Function Rule6575(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    Dim strCompanyName As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlThirdPartyAccount As IXMLDOMNode
    Dim xmlNameAndAddressDir As IXMLDOMNode
    
    Dim blnSuccess As Boolean
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6575"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")

        For Each xmlThirdPartyAccount In xmlApplicant.selectNodes("CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/THIRDPARTY")
                           
            If Not xmlThirdPartyAccount.Attributes.getNamedItem("COMPANYNAME") Is Nothing Then
                
                strCompanyName = UCase$(xmlThirdPartyAccount.Attributes.getNamedItem("COMPANYNAME").nodeValue)
                    
                If HasExistingBorrowings(strCompanyName) Then
                            
                        blnSuccess = False
                        Exit For
                End If
                    
            End If

        Next xmlThirdPartyAccount
        'AW EP923   30/06/2006
        For Each xmlNameAndAddressDir In xmlApplicant.selectNodes("CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/NAMEANDADDRESSDIRECTORY")
                           
            If Not xmlNameAndAddressDir.Attributes.getNamedItem("COMPANYNAME") Is Nothing Then
                
                strCompanyName = UCase$(xmlNameAndAddressDir.Attributes.getNamedItem("COMPANYNAME").nodeValue)
                    
                If HasExistingBorrowings(strCompanyName) Then
                            
                        blnSuccess = False
                        Exit For
                End If
                    
            End If

        Next xmlNameAndAddressDir
        'AW EP923   30/06/2006 - End
    Next xmlApplicant
    
    strRuleValue = ""
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6575 = intResult

End Function

Private Function Rule6580(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer

    Dim lngValuation As Long, _
        lngMinPurchasePriceRes As Long, _
        lngMinPurchasePriceBTL As Long

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean
    Dim strScoreCardInd As String
    
    Dim dblPurchasePriceOrEstimatedValue As Double
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6580"
    'Rule 6580: Property value / PP less than minimum
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    'Get minimum purchase price residential from global parameter
    lngMinPurchasePriceRes = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceRes", Amount)

    'Get minimum purchase price buy to let from global parameter
    Set xmlNodeList = xmlApplicationNode.selectNodes("GLOBALPARAMETER[@NAME=""MinPurchPriceBTL""]")
    lngMinPurchasePriceBTL = GetGlobalParameter(xmlApplicationNode, "MinPurchPriceBTL", Amount)
    
    'Check to see if a valuation report exists
    Set xmlNode = xmlApplicationNode.selectSingleNode("VALUERINSTRUCTION")
    'if valuation report exists
    If Not xmlNode Is Nothing Then
        'Read the ValnRepValuation record with the highest value of InstructionSequenceNumber for the Application
        Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/VALNREPVALUATION")
        If xmlNodeList.length > 0 Then
            If Not xmlNodeList.Item(0).Attributes.getNamedItem("POSTWORKSVALUATION") Is Nothing Then
                If xmlNodeList.Item(0).Attributes.getNamedItem("POSTWORKSVALUATION").nodeValue <> "0" Then
                    lngValuation = xmlNodeList.Item(0).Attributes.getNamedItem("POSTWORKSVALUATION").nodeValue
                Else
                    If Not xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION") Is Nothing Then
                        lngValuation = xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION").nodeValue
                    End If
                End If
            Else
                If Not xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION") Is Nothing Then
                    lngValuation = xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION").nodeValue
                End If
            End If
        End If
    Else
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            lngValuation = xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").nodeValue
        End If
    End If

    If Not xmlApplicationNode Is Nothing Then
        'residential
        If (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Or _
        Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing) And _
        lngValuation < lngMinPurchasePriceRes Then
            'if valuation < minimum purchase price then [FAIL]
            blnSuccess = False
            strRuleValue = "Value: £" & lngValuation
        'let to buy
        ElseIf (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
        Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing) And _
        lngValuation < lngMinPurchasePriceBTL Then
            'if valuation < minimum purchase price then [FAIL]
            blnSuccess = False
            strRuleValue = "Value: £" & lngValuation
        End If
    End If

    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6580 = intResult
End Function

Private Function Rule6590(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer
    
    Dim dblLTV As Double
    Dim dblAmountRequested As Double
    Dim dblPurchPriceEstVal As Double
    Dim dblMaxBTLRemortgageLTV As Double
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6590"
    
    'Rule 6590: LTV maximum exceeded for BTL remortgage
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF (ApplicationFactFind.NatureOfLoan not = "Buy-to-Let (Income)" AND not = "Buy-to-Let (Rental)" (combo "NatureOfLoan" ValidationType not = "BI" AND not = "BR")) OR ApplicationFactFind.TypeOfApplication not = "Remortgage" (combo "TypeOfMortgage" ValidationType = "R") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_R") Is Nothing Then

        dblMaxBTLRemortgageLTV = GetGlobalParameter(xmlApplicationNode, "MaxBTLRemortgageLTV", Percentage)
        
            
        'IF the Quotation is not found then calculate LTV = NewLoan.AmountRequested / ApplicationFactFind.PurchasePriceOrEstimatedValue * 100 (rounded to 2dp)
        intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
        If intMortgageSubquoteNumber = 0 Then
            If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
                dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
            End If
                
            If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
                dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
            End If
            
            
            dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
            
        'ELSE set LTV = MortgageSubQuote.LTV.
        Else
            Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("LTV") Is Nothing Then
                    dblLTV = xmlNode.Attributes.getNamedItem("LTV").nodeValue
                End If
            End If
        End If
        
        If dblLTV > dblMaxBTLRemortgageLTV Then
            blnSuccess = False
        End If
        
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
    
        strRuleValue = "LTV: " & dblLTV
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
        
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6590 = intResult
End Function

Private Function Rule6595(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intProdScheme As Integer
    Dim intRuleValidationType As Integer
    
    Dim dblBandAmount As Double
    Dim dblLTV As Double
    Dim dblAmountRequested As Double
    Dim dblPurchPriceEstVal As Double
    
    Dim bRes As Boolean
    Dim bLTB As Boolean
    Dim bBTL As Boolean
    
    Dim strEligInd As String
    Dim strSeek As String
    
    Dim xmlNode As IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6595"
    
    'Rule 6595 : Loan greater than LTV maximum for scheme
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    
    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If
    
    If strEligInd <> "" Then
    
        If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing Then
            strSeek = "MaxLoanByLTVRes"
        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_LT") Is Nothing Then
            strSeek = "MaxLoanByLTVLTB"
        ElseIf Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing Or Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
            strSeek = "MaxLoanByLTVBTL"
        End If
            
        
        Select Case strEligInd
                Case "ST":  strSeek = strSeek & "PR"
                Case "UL":  strSeek = strSeek & "NP"
                Case Else:  strSeek = strSeek & strEligInd
        End Select
             
             
        
        If Not xmlApplicationNode.selectSingleNode("QUOTATION") Is Nothing Then
            If Not xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@LTV") Is Nothing Then
                dblLTV = xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@LTV").Text
            End If
            
            If Not xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED") Is Nothing Then
            
              dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED").Text)
            End If
            
        Else
            If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
                dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
            End If
        
            If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
                dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
            End If
            
            dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
        End If
        
   
      ' Locate the global parameter band where High Band = smallest value >= LTV.
        If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='" & strSeek & "'][@HIGHBAND>='" & CStr(dblLTV) & "']") Is Nothing Then

            Set xmlNode = xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME='" & strSeek & "'][@HIGHBAND>='" & CStr(dblLTV) & "']")
            dblBandAmount = CDbl(xmlNode.Attributes.getNamedItem("AMOUNT").Text)

        End If
     
        'IF Loan > band Amount then [FAIL]
        If dblAmountRequested > dblBandAmount Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

    End If
    
    
    strRuleValue = "Loan: " & dblAmountRequested
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Rule6595 = intResult
End Function

Private Function Rule6600( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim objApplicant As IXMLDOMNode
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim dblAmountRequested As Double, _
        dblMaxAutoAcceptLimit As Double
        
    Dim blnSuccess As Boolean
    
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6600"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
        
    'RULE 6600: Loan Amount greater than auto accept limit
    blnSuccess = True
    
    'Get Max Auto Accept Limit
    dblMaxAutoAcceptLimit = GetGlobalParameter(xmlApplicationNode, "MaxAutoAcceptLimit", Amount)
    
    'IF the Quotation is not found then set loan = NewLoan.AmountRequested ELSE set loan = MortgageSubQuote.AmountRequested.
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    If intMortgageSubquoteNumber = 0 Then
        'get loan amount for the application
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
    Else
        Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("AMOUNTREQUESTED") Is Nothing Then
                dblAmountRequested = xmlNode.Attributes.getNamedItem("AMOUNTREQUESTED").nodeValue
            End If
        End If
    End If
    
    'If NewLoan.AmountRequested > global parameter MaxAutoAcceptLimit (£250,000) then [FAIL]
    If dblAmountRequested > dblMaxAutoAcceptLimit Then
        blnSuccess = False
    End If
                
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Loan: £" & dblAmountRequested
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6600 = intResult
    
    Set xmlNodeList = Nothing
        
End Function

Private Function Rule6605( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean
    
    Dim strScoreCardInd As String
    
    Dim lngValuation As Long, _
        lngPurchasePriceOrEstimatedValue As Long
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6605"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6605: Under-valued or gift
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
        
    'Read the VALNREPSUMMARY record within the ValnRepValuation record with the highest value of InstructionSequenceNumber for the Application
    Set xmlNode = xmlApplicationNode.selectSingleNode("VALNREPVALUATION")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("POSTWORKSVALUATION") Is Nothing Then
        
            If xmlNode.Attributes.getNamedItem("POSTWORKSVALUATION").nodeValue <> 0 Then
                lngValuation = xmlNode.Attributes.getNamedItem("POSTWORKSVALUATION").nodeValue
            End If
            
        Else
            If Not xmlNode.Attributes.getNamedItem("PRESENTVALUATION") Is Nothing Then
                lngValuation = xmlNode.Attributes.getNamedItem("PRESENTVALUATION").nodeValue
            End If
        End If
        
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            lngPurchasePriceOrEstimatedValue = xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").nodeValue
        End If
        
        'IF Valuation > ApplicationFactFind.PurchasePriceOrEstimatedValue then [FAIL]
        If lngValuation > lngPurchasePriceOrEstimatedValue Then
            blnSuccess = False
        End If
            
        strRuleValue = "Purchase Price: £" & lngPurchasePriceOrEstimatedValue & ", Valuation: £" & lngValuation
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6605 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6610( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        dblMaxLTVAllowed As Double, _
        intPropertyDepositCount_BIS As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer

    Dim dblAmountRequested As Double, _
        dblPurchPriceEstVal As Double, _
        dblLTV As Double
        
    Dim strRuleValue As String, _
        strRuleDescription As String

        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6610"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6610: Deposit includes Builders Incentive and max LTV
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
  
    'Get max LTV limit 1
    dblMaxLTVAllowed = GetGlobalParameter(xmlApplicationNode, "MaximumLTVAllowed", Percentage)
    
    'IF the Quotation is not found then calculate LTV = NewLoan.AmountRequested / ApplicationFactFind.PurchasePriceOrEstimatedValue * 100 (rounded to 2dp)
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    If intMortgageSubquoteNumber = 0 Then
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
            
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
        End If
        
        
        dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
        
    'ELSE set LTV = MortgageSubQuote.LTV.
    Else
        Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("LTV") Is Nothing Then
                dblLTV = xmlNode.Attributes.getNamedItem("LTV").nodeValue
            End If
        End If
    End If
        
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT[@SOURCEOFFUNDING_TYPE_BIS = 'true']")
    intPropertyDepositCount_BIS = xmlNodeList.length
        
    'IF LTV = global parameter "MaximumLTVAllowed" AND at least one "Builders Incentive Scheme" record is found then [FAIL].
    If intPropertyDepositCount_BIS > 0 And dblLTV = dblMaxLTVAllowed Then
        blnSuccess = False
    End If
    
    strRuleValue = "LTV: " & dblLTV
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6610 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6615( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMaxBISDeposit As Integer, _
        intMortgageSubquoteNumber As Integer
    
    Dim lngBISDepositTotalAmount As Long
        
    Dim dblThreshold As Double, _
        dblMaxLTVAllowed As Double, _
        dblAmountRequested As Double, _
        dblPurchPriceEstVal As Double, _
        dblLTV As Double, _
        dblValuation As Double
    
    Dim strRuleValue As String, _
        strRuleDescription As String
       
    Dim blnSuccess As Boolean
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6615"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6615: Builder gifted deposit greater than maximum
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True

    'Global parameter "MaxBISDeposit" = 5% that represents the maximum percentage of the valuation that can be from a builder's incentive scheme.
    intMaxBISDeposit = GetGlobalParameter(xmlApplicationNode, "MaxBISDeposit", Amount)

    'Get max LTV allowed
    dblMaxLTVAllowed = GetGlobalParameter(xmlApplicationNode, "MaximumLTVAllowed", Percentage)
    
    'Check to see if a valuation report exists
    Set xmlNode = xmlApplicationNode.selectSingleNode("VALUERINSTRUCTION")
    'if valuation report exists
    If Not xmlNode Is Nothing Then
        'Read the ValnRepValuation record with the highest value of InstructionSequenceNumber for the Application
        Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/VALNREPVALUATION")
        If xmlNodeList.length > 0 Then
            If Not xmlNodeList.Item(0).Attributes.getNamedItem("POSTWORKSVALUATION") Is Nothing Then
                If xmlNodeList.Item(0).Attributes.getNamedItem("POSTWORKSVALUATION").nodeValue <> "0" Then
                    dblValuation = xmlNodeList.Item(0).Attributes.getNamedItem("POSTWORKSVALUATION").nodeValue
                Else
                    If Not xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION") Is Nothing Then
                        dblValuation = xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION").nodeValue
                    End If
                End If
            Else
                If Not xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION") Is Nothing Then
                    dblValuation = xmlNodeList.Item(0).Attributes.getNamedItem("PRESENTVALUATION").nodeValue
                End If
            End If
        End If
    Else
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            dblValuation = xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").nodeValue
        End If
    End If
    
    'Set threshold
    dblThreshold = Round(dblValuation * intMaxBISDeposit / 100, 2)
    
    'Read all Builders Incentive Scheme deposit records for the case (see database notes)
    'Set BISDeposit = Sum NewPropertyDeposit.Amount for the records read
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT[@SOURCEOFFUNDING_TYPE_BIS = 'true']")
    For Each xmlNode In xmlNodeList
        If Not xmlNode.Attributes.getNamedItem("AMOUNT") Is Nothing Then
            lngBISDepositTotalAmount = lngBISDepositTotalAmount + xmlNode.Attributes.getNamedItem("AMOUNT").nodeValue
        End If
    Next xmlNode
    
    'IF the Quotation is not found then calculate LTV = NewLoan.AmountRequested / ApplicationFactFind.PurchasePriceOrEstimatedValue * 100 (rounded to 2dp)
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    If intMortgageSubquoteNumber = 0 Then
        If Not xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED") Is Nothing Then
            dblAmountRequested = CDbl(xmlApplicationNode.selectSingleNode("NEWLOAN/@AMOUNTREQUESTED").Text)
        End If
            
        If Not xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE") Is Nothing Then
            dblPurchPriceEstVal = CDbl(xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE").Text)
        End If
        
        
        dblLTV = Round((dblAmountRequested / dblPurchPriceEstVal) * 100, 2)
        
    'ELSE set LTV = MortgageSubQuote.LTV.
    Else
        Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("LTV") Is Nothing Then
                dblLTV = xmlNode.Attributes.getNamedItem("LTV").nodeValue
            End If
        End If
    End If
    
    'IF LTV < global parameter "MaximumLTVAllowed"  AND deposit > threshold then [FAIL]
    If dblLTV < dblMaxLTVAllowed And lngBISDepositTotalAmount > dblThreshold Then
        blnSuccess = False
    End If
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Deposit: £" & lngBISDepositTotalAmount & "; Valuation: £" & dblValuation
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6615 = intResult
        
End Function

Private Function Rule6620( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intPropertyDepositCount As Integer, _
        intPropertyDepositCount_GFT As Integer, _
        intRuleValidationType As Integer

    Dim dblMaxLTV2 As Double, _
        dblSubQuoteLTV As Double

    Dim strRuleValue As String, _
        strPropertyType As String, _
        strExLocalAuthorityInd As String, _
        strRuleDescription As String


    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6620"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6620: Gifted deposit
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
        
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT")
    intPropertyDepositCount = xmlNodeList.length
    
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/NEWPROPERTYDEPOSIT[@SOURCEOFFUNDING_TYPE_GFT = 'true']")
    intPropertyDepositCount_GFT = xmlNodeList.length
        
    'IF at least one NewPropertyDeposit record is read AND any NewPropertyDeposit records read have SourceOfFunding = "Deed of Gift" (combo "DepositSource" ValidationType = "GFT") then [FAIL].
    If intPropertyDepositCount > 0 And intPropertyDepositCount_GFT > 0 Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6620 = intResult
    
    Set xmlNode = Nothing
        
End Function

'======================================================================================
'Added By:Dhira Singh
'Application Rule : 6625
'Date : 05 Dec 2006
'======================================================================================
Private Function Rule6625( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlMortgageSubQuoteNode As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String, _
        strPurpose As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6625"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6625: Remortgage where purpose is "Other"
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'IF ApplicationFactFind.TypeOfApplication not = "Remortgage" (combo "TypeOfMortgage" ValidationType not = "R") then do not run the rule.
    If Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing Then
        Set xmlMortgageSubQuoteNode = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
        If Not xmlMortgageSubQuoteNode Is Nothing Then
            'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
             For Each xmlNode In xmlMortgageSubQuoteNode
                 'IF any LoanComponent has LoanComponent.SubPurposeOfLoan = "Other" (combo "SubPurposeOfLoan" ValidationType = "O") then [FAIL].
                 If Not xmlNode Is Nothing Then
                     If Not xmlNode.Attributes.getNamedItem("SUBPURPOSEOFLOAN_TYPE_O") Is Nothing Then
                         blnSuccess = False
                     End If
                 End If
             Next xmlNode
        End If
                    
        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6625 = intResult
    
    Set xmlNode = Nothing
        
End Function
'======================================================================================
'End : Application Rule 6625
'======================================================================================

Private Function Rule6635( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponentsNew As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6635"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6635: Total number of components greater than maximum
    'Multi-component loan and invalid number of components for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") then do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponentsNew = GetGlobalParameter(xmlApplicationNode, "MaxComponentsNew", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsNew" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponentsNew Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6635 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6640(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlApplicant As IXMLDOMNode, _
        xmlAccountRelationships As IXMLDOMNode, _
        xmlMortgageAccounts As IXMLDOMNode, _
        xmlMortgageLoans As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
    
    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim dblDisbursedAmount As Double
    Dim lngOriginalLoanAmount As Long
    
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6640"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6640: FA/CLI with outstanding retention
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
   
    'Read each applicant for the application
    For Each xmlApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        'FOR each MortgageAccount record where Account.IMPORTEDINDICATOR = true, read all related MortgageLoan records.
        For Each xmlAccountRelationships In xmlApplicant.selectNodes("CUSTOMERVERSION/ACCOUNTRELATIONSHIP")
            For Each xmlMortgageAccounts In xmlAccountRelationships.selectNodes("ACCOUNT[@IMPORTEDINDICATOR='1']")
                
                'IF the Account record is found then read all MortgageLoan records for the AccountGuid.
                For Each xmlMortgageLoans In xmlAccountRelationships.selectNodes("./MORTGAGELOAN")
                    If Not xmlMortgageLoans.Attributes.getNamedItem("DISBURSEDAMOUNT") Is Nothing Then
                        dblDisbursedAmount = xmlMortgageLoans.Attributes.getNamedItem("DISBURSEDAMOUNT").nodeValue
                    End If
                    
                    If Not xmlMortgageLoans.Attributes.getNamedItem("ORIGINALLOANAMOUNT") Is Nothing Then
                        lngOriginalLoanAmount = xmlMortgageLoans.Attributes.getNamedItem("ORIGINALLOANAMOUNT").nodeValue
                    End If
                    
                    'IF any MortgageLoan has DisbursedAmount < OriginalLoanAmount then [FAIL].
                    If dblDisbursedAmount < lngOriginalLoanAmount Then
                        blnSuccess = False
                        Set xmlApplicant = Nothing  'set xml to nothing so we exit loops
                        Set xmlMortgageLoans = Nothing
                        Set xmlMortgageAccounts = Nothing
                        Set xmlAccountRelationships = Nothing
                    End If
                Next xmlMortgageLoans
            Next xmlMortgageAccounts
        Next xmlAccountRelationships
    Next xmlApplicant
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
    End If
    
    strRuleValue = "Disbursed: " & dblDisbursedAmount & "; Loan: " & lngOriginalLoanAmount
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
  
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6640 = intResult
    
    Set xmlApplicant = Nothing
    Set xmlMortgageLoans = Nothing
    Set xmlMortgageAccounts = Nothing
    Set xmlAccountRelationships = Nothing
        
End Function

Private Function Rule6645(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode

    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim strEligInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6645"

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    ' Experian unable to allocate product scheme
    ' IF BespokeDecision.EligibilityIndicator = "NB" OR = "NR" then [FAIL].

    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If

    If strEligInd = "NB" Or strEligInd = "NR" Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Product Scheme: " & strEligInd
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6645 = intResult

End Function
Private Function Rule6650(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim objApplicant As IXMLDOMNode
    
    Dim xmlNode As IXMLDOMNode
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strEligInd As String
        
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bFail As Boolean
    Dim objNode As IXMLDOMNode
        
    'RULE 6650: Invalid product scheme returned from Experian
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6650"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule6650() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'Locate the entry in combo "ProductScheme" where ValidationType = BespokeDecision.EligibilityIndicator.
    'IF the combo entry is not found then [FAIL]
    Set xmlNode = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("ELIGIBILITYINDICATOR") Is Nothing Then
            strEligInd = xmlNode.Attributes.getNamedItem("ELIGIBILITYINDICATOR").nodeValue
        End If
        
        
        Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='ProductScheme']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE='" & strEligInd & "']")
        If xmlNode Is Nothing Then
            bFail = True
        End If
    End If
    
    
    If bFail Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    strRuleValue = "Product Scheme: " & strEligInd
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6650 = intResult
End Function

Private Function Rule6660(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim xmlGlobalParamValue As IXMLDOMNode
    Dim objApplicant As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bFail As Boolean
    Dim objNode As IXMLDOMNode
        
    'Rule 6660: Experian failure or time-out
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6660"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
  
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' IF the ApplicationCreditCheck record is not found
    '    OR the KnowYourCustomerCheck record for any applicant is not found then [FAIL].
    
    If xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK") Is Nothing Then
        bFail = True
    End If
    
    For Each objNode In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If objNode.selectSingleNode("./KNOWYOURCUSTOMERCHECK") Is Nothing Then
            bFail = True
        End If
    Next objNode
    
    If bFail Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6660 = intResult
End Function

Private Function Rule6695(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    Dim objApplicant As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim bFail As Boolean
    Dim objNode As IXMLDOMNode
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6695"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' rule 6695
    ' Applicant authentication documentation required
    ' Run the rule for each applicant on the case.
    ' // A fail does not affect the overall DIP but is for information to the user.
    ' IF KnowYourCustomerAuthenticateResults.AuthDecision = "NA00" then [FAIL].
    
    For Each objNode In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If Not objNode.selectSingleNode("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS/@AUTHDECISION") Is Nothing Then
            If objNode.selectSingleNode("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS/@AUTHDECISION").Text = "NA00" Then
                bFail = True
                Exit For
            End If
        End If
    Next objNode
    
    If bFail Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6695 = intResult
End Function

Private Function Rule6665( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim bHighRisk As Boolean, _
        bHighRiskFound As Boolean
    
    Dim xmlCustomer As IXMLDOMNode, _
        xmlComboHighPolicy As IXMLDOMNode, _
        xmlHighPolicy As IXMLDOMNode, _
        xmlHighRiskCode As IXMLDOMNode
        
    Dim xmlHighPolicies As IXMLDOMNodeList, _
        xmlHighRiskCodes As IXMLDOMNodeList
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strHighRiskPolicyRule As String, _
        strHighRiskCodeDesc As String
        
    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6665"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6665: Authenticate "High Risk" code - Experian
    'Note: code taken from DIP rule 3310
                    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'apply this rule to each applicant
    For Each xmlCustomer In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        If Not xmlCustomer Is Nothing Then
            ' Read all KnowYourCustomerHighRiskPolicyRules records for the
            ' KnowYourCustomerAuthenticateResults record.
            ' FOR each record read:
            ' IF HighRiskPolicyRule = ValidationType from any entry in combo "AuthHighRiskCodes" then [FAIL]
            
            'make sure we are searching against the latest knowyourcustomercheck record - use last()
            Set xmlHighPolicies = xmlCustomer.selectNodes("./KNOWYOURCUSTOMERCHECK/KNOWYOURCUSTOMERAUTHENTICATERESULTS/KNOWYOURCUSTOMERHIGHRISKPOLICYRULES")
            For Each xmlHighPolicy In xmlHighPolicies
                If Not xmlHighPolicy Is Nothing Then
                    If Not xmlHighPolicy.Attributes.getNamedItem("HIGHRISKPOLICYRULE") Is Nothing Then
                        strHighRiskPolicyRule = xmlHighPolicy.Attributes.getNamedItem("HIGHRISKPOLICYRULE").nodeValue
                        Set xmlComboHighPolicy = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='AuthHighRiskCodes']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE='" & strHighRiskPolicyRule & "']")
                        If Not xmlComboHighPolicy Is Nothing Then
                            Set xmlNode = xmlComboHighPolicy.selectSingleNode("../@VALUENAME")
                            If Not xmlNode Is Nothing Then
                                'get SM value name
                                strHighRiskCodeDesc = xmlNode.Text
                            End If

                            'search for rule validation type
                            Set xmlHighRiskCodes = xmlApplicationNode.selectNodes("./COMBOGROUP[@GROUPNAME='AuthHighRiskCodes']/COMBOVALUE[@VALUENAME='" & strHighRiskCodeDesc & "']/COMBOVALIDATION/@VALIDATIONTYPE")
                            For Each xmlHighRiskCode In xmlHighRiskCodes
                                'if app rule found
                                If InStr(1, xmlHighRiskCode.Text, "APP_RULE_") > 0 Then
                                    'get App rule number only
                                    gstrRuleNumber = Mid(xmlHighRiskCode.Text, 10)
                                    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
                                    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
                                    bHighRiskFound = True
                                    
                                    'report fail for each high risk policy code found
                                    strRuleValue = gstrRuleNumber
                                    intScore = intRuleValidationType
                                    intResult = RiskAssessmentDecision.Referral
                                    CountReferral (intScore) 'add to appropriate referral count
                                    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
                                End If
                            Next xmlHighRiskCode

                        End If
                    End If
                End If
            Next xmlHighPolicy
        End If
    
    Next xmlCustomer
    
    'if a high risk has not been found
    If Not bHighRiskFound Then
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6665 = intResult
    
    Set xmlNodeList = Nothing
        
End Function

Private Function Rule6700( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer

    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlCreditCheckReasonCodes As IXMLDOMNodeList, _
        xmlSMReasonCodes As IXMLDOMNodeList

    Dim xmlNode As IXMLDOMNode, _
        objApplicant As IXMLDOMNode, _
        xmlComboValidation As IXMLDOMNode, _
        xmlSMReasonCode As IXMLDOMNode

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intRuleNumberCount As Integer, _
        intRuleCount As Integer

    Dim bSMReasonCodeFound As Boolean
    Dim bSMRuleFound As Boolean
    Dim bReturnRuleResult As Boolean

    Dim strSMRuleNumbers() As String


    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strReasonCode As String, _
        strSMValueName As String

    Dim strScoreCardInd As String


    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6700"
    strScoreCardInd = 0 'False

    'RULE 6700: Strategy Manager referral - Experian
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'get most recent application credit check record
    Set xmlCreditCheckReasonCodes = xmlApplicationNode.selectNodes("APPLICATIONCREDITCHECK[1]/CREDITCHECKREASONCODE")
    For Each xmlNode In xmlCreditCheckReasonCodes
        'get reason code
        If Not xmlNode.Attributes.getNamedItem("REASONCODE") Is Nothing Then
            strReasonCode = xmlNode.Attributes.getNamedItem("REASONCODE").nodeValue
        End If

        'IF ReasonCode = ValidationType from any entry in combo "SMReasonCode"
        'AND combo entry also has ValidationType = "Refer" then
        Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='SMReasonCode']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE='" & strReasonCode & "']")
        If Not xmlNode Is Nothing Then
            Set xmlNode = xmlNode.selectSingleNode("../COMBOVALIDATION[@VALIDATIONTYPE='" & "Refer" & "']")
            If Not xmlNode Is Nothing Then
                Set xmlNode = xmlNode.selectSingleNode("../@VALUENAME")
                If Not xmlNode Is Nothing Then
                    'get SM value name
                    strSMValueName = xmlNode.Text
                End If

                'search for rule validation type
                Set xmlSMReasonCodes = xmlApplicationNode.selectNodes("./COMBOGROUP[@GROUPNAME='SMReasonCode']/COMBOVALUE[@VALUENAME='" & strSMValueName & "']/COMBOVALIDATION/@VALIDATIONTYPE")
                For Each xmlSMReasonCode In xmlSMReasonCodes
                    'return the rule result in the XML as default
                    bReturnRuleResult = True

                    'if app rule found
                    If InStr(1, xmlSMReasonCode.Text, "APP_RULE_") Then
                        'get app rule number only
                        gstrRuleNumber = Mid(xmlSMReasonCode.Text, 10)

                        If intRuleNumberCount > 0 Then
                            For intRuleCount = 1 To UBound(strSMRuleNumbers)
                                If strSMRuleNumbers(intRuleCount) = gstrRuleNumber Then
                                    bReturnRuleResult = False
                                End If
                            Next intRuleCount
                        End If

                        'only return the rule in the XML if it hasn't already been returned
                        If bReturnRuleResult Then

                            strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
                            intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
                            bSMReasonCodeFound = True

                            'report fail for each sm reason code found
                            strRuleValue = gstrRuleNumber
                            intScore = intRuleValidationType
                            intResult = RiskAssessmentDecision.Referral
                            CountReferral (intScore) 'add to appropriate referral count
                            AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

                            'increment rule number array counter
                            intRuleNumberCount = intRuleNumberCount + 1

                            ReDim Preserve strSMRuleNumbers(1 To intRuleNumberCount) As String
                            'add the rule number to the array.  This indicates the rule has been returned in the XML.
                            strSMRuleNumbers(intRuleNumberCount) = gstrRuleNumber

                        End If
                    End If
                Next xmlSMReasonCode

            End If
        End If
    Next xmlNode

    If Not bSMReasonCodeFound Then
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6700 = intResult

    Set xmlNodeList = Nothing

End Function

Private Function Rule6745( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer

    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer

    Dim dCurrentAffordabilityIndex As Double
    Dim dMinAffordabilityIndex As Double

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6745"

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' RULE 6745: Affordability Index below minimum - Experian.  Based on rule 3350.
    ' IF BXApplicationBureau.RRP2_CurrentAffordabilityIndex < global parameter "MinAffordabilityIndex" then
    '  [FAIL]

    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BXAPPLICATIONBUREAU/@RRP2_CURRENTAFFORDABILITYINDEX") Is Nothing Then
        dMinAffordabilityIndex = CDbl(GetGlobalParameter(xmlApplicationNode, "MinAffordabilityIndex", Amount))
        If Not xmlApplicationNode.selectSingleNode("./GLOBALPARAMETER[@NAME=""MinAffordabilityIndex""]") Is Nothing Then
            
            If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BXAPPLICATIONBUREAU/@RRP2_CURRENTAFFORDABILITYINDEX") Is Nothing Then
                dCurrentAffordabilityIndex = CDbl(xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BXAPPLICATIONBUREAU/@RRP2_CURRENTAFFORDABILITYINDEX").Text)
            End If
            
            If dCurrentAffordabilityIndex < dMinAffordabilityIndex Then
                intScore = intRuleValidationType
                strRuleValue = "AI: " & dCurrentAffordabilityIndex
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
            End If
        End If
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6745 = intResult

End Function

Private Function Rule6755(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim blnSuccess As Boolean
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageSubquoteNumber As Integer
    Dim intGroupSequenceNumber As Integer
    
    Dim strValidationType As String
    Dim strEligInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6755"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    ' Rule 6755: Product scheme cascade
    
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
    
    'Locate any LoanComponent record where LoanComponent.ManualPortedLoanInd = False AND LoanComponent.ProductSwitchRetainProductInd = False.
    Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
    If Not xmlNode Is Nothing Then
        'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber.
        Set xmlNode = xmlNode.selectSingleNode("SPECIALGROUP")
        If Not xmlNode Is Nothing Then
            'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber
            If Not xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER") Is Nothing Then
                intGroupSequenceNumber = xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER").nodeValue
                Set xmlNode = xmlApplicationNode.selectSingleNode("./COMBOGROUP[@GROUPNAME='SpecialGroup']/COMBOVALUE[@VALUEID='" & intGroupSequenceNumber & "']")
                If Not xmlNode Is Nothing Then
                    Set xmlNode = xmlNode.selectSingleNode("COMBOVALIDATION")
                    If Not xmlNode.Attributes.getNamedItem("VALIDATIONTYPE") Is Nothing Then
                        strValidationType = xmlNode.Attributes.getNamedItem("VALIDATIONTYPE").nodeValue
                    End If
                End If
            End If
        End If
    End If
    
    'IF ValidationType not = BespokeDecision.EligibilityIndicator then [FAIL].
    If Not xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR") Is Nothing Then
        strEligInd = xmlApplicationNode.selectSingleNode("./APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR").Text
    End If
    
    'If the special group does not match the value determined by Strategy Manager then the rule fails.
    If strValidationType <> strEligInd Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
        intScore = intRuleValidationType
    End If
    
    strRuleValue = "Requested: " & strValidationType & ", Strat Man: " & strEligInd
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6755 = intResult
End Function

Private Function Rule6760( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intSelfBuildInd As Integer, _
        intRuleValidationType As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

        
    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6760"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6760: The property being mortgaged is self-built.
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
        
    Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("SELFBUILDINDICATOR") Is Nothing Then
            intSelfBuildInd = xmlNode.Attributes.getNamedItem("SELFBUILDINDICATOR").nodeValue
        End If
    End If
    
    'IF NewProperty.SelfBuildIndicator = True then [FAIL]
    If intSelfBuildInd = 1 Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6760 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6765( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer

    Dim xmlNode As IXMLDOMNode

    Dim intResult As Integer, _
        intScore As Integer, _
        intDemandInArea As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer

    Dim lngRentalCoverPercentage As Long

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim dblMonthlyRentalIncome As Double, _
        dblTotalNetMonthlyCost As Double, _
        dblBandAmount As Double, _
        dblLTV As Double, _
        dblEarnedIncomeAmount As Double, _
        dblRentalIncome As Double

    Dim blnSuccess As Boolean, _
        blValRepPropertyPercentage_Calculated As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6765"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6765: For BTL-R % rental income cover < min LTV allowed

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF ApplicationFactFind.NatureOfLoan = "Buy-to-Let Rental Based" (combo "NatureOfLoan" ValidationType = "BR")
    If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing Then
        
        Set xmlNode = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TOTALNETMONTHLYCOST") Is Nothing Then
                dblTotalNetMonthlyCost = xmlNode.Attributes.getNamedItem("TOTALNETMONTHLYCOST").nodeValue
            End If
            If Not xmlNode.Attributes.getNamedItem("LTV") Is Nothing Then
                dblLTV = xmlNode.Attributes.getNamedItem("LTV").nodeValue
            End If
        End If
                
        'Read the most recent valuation report
        Set xmlNode = xmlApplicationNode.selectSingleNode("VALNREPPROPERTYDETAILS[last()]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("RENTALINCOME") Is Nothing Then
                If xmlNode.Attributes.getNamedItem("RENTALINCOME").nodeValue <> 0 Then
                    dblRentalIncome = xmlNode.Attributes.getNamedItem("RENTALINCOME").nodeValue
                    
                    'avoid an overflow error
                    If dblTotalNetMonthlyCost > 0 Then
                        'Calculate RentalCover% = (ValnRepPropertyDetails.RentalIncome / MortgageSubQuote.TotalNetMonthlyCost * 100) truncated to nearest integer.
                        lngRentalCoverPercentage = Fix((dblRentalIncome / dblTotalNetMonthlyCost) * 100)
                        blValRepPropertyPercentage_Calculated = True
                    End If
                End If
            End If
        End If
        
        'if ValnRepPropertyDetails.RentalIncome not available,
        'calculate rental % based on NEWPROPERTY details
        If Not blValRepPropertyPercentage_Calculated Then
            'calculate RentalCover% = (NewProperty.MonthlyRentalIncome / MortgageSubQuote.TotalNetMonthlyCost * 100) truncated to nearest integer
            Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
            If Not xmlNode Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("MONTHLYRENTALINCOME") Is Nothing Then
                    dblMonthlyRentalIncome = xmlNode.Attributes.getNamedItem("MONTHLYRENTALINCOME").nodeValue
                End If
                
                If dblMonthlyRentalIncome > 0 And dblTotalNetMonthlyCost > 0 Then 'avoid an overflow error
                    'truncate (round down) % to nearest whole number
                    lngRentalCoverPercentage = Fix((dblMonthlyRentalIncome / dblTotalNetMonthlyCost) * 100)
                End If
            End If
        End If

        'Locate the band in global banded parameter "LPMaxLTVBTLRental" where MortgageSubQuote.LTV <= smallest value of High Band.
        If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME=""LPMaxLTVBTLRental""][@HIGHBAND>=" & CStr(dblLTV) & "]") Is Nothing Then
            If Not xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME=""LPMaxLTVBTLRental""][@HIGHBAND>=" & CStr(dblLTV) & "]").Attributes.getNamedItem("PERCENTAGE") Is Nothing Then
                dblBandAmount = CDbl(xmlApplicationNode.selectSingleNode("GLOBALBANDEDPARAMETER[@NAME=""LPMaxLTVBTLRental""][@HIGHBAND>=" & CStr(dblLTV) & "]").Attributes.getNamedItem("PERCENTAGE").Text)
            End If
        End If

        'IF RentalCover% < Amount from parameter band then [FAIL]
        If lngRentalCoverPercentage < dblBandAmount Then
            blnSuccess = False
        End If
    End If
    
    strRuleValue = "RentalCover%: " & lngRentalCoverPercentage
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6765 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6770( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode, _
        objApplicant As IXMLDOMNode, _
        xmlCustomerAddress As IXMLDOMNode, _
        xmlCustomerAddressDetails As IXMLDOMNode, _
        xmlNewPropertyAddress As IXMLDOMNode, _
        xmlValnPropertyDetails As IXMLDOMNode

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strTenantedPropertyInd As String, _
        strRuleDescription As String
    
    'current property and new property details - used for address comparison
    Dim strCurrentBuildingOrHouseName As String, _
        strCurrentBuildingOrHouseNumber As String, _
        strCurrentFlatNumber As String, _
        strCurrentStreet As String, _
        strCurrentDistrict As String, _
        strCurrentTown As String, _
        strCurrentCounty As String, _
        strCurrentCountry As String, _
        strCurrentPostcode As String, _
        strNewBuildingOrHouseName As String, _
        strNewBuildingOrHouseNumber As String, _
        strNewFlatNumber As String, _
        strNewStreet As String, _
        strNewDistrict As String, _
        strNewTown As String, _
        strNewCounty As String, _
        strNewCountry As String, _
        strNewPostcode As String

    Dim strValnBuildingOrHouseName As String, _
        strValnBuildingOrHouseNumber As String, _
        strValnFlatNumber As String, _
        strValnStreet As String, _
        strValnDistrict As String, _
        strValnTown As String, _
        strValnCounty As String, _
        strValnCountry As String, _
        strValnPostcode As String


    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6770"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6770: Unencumbered property
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
            
    'FOR each CustomerRole record linked to the ApplicationFactFind, read the current CustomerAddress record where AddressType = "Home" OR "HomeTargeted" AND NatureOfOccupancy = "Home Owner - No Mortgage" (combo "NatureOfOccupancy" ValidationType = "ONM").
    For Each objApplicant In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        Set xmlCustomerAddress = objApplicant.selectSingleNode("./CUSTOMERVERSION/CUSTOMERADDRESS[@NATUREOFOCCUPANCY_TYPE_ONM='true']")
        'IF the CustomerAddress record is found then read the Address record.
        If Not xmlCustomerAddress Is Nothing Then
        
            If Not xmlCustomerAddress.Attributes.getNamedItem("ADDRESSTYPE_TYPE_H") Is Nothing Or _
            Not xmlCustomerAddress.Attributes.getNamedItem("ADDRESSTYPE_TYPE_TH") Is Nothing Then
                    
                strCurrentBuildingOrHouseName = ""
                strCurrentBuildingOrHouseNumber = ""
                strCurrentFlatNumber = ""
                strCurrentStreet = ""
                strCurrentTown = ""
                strCurrentCounty = ""
                strCurrentCountry = ""
                strCurrentPostcode = ""
                strCurrentDistrict = ""
                                                  
                'find current address
                Set xmlCustomerAddressDetails = xmlCustomerAddress.selectSingleNode("ADDRESS")
                If Not xmlCustomerAddressDetails Is Nothing Then
                    'get current address details
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("BUILDINGORHOUSENAME") Is Nothing Then
                        strCurrentBuildingOrHouseName = xmlCustomerAddressDetails.Attributes.getNamedItem("BUILDINGORHOUSENAME").nodeTypedValue
                    End If
                                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("BUILDINGORHOUSENUMBER") Is Nothing Then
                        strCurrentBuildingOrHouseNumber = xmlCustomerAddressDetails.Attributes.getNamedItem("BUILDINGORHOUSENUMBER").nodeTypedValue
                    End If
                  
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("FLATNUMBER") Is Nothing Then
                        strCurrentFlatNumber = xmlCustomerAddressDetails.Attributes.getNamedItem("FLATNUMBER").nodeTypedValue
                    End If
                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("STREET") Is Nothing Then
                        strCurrentStreet = xmlCustomerAddressDetails.Attributes.getNamedItem("STREET").nodeTypedValue
                    End If
                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("DISTRICT") Is Nothing Then
                        strCurrentDistrict = xmlCustomerAddressDetails.Attributes.getNamedItem("DISTRICT").nodeTypedValue
                    End If
                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("TOWN") Is Nothing Then
                        strCurrentTown = xmlCustomerAddressDetails.Attributes.getNamedItem("TOWN").nodeTypedValue
                    End If
                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("COUNTY") Is Nothing Then
                        strCurrentCounty = xmlCustomerAddressDetails.Attributes.getNamedItem("COUNTY").nodeTypedValue
                    End If
                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("COUNTRY") Is Nothing Then
                        strCurrentCountry = xmlCustomerAddressDetails.Attributes.getNamedItem("COUNTRY").nodeTypedValue
                    End If
                    
                    If Not xmlCustomerAddressDetails.Attributes.getNamedItem("POSTCODE") Is Nothing Then
                        strCurrentPostcode = xmlCustomerAddressDetails.Attributes.getNamedItem("POSTCODE").nodeTypedValue
                    End If
                                    
                    Set xmlValnPropertyDetails = xmlApplicationNode.selectSingleNode("VALNREPPROPERTYDETAILS[last()]")
                    If Not xmlValnPropertyDetails Is Nothing Then
                        
                        strValnBuildingOrHouseName = ""
                        strValnBuildingOrHouseNumber = ""
                        strValnFlatNumber = ""
                        strValnStreet = ""
                        strValnTown = ""
                        strValnCounty = ""
                        strValnCountry = ""
                        strValnPostcode = ""
                        strValnDistrict = ""

                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("BUILDINGORHOUSENAME") Is Nothing Then
                            strValnBuildingOrHouseName = xmlValnPropertyDetails.Attributes.getNamedItem("BUILDINGORHOUSENAME").nodeTypedValue
                        End If
                                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("BUILDINGORHOUSENUMBER") Is Nothing Then
                            strValnBuildingOrHouseNumber = xmlValnPropertyDetails.Attributes.getNamedItem("BUILDINGORHOUSENUMBER").nodeTypedValue
                        End If
                    
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("FLATNUMBER") Is Nothing Then
                            strValnFlatNumber = xmlValnPropertyDetails.Attributes.getNamedItem("FLATNUMBER").nodeTypedValue
                        End If
                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("STREET") Is Nothing Then
                            strValnStreet = xmlValnPropertyDetails.Attributes.getNamedItem("STREET").nodeTypedValue
                        End If
                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("DISTRICT") Is Nothing Then
                            strValnDistrict = xmlValnPropertyDetails.Attributes.getNamedItem("DISTRICT").nodeTypedValue
                        End If
                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("TOWN") Is Nothing Then
                            strValnTown = xmlValnPropertyDetails.Attributes.getNamedItem("TOWN").nodeTypedValue
                        End If
                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("COUNTY") Is Nothing Then
                            strValnCounty = xmlValnPropertyDetails.Attributes.getNamedItem("COUNTY").nodeTypedValue
                        End If
                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("COUNTRY") Is Nothing Then
                            strValnCountry = xmlValnPropertyDetails.Attributes.getNamedItem("COUNTRY").nodeTypedValue
                        End If
                      
                        If Not xmlValnPropertyDetails.Attributes.getNamedItem("POSTCODE") Is Nothing Then
                            strValnPostcode = xmlValnPropertyDetails.Attributes.getNamedItem("POSTCODE").nodeTypedValue
                        End If
                                    
                        If strCurrentBuildingOrHouseName = strValnBuildingOrHouseName And _
                          strCurrentBuildingOrHouseNumber = strValnBuildingOrHouseNumber And _
                          strCurrentFlatNumber = strValnFlatNumber And _
                          strCurrentStreet = strValnStreet And _
                          strCurrentTown = strValnTown And _
                          strCurrentCounty = strValnCounty And _
                          strCurrentCountry = strValnCountry And _
                          strCurrentPostcode = strValnPostcode And _
                          strCurrentDistrict = strValnDistrict Then
                            blnSuccess = False
                            Exit For
                        End If
                                                                        
                    Else
                        strNewBuildingOrHouseName = ""
                        strNewBuildingOrHouseNumber = ""
                        strNewFlatNumber = ""
                        strNewStreet = ""
                        strNewTown = ""
                        strNewCounty = ""
                        strNewCountry = ""
                        strNewPostcode = ""
                        strNewDistrict = ""

                        Set xmlNewPropertyAddress = xmlApplicationNode.selectSingleNode("./NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS")
                        'Compare the following Address fields: BuildingOrHouseName, BuildingOrHouseNumber, FlatNumber, Street, District, Town, County, Country, Postcode.
                        If Not xmlNewPropertyAddress Is Nothing Then
                          'get new address details
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("BUILDINGORHOUSENAME") Is Nothing Then
                                strNewBuildingOrHouseName = xmlNewPropertyAddress.Attributes.getNamedItem("BUILDINGORHOUSENAME").nodeTypedValue
                            End If
                                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("BUILDINGORHOUSENUMBER") Is Nothing Then
                                strNewBuildingOrHouseNumber = xmlNewPropertyAddress.Attributes.getNamedItem("BUILDINGORHOUSENUMBER").nodeTypedValue
                            End If
                        
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("FLATNUMBER") Is Nothing Then
                                strNewFlatNumber = xmlNewPropertyAddress.Attributes.getNamedItem("FLATNUMBER").nodeTypedValue
                            End If
                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("STREET") Is Nothing Then
                                strNewStreet = xmlNewPropertyAddress.Attributes.getNamedItem("STREET").nodeTypedValue
                            End If
                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("DISTRICT") Is Nothing Then
                                strNewDistrict = xmlNewPropertyAddress.Attributes.getNamedItem("DISTRICT").nodeTypedValue
                            End If
                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("TOWN") Is Nothing Then
                                strNewTown = xmlNewPropertyAddress.Attributes.getNamedItem("TOWN").nodeTypedValue
                            End If
                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("COUNTY") Is Nothing Then
                                strNewCounty = xmlNewPropertyAddress.Attributes.getNamedItem("COUNTY").nodeTypedValue
                            End If
                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("COUNTRY") Is Nothing Then
                                strNewCountry = xmlNewPropertyAddress.Attributes.getNamedItem("COUNTRY").nodeTypedValue
                            End If
                          
                            If Not xmlNewPropertyAddress.Attributes.getNamedItem("POSTCODE") Is Nothing Then
                                strNewPostcode = xmlNewPropertyAddress.Attributes.getNamedItem("POSTCODE").nodeTypedValue
                            End If
                          
                            'IF the applicant's CurrentAddress record is found AND CurrentAddress.Address = NewPropertyAddress.Address then [FAIL].
                            If strCurrentBuildingOrHouseName = strNewBuildingOrHouseName And _
                              strCurrentBuildingOrHouseNumber = strNewBuildingOrHouseNumber And _
                              strCurrentFlatNumber = strNewFlatNumber And _
                              strCurrentStreet = strNewStreet And _
                              strCurrentTown = strNewTown And _
                              strCurrentCounty = strNewCounty And _
                              strCurrentCountry = strNewCountry And _
                              strCurrentPostcode = strNewPostcode And _
                              strCurrentDistrict = strNewDistrict Then
                                blnSuccess = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next objApplicant
        
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6770 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule6775( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim dblTotalAllowableIncome As Double, _
        dblTotalApproximateMonthlyCost As Double
    
    Dim blnSuccess As Boolean
        
    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6775"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6775: Total Allowable Income < Total Approximate Monthly Cost
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)
        
    'Please note: totalallowableincome & total approximatemonthlycost are calculated fields called from omRA
    If Not xmlApplicationNode.Attributes.getNamedItem("TOTALALLOWABLEINCOME") Is Nothing Then
        dblTotalAllowableIncome = xmlApplicationNode.Attributes.getNamedItem("TOTALALLOWABLEINCOME").nodeValue
    End If
    
    If Not xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST") Is Nothing Then
        dblTotalApproximateMonthlyCost = xmlApplicationNode.Attributes.getNamedItem("TOTALAPPROXIMATEMONTHLYCOST").nodeValue
    End If
    
    'IF Passed in "TotalAllowableIncome" < "TotalApproximateMonthlyCost" THEN Fail
    If dblTotalAllowableIncome < dblTotalApproximateMonthlyCost Then
        blnSuccess = False
    End If

    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6775 = intResult

    Set xmlNode = Nothing

End Function

'========================================================
'OS 01/12/2006
'========================================================

Private Function Rule6780( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMinEmployHistMonths As Integer
    
    Dim blnSuccess As Boolean
    
    Dim ndIntroducer As MSXML2.IXMLDOMNode
            
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6780"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6780: Grey-listed broker
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
                
    Set ndIntroducer = xmlApplicationNode.selectSingleNode("./APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE_TYPE_B='true' and @LISTINGSTATUS_TYPE_GL='true']")

    If Not ndIntroducer Is Nothing Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6780 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndIntroducer = Nothing
        
End Function


'========================================================
'OS 29/11/2006
'========================================================

Private Function Rule6785( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMinEmployHistMonths As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim blnEmploymentExists As Boolean
    
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    
    Dim dtmStartedOrEst As Date
    Dim dtmLeftOrCeased As Date
    Dim dtmToday As Date
        
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6785"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 6785: Employment history < minimum period
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    intMinEmployHistMonths = CInt(GetGlobalParameter(xmlApplicationNode, "MinEmploymentHist", Amount))
    
    'Iterate through each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        blnEmploymentExists = False
    
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        If Not ndEmployment Is Nothing Then
            
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
            'Loop through all Employment records. If the customer was employed intMinEmployHistMonths ago then rule passes
            For Each ndEmployment In ndCustomerRole.selectNodes("./CUSTOMERVERSION/EMPLOYMENT[not(@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true')]")
                Set attStartedOrEst = ndEmployment.Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
                Set attLeftOrCeased = ndEmployment.Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                
                If Not (attStartedOrEst Is Nothing) Then
                    dtmStartedOrEst = attStartedOrEst.Text
                    
                    'Is the dtmStartedOrEst over intMinEmployHistMonths (12 months) ago
                    If DateDiffInMonthsRoundedDown(dtmStartedOrEst, dtmToday) >= intMinEmployHistMonths Then
                        If Not (attLeftOrCeased Is Nothing) Then
                            dtmLeftOrCeased = attLeftOrCeased.Text
                                
                            'Is the dtmLeftOrCeased more recent than intMinEmployHistMonths (12 months) ago
                            If DateDiffInMonthsRoundedDown(dtmLeftOrCeased, dtmToday) < intMinEmployHistMonths Then
                                blnEmploymentExists = True
                                Exit For
                            End If
                        Else
                            'No DATELEFTORCEASEDTRADING found so still employed
                            blnEmploymentExists = True
                            Exit For
                        End If
                    End If
                Else
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED missing for Customer EMPLOYMENT, return a fail"
                End If
            Next
                    
            If Not (blnEmploymentExists) Then
                'Valid Employment record for Customer didn't exist so fail rule and exit loop
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                Exit For
            End If
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6785 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
        
End Function

'========================================================
'OS 30/11/2006
'========================================================

Private Function Rule6790(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
            
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6790"
    
    'Rule 6790: Non-permanent employment
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        'Locate the Employment record where Employment.MainStatus = True AND Employment.EmploymentStatus = "Employed - Temporary" (combo "EmploymentStatus" ValidationType = "EMPT").
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        'IF the record is found then [FAIL].
        If Not ndEmployment Is Nothing Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
            Exit For
        End If

    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6790 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
        
End Function

'========================================================
'OS 30/11/2006
'========================================================

Private Function Rule6795(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMinTimeEmployed As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    
    Dim dtmStartedOrEst As Date
    Dim dtmLeftOrCeased As Date
    Dim dtmToday As Date
        
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6795"
    
    'Rule 6795: Time in current employment less than minimum
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    intMinTimeEmployed = CInt(GetGlobalParameter(xmlApplicationNode, "MinTimeEmployed", Amount))
    
    'Iterate through each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        If Not ndEmployment Is Nothing Then
            
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
            
            Set attStartedOrEst = ndEmployment.Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
                
            If Not (attStartedOrEst Is Nothing) Then
                If DateDiffInMonthsRoundedDown(attStartedOrEst.Text, dtmToday) < intMinTimeEmployed Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore) 'add to appropriate referral count
                    strRuleValue = "Time (months): " & CStr(DateDiffInMonthsRoundedDown(attStartedOrEst.Text, dtmToday))
                    Exit For
                End If
            Else
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED missing for Customer EMPLOYMENT, return a fail"
            End If
            
        End If
    Next
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6795 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function


'========================================================
'OS 30/11/2006
'========================================================

Private Function Rule6800(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMinEmployHistMonths As Integer
    Dim intMaxNumEmployments As Integer
    Dim intEmploymentCount As Integer
    
    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    
    Dim dtmLeftOrCeased As Date
    Dim dtmToday As Date
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6800"
    
    'Rule 6800: Number or employers greater than maximum
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    intMinEmployHistMonths = CInt(GetGlobalParameter(xmlApplicationNode, "MinEmploymentHist", Amount))
    intMaxNumEmployments = CInt(GetGlobalParameter(xmlApplicationNode, "MaxNumberEmployments", Amount))
    
    'Iterate through each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
        
        'default for each customer
        intEmploymentCount = 0
    
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
    
        If Not ndEmployment Is Nothing Then
            
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
            
            'Loop through all Employment records. if the customer was employed intMinEmployHistMonths ago then rule passes
            For Each ndEmployment In ndCustomerRole.selectNodes("./CUSTOMERVERSION/EMPLOYMENT[not(@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true')]")
                Set attLeftOrCeased = ndEmployment.Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                
                If Not (attLeftOrCeased Is Nothing) Then
                    dtmLeftOrCeased = attLeftOrCeased.Text
                    
                    'Is the dtmLeftOrCeased more recent than (or in the future) intMinEmployHistMonths (12 months) ago
                    If DateDiffInMonthsRoundedUp(dtmLeftOrCeased, dtmToday) <= intMinEmployHistMonths Then
                        intEmploymentCount = intEmploymentCount + 1
                    End If
                Else
                    'Current employment won't have an end date
                    intEmploymentCount = intEmploymentCount + 1
                End If
            Next
            
            'If there are more than intMaxNumEmployments in last intMinEmployHistMonths
            'for any customer rule fails. Exit loop
            If intEmploymentCount > intMaxNumEmployments Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                Exit For
            End If
        End If
    Next
       
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6800 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Set ndCustomerRole = Nothing
    
End Function




'======================================================================================
'Added By:Dhira Singh
'Application Rule : 6805
'Date : 30th Nov 2006
'======================================================================================
Private Function Rule6805(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim strRuleValue As String, strRuleDescription As String

    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer

    Dim ndCustomerRole As MSXML2.IXMLDOMNode
    Dim ndEmployment As MSXML2.IXMLDOMNode
    Dim lstEmployment  As MSXML2.IXMLDOMNodeList
    Dim iEmployment  As Integer
    Dim iMinEmployHistMonths As Integer
    Dim iMaxEmploymentGapWeeks As Integer
   
    
    Dim attLeftOrCeased As MSXML2.IXMLDOMAttribute
    Dim attStartedOrEst As MSXML2.IXMLDOMAttribute
    Dim dtmLeftOrCeased As Date
    Dim dtmStartedOrEst As Date
    Dim dtmToday As Date

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6805"
    
    'Rule 6805 : For employed applicants where gap between any change of employment in last 12 months > 4 weeks.
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    iMinEmployHistMonths = CInt(GetGlobalParameter(xmlApplicationNode, "MinEmploymentHist", Amount))
    iMaxEmploymentGapWeeks = CInt(GetGlobalParameter(xmlApplicationNode, "MaxEmploymentGap", Amount))
    
    'Iterate each CUSTOMERROLE
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")

        'Check there's a main employment record
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
                
        If Not ndEmployment Is Nothing Then
            
            '---------------------------------------------------
            'continue
            
            iEmployment = 0
            
            dtmToday = GetDateTodayFromIngestion(xmlApplicationNode)
    
            Set lstEmployment = ndCustomerRole.selectNodes("./CUSTOMERVERSION/EMPLOYMENT[not(@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER='true')]")
                
            
            
            'Check gap between today and dtmLeftOrCeased of most recent employment
            Set attLeftOrCeased = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
            If Not (attLeftOrCeased Is Nothing) Then
                dtmLeftOrCeased = CDate(attLeftOrCeased.Text)
                
                'Is Gap between Today and last dtmLeftOrCeased more then 4 weeks?
                If TrueDateDiff("w", dtmLeftOrCeased, dtmToday) > iMaxEmploymentGapWeeks Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore) 'add to appropriate referral count
                    Exit For
                End If
            Else
                'Current employment so there's no DateLeftOrCeasedTrading
                dtmLeftOrCeased = dtmToday
            End If
            
            'Get dtmStartedOrEst for most recent employment
            Set attStartedOrEst = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
            If Not (attStartedOrEst Is Nothing) Then
                dtmStartedOrEst = CDate(attStartedOrEst.Text)
                iEmployment = iEmployment + 1
                
                'Continue if we're still processing records within the last MinEmploymentHist months else exit
                Do While (iEmployment <= (lstEmployment.length - 1)) And (DateDiffInMonthsRoundedUp(dtmLeftOrCeased, dtmToday) <= iMinEmployHistMonths)
                    
                    Set attLeftOrCeased = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATELEFTORCEASEDTRADING")
                    If Not (attLeftOrCeased Is Nothing) Then
                        dtmLeftOrCeased = CDate(attLeftOrCeased.Text)
                        
                        If DateDiffInMonthsRoundedUp(dtmLeftOrCeased, dtmToday) <= iMinEmployHistMonths Then
                            'Is Gap between the previous Employment dtmStartedOrEst and the current
                            'Employment dtmLeftOrCeased > iMaxEmploymentGapWeeks (weeks)
                            If TrueDateDiff("w", dtmLeftOrCeased, dtmStartedOrEst) > iMaxEmploymentGapWeeks Then
                                intScore = intRuleValidationType
                                intResult = RiskAssessmentDecision.Referral
                                CountReferral (intScore) 'add to appropriate referral count
                                Exit For
                            End If
                        End If
                    End If
                    
                    'Get dtmStartedOrEst of current Employment record
                    Set attStartedOrEst = lstEmployment.Item(iEmployment).Attributes.getNamedItem("DATESTARTEDORESTABLISHED")
                    If Not (attStartedOrEst Is Nothing) Then
                        dtmStartedOrEst = CDate(attStartedOrEst.Text)
                    End If
                    
                    iEmployment = iEmployment + 1
                Loop
            
            Else
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute missing: DATESTARTEDORESTABLISHED, return a fail"
            End If
        End If
    Next

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    Rule6805 = intResult

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Set ndCustomerRole = Nothing
    Set lstEmployment = Nothing

End Function
'======================================================================================
'End: App Rule 6805
'======================================================================================

'======================================================================================
'Added By:Dhira Singh
'Application Rule : 6810
'Date : 29th Nov 2006
'======================================================================================
Private Function Rule6810(ByVal xmlApplicationNode As IXMLDOMNode, ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    
    Dim ndEmployment As IXMLDOMNode
    Dim ndEmployedDetails As IXMLDOMNode
    Dim ndCustomerRole As IXMLDOMNode
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6810"
    
    'RULE 6810: For employed applicants where outstanding probationary period in current employment.
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
   ' LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
       
    blnSuccess = True
       
    For Each ndCustomerRole In xmlApplicationNode.selectNodes("./CUSTOMERROLE")
    
        'Locate the Employment record where Employment.MainStatus = True AND Employment.EmploymentStatus = "Employed " (combo "EmploymentStatus" ValidationType = "EMP").
        Set ndEmployment = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1' and @EMPLOYMENTSTATUS_TYPE_EMP='true' or @EMPLOYMENTSTATUS_TYPE_EMPT='true']")
       
       'For the employment record, locate the associated EmploymentDetails record.
        Set ndEmployedDetails = ndCustomerRole.selectSingleNode("./CUSTOMERVERSION/EMPLOYMENT/EMPLOYEDDETAILS[@PROBATIONARYINDICATOR='1']")
        
    
        'IF EmployedDetails.ProbationaryIndicator is TRUE THEN [FAIL]
        If Not ndEmployment Is Nothing Then
            If Not ndEmployedDetails Is Nothing Then
                blnSuccess = False
                Exit For
            End If
        End If
    Next
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
        
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    'LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule6810 = intResult
    
End Function
'======================================================================================
'End: App Rule 6810
'======================================================================================


Private Function Rule6815( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean, _
        blFreeLegalFees As Boolean
        
    Dim strScoreCardInd As String

    Dim intFreeLegalFees As Integer
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6815"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6815: Free Legals Product Selected
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'Locate all LoanComponent records for the active quote
    'Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT/MORTGAGEPRODUCT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']")
    'JD EP2_1282 correct search string
    Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0' and @PRODUCTSWITCHRETAINPRODUCTIND='0']/MORTGAGEPRODUCT")

    'FOR each mortgage product, search for freelegalfees = 1
    For Each xmlNode In xmlNodeList
        If Not xmlNode.Attributes.getNamedItem("FREELEGALFEES") Is Nothing Then
            intFreeLegalFees = xmlNode.Attributes.getNamedItem("FREELEGALFEES").nodeValue
        End If
        
        If intFreeLegalFees = 1 Then
            blFreeLegalFees = True
        End If
    Next xmlNode
    
    'IF freelegalfees = 1 found then [FAIL].
    If blFreeLegalFees Then
        blnSuccess = False
    End If

    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6815 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6820( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponentsFA As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6820"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6820: Total number of components greater than maximum
    'Multi-component loan and invalid number of components for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") then do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponentsFA = GetGlobalParameter(xmlApplicationNode, "MaxComponentsFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False.
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsFA" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponentsFA Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6820 = intResult

    Set xmlNode = Nothing

End Function















Private Function Rule6825( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponentsNew As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6825"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6825: Number of C&I components greater than maximum
    'Multi-component loan and number of C&I components greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponentsNew = GetGlobalParameter(xmlApplicationNode, "MaxComponentsCINew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsCINew" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponentsNew Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6825 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6830( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponents As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6830"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6830: Number of IO components greater than maximum

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponents = GetGlobalParameter(xmlApplicationNode, "MaxComponentsIONew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsIONew" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponents Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6830 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6835( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponents As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6835"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6835: Number of C&I components greater than maximum
    'Multi-component loan and number of C&I components greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponents = GetGlobalParameter(xmlApplicationNode, "MaxComponentsCIFA", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsCIFA" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponents Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6835 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6840( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxComponents As Integer, _
        intLoanCount As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6840"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6840: Number of IO components greater than maximum
    'Multi-component loan and number of IO components greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxComponents = GetGlobalParameter(xmlApplicationNode, "MaxComponentsIOFA", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        If Not xmlNodeList Is Nothing Then
            intLoanCount = xmlNodeList.length
        End If

        'IF number of LoanComponent records located > global parameter "MaxComponentsIOFA" then set number = number of LoanComponent records and [FAIL].
        If intLoanCount > intMaxComponents Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6840 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6845( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6845"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6845: Number of C&I terms greater than maximum
    'Multi-component loan and number of C&I terms greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsCINew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsCINew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6845 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6850( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6850"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6850: Number of IO terms greater than maximum
    'Multi-component loan and number of IO terms greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsIONew", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsIONew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6850 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6855( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6855"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6855: Number of C&I terms greater than maximum
    'Multi-component loan and number of C&I terms greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsCIFA", Amount)

        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsCIFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6855 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6860( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        intLoanTermInMonths As Integer, _
        intLoanTermInYears As Integer, _
        intPreviousLoanTermInYears As Integer, _
        intPreviousLoanTermInMonths As Integer

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6860"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6860: Number of IO terms greater than maximum
    'Multi-component loan and number of IO terms greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxTermsIOFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_TERM[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']")
        
        'FOR each LoanComponent, examine LoanComponent.TermInYears and LoanComponent.TermInMonths and count the number of different terms found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("TERMINYEARS") Is Nothing Then
                intLoanTermInYears = xmlNode.Attributes.getNamedItem("TERMINYEARS").nodeValue
            End If

            If Not xmlNode.Attributes.getNamedItem("TERMINMONTHS") Is Nothing Then
                intLoanTermInMonths = xmlNode.Attributes.getNamedItem("TERMINMONTHS").nodeValue
            End If
            
            If intPreviousLoanTermInYears <> intLoanTermInYears Or _
            intPreviousLoanTermInMonths <> intLoanTermInMonths Then
                intLoanCount = intLoanCount + 1
            End If
            
            intPreviousLoanTermInYears = intLoanTermInYears
            intPreviousLoanTermInMonths = intLoanTermInMonths
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxTermsIOFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6860 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6865( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6865"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6865: Number of C&I products greater than maximum
    'Multi-component loan and number of C&I products greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsCINew", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsCINew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6865 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6870( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6870"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6870: Number of IO products greater than maximum
    'Multi-component loan and number of IO products greater than maximum allowed for initial advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication not = "Purchase" AND not = "Remortgage" AND not = "RTB Purchase" AND not = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType not = "PRCH" AND not = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing Or _
    Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsIONew", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsIONew" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6870 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6875( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6875"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6875: Number of C&I products greater than maximum
    'Multi-component loan and number of C&I products greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsCIFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Capital & Interest" (combo "RepaymentType" ValidationType = "C")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_C='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsCIFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6875 = intResult

    Set xmlNode = Nothing

End Function

Private Function Rule6880( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) As Integer

    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList

    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer, _
        intMaxTerms As Integer, _
        intLoanCount As Integer, _
        intDifferentTerms As Integer, _
        strPreviousProductCode As String, _
        strProductCode As String

    Dim strRuleValue As String, _
        strRuleDescription As String

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String

    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "6880"
    strScoreCardInd = 0 'False

    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))

    'RULE 6880: Number of IO products greater than maximum
    'Multi-component loan and number of IO products greater than maximum allowed for further advance
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True

    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.All)

    'IF the Quotation is not found then do not run the rule
    'and IF ApplicationFactFind.TypeOfApplication = "Purchase" OR = "Remortgage" OR = "RTB Purchase" OR = "RTB Remortgage" (combo "TypeOfMortgage" ValidationType = "PRCH" OR = "R") THEN do not run the rule.
    If intMortgageSubquoteNumber > 0 And _
    (xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_PRCH") Is Nothing And _
    xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing) Then

        'Get max components
        intMaxTerms = GetGlobalParameter(xmlApplicationNode, "MaxProductsIOFA", Amount)
        
        'Locate all LoanComponent records where LoanComponent.PortedLoan = False AND LoanComponent.ProductSwitchRetainProductInd = False AND  LoanComponent.RepaymentMethod = "Interest Only" (combo "RepaymentType" ValidationType = "I")
        Set xmlNodeList = xmlApplicationNode.selectNodes("./QUOTATION/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intMortgageSubquoteNumber & "]" & "/LOANCOMPONENT_ORDERBY_MORTGAGEPRODUCTCODE[@MANUALPORTEDLOANIND='0'][@PRODUCTSWITCHRETAINPRODUCTIND='0'][@REPAYMENTMETHOD_TYPE_I='true']/MORTGAGEPRODUCT")

        'FOR each LoanComponent, examine LoanComponent.MortgageProductCode and count the number of different products found.
        For Each xmlNode In xmlNodeList
            If Not xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE") Is Nothing Then
                strProductCode = xmlNode.Attributes.getNamedItem("MORTGAGEPRODUCTCODE").nodeValue
            End If
            
            If strPreviousProductCode <> strProductCode Then
                intLoanCount = intLoanCount + 1
            End If
            
            strPreviousProductCode = strProductCode
        Next xmlNode
        
        'IF number of terms found > global parameter "MaxProductsIOFA" then [FAIL].
        If intLoanCount > intMaxTerms Then
            blnSuccess = False
        End If

        If Not blnSuccess Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore) 'add to appropriate referral count
        End If

        strRuleValue = "No: " & intLoanCount
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult

    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult

    Rule6880 = intResult

    Set xmlNode = Nothing

End Function


