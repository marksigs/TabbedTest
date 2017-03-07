VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValuationRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
'PROG   Date            Description
'LH     13-Nov-2006     Creation
'LH     16/11/2006      EP2_78  - Added Mortgage Type Rule Triggers facility
'LH     17/11/2006      EP2_129 - Merged Dhira & Ovais' valuation rules.
'DS     17/11/2006      EP2_129 - Added & modified construction rules.
'MCh    20/11/2006      EP2_129 - added function GetCurrentCompletedValuationElement()
'                                 Rules 8195
'DS     20/11/2006      EP2_129 - Modified construction rules and added Legal Issues Rules
'MCh    20/11/2006      EP2_129 - Rules added: 8200, 8210, 8215, 8220, modified 8005
'OS     21/11/2006      EP2_129 - Modified Property type and Tenure rules and added Locality rules
'MCh    21/11/2006      EP2_129 - Rules 8180, 8185, 8190
'OS     21/11/2006      EP2_129 - Updated RunRulesValuation
'DS     21/11/2006      EP2_129 - Modified construction rules and Legal Issues Rules
'OS     22/11/2006      EP2_129 - Finalized New Build rules and added funtion calls
'DS     22/11/2006      EP2_129 - Modified Rule 8065 and added Rules 8155 and 8160
'MCh    22/11/2006      EP2_129 -
'LH     06/12/2006      EP2_129 - Errors now correctly handled back to calling RARulesBO component
'LH     06/12/2006      EP2_332 - Changed values for Enum RiskAssessmentDecision
'OS     07/12/2006      EP2_332 - Removed redundant comment from Rule 8025, and added null checks
'LH     08/12/2006      EP2_332 - Changed raiseerror() calls in rules to a fail and writing a line to the log file
'LH     08/12/2006      EP2_332 - Added extra xml node check to 8035, 8205, 8215
'LH     18/12/2006      EP2_547 - Fixed rule 8205
'LH     19/12/2006      EP2_578 - Fixed rule 8205
'LH     20/12/2006      EP2_602 - Valuation Rules Rework
'LH     18/01/2007      EP2_909 - Fixed rule 8035
'LH     18/01/2007      EP2_910 - Fixed rule 8120
'LH     24/01/2007      EP2_971 - Fixed logic around calling valuation rules
'LH     29/01/2007      EP2_759 - Fixed GetCurrentCompletedValuationElement()
'LH     31/01/2007      EP2_1149 - Fixed rule 8175
'LH     02/01/2007      EP2_1185 - Fixed rule 8210, 8215, 8220
'LH     22/02/2007      EP2_1607 - Changed rule 8230 according to spec change
'OS     02/03/2007      EP2_1740 - Changed rule 8185 according to spec change
'LDM    07/03/2007      EP2_1648 - Use the stagesequence no from the stage table to know how far the case has progressed
'-----------------------------------------------------------------------------------

Option Explicit

Private gstrRuleNumber As String
Private gxmlRequest As MSXML2.DOMDocument40

Private gintTermAsMonths As Integer
Private gintUnderWriterCounter As Integer
Private gintProcessorCounter As Integer


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

Private Enum ValuationDataNode
    eVALUERINSTRUCTION = 1
    eVALNREPVALUATION = 2
    eVALNREPPROPERTYRISKS = 3
    eVALNREPPROPERTYDETAILS = 4
    eVALNREPPROPERTYSERVICES = 5
    eVALNREPSUMMARY = 6
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

Public Function RunRulesValuation( _
    ByVal xmlrequest As DOMDocument40, _
    ByRef xmlResponseNode As IXMLDOMNode, _
    ByRef intFailLevel As Integer, _
    ByRef intUnderWriterCounter As Integer, _
    ByRef intProcessorCounter As Integer, _
    ByRef strRuleNumber As String, _
    ByRef bErrFlag As Boolean, _
    ByRef intErrNumber As Integer, _
    ByRef strErrDescription As String) _
    As Integer
    
    Dim xmlApplicationNode As IXMLDOMNode
    Dim blRule8230 As Boolean
    Dim intResult As Integer
    
    On Error GoTo err_RunRulesValuation
    
    'This function is called from RARulesBO.RunRulesValuation()
    
    'as application rules may of already been run in RARulesBO, we need to set
    'the new underwriter counter and processor counter to the values that
    'were set in RARulesBO, then return the new count back to RARulesBO
    'after the valuation rules have finished running.
    'intFailLevel is passed into this function as a parameter for the same reason.
    gintUnderWriterCounter = intUnderWriterCounter
    gintProcessorCounter = intProcessorCounter
    gstrRuleNumber = strRuleNumber
    
    Set gxmlRequest = xmlrequest
    
    Set xmlApplicationNode = gxmlRequest.selectSingleNode("REQUEST/APPLICATION")
    
    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "START: RunRulesValuation()"
    
    'only run 8230 if Stage >= "Data Verification"
    If Not StageLessThanValnStage(xmlApplicationNode) Then
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8230) Then
            intResult = Rule8230(xmlApplicationNode, xmlResponseNode)
            'if rule 8230 was successful, flag it
            If intResult = RiskAssessmentDecision.Accept Then
                blRule8230 = True
            End If
            intFailLevel = EpsomMax(intFailLevel, intResult)
        End If
    End If
    
    'only run valuation rules if (case stage < valuation stage) OR
    '(valuation stage >= data verification stage AND valuation report exists)
    If StageLessThanValnStage(xmlApplicationNode) Or blRule8230 Then
        'Insert Below: The 12 or so rules here that can run before, during & after the data verification stage
        'They are marked as: App Rule? "YES" in the valuation rules spec (xls)

        'Rule 8000
        'Property type = "Back-to-Back" and LTV > 85%
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8000) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8000(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8005: Property type = "Converted Studio Flat"
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8005) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8005(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8010: Property type = "Other"
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8010) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8010(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8015: Property type = ("Purpose Built Flat" or "Converted Flat" or "Studio Flat") and number of floors in block = 0
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8015) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8015(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8025: Number of kitchens not = 1
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8025) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8025(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8035 : New build & builder's certificate not acceptable
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8035) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8035(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8040 : Wall construction is not conventional
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8040) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8040(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8045 : Roof construction is not conventional tile/slate
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8045) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8045(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8070: Tenure = "Freehold" and property type = "Purpose Built Flat" or "Converted Flat" or "Maisonette" or "Studio Flat"
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8070) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8070(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8075: Tenure = "Leasehold" and lease term remaining at end of mortgage term < 40 years
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8075) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8075(xmlApplicationNode, xmlResponseNode))
        End If
        
        'Rule 8120 : Plot is greater than 10 acres
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8120) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8120(xmlApplicationNode, xmlResponseNode))
        End If
            
        'Rule 8140: Above commercial premises = "Yes"
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8140) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8140(xmlApplicationNode, xmlResponseNode))
        End If
                
        'Rule 8205: BTL rental cover less than minimum allowed LTV
        If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8205) Then
            intFailLevel = EpsomMax(intFailLevel, Rule8205(xmlApplicationNode, xmlResponseNode))
        End If
        
        'only run these rules if a valuation report exists and
        'current stage of the case >= "Data Verification"
        If blRule8230 Then
            'Insert Below: All the other valuation rules
            'They are marked as: App Rule? "NO" in the valuation rules spec (xls)
            ' ...
            
            '==============================================
            'Added by Ovais Shah
            '==============================================
            'Rule 8020: Property type = ("Purpose Built Flat" or "Converted Flat" or "Studio Flat")
            'and number of floors in block >= 5 and ex-local authority = "Yes"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8020) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8020(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8030: Balcony access = "Yes"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8030) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8030(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8125: Property type = ("Purpose Built Flat" or "Converted Flat" or "Maisonette" or "Studio Flat") and ex-local authority = "Yes"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8125) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8125(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8130: Proportion of properties in private ownership on local authority development < 75%
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8130) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8130(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8135: Boarded-up or damaged properties in the local area = "Yes"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8135) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8135(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8145: Property in a mining area = "Yes"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8145) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8145(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8150: Property affected by radon gas = "Yes" and taken into account in valuation = "No"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8150) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8150(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8165: Tenanted residential property for remortgage
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8165) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8165(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8170: Property affected by radon gas = "Yes" and taken into account in valuation = "No"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8170) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8170(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8175: PP includes sales incentive for the new build
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8175) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8175(xmlApplicationNode, xmlResponseNode))
            End If
            '==============================================
            
            '======================================================================================
            'Added By Dhira Singh: Valuation rules based on construction
            '======================================================================================
            
            'Rule 8050 : Structural issues but not long-standing
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8050) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8050(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8055 : Property at risk of mundic problems
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8055) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8055(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8060 : Live work unit with residential element < 70%
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8060) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8060(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8065 : BTL, Live work unit with residential element > 70%
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8065) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8065(xmlApplicationNode, xmlResponseNode))
            End If
                    
            '======================================================================================
            '======================================================================================
            'Added By Dhira Singh: Valuation rules based on Legal Issues
            '======================================================================================
            'Rule 8080 : Legal issue "Rights of Way" = "Yes"
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8080) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8080(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8085 : Property has Shared Drive / Access issues
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8085) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8085(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8090 : Property has garage on separate site
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8090) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8090(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8095 : Property has flying freehold issue
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8095) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8095(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8100 :Property has ill-defined boundaries
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8100) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8100(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8105 : Surveyor to inspect title plans before Valn
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8105) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8105(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8110 : Legal issue "Other" = "Yes".
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8110) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8110(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8115 : Legal issue "Shared Services" = "Yes".
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8115) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8115(xmlApplicationNode, xmlResponseNode))
            End If
                
            '======================================================================================
            
            '======================================================================================
            'Added By Dhira Singh: Valuation rules based on Essential Repairs
            '======================================================================================
            'Rule 8155 : Recommended retention amount more than £1000.
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8155) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8155(xmlApplicationNode, xmlResponseNode))
            End If
            '======================================================================================
            '======================================================================================
            'Added By Dhira Singh: Valuation rules based on Alterations & Extensions
            '======================================================================================
            'Rule 8160 : Alterations/extensions require planning permission.
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8160) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8160(xmlApplicationNode, xmlResponseNode))
            End If
            '======================================================================================
            
            'Saleability
            
            'Rule 8180 : Property doesn’t represent suitable security
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8180) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8180(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8185 : Poor demand for property
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8185) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8185(xmlApplicationNode, xmlResponseNode))
            End If
                
            'Rule 8190 : Factors affect future saleability
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8190) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8190(xmlApplicationNode, xmlResponseNode))
            End If
            
            '========================================================================================
            'But-to-let Mortgages
    
            'Rule 8195 : Poor rental demand for this BTL property
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8195) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8195(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8200: BTL property not suitable for letting
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8200) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8200(xmlApplicationNode, xmlResponseNode))
            End If
            
            '========================================================================================
            'Valuation
            
            'Rule 8210 : Property valn not same to purchase/estimated price
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8210) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8210(xmlApplicationNode, xmlResponseNode))
            End If
            
            '========================================================================================
            'Important Information
            
            'Rule 8215 : General remarks is not empty
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8215) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8215(xmlApplicationNode, xmlResponseNode))
            End If
            
            '========================================================================================
            'Valuer's Details
            
            'Rule 8220 :
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8220) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8220(xmlApplicationNode, xmlResponseNode))
            End If
            
            'Rule 8225: Valuer details not found
            If RunRule(RuleSet.ValnSet, xmlApplicationNode, 8225) Then
                intFailLevel = EpsomMax(intFailLevel, Rule8225(xmlApplicationNode, xmlResponseNode))
            End If
                
        End If
    
    End If
    
    'make sure the accumulated counters are returned back to RARulesBO
    intUnderWriterCounter = gintUnderWriterCounter
    intProcessorCounter = gintProcessorCounter
    strRuleNumber = gstrRuleNumber
    
    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "END: RunRulesValuation() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intFailLevel
    
    RunRulesValuation = intFailLevel
    
    Exit Function
    
err_RunRulesValuation:
    strRuleNumber = gstrRuleNumber
    bErrFlag = True
    intErrNumber = Err.number
    strErrDescription = Err.Description
End Function

Private Function ValnReportExists(ByVal xmlApplicationNode As IXMLDOMNode) As Boolean
    'Purpose: Checks whether a valuation report exists
    '         returns true or false accordingly

    Dim xmlNode As IXMLDOMNode
    Dim blSuccess As Boolean

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: ValnReportExists()"

    blSuccess = True

    Set xmlNode = xmlApplicationNode.selectSingleNode("./VALUERINSTRUCTION[position()=1]")
    If xmlNode Is Nothing Then
        blSuccess = False
    End If

    'log details to log file
    LogDetails 1, String(2, Chr(9)) & "END: ValnReportExists() returned: " & blSuccess

    ValnReportExists = blSuccess

    Set xmlNode = Nothing

End Function


Private Sub CountReferral(ByVal intScore As Integer)
'decides the referral type and adds to the appropriate counter

If intScore = RiskAssessmentReferralType.Processor Then
    gintProcessorCounter = gintProcessorCounter + 1
ElseIf intScore = RiskAssessmentReferralType.Underwriter Then
    gintUnderWriterCounter = gintUnderWriterCounter + 1
End If

End Sub
Private Function GetMortgageSubQuoteNumber(eQuoteType As QuoteType) As Integer
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intAcceptedQuoteNumber As Integer, _
        intMortgageSubquoteNumber As Integer
    
    LogDetails 1, String(2, Chr(9)) & "START: GetMortgageSubQuoteNumber(eQuoteType=" & eQuoteType & ") 1=accepted quote, 2=active, 3=either"
  
    If eQuoteType = Accepted Or eQuoteType = All Then
        'Get Accepted Quote Number
        Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION[@ACCEPTEDQUOTENUMBER]")
        If xmlNodeList.length > 0 Then
            intAcceptedQuoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("ACCEPTEDQUOTENUMBER").Text
        End If
    End If
    
    If eQuoteType = Active Or eQuoteType = All Then
        'if quote type is 'all' and accepted quote number was not found then find the active quote number
        If Len(intAcceptedQuoteNumber) = 0 Then
            'Get Active Quote Number
            Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION[@ACTIVEQUOTENUMBER]")
            If xmlNodeList.length > 0 Then
                intAcceptedQuoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("ACTIVEQUOTENUMBER").Text
            End If
        End If
    End If
    
    'Get Mortgage Subquote Number
    Set xmlNodeList = gxmlRequest.selectNodes("REQUEST/APPLICATION/QUOTATION[@QUOTATIONNUMBER=" & intAcceptedQuoteNumber & "]")
    If xmlNodeList.length > 0 Then
        intMortgageSubquoteNumber = xmlNodeList.Item(0).Attributes.getNamedItem("MORTGAGESUBQUOTENUMBER").Text
    End If
    
    LogDetails 1, String(2, Chr(9)) & "END: GetMortgageSubQuoteNumber() returned: " & intMortgageSubquoteNumber

    GetMortgageSubQuoteNumber = intMortgageSubquoteNumber

End Function

' Returns the correct valuation element for the max sequenceno completed valuation
' for the following:
'       VALUERINSTRUCTION
'       VALNREPVALUATION
'       VALNREPPROPERTYRISKS
'       VALNREPPROPERTYDETAILS
'       VALNREPPROPERTYSERVICES
Private Function GetCurrentCompletedValuationElement(ByVal xmlApplicationNode As MSXML2.IXMLDOMNode, _
                                                    ByVal eValuationDataNode As ValuationDataNode) As MSXML2.IXMLDOMNode

    Dim lstCompletedInstruction As MSXML2.IXMLDOMNodeList
    Dim ndValuerInstruction As MSXML2.IXMLDOMNode
    Dim iMaxSequenceNo As Integer
    Dim ndReturn As MSXML2.IXMLDOMNode
    
    'Get completed ValuerInstruction nodes
    Set lstCompletedInstruction = xmlApplicationNode.selectNodes("./VALUERINSTRUCTION[@VALUATIONSTATUS='30']")
    
    If Not (lstCompletedInstruction Is Nothing) Then
    
        'ValuerInstruction nodes are in desc order so get first for MaxSequenceNo
        Set ndValuerInstruction = lstCompletedInstruction.Item(0)
           
        If Not (ndValuerInstruction Is Nothing) Then
    
            iMaxSequenceNo = ndValuerInstruction.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO").Text
    
            Select Case UCase(eValuationDataNode)
                Case eVALUERINSTRUCTION
                    Set ndReturn = ndValuerInstruction
                Case eVALNREPVALUATION
                    Set ndReturn = xmlApplicationNode.selectSingleNode("./VALNREPVALUATION[@INSTRUCTIONSEQUENCENO='" & iMaxSequenceNo & "']")
                Case eVALNREPPROPERTYRISKS
                    Set ndReturn = xmlApplicationNode.selectSingleNode("./VALNREPPROPERTYRISKS[@INSTRUCTIONSEQUENCENO='" & iMaxSequenceNo & "']")
                Case eVALNREPPROPERTYDETAILS
                    Set ndReturn = xmlApplicationNode.selectSingleNode("./VALNREPPROPERTYDETAILS[@INSTRUCTIONSEQUENCENO='" & iMaxSequenceNo & "']")
                Case eVALNREPPROPERTYSERVICES
                    Set ndReturn = xmlApplicationNode.selectSingleNode("./VALNREPPROPERTYSERVICES[@INSTRUCTIONSEQUENCENO='" & iMaxSequenceNo & "']")
                Case eVALNREPSUMMARY
                    Set ndReturn = xmlApplicationNode.selectSingleNode("./VALNREPSUMMARY[@INSTRUCTIONSEQUENCENO='" & iMaxSequenceNo & "']")
                Case Else
                    Set ndReturn = Nothing
            End Select
            
        End If
    
    End If
    
    Set GetCurrentCompletedValuationElement = ndReturn
    
End Function
Private Function EpsomMax(CurrentLevel As Integer, NewLevel As Integer) As Integer
    'returns highest level of RiskAssessmentDecision
    If NewLevel > CurrentLevel Then
        EpsomMax = NewLevel
    Else
        EpsomMax = CurrentLevel
    End If

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
                        
                        'If validation type is 'OTHER' and there are no other validation types for the rule
                        'then we exit the loop and indicate that the rule should be
                        'attempted to be run; subject to any other limiting conditions (after this function has been called).
                        If strRuleValidationType = "OTHER" And xmlRuleAppTypes.length = 1 Then
                            blSuccess = True
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


Private Sub RaiseError(ByVal vstrDescription)

    Err.Raise vbObjectError + 512 + 999, "omRARulesBO", vstrDescription

End Sub

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
'LDM    07/03/2007  EP2_1648 - Use the stagesequence no from the stage table to know how far the case has progressed
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

Private Function StageLessThanValnStage(ByVal xmlApplicationNode As IXMLDOMNode) As Boolean
    
    'Decides whether the case stage in the request XML is less than the valuation report stage
    
    Dim bLessThanValnStage As Boolean
    Dim xmlNodeList As IXMLDOMNodeList
    Dim intValuationReportStage As Integer
    Dim intStageSeqNo As Integer

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: StageLessThanValnStage()"
    
    Set xmlNodeList = xmlApplicationNode.selectNodes("GLOBALPARAMETER[@NAME=""ValuationReportStage""]")
    If xmlNodeList.length > 0 Then
        If Not xmlNodeList.Item(0).Attributes.getNamedItem("AMOUNT") Is Nothing Then
            intValuationReportStage = xmlNodeList.Item(0).Attributes.getNamedItem("AMOUNT").nodeValue
        End If
    Else
        RaiseError "Globalparameter 'ValuationReportStage' not found"
    End If

    'LDM 07/03/2007 EP2_1648
    intStageSeqNo = GetCaseStageSeqNo
    
    If intStageSeqNo > 0 And intStageSeqNo < intValuationReportStage Then
        bLessThanValnStage = True
    End If

    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: StageLessThanValnStage() returned: " & bLessThanValnStage

    StageLessThanValnStage = bLessThanValnStage
    
End Function

Private Function GetGlobalParameter(ByVal xmlApplicationNode As IXMLDOMNode, _
                                    ByVal sName As String, _
                                    ByVal eGlobalParameterType As GlobalParameterType) As String
   
    Dim xmlNode As IXMLDOMNode
    Dim sParamValue As String
    Dim sGlobalParameterType As String
    
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

'========================================================================
'Start: Property type rules
'========================================================================

Private Function Rule8000( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intMortgageSubquoteNumber As Integer
        
    Dim strRuleValue As String, _
        strPropertyType As String, _
        strTenureType As String, _
        strRuleDescription As String

    Dim dblMaxLTVBacktoBack As Double, _
        dblLTV As Double, _
        dblAmountRequested As Double, _
        dblPurchPriceEstVal As Double

    Dim blnSuccess As Boolean

    Dim strScoreCardInd As String
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8000"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8000: Back to Back Housing
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    blnSuccess = True
        
    intMortgageSubquoteNumber = GetMortgageSubQuoteNumber(QuoteType.Accepted)
    
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
    
    'Get global parameter
    dblMaxLTVBacktoBack = GetGlobalParameter(xmlApplicationNode, "LPMaxLTVBacktoBack", Percentage)
        
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
    
        'IF NewProperty.DescriptionOfProperty from Combo PropertyDescription = "Back-to-Back" (ValidationType = "BB") AND LTV > LPMaxLTVBacktoBack [FAIL]
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_BB") Is Nothing Then
                If dblLTV > dblMaxLTVBacktoBack Then
                    blnSuccess = False
                End If
            End If
        End If
    Else 'rule 8230 passed
       'Call function GetCurrentCompletedValuationElement
       Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
       'Check whether a node is returned
       If Not xmlNode Is Nothing Then
           'IF NewProperty.DescriptionOfProperty from Combo PropertyDescription = "Back-to-Back" (ValidationType = "BB") AND LTV > LPMaxLTVBacktoBack [FAIL]
           If Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_BB") Is Nothing Then
               If dblLTV > dblMaxLTVBacktoBack Then
                   blnSuccess = False
               End If
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
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=referral): " & intResult
    
    Rule8000 = intResult
    
    Set xmlNode = Nothing

End Function

Private Function Rule8005( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8005"
    
    'Rule 8005: Property type = "Converted Studio Flat"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_F") Is Nothing _
            And Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_CS") Is Nothing Then
        
                blnSuccess = False
            
            End If
        End If
        
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
       Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
       'Check whether a node is returned
       If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_F") Is Nothing _
            And Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_CS") Is Nothing Then
        
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
    
    Rule8005 = intResult

End Function


Private Function Rule8010( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8010"
    
    'Rule 8010: Property type = "Other"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_O") Is Nothing Then
                blnSuccess = False
            End If
        End If
        
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
       'Check whether a node is returned
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_O") Is Nothing Then
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
    
    Rule8010 = intResult

End Function


Private Function Rule8015( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intFloorsInBlock As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8015"
    
    'Rule 8015: Property type = ("Purpose Built Flat" or "Converted Flat" or "Studio Flat") and number of floors in block = 0.
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_F") Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("NUMBEROFSTOREYS") Is Nothing Then
                    intFloorsInBlock = CInt(xmlNode.Attributes.getNamedItem("NUMBEROFSTOREYS").Text)
                End If
                If (Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_C") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_PB") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_CS") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_PBS") Is Nothing) _
                And intFloorsInBlock = 0 Then
                    blnSuccess = False
                End If
            End If
        End If
    Else 'rule 8230 passed
       'Call function GetCurrentCompletedValuationElement
       Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
       'Check whether a node is returned
       If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_F") Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("STOREYSINBLOCK") Is Nothing Then
                    intFloorsInBlock = CInt(xmlNode.Attributes.getNamedItem("STOREYSINBLOCK").Text)
                End If
                If (Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_C") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PB") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_CS") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PBS") Is Nothing) _
                And intFloorsInBlock = 0 Then
                    blnSuccess = False
                End If
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
    
    Rule8015 = intResult

End Function


Private Function Rule8020( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intFloorsInBlock As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8020"
    
    'Rule 8020: Property type = ("Purpose Built Flat" or "Converted Flat" or "Studio Flat") and number of floors in block >= 5 and ex-local authority = "Yes".
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_F") Is Nothing Then
            
            If Not xmlNode.Attributes.getNamedItem("STOREYSINBLOCK") Is Nothing Then
                intFloorsInBlock = CInt(xmlNode.Attributes.getNamedItem("STOREYSINBLOCK").Text)
            End If
            
            If (Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_C") Is Nothing _
            Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PB") Is Nothing _
            Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_CS") Is Nothing _
            Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PBS") Is Nothing) _
            And intFloorsInBlock >= 5 _
            And Not xmlNode.Attributes.getNamedItem("EXLOCALAUTHORITY") Is Nothing Then
                If xmlNode.Attributes.getNamedItem("EXLOCALAUTHORITY").Text = "1" Then
                blnSuccess = False
                End If
            End If
        End If
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
        
    strRuleValue = "Floors in block: " & intFloorsInBlock
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8020 = intResult

End Function


Private Function Rule8025( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intKitchens As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8025"
    
    'Rule 8025: Number of kitchens not = 1.
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTYROOMTYPE[@ROOMTYPE_TYPE_K]")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("NUMBEROFROOMS") Is Nothing Then
                intKitchens = CInt(xmlNode.Attributes.getNamedItem("NUMBEROFROOMS").Text)
            End If
        End If
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
        'Check whether a node is returned
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("NUMBEROFKITCHENS") Is Nothing Then
                intKitchens = CInt(xmlNode.Attributes.getNamedItem("NUMBEROFKITCHENS").Text)
            End If
        End If
    End If
    
    'Number of kitchens not = 1
    If intKitchens <> 1 Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
        
    strRuleValue = "Kitchens: " & intKitchens
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8025 = intResult

End Function


Private Function Rule8030( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intKitchens As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8030"
    
    'Rule 8030: Balcony access = "Yes".
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("BALCONYACCESS") Is Nothing Then
            If xmlNode.Attributes.getNamedItem("BALCONYACCESS").Text = "1" Then
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
    
    Rule8030 = intResult

End Function

'========================================================================
'End: Property type rules
'========================================================================

'===============================================================================================
'Begin : Construction Rules Code
'Added By : Dhira Singh
'===============================================================================================

Private Function Rule8035( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intSignatureReturned As Integer
    Dim intNewPropertyInd As Integer

    Dim strRuleValue As String
    Dim strRuleDescription As String
            
    Dim blnSuccess As Boolean
        
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8035"
    
    
    'Rule 8035: New build = "Yes" and builder's certificate not one of the acceptable types.
    'IF ValnRepPropertyDetails.NEWPROPERTYINDICATOR  is TRUE and
    'Then IF ValnRepValuation.CertificationType from Combo
    'BuildingCertificationType <> "NHBC" OR "Zurich Municipal Building Guarantee"
    'OR "Architect Certificate" OR "Premier Guarantee"  (ValidationTypes = "N" OR "Z" OR "A OR "P")
    
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
        
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("NEWPROPERTYINDICATOR") Is Nothing Then
                intNewPropertyInd = CInt(xmlNode.Attributes.getNamedItem("NEWPROPERTYINDICATOR").nodeValue)
            End If
        
            If (intNewPropertyInd = 1) And (xmlNode.Attributes.getNamedItem("HOUSEBUILDERSGUARANTEE_TYPE_N") Is Nothing And _
                xmlNode.Attributes.getNamedItem("HOUSEBUILDERSGUARANTEE_TYPE_Z") Is Nothing And _
                xmlNode.Attributes.getNamedItem("HOUSEBUILDERSGUARANTEE_TYPE_A") Is Nothing And _
                xmlNode.Attributes.getNamedItem("HOUSEBUILDERSGUARANTEE_TYPE_P") Is Nothing) Then
                
                blnSuccess = False
            End If
        End If
    Else 'rule 8230 passed
        
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
        'Check whether a node is returned
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("NEWPROPERTYINDICATOR") Is Nothing Then
               intNewPropertyInd = CInt(xmlNode.Attributes.getNamedItem("NEWPROPERTYINDICATOR").nodeValue)
            End If
        End If
        
       'Call function GetCurrentCompletedValuationElement
       Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)
       
       'Check whether a node is returned
       If Not xmlNode Is Nothing Then
           If (intNewPropertyInd = 1) And (xmlNode.Attributes.getNamedItem("CERTIFICATIONTYPE_TYPE_N") Is Nothing And _
                xmlNode.Attributes.getNamedItem("CERTIFICATIONTYPE_TYPE_Z") Is Nothing And _
                xmlNode.Attributes.getNamedItem("CERTIFICATIONTYPE_TYPE_A") Is Nothing And _
                xmlNode.Attributes.getNamedItem("CERTIFICATIONTYPE_TYPE_P") Is Nothing) Then
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
    
    Rule8035 = intResult

End Function


Private Function Rule8040( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
        
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8040"
    
    'Rule 8040: Wall construction not = conventional brick/block/stone
    'If NewProperty.BuildingConstructionType from Combo BuildingConstruction is not = "Brick/Block/Stone (Conventional)" (Validation type = "TC")
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF Wall construction not = conventional brick/block/stone
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then 'Application data
            If xmlNode.Attributes.getNamedItem("BUILDINGCONSTRUCTIONTYPE_TYPE_TC") Is Nothing Then
                blnSuccess = False
            End If
        End If
        
    Else 'rule 8230 passed
        
       'Call function GetCurrentCompletedValuationElement
       Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
       'Check whether a node is returned
       If Not xmlNode Is Nothing Then
           If xmlNode.Attributes.getNamedItem("STRUCTURE_TYPE_TC") Is Nothing Then
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
    
    Rule8040 = intResult

End Function


Private Function Rule8045( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
          
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8045"
    
    'Rule 8045: Wall construction not = conventional brick/block/stone
    'If NewProperty.BuildingConstructionType from Combo BuildingConstruction is not = "Brick/Block/Stone (Conventional)" (Validation type = "TC")
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF Wall construction not = conventional brick/block/stone
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then 'Application data
            If xmlNode.Attributes.getNamedItem("ROOFCONSTRUCTIONTYPE_TYPE_S") Is Nothing Then
                 blnSuccess = False
            End If
        End If
    Else 'rule 8230 passed
       
       'Call function GetCurrentCompletedValuationElement
       Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
       'Check whether a node is returned
       If Not xmlNode Is Nothing Then
           If xmlNode.Attributes.getNamedItem("MAINROOF_TYPE_S") Is Nothing Then
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
    
    Rule8045 = intResult

End Function

Private Function Rule8050( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
    Dim strRoofConstructionTypeInd As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intNonProgStructIssuesInd As Integer
    Dim intStructuralIssuesInd As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8050"
    
    'Rule 8050: Structural issues = "Yes" and long-standing / non-progressive = "No
    'IF ValnRepPropertyRisks.StructuralIssues is TRUE and ValnRepPropertyRisks.NonProgressiveStructuralIssues is FALSE THEN [REFER]
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF Structural issues = "Yes" and long-standing / non-progressive = "No
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("STRUCTURALISSUES") Is Nothing Then
            intStructuralIssuesInd = CInt(xmlNode.Attributes.getNamedItem("STRUCTURALISSUES").nodeValue)
        End If
        
        If Not xmlNode.Attributes.getNamedItem("NONPROGRESSIVESTRUCTURALISSUES") Is Nothing Then
            intNonProgStructIssuesInd = CInt(xmlNode.Attributes.getNamedItem("NONPROGRESSIVESTRUCTURALISSUES").nodeValue)
        End If
    End If
    
      
    If (intStructuralIssuesInd = 1 And intNonProgStructIssuesInd = 0) Then
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
    
    Rule8050 = intResult

End Function


Private Function Rule8055( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
    Dim strRoofConstructionTypeInd As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intMundicRiskInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8055"
    
    'Rule 8055: Property at risk of mundic problems = "Yes".
    ' IF ValnRepPropertyRisks.MundicRisk is TRUE THEN [REFER]
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF Property at risk of mundic problems = "Yes".
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("MUNDICRISK") Is Nothing Then
            intMundicRiskInd = CInt(xmlNode.Attributes.getNamedItem("MUNDICRISK").nodeValue)
        End If
        
    End If
          
    If (intMundicRiskInd = 1) Then
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
    
    Rule8055 = intResult

End Function

Private Function Rule8060( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intLiveWorkUnitResElementInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8060"
    
    'Rule 8060: Live work unit with residential element > 70% = "No".
    'IF ValnRepPropertyDetails.LiveWorkUnitResElement is FALSE THEN [REFER]
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyDetails.LiveWorkUnitResElement is FALSE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("LIVEWORKUNITRESELEMENT") Is Nothing Then
            intLiveWorkUnitResElementInd = CInt(xmlNode.Attributes.getNamedItem("LIVEWORKUNITRESELEMENT").nodeValue)
        End If
        
    End If
          
    If (intLiveWorkUnitResElementInd = 0) Then
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
    
    Rule8060 = intResult

End Function


Private Function Rule8065( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
    Dim strNatureOfLoan As String
       
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intLiveWorkUnitResElementInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8065"
    
    'Rule 8065: Live work unit with residential element > 70% = "No".
    'IF ValnRepPropertyDetails.LiveWorkUnitResElement is TRUE
    'Then IF (ApplicationFactFind.NatureOfLoan = "Buy-to-Let" (combo "NatureOfLoan" ValidationType = "BI" OR "BR") THEN [REFER]
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyDetails.LiveWorkUnitResElement is TRUE
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("LIVEWORKUNITRESELEMENT") Is Nothing Then
            intLiveWorkUnitResElementInd = CInt(xmlNode.Attributes.getNamedItem("LIVEWORKUNITRESELEMENT").nodeValue)
        End If
        
    End If
                    
    If (intLiveWorkUnitResElementInd = 1) And ((Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BI") Is Nothing) _
        Or (Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR") Is Nothing)) Then
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
    
    Rule8065 = intResult

End Function
'========================================================================
'End of Construction Rules Code
'========================================================================


'========================================================================
'Start: Tenure rules
'========================================================================

Private Function Rule8070( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8070"
    
    'Rule 8070: Tenure = "Freehold" and property type = "Purpose Built Flat" or "Converted Flat" or "Maisonette" or "Studio Flat".
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TENURETYPE_TYPE_F") Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_C") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_PB") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_CS") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_PBS") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("DESCRIPTIONOFPROPERTY_TYPE_M") Is Nothing Then
                    blnSuccess = False
                End If
            End If
        End If
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
        'Check whether a node is returned
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TENURE_TYPE_F") Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_C") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PB") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_CS") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PBS") Is Nothing _
                Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_M") Is Nothing Then
                    blnSuccess = False
                End If
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
    
    Rule8070 = intResult

End Function


Private Function Rule8075( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
 Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intMortgageTerm As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8075"
    
    'Rule 8075: Tenure = "Leasehold" and lease term remaining at end of mortgage term < 40 years.
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TENURETYPE_TYPE_L") Is Nothing Then
                If Not xmlApplicationNode.selectSingleNode("NEWPROPERTYLEASEHOLD/UNEXPIREDTERMOFLEASEYEARS") Is Nothing Then
                    intMortgageTerm = CInt(xmlApplicationNode.selectSingleNode("NEWPROPERTYLEASEHOLD/UNEXPIREDTERMOFLEASEYEARS").Text)
                    'Mortgage term < 40 years
                    If intMortgageTerm < 40 Then
                        blnSuccess = False
                    End If
                End If
            End If
        End If
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
        'Check whether a node is returned
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TENURE_TYPE_L") Is Nothing Then
                If Not xmlNode.Attributes.getNamedItem("UNEXPIREDLEASE") Is Nothing Then
                    intMortgageTerm = CInt(xmlNode.Attributes.getNamedItem("UNEXPIREDLEASE").Text)
                    'Mortgage term < 40 years
                    If intMortgageTerm < 40 Then
                        blnSuccess = False
                    End If
                End If
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
    
    Rule8075 = intResult

End Function

'========================================================================
'End: Tenure rules
'========================================================================

'========================================================================
'Start : Legal Issues Rules
'Added By : Dhira Singh
'========================================================================

Private Function Rule8080( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intLiveWorkUnitResElementInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8080"
    
    'Rule 8080: Legal issue "Rights of Way" = "Yes".
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.RightsOfWay is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("RIGHTSOFWAY") Is Nothing Then
            intLiveWorkUnitResElementInd = CInt(xmlNode.Attributes.getNamedItem("RIGHTSOFWAY").nodeValue)
        End If
        
    End If
          
    If (intLiveWorkUnitResElementInd = 1) Then
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
    
    Rule8080 = intResult

End Function

Private Function Rule8085( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intSharedDriveOrAccessInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8085"
    
    'Rule 8085: Legal issue "Shared Drive / Access" = "Yes"..
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.SharedDriveOrAccess is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("SHAREDDRIVEORACCESS") Is Nothing Then
            intSharedDriveOrAccessInd = CInt(xmlNode.Attributes.getNamedItem("SHAREDDRIVEORACCESS").nodeValue)
        End If
        
    End If
          
    If (intSharedDriveOrAccessInd = 1) Then
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
    
    Rule8085 = intResult

End Function

Private Function Rule8090( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intGarageOnSeperateSiteInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8090"
    
    'Rule 8090: Legal issue "Garage on Separate Site" = "Yes".
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.GarageOnSeparateSite is TRUE THEN [REFER]
     Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("GARAGEONSEPARATESITE") Is Nothing Then
            intGarageOnSeperateSiteInd = CInt(xmlNode.Attributes.getNamedItem("GARAGEONSEPARATESITE").nodeValue)
        End If
        
    End If
          
    If (intGarageOnSeperateSiteInd = 1) Then
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
    
    Rule8090 = intResult
       
End Function

Private Function Rule8095( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intFlyingFreeholdInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8095"
    
    'Rule 8095: Legal issue "Flying Freehold" = "Yes".
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.FlyingFreehold is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("FLYINGFREEHOLD") Is Nothing Then
            intFlyingFreeholdInd = CInt(xmlNode.Attributes.getNamedItem("FLYINGFREEHOLD").nodeValue)
        End If
        
    End If
          
    If (intFlyingFreeholdInd = 1) Then
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
    
    Rule8095 = intResult
       
End Function

Private Function Rule8100( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intBoundaryIllDefinedInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8100"
    
    'Rule 8100: Legal issue "Ill-Defined Boundaries" = "Yes".
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.BoundaryIllDefined is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("BOUNDARYILLDEFINED") Is Nothing Then
            intBoundaryIllDefinedInd = CInt(xmlNode.Attributes.getNamedItem("BOUNDARYILLDEFINED").nodeValue)
        End If
        
    End If
          
    If (intBoundaryIllDefinedInd = 1) Then
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
    
    Rule8100 = intResult
       
End Function

Private Function Rule8105( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intSurveyorToInspectTitleInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8105"
    
    'Rule 8105: Legal issue "Surveyor Needs to Inspect Title Plans prior to Confirming Valuation" = "Yes"..
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.SurveyorToInspectTitle is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("SURVEYORTOINSPECTTITLE") Is Nothing Then
            intSurveyorToInspectTitleInd = CInt(xmlNode.Attributes.getNamedItem("SURVEYORTOINSPECTTITLE").nodeValue)
        End If
        
    End If
          
    If (intSurveyorToInspectTitleInd = 1) Then
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
    
    Rule8105 = intResult
       
End Function

Private Function Rule8110( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intOtherLegalIssuesInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8110"
    
    'Rule 8110: Legal issue "Other" = "Yes.
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.OtherLegalIssues is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("OTHERLEGALISSUES") Is Nothing Then
            intOtherLegalIssuesInd = CInt(xmlNode.Attributes.getNamedItem("OTHERLEGALISSUES").nodeValue)
        End If
    End If
          
    If (intOtherLegalIssuesInd = 1) Then
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
    
    Rule8110 = intResult
       
End Function


Private Function Rule8115( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intSharedServicesInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8115"
    
    'Rule 8115: Legal issue "Shared Services" = "Yes".
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF  ValnRepPropertyRisks.SharedServices is TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("SHAREDSERVICES") Is Nothing Then
            intSharedServicesInd = CInt(xmlNode.Attributes.getNamedItem("SHAREDSERVICES").nodeValue)
        End If
        
    End If
          
    If (intSharedServicesInd = 1) Then
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
    
    Rule8115 = intResult
       
End Function


Private Function Rule8120( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    
    Dim intLargePlotInd As Integer
    Dim intRuleValidationType As Integer
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8120"
    
    'Rule 8120: Plot > 10 acres.
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
     
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'if Plot > 10 acres then fail, otherwise pass
       
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("TOTALLANDTYPE_TYPE_GT10") Is Nothing Then
                blnSuccess = False
            End If
        End If
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
        If Not xmlNode Is Nothing Then 'Valuation report data
            If Not xmlNode.Attributes.getNamedItem("LARGEPLOT") Is Nothing Then
                intLargePlotInd = CInt(xmlNode.Attributes.getNamedItem("LARGEPLOT").nodeValue)
                If (intLargePlotInd = 1) Then
                    blnSuccess = False
                End If
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
    
    Rule8120 = intResult
       
End Function
'===============================================================================================
'End :  Legal Issues Rules
'===============================================================================================


'========================================================================
'Start: Locality rules
'========================================================================

Private Function Rule8125( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8125"
    
    'Rule 8125: Property type = ("Purpose Built Flat" or "Converted Flat" or "Maisonette" or "Studio Flat")
    'and ex-local authority = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If (Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_M") Is Nothing _
        Or Not xmlNode.Attributes.getNamedItem("TYPEOFPROPERTY_TYPE_F") Is Nothing) _
        And (Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_F") Is Nothing _
        Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PB") Is Nothing _
        Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_C") Is Nothing _
        Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_F") Is Nothing _
        Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_PBS") Is Nothing _
        Or Not xmlNode.Attributes.getNamedItem("PROPERTYDESCRIPTION_TYPE_CS") Is Nothing) Then
            blnSuccess = False
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
    
    Rule8125 = intResult

End Function


Private Function Rule8130( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
   
Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intPercentagePrivateOwner As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8130"
    
    'Rule 8130: Proportion of properties in private ownership on local authority development < 75%
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("EXLAPERCENTAGEPRIVATEOWNER") Is Nothing Then
            intPercentagePrivateOwner = CInt(xmlNode.Attributes.getNamedItem("EXLAPERCENTAGEPRIVATEOWNER").Text)
        End If
            
        If intPercentagePrivateOwner < 75 _
        And Not xmlNode.Attributes.getNamedItem("EXLOCALAUTHORITY") Is Nothing Then
            If xmlNode.Attributes.getNamedItem("EXLOCALAUTHORITY").Text = "1" Then
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
    
    Rule8130 = intResult
   
End Function


Private Function Rule8135( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8135"
    
    'Rule 8135: Boarded-up or damaged properties in the local area = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("BOARDEDUPORDAMAGEDPROPERTY") Is Nothing Then
           If xmlNode.Attributes.getNamedItem("BOARDEDUPORDAMAGEDPROPERTY").Text = "1" Then
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
    
    Rule8135 = intResult
    
End Function


Private Function Rule8140( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8140"
    
    'Rule 8140: Above commercial premises = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    'IF GetCaseStage() < GP "ValuationReportStage
    If StageLessThanValnStage(xmlApplicationNode) Then
        Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("FLATABOVECOMMERCIAL") Is Nothing Then
                If xmlNode.Attributes.getNamedItem("FLATABOVECOMMERCIAL").Text = "1" Then
                    blnSuccess = False
                End If
            End If
        End If
    Else 'rule 8230 passed
        'Call function GetCurrentCompletedValuationElement
        Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
       
        'Check whether a node is returned
        If Not xmlNode Is Nothing Then
            If Not xmlNode.Attributes.getNamedItem("BUSINESSINBLOCK") Is Nothing Then
                If xmlNode.Attributes.getNamedItem("BUSINESSINBLOCK").Text = "1" Then
                    blnSuccess = False
                End If
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
    
    Rule8140 = intResult
    
End Function


Private Function Rule8145( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8145"
    
    'Rule 8145: Property in a mining area = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("MININGAREA") Is Nothing Then
           If xmlNode.Attributes.getNamedItem("MININGAREA").Text = "1" Then
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
    
    Rule8145 = intResult

End Function


Private Function Rule8150( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8150"
    
    'Rule 8150: Property affected by radon gas = "Yes" and taken into account in valuation = "No"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("RADONGASAREA") Is Nothing _
        And Not xmlNode.Attributes.getNamedItem("RADONGASALLOWEDFOR") Is Nothing Then
           If xmlNode.Attributes.getNamedItem("RADONGASAREA").Text = "1" _
           And xmlNode.Attributes.getNamedItem("RADONGASALLOWEDFOR").Text = "0" Then
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
    
    Rule8150 = intResult
    
End Function
'========================================================================
'End: Locality rules
'========================================================================

'========================================================================
'Start: Essential Repairs Rules
'========================================================================

Private Function Rule8155( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
        
    Dim dblRetentionWorks As Double
    Dim dblValuationRetentionAmt As Double
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8155"
    
    'Get global parameter
    dblValuationRetentionAmt = GetGlobalParameter(xmlApplicationNode, "ValuationRetentionAmt", Amount)
    
    
    
    'Rule 8155: Recommended retention amount > £1000
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepValuation.RetentionWorks > New GP ValuationRetentionAmt (£1000) THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("RETENTIONWORKS") Is Nothing Then
            dblRetentionWorks = CDbl(xmlNode.Attributes.getNamedItem("RETENTIONWORKS").nodeValue)
        End If
    End If
          
    If (dblRetentionWorks > dblValuationRetentionAmt) Then
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If

    strRuleValue = "Retention Amount: " & dblRetentionWorks
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
      
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8155 = intResult
       
End Function
'========================================================================
'End : Essential Repairs Rules
'========================================================================

'========================================================================
'Start: Alterations and Extension Rules
'========================================================================
Private Function Rule8160( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String
    Dim strRuleDescription As String
       
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intExtensionsOrAlterationsInd As Integer
    
        
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8160"
    
    
    'Rule 8160: Alterations or extensions require planning permission = "Yes".
        
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
   
    'Check IF ValnRepPropertyRisks.ExtensionsOrAlterations = TRUE THEN [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    If Not xmlNode Is Nothing Then 'Valuation report data
        If Not xmlNode.Attributes.getNamedItem("EXTENSIONSORALTERATIONS") Is Nothing Then
            intExtensionsOrAlterationsInd = CDbl(xmlNode.Attributes.getNamedItem("EXTENSIONSORALTERATIONS").nodeValue)
        End If
    End If
          
    If (intExtensionsOrAlterationsInd = 1) Then
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
    
    Rule8160 = intResult
       
End Function
'========================================================================
'End : Alterations and Extension Rules
'========================================================================

'========================================================================
'Start: Occupancy rules
'========================================================================

Private Function Rule8165( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8165"
    
    'Rule 8165: Product category = "Residential" and mortgage type = "Remortgage" and property tenanted = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_RS") Is Nothing _
        And Not xmlApplicationNode.Attributes.getNamedItem("TYPEOFAPPLICATION_TYPE_R") Is Nothing _
        And Not xmlNode.Attributes.getNamedItem("TENANTEDPROPERTYIND") Is Nothing Then
            If xmlNode.Attributes.getNamedItem("TENANTEDPROPERTYIND").Text = "1" Then
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
    
    Rule8165 = intResult
    
End Function


Private Function Rule8170( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer

Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    
    Dim blnSuccess As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8170"
    
    'Rule 8170: Property subject to occupancy restrictions = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYRISKS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("OCCUPANCYRESTRICTIONS") Is Nothing Then
            If xmlNode.Attributes.getNamedItem("OCCUPANCYRESTRICTIONS").Text = "1" Then
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
    
    Rule8170 = intResult
    
End Function
'========================================================================
'End: Occupancy rules
'========================================================================


'========================================================================
'Start: New build rules
'========================================================================

Private Function Rule8175( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim strRuleValue As String, strRuleDescription As String
    
    Dim intResult As Integer
    Dim intScore As Integer
    Dim intRuleValidationType As Integer
    Dim intNewPropertyIndicator As Integer
    
    Dim blnSuccess As Boolean
    
    Dim dblYearBuilt As Double
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8175"
    intNewPropertyIndicator = 0
    dblYearBuilt = -1
    
    'Rule 8175: Property subject to occupancy restrictions = "Yes"
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    blnSuccess = True
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    
    'Check whether a node is returned
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("NEWPROPERTYINDICATOR") Is Nothing Then
            intNewPropertyIndicator = CInt(xmlNode.Attributes.getNamedItem("NEWPROPERTYINDICATOR").Text)
        End If
        If Not xmlNode.Attributes.getNamedItem("YEARBUILT") Is Nothing Then
            dblYearBuilt = CDbl(xmlNode.Attributes.getNamedItem("YEARBUILT").Text)
        End If
        
        If intNewPropertyIndicator = 1 Or (dblYearBuilt < 10 And dblYearBuilt > 0) Then
            If Not xmlNode.Attributes.getNamedItem("PPINCLSALESINCENTIVES") Is Nothing Then
                If xmlNode.Attributes.getNamedItem("PPINCLSALESINCENTIVES").nodeValue = "1" Then
                    blnSuccess = False
                End If
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
    
    Rule8175 = intResult
    
End Function

'========================================================================
'End: New build rules
'========================================================================

Private Function Rule8180(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
                
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
            
    Dim ndValNRepValuation As MSXML2.IXMLDOMNode, _
        attOverallCondition As MSXML2.IXMLDOMAttribute
        
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8180"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8180: Property doesn’t represent suitable security
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType


    Set ndValNRepValuation = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)

    If Not (ndValNRepValuation Is Nothing) Then
    
        Set attOverallCondition = ndValNRepValuation.Attributes.getNamedItem("OVERALLCONDITION_TYPE_BA")
        If Not (attOverallCondition Is Nothing) Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore)
        End If
    
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8180 = intResult
        
End Function

Private Function Rule8185(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
                
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
            
    Dim ndValNRepValuation As MSXML2.IXMLDOMNode, _
        attSaleability As MSXML2.IXMLDOMAttribute
                
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8185"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8185: Poor demand for property
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    Set ndValNRepValuation = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)

    If Not (ndValNRepValuation Is Nothing) Then
        
        Set attSaleability = ndValNRepValuation.Attributes.getNamedItem("SALEABILITY_TYPE_BA")
        If Not (attSaleability Is Nothing) Then
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore)
        End If
        
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8185 = intResult
        
End Function

Private Function Rule8190(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
                
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
            
    Dim ndValNRepPropertyServices As MSXML2.IXMLDOMNode, _
        attFutureSaleability As MSXML2.IXMLDOMAttribute
                
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8190"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8190: Factors affect future saleability
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    Set ndValNRepPropertyServices = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYSERVICES)

    If Not (ndValNRepPropertyServices Is Nothing) Then
        
        Set attFutureSaleability = ndValNRepPropertyServices.Attributes.getNamedItem("FUTURESALEABILITY")
        If Not (attFutureSaleability Is Nothing) Then
        
            If CInt(attFutureSaleability.Text) = 1 Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
            End If
        
        End If
        
    End If

    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8190 = intResult
        
End Function

Private Function Rule8195(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
                
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
            
    Dim ndValNRepPropertyDetails As MSXML2.IXMLDOMNode, _
        attNatureOfLoan_BR As MSXML2.IXMLDOMAttribute, _
        attRentalDemand_P As MSXML2.IXMLDOMAttribute
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8195"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8195: Poor rental demand for this BTL property
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    Set attNatureOfLoan_BR = xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR")
    If Not (attNatureOfLoan_BR Is Nothing) Then
        Set ndValNRepPropertyDetails = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
        
        If Not (ndValNRepPropertyDetails Is Nothing) Then
            Set attRentalDemand_P = ndValNRepPropertyDetails.Attributes.getNamedItem("RENTALDEMAND_TYPE_P")
            
            If Not (attRentalDemand_P Is Nothing) Then
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
                
                Set attRentalDemand_P = Nothing
            End If
            
            Set ndValNRepPropertyDetails = Nothing
        End If
        
        Set attNatureOfLoan_BR = Nothing
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8195 = intResult
        
End Function

Private Function Rule8200(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
    
    Dim ndValNRepPropertyDetails As MSXML2.IXMLDOMNode, _
        attNatureOfLoan_BR As MSXML2.IXMLDOMAttribute, _
        attRentableCondition As MSXML2.IXMLDOMAttribute
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8200"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8200: Property valn not same to purchase/estimated price
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    Set attNatureOfLoan_BR = xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR")
    If Not (attNatureOfLoan_BR Is Nothing) Then
    
        Set ndValNRepPropertyDetails = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPPROPERTYDETAILS)
    
        If Not (ndValNRepPropertyDetails Is Nothing) Then
        
            Set attRentableCondition = ndValNRepPropertyDetails.Attributes.getNamedItem("RENTABLECONDITION")
            
            If Not (attRentableCondition Is Nothing) Then
                
                If Not (CBool(attRentableCondition.Text)) Then
                    'RENTABLECONDITION is false - Refer
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore) 'add to appropriate referral count
                End If
                
                Set attRentableCondition = Nothing
            Else
                'RENTABLECONDITION does not exist so false - Refer
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore) 'add to appropriate referral count
            End If
        
            Set ndValNRepPropertyDetails = Nothing
        End If
        
        Set attNatureOfLoan_BR = Nothing
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8200 = intResult
        
End Function


Private Function Rule8205(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
    
    Dim attNatureOfLoan_BR As MSXML2.IXMLDOMAttribute, _
        attMonthlyRentalIncome As MSXML2.IXMLDOMAttribute, _
        dblMonthlyRentalIncome As Double, _
        attTotalNetMonthlyCost As MSXML2.IXMLDOMAttribute, _
        dblTotalNetMonthlyCost As Double, _
        attLTV As MSXML2.IXMLDOMAttribute, _
        dblLTV As Double, _
        attRentalIncome As MSXML2.IXMLDOMAttribute, _
        dblRentalIncome As Double
        
    Dim dblRentalCover As Double, _
        lstPercentages As MSXML2.IXMLDOMNodeList, _
        attPercentage As MSXML2.IXMLDOMAttribute
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8205"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8205: Property valn not same to purchase/estimated price
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType
    
    
    If StageLessThanValnStage(xmlApplicationNode) Then
        'Use application data
        Set attNatureOfLoan_BR = xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR")
        
        If Not (attNatureOfLoan_BR Is Nothing) Then
            
            If Not xmlApplicationNode.selectSingleNode("./VALNREPPROPERTYDETAILS") Is Nothing Then
                Set attMonthlyRentalIncome = xmlApplicationNode.selectSingleNode("./NEWPROPERTY").Attributes.getNamedItem("MONTHLYRENTALINCOME")
            End If
            
            If Not xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE") Is Nothing Then
                Set attTotalNetMonthlyCost = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE").Attributes.getNamedItem("TOTALNETMONTHLYCOST")
            End If
            
            If Not (attMonthlyRentalIncome Is Nothing) And _
               Not (attTotalNetMonthlyCost Is Nothing) Then
                              
                dblMonthlyRentalIncome = CDbl(attMonthlyRentalIncome.Text)
                dblTotalNetMonthlyCost = CDbl(attTotalNetMonthlyCost.Text)
                
                If dblTotalNetMonthlyCost = 0 Then
                    LogDetails 1, String(4, Chr(9)) & "TotalNetMonthlyCost = 0, return a fail"
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    Exit Function
                Else
                    dblRentalCover = Fix(dblMonthlyRentalIncome / (dblTotalNetMonthlyCost * 100))
                End If
            End If
        End If
    Else
        'Use valuation data
        Set attNatureOfLoan_BR = xmlApplicationNode.Attributes.getNamedItem("NATUREOFLOAN_TYPE_BR")
        
        If Not (attNatureOfLoan_BR Is Nothing) Then
            
            If Not xmlApplicationNode.selectSingleNode("./VALNREPPROPERTYDETAILS") Is Nothing Then
                Set attRentalIncome = xmlApplicationNode.selectSingleNode("./VALNREPPROPERTYDETAILS").Attributes.getNamedItem("RENTALINCOME")
            End If
            
            If Not xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE") Is Nothing Then
                Set attTotalNetMonthlyCost = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE").Attributes.getNamedItem("TOTALNETMONTHLYCOST")
            End If
            
            If Not (attRentalIncome Is Nothing) And _
               Not (attTotalNetMonthlyCost Is Nothing) Then
                              
                dblRentalIncome = CDbl(attRentalIncome.Text)
                dblTotalNetMonthlyCost = CDbl(attTotalNetMonthlyCost.Text)
                
                If dblTotalNetMonthlyCost = 0 Then
                    LogDetails 1, String(4, Chr(9)) & "TotalNetMonthlyCost = 0, return a fail"
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore)
                    Exit Function
                Else
                    dblRentalCover = Fix(dblRentalIncome / (dblTotalNetMonthlyCost * 100))
                End If
            End If
        End If
    End If
    
    If Not (attNatureOfLoan_BR Is Nothing) Then
        If Not xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE") Is Nothing Then
            'Get LTV and get Percentage from LPMaxLTVBTLRental
            Set attLTV = xmlApplicationNode.selectSingleNode("./QUOTATION/MORTGAGESUBQUOTE").Attributes.getNamedItem("LTV")
        End If
        
        If Not (attLTV Is Nothing) Then
            
            dblLTV = CDbl(attLTV.Text)
        
            Set lstPercentages = xmlApplicationNode.selectNodes("./GLOBALBANDEDPARAMETER[@NAME='LPMaxLTVBTLRental'][@HIGHBAND>='" & dblLTV & "']")
            
            If lstPercentages.length > 0 Then
                
                Set attPercentage = lstPercentages.Item(0).Attributes.getNamedItem("PERCENTAGE")
                If Not (attPercentage Is Nothing) Then
                     If dblRentalCover < CInt(attPercentage.Text) Then
                        intScore = intRuleValidationType
                        intResult = RiskAssessmentDecision.Referral
                        CountReferral (intScore)
                     End If
                End If
            
            End If
        
        End If
        
        AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
    End If
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8205 = intResult
        
End Function
Private Function Rule8210(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intSignatureReturned As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
        
    Dim ndValNRepValuation As MSXML2.IXMLDOMNode, _
        attPresentValuation As MSXML2.IXMLDOMAttribute, _
        attPurchasePriceOrEstValue As MSXML2.IXMLDOMAttribute
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8210"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    Set attPurchasePriceOrEstValue = xmlApplicationNode.Attributes.getNamedItem("PURCHASEPRICEORESTIMATEDVALUE")
    
    Set ndValNRepValuation = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)
    
    If Not (ndValNRepValuation Is Nothing) Then
    
        Set attPresentValuation = ndValNRepValuation.Attributes.getNamedItem("PRESENTVALUATION")
        
        If Not (attPurchasePriceOrEstValue Is Nothing) Then
        
            If Not (attPresentValuation Is Nothing) Then
                
                If attPurchasePriceOrEstValue.Text <> attPresentValuation.Text Then
                    intScore = intRuleValidationType
                    intResult = RiskAssessmentDecision.Referral
                    CountReferral (intScore) 'add to appropriate referral count
                End If
            Else
                LogDetails 1, String(4, Chr(9)) & "Mandatory attribute PURCHASEPRICEORESTIMATEDVALUE missing, return a fail"
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
            End If
        
        Else
            LogDetails 1, String(4, Chr(9)) & "Mandatory attribute PRESENTVALUATION missing, return a fail"
            intScore = intRuleValidationType
            intResult = RiskAssessmentDecision.Referral
            CountReferral (intScore)
        End If
    
    End If
    
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8210 = intResult
        
End Function

Private Function Rule8215(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
        
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intSignatureReturned As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
        
    Dim ndValuation As MSXML2.IXMLDOMNode
    Dim attGeneralRemarks As MSXML2.IXMLDOMAttribute
        
    Dim ndValNRepValuation As MSXML2.IXMLDOMNode
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8215"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    Set ndValNRepValuation = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)
    If Not ndValNRepValuation Is Nothing Then
        Set attGeneralRemarks = ndValNRepValuation.Attributes.getNamedItem("GENERALREMARKS")
    End If
    
    If Not (attGeneralRemarks Is Nothing) Then
        intScore = intRuleValidationType
        intResult = RiskAssessmentDecision.Referral
        CountReferral (intScore) 'add to appropriate referral count
    End If
        
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8215 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule8220(ByVal xmlApplicationNode As IXMLDOMNode, _
                          ByVal xmlResponseNode As IXMLDOMNode) As Integer
    
    Dim xmlNode As IXMLDOMNode
        
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intSignatureReturned As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String, _
        strScoreCardInd As String
        
    Dim ndValuerInstruction As MSXML2.IXMLDOMNode, _
        attAppointmentDate As MSXML2.IXMLDOMAttribute
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8220"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    Set ndValuerInstruction = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALUERINSTRUCTION)
    
    If Not (ndValuerInstruction Is Nothing) Then
        
        Set attAppointmentDate = ndValuerInstruction.Attributes.getNamedItem("APPOINTMENTDATE")
        
        If Not (attAppointmentDate Is Nothing) Then
        
            If TrueDateDiff("m", CDate(attAppointmentDate.Text), Now()) > 3 Then
            
                intScore = intRuleValidationType
                intResult = RiskAssessmentDecision.Referral
                CountReferral (intScore)
            
            End If
        
        End If
        
    End If
        
    AddResponse xmlResponseNode, CInt(gstrRuleNumber), strRuleDescription, strRuleValue, intScore, intResult
       
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "END: Rule" & gstrRuleNumber & "() returned (1=pass, 2=acceptcascade, 3= referral, 4=decline): " & intResult
    
    Rule8220 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule8225( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
        
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer, _
        intSignatureReturned As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String
        
    Dim blnSuccess As Boolean, _
        strScoreCardInd As String
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8225"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'RULE 8225: Valuer details not found
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    blnSuccess = True
    
    'If ValnRepValuation.SignatureReturned Is False Then [REFER]
    Set xmlNode = GetCurrentCompletedValuationElement(xmlApplicationNode, eVALNREPVALUATION)
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("SIGNATURERETURNED") Is Nothing Then
            intSignatureReturned = CInt(xmlNode.Attributes.getNamedItem("SIGNATURERETURNED").nodeTypedValue)
            If intSignatureReturned = 0 Then
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
    
    Rule8225 = intResult
    
    Set xmlNode = Nothing
        
End Function

Private Function Rule8230( _
    ByVal xmlApplicationNode As IXMLDOMNode, _
    ByVal xmlResponseNode As IXMLDOMNode) _
    As Integer
    
    Dim xmlNode As IXMLDOMNode
        
    Dim intResult As Integer, _
        intScore As Integer, _
        intRuleValidationType As Integer
        
    Dim strRuleValue As String, _
        strRuleDescription As String

        
    Dim blnSuccess As Boolean, _
        blnValnReportFound As Boolean

    Dim strScoreCardInd As String
        
    intResult = RiskAssessmentDecision.Accept
    gstrRuleNumber = "8230"
    strScoreCardInd = 0 'False
    
    strRuleDescription = GetRuleDescription(xmlApplicationNode, CInt(gstrRuleNumber))
    intRuleValidationType = GetRuleValidationType(xmlApplicationNode, CInt(gstrRuleNumber))
    
    'log details to log file
    LogDetails 1, String(3, Chr(9)) & "START: Rule" & gstrRuleNumber & "() - Description: " & strRuleDescription & ", Validation type: " & intRuleValidationType

    'RULE 8230: No valuation report; re-run Case Assessment
    
    blnSuccess = True
    blnValnReportFound = True
    
    Set xmlNode = xmlApplicationNode.selectSingleNode("NEWPROPERTY")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("VALUATIONTYPE_TYPE_NONE") Is Nothing Then
            blnValnReportFound = False
        End If
    End If
    
    If blnValnReportFound Then
        'IF the Stage if >= GP ValuationReportStage and the valuation report doesn't exist THEN [FAIL]
        If GetCurrentCompletedValuationElement(xmlApplicationNode, eVALUERINSTRUCTION) Is Nothing _
        And (StageLessThanValnStage(xmlApplicationNode) = False) Then
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
    End If
        
    Rule8230 = intResult
    
    Set xmlNode = Nothing
        
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
    
    iNumOfDays = DateDiff("d", dtmStart, dtmEnd)
    iNumOfDays = IIf(iNumOfDays < 0, iNumOfDays + -1, iNumOfDays + 1)
    
    Select Case LCase(sInterval)
        Case "d"
            
            iResult = iNumOfDays
            
        Case "w"

            iResult = Fix(iNumOfDays / 7)
            
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

    TrueDateDiff = iResult

End Function




