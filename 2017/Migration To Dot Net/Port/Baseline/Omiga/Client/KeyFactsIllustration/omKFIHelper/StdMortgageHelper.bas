Attribute VB_Name = "StdMortgageHelper"
'********************************************************************************
'** Module:         StdMortgageHelper
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    This module contains functions/procedures that are common to
'**                 standard mortgage KFI documents (KFI and Offer).
'********************************************************************************
'Prog   Date        Description
'BC     05/01/2005  MAR88 - Project MARS
'TW     09/11/2005  MAR442  - A number of changes to allow omCK.CreateKFI to work
'                             Note that these changes prevent errors where certain data is missing
'                             See the specific methods/functions for the changes
'BC     16/01/2006  MAR1058  - Additional text in Section 11 from Fixed Rate products
'HMA    02/02/2006  MAR1071  - Use global parameter for Maximum LTV percentage.
'BC     01/05/2006  MAR1685    Inhibit generation of TERMINMONTHS if months = zero
'BC     16/05/2006  MAR1736    Show LC Seq Number correctly in AddMortgageComponentsDetails
'PM     19/05/2006  EP584      Amended AddServicesProvided
'PB     05/06/2006  EP651      MAR1590 modified AddMortgageComponentsDetails - show two rows if a loan is Part & Part
'PB     12/06/2006  EP730      MAR1831 Change in description of 'Repayment' payment method
'PM     29/06/2006  EP917      Updated BuildSection7Common to add REGULATEDMORTGAGE element if appropriate
'DRC    12/07/06    EP983      Reverse MAR1071 - so MaxLTV is product based, rather than from a globalparameter
'INR    29/11/2006  EP2_139    Offer and KFI document changes.
'PB     30/11/2006  EP2_139    Added INTERESTONLY attribute to section 7a.
'PB     12/12/2006  EP2_422    Changes or KFI
'INR    18/12/2006  EP2_545    Fixed buildsection7common typo
'INR    07/02/2007  EP2_583    KFI changes
'INR    22/02/2007  EP2_1450   Rework BuildSection7ACommon
'INR    28/02/2007  EP2_1449   Section4 "Prodcut may not be available" only if not additional borrowing
'PB     05/03/2007  EP2_1629   Make sure section number references in text match up correctly.
'PB     09/03/2007  EP2_1627   Utilised gblnSection2 flag
'INR    20/03/2007  EP2_1977  Deal with REGULATIONINDICATOR_VALIDID
'GHun   15/04/2007  EP2_2395   Changed BuildSection7Commmon
'INR    17/04/2007  EP2_2429    Always have section2 now
'INR    20/04/2007  EP2_2478    fixed rate period in section7 instead of end date
'********************************************************************************
Option Explicit

Private Const mcstrModuleName As String = "stdMortgageHelper"

Private Type PrefRate
    Part As Integer
    Increase As Double
    Period As Integer
    PrefRate As String
    Rate As String
End Type

'********************************************************************************
'** Function:       BuildSection7Common
'** Created by:     Andy Maggs
'** Date:           12/07/2004
'** Description:    Sets the elements and attributes for the compulsory section7
'**                 (Are you comfortable with the risks?)
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section7 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildSection7Common(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7Common"
    Dim xmlLoanComps As IXMLDOMNodeList
    Dim xmlLoanComp As IXMLDOMNode
    Dim intNoOfLCs As Integer
    Dim intCount As Integer
    Dim xmlComponent As IXMLDOMNode
    Dim xmlDetail As IXMLDOMNode
    Dim xmlPart As IXMLDOMNode
    Dim xmlPartNum As IXMLDOMNode
    Dim xmlInterestRates As IXMLDOMNodeList
    Dim xmlInterestRate As IXMLDOMNode
    Dim objHelper As IntRateTypesHelper
    Dim eType As MortgageInterestRateType
    Dim intIndex As Integer
    Dim blnAddMaxMinRate As Boolean
    'INR CORE82
    Dim xmlGen1 As IXMLDOMNode
    'PM 29/06/2006 EP917 - Start
    Dim xmlAFF As IXMLDOMNode
    Dim strRegInd As String
    'PM 29/06/2006 EP917 - End
    
    'PB 04/12/2006 EP2_139
    Dim xmlThisPart As IXMLDOMNode
    Dim xmlPaymentSchedules As IXMLDOMNodeList
    Dim blnIsVariable As Boolean
    Dim blnIsFixed As Boolean
    Dim xmlDiscount As IXMLDOMNode
    Dim xmlFixed As IXMLDOMNode
    Dim xmlVariable As IXMLDOMNode
    Dim strVariableParts As String
    Dim strFixedParts As String
    Dim intCommaPos As Integer
    Dim blnFixed As Boolean
    Dim a_PrefRates() As PrefRate
    Dim intPrefCount As Integer
    Dim strRate As String
    Dim strPrefRate As String
    Dim xmlAfterPrefrate As IXMLDOMNode
    Dim xmlSchedules As IXMLDOMNodeList
     'EP2_2395 GHun
    Dim strFixedInterestRate As String
    Dim strVariableInterestRate As String
     'EP2_2395 End
    Dim dblFixedIncrease As Double
    Dim dblVariableIncrease As Double
    Dim xmlInterestRateType As IXMLDOMNode
    Dim strPreviousStepStartDate As String
    Dim strThisStepStartDate
    Dim intStepLoop As Integer
    Dim intPrefPeriod As Integer
    Dim strPrefPeriodEndDate As String
    Dim dblIncrease As Double
    Dim intRatePeriondInMonths As Integer, strPrefPeriod As String
    
    On Error GoTo ErrHandler
   
    '*-add GEN1 and GEN3 element
    Call Section7_AddGEN1andGEN3(vobjCommon, vxmlNode)
          
    'PM 29/06/2006 EP917 - Start
    '*-If this a regulated mortgage create the REGULATEDMORTGAGE element
    Set xmlAFF = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlAFF Is Nothing Then
        'EP2_1977 REGULATIONINDICATOR now holds ValidationId
        strRegInd = xmlGetAttributeText(xmlAFF, "REGULATIONINDICATOR")
        If CheckForValidationType(strRegInd, "R") Then
            Call vobjCommon.CreateNewElement("REGULATEDMORTGAGE", vxmlNode)
        End If
    End If
    'PM 29/06/2006 EP917 - End
    
    Set xmlLoanComps = vobjCommon.LoanComponents
    intNoOfLCs = xmlLoanComps.length
    
        '*-create the COMPONENT element
    Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", vxmlNode)
    '*-create the mandatory COMPONENTDETAIL element
    Set xmlDetail = vobjCommon.CreateNewElement("COMPONENTDETAIL", xmlComponent)
    
    If intNoOfLCs > 1 Then
        '*-create the MULTIPART element
        Set xmlPart = vobjCommon.CreateNewElement("MULTIPART", xmlDetail)
        '
        ' Description
        ' --------------
        '
        ' We need to work out which components we can group together by variable/fixed.
        ' This is because the text in part 7 needs to appear as follows:
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        ' (example 1)
        '
        ' What if interest rates go up?
        ' The monthly payments shown in this illustration could be considerably different if interest
        ' rates change.
        '
        ' Parts 1 and 2
        '         For example, for one percentage point increase in the standard variable rate, your
        '         monthly payment will increase by around £90.84
        '
        ' Part 3
        '         The interest rate is fixed until 31/3/2008. Your monthly payment on this list will not
        '         change with interest rates until after this date.
        '
        ' After the end of the fixed rate period which applies to part 3, then for one percentage point
        ' increase in the standard variable rate, your monthly payment will increase by around £115.62.
        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        ' (example 2)
        '
        ' What if interest rates go up?
        ' The monthly payments shown in this illustration could be considerably different if interest
        ' rates change.
        '
        ' Part 1
        '         For example, for one percentage point increase in the standard variable rate, your
        '         monthly payment will increase by around £90.84
        '
        ' Parts 2 and 3
        '         The interest rate is fixed until 31/3/2008. Your monthly payment on this list will not
        '         change with interest rates until after this date.
        '
        ' After the end of the fixed rate period which applies to parts 2 and 3, then for one percentage
        ' point increase in the standard variable rate, your monthly payment will increase by around £115.62.
        '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        '*-loop thru components
        For intCount = 0 To intNoOfLCs - 1
            '
            Set xmlLoanComp = xmlLoanComps.Item(intCount)
            '
            '*-create the PART element
            'Set xmlPartNum = vobjCommon.CreateNewElement("PART", xmlPart)
            '*-create the PARTNUM attribute
            'Call xmlSetAttributeValue(xmlPartNum, "PARTNUM", CStr(intCount + 1))
            '
            '*-Add either DISCOUNT or FIXED element accordingly...
            '
            'If the first interest rate has ratetype F then this is fixed.
            '
            blnFixed = (xmlGetAttributeText(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]"), "RATETYPE") = "F")
            
            If blnFixed Then
                strFixedInterestRate = xmlGetAttributeText(xmlLoanComp.selectSingleNode(".//BASERATE"), "RATEDESCRIPTION")   'EP2_2395 GHun
                'Fixed
                If strFixedParts = "" Then
                    strFixedParts = " " & CStr(intCount + 1)
                Else
                    If Left(strFixedParts, 1) = "s" Then
                        strFixedParts = strFixedParts & ", " & CStr(intCount + 1)
                    Else
                        strFixedParts = "s " & strFixedParts & ", " & CStr(intCount + 1)
                    End If
                End If
                'SR EP2_2159
                intRatePeriondInMonths = xmlGetAttributeAsInteger( _
                        xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]"), _
                        "INTERESTRATEPERIOD")
                'SR EP2_2159 - End
                ' Add increase value 'SR EP2_2159 - Get INCREASEDMONTHLYCOSTDIFFERENCE from LOANCOMPONENTPAYMENTSCHEDULE
                dblFixedIncrease = dblFixedIncrease + xmlGetAttributeAsDouble( _
                    xmlLoanComp.selectSingleNode("LOANCOMPONENTPAYMENTSCHEDULE"), "INCREASEDMONTHLYCOSTDIFFERENCE")
            Else
                strVariableInterestRate = xmlGetAttributeText(xmlLoanComp.selectSingleNode(".//BASERATE"), "RATEDESCRIPTION")    'EP2_2395 GHun
                'Variable
                If strVariableParts = "" Then
                    strVariableParts = " " & CStr(intCount + 1)
                Else
                    If Left(strVariableParts, 1) = "s" Then
                        strVariableParts = strVariableParts & ", " & CStr(intCount + 1)
                    Else
                        strVariableParts = "s" & strVariableParts & ", " & CStr(intCount + 1)
                    End If
                End If
                ' Add increase value 'SR EP2_2159 - Get INCREASEDMONTHLYCOSTDIFFERENCE from LOANCOMPONENTPAYMENTSCHEDULE
                dblVariableIncrease = dblVariableIncrease + xmlGetAttributeAsDouble( _
                        xmlLoanComp.selectSingleNode("LOANCOMPONENTPAYMENTSCHEDULE"), "INCREASEDMONTHLYCOSTDIFFERENCE")
            End If
            '
        Next intCount
        '
        'Replace final comma with 'and'
        '
        intCommaPos = InStrRev(strFixedParts, ",")
        If intCommaPos > 0 Then
            strFixedParts = Left(strFixedParts, intCommaPos - 1) & " and" & Mid(strFixedParts, intCommaPos + 1)
        End If
        intCommaPos = InStrRev(strVariableParts, ",")
        If intCommaPos > 0 Then
            strVariableParts = Left(strVariableParts, intCommaPos - 1) & " and" & Mid(strVariableParts, intCommaPos + 1)
        End If
        '
        'Now add VARIABLE and FIXED elements if required
        'SR EP2_2159
        If (intRatePeriondInMonths >= 12) Then
            strPrefPeriod = CStr(Int(intRatePeriondInMonths / 12)) & IIf(intRatePeriondInMonths > 12, " Years", " Year")
        End If
        If (intRatePeriondInMonths Mod 12) <> 0 Then
            intRatePeriondInMonths = intRatePeriondInMonths Mod 12
            strPrefPeriod = strPrefPeriod & " and " & CStr(intRatePeriondInMonths) & IIf(intRatePeriondInMonths > 1, " Months", " Month")
        End If
        'SR EP2_2159 - End

        If strFixedParts <> "" Then
            Set xmlFixed = vobjCommon.CreateNewElement("FIXED", xmlPart)
            xmlSetAttributeValue xmlFixed, "PARTS", strFixedParts
            xmlSetAttributeValue xmlFixed, "RATE", strFixedInterestRate 'EP2_2395 GHun
            xmlSetAttributeValue xmlFixed, "INCREASE", dblFixedIncrease
            xmlSetAttributeValue xmlFixed, "PREFPERIOD", strPrefPeriod
        End If
        If strVariableParts <> "" Then
            Set xmlVariable = vobjCommon.CreateNewElement("VARIABLE", xmlPart)
            xmlSetAttributeValue xmlVariable, "PARTS", strVariableParts
            xmlSetAttributeValue xmlVariable, "RATE", strVariableInterestRate  'EP2_2395 GHun
            xmlSetAttributeValue xmlVariable, "INCREASE", dblVariableIncrease
        End If
        '
        'Now add the bit which describes what happens when the preferential rate period ends
        '
        For intCount = 0 To intNoOfLCs - 1
            '
            Set xmlLoanComp = xmlLoanComps.Item(intCount)
            'Count the payment schedules
            Set xmlSchedules = xmlLoanComp.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
            If xmlSchedules.length > 1 Then
                'More than 1 schedule, so first one must be a preferential period
                intPrefCount = intPrefCount + 1
                'Get the rate name
                strPrefRate = GetInterestRateTypeText(GetInterestRateType(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]")))
                strRate = GetInterestRateTypeText(GetInterestRateType(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=2]")))
                strRate = xmlGetAttributeText(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]/BASERATESETDATA/BASERATE"), "RATEDESCRIPTION")
                '
                ReDim Preserve a_PrefRates(intPrefCount)
                '
                a_PrefRates(intPrefCount).Part = intCount + 1
                '
                'Try to get the monthly increase step
                '
                intPrefPeriod = 0 ' reset from previous
                For intStepLoop = 1 To xmlSchedules.length - 1
                    '
                    If xmlGetAttributeAsDouble(xmlSchedules(intStepLoop), "INCREASEDMONTHLYCOSTDIFFERENCE") Then
                        '
                        a_PrefRates(intPrefCount).Increase = xmlGetAttributeAsDouble(xmlSchedules(intStepLoop), "INCREASEDMONTHLYCOSTDIFFERENCE")
                        a_PrefRates(intPrefCount).Period = intPrefPeriod
                        Exit For
                        '
                    Else
                        ' get the start date of the previous payment period
                        strPreviousStepStartDate = xmlGetAttributeText(xmlSchedules(intStepLoop - 1), "STARTDATE")
                        ' get the start date of the current step period
                        strThisStepStartDate = xmlGetAttributeText(xmlSchedules(intStepLoop), "STARTDATE")
                        ' get the difference of the two dates and add to the total
                        intPrefPeriod = intPrefPeriod + DateDiff("m", strPreviousStepStartDate, strThisStepStartDate)
                        '
                    End If
                    '
                Next intStepLoop
                a_PrefRates(intPrefCount).Rate = strRate
                a_PrefRates(intPrefCount).PrefRate = strPrefRate
                '
            End If
            '
        Next intCount
        '
        'Build AFTERPREFRATE XML
        '
        For intCount = 1 To intPrefCount
            ' Check we have the 'end-of-preferential-period-one-percent' increase value,
            ' otherwise display nothing.
            ' If this value is missing then this is not fixed/capped/collared so this text is not required.
            If a_PrefRates(intCount).Increase Then
                Set xmlAfterPrefrate = vobjCommon.CreateNewElement("AFTERPREFRATE", xmlPart)
                xmlSetAttributeValue xmlAfterPrefrate, "PART", a_PrefRates(intCount).Part
                xmlSetAttributeValue xmlAfterPrefrate, "RATE", a_PrefRates(intCount).Rate
                xmlSetAttributeValue xmlAfterPrefrate, "PREFRATE", a_PrefRates(intCount).PrefRate
                xmlSetAttributeValue xmlAfterPrefrate, "INCREASE", a_PrefRates(intCount).Increase
            End If
        Next intCount
        '
    Else
        '
        '<COMPONENTDETAIL>
        '   <SINGLEPART>
        '       <VARIABLE INCREASE=''/>
        '   ...or...
        '       <FIXED PREFPERIOD='' RATE='' INCREASE=''/>
        '   </SINGLEPART>
        '</COMPONENTDETAIL>
        '
        '*-create the SINGLEPART element
        Set xmlPart = vobjCommon.CreateNewElement("SINGLEPART", xmlDetail)
        '
        Set xmlLoanComp = xmlLoanComps.Item(0)
        '
        'EP2_583 If we don't have xmlLoanComp, no point trying anything below here
        If Not xmlLoanComp Is Nothing Then
            Set xmlSchedules = xmlLoanComp.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
            'EP2_2478
            intRatePeriondInMonths = xmlGetAttributeAsInteger( _
                    xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE"), "INTERESTRATEPERIOD")
    
            If (intRatePeriondInMonths >= 12) Then
                strPrefPeriod = CStr(Int(intRatePeriondInMonths / 12)) & IIf(intRatePeriondInMonths > 12, " Years", " Year")
            End If
            If (intRatePeriondInMonths Mod 12) <> 0 Then
                intRatePeriondInMonths = intRatePeriondInMonths Mod 12
                strPrefPeriod = strPrefPeriod & " and " & CStr(intRatePeriondInMonths) & IIf(intRatePeriondInMonths > 1, " Months", " Month")
            End If
            If xmlSchedules.length > 1 Then
                'More than 1 schedule, so first one must be a preferential period
                intPrefCount = intPrefCount + 1
                'Get the rate name
                strPrefRate = GetInterestRateTypeText(GetInterestRateType(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]")))
                'strRate = GetInterestRateTypeText(GetInterestRateType(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=2]")))
                '
                ReDim Preserve a_PrefRates(intPrefCount)
                '
                a_PrefRates(intPrefCount).Part = intCount + 1
                '
                'Try to get the monthly increase step
                '
                intPrefPeriod = 0 ' reset from previous
                For intStepLoop = 0 To xmlSchedules.length - 1
                    '
                    If xmlGetAttributeAsDouble(xmlSchedules(intStepLoop), "INCREASEDMONTHLYCOSTDIFFERENCE") Then
                        '
                        dblIncrease = xmlGetAttributeAsDouble(xmlSchedules(intStepLoop), "INCREASEDMONTHLYCOSTDIFFERENCE")
                        strPrefPeriodEndDate = xmlGetAttributeText(xmlSchedules(intStepLoop), "STARTDATE")
                        strRate = xmlGetAttributeText(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=" & intStepLoop + 1 & "]/BASERATESETDATA/BASERATE"), "RATEDESCRIPTION")
                        Exit For
                        '
                    End If
                    '
                Next intStepLoop
                '
            End If
        '
            blnFixed = (xmlGetAttributeText(xmlLoanComp.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]"), "RATETYPE") = "F")
        
            If blnFixed Then
                'Fixed
                Set xmlFixed = vobjCommon.CreateNewElement("FIXED", xmlPart)
'                xmlSetAttributeValue xmlFixed, "PREFPERIODEND", strPrefPeriodEndDate
                xmlSetAttributeValue xmlFixed, "PREFPERIOD", strPrefPeriod
                xmlSetAttributeValue xmlFixed, "RATE", strRate
                xmlSetAttributeValue xmlFixed, "INCREASE", dblIncrease
            Else
                'Variable
                Set xmlVariable = vobjCommon.CreateNewElement("VARIABLE", xmlPart)
                'EP2_545
                xmlSetAttributeValue xmlVariable, "INCREASE", dblIncrease
                'EP2_583
                xmlSetAttributeValue xmlVariable, "RATE", strRate
            End If
            '
            '*-create a COMPONENT element
            '
            Set xmlLoanComp = xmlLoanComps.Item(0)
            '*-get the interest rate information for this loan component
            Set xmlInterestRates = vobjCommon.GetLoanComponentInterestRates(xmlLoanComp)
            Set objHelper = New IntRateTypesHelper
            Call objHelper.Initialise(xmlLoanComp, xmlInterestRates)
    
            Section7_AddLCBasedIncreasePercentageElement vobjCommon, objHelper, xmlLoanComp, xmlDetail
            'SR 21/09/2004 : CORE82
            'MAR88 BC - reinstigate this loop (previously commented out)
            intIndex = 0
            For Each xmlInterestRate In xmlInterestRates
                intIndex = intIndex + 1
                            'intIndex = 1
                'SR 21/09/2004 : CORE82 - End
                Set xmlInterestRate = xmlInterestRates.Item(intIndex - 1)   'SR 21/09/2004 : CORE82
                '*-get the type of this interest rate
                eType = GetInterestRateType(xmlInterestRate)
    'SR 21/09/2004 : CORE82
    'MAR88 - comment out this call, processing has just been done in Section7_AddLCBasedIncreasePercentageElement
    '            Call Section7_AddIncreasePercentageElement(vobjCommon, objHelper, _
                        xmlInterestRate, eType, xmlDetail)
    'SR 21/09/2004 : CORE82 - End
                blnAddMaxMinRate = False
                If eType = mrtFixedRate Then
                    blnAddMaxMinRate = True
                        
                    'INR CORE82 Need to check that Gen1 exists
                    Set xmlGen1 = vobjCommon.Document.selectSingleNode("//GEN1")
                    If Not xmlGen1 Is Nothing Then
                        Call Section7_AddFixedRateElement(vobjCommon, objHelper, _
                                xmlInterestRate, xmlDetail)
                    End If
                ElseIf eType = mrtCappedAndCollaredRate Or eType = mrtCappedRate _
                        Or eType = mrtCollaredRate Then
                    blnAddMaxMinRate = True
                    Call Section7_AddCappedCollaredElement(vobjCommon, objHelper, _
                            xmlDetail)
                End If
                         
                If blnAddMaxMinRate Then
                    '*-create the MAXMINRATE element
                    Call Section7_CreateMaxMinRate(vobjCommon, objHelper, xmlInterestRate, _
                            eType, intIndex, xmlDetail)
                End If
            
            Next xmlInterestRate  'SR 21/09/2004 : CORE82
        End If
        '
    End If

    Set xmlLoanComps = Nothing
    Set xmlLoanComp = Nothing
    Set xmlComponent = Nothing
    Set xmlDetail = Nothing
    Set xmlPart = Nothing
    Set xmlPartNum = Nothing
    Set xmlInterestRates = Nothing
    Set xmlInterestRate = Nothing
    Set xmlGen1 = Nothing
    'PM 29/06/2006 EP917 - Start
    Set xmlAFF = Nothing
    'PM 29/06/2006 EP917 - End
Exit Sub
ErrHandler:
    Set xmlLoanComps = Nothing
    Set xmlLoanComp = Nothing
    Set xmlComponent = Nothing
    Set xmlDetail = Nothing
    Set xmlPart = Nothing
    Set xmlPartNum = Nothing
    Set xmlInterestRates = Nothing
    Set xmlInterestRate = Nothing
    Set xmlGen1 = Nothing
    'PM 29/06/2006 EP917 - Start
    Set xmlAFF = Nothing
    'PM 29/06/2006 EP917 - End
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       Section7_AddGEN1andGEN3
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Procedure to add GEN1 and GEN3 elements to Section7
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - Node to attach the elements to
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_AddGEN1andGEN3(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "Section7_AddGEN1andGEN3"
    
    Dim xmlItems As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlInterestRate As IXMLDOMNode
    Dim eRateType As MortgageInterestRateType
    Dim blnGen1 As Boolean
    Dim blnGen3 As Boolean
    Dim intPeriod As Integer
    Dim lngCeilingRate As Long
    Dim dblActualBaseRate As Double
    
    On Error GoTo ErrHandler
    
    blnGen1 = False
    blnGen3 = False
    
    Set xmlItems = vobjCommon.LoanComponents
    For Each xmlItem In xmlItems
        '*-get the first interest rate type record for this loan component
        Set xmlInterestRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlItem)
        
'TW 09/11/2005 MAR442
        If Not xmlInterestRate Is Nothing Then
'TW 09/11/2005 MAR442 End
            eRateType = GetInterestRateType(xmlInterestRate)
'TW 09/11/2005 MAR442
        End If
'TW 09/11/2005 MAR442 End
        If eRateType = mrtFixedRate Then
            intPeriod = xmlGetAttributeAsInteger(xmlInterestRate, "INTERESTRATEPERIOD")
            If intPeriod <> -1 Then
                '*-interest rate is not fixed for the whole term
                blnGen1 = True
                blnGen3 = True
                '*-we only have to find 1 so stop looking
                Exit For
            End If
        
        ElseIf eRateType = mrtCappedRate Or eRateType = mrtCappedAndCollaredRate Then
            lngCeilingRate = xmlGetAttributeAsLong(xmlInterestRate, "CEILINGRATE")
            dblActualBaseRate = GetBaseInterestRate(vobjCommon, xmlInterestRate, "")
            
            If lngCeilingRate > dblActualBaseRate + 1 Then
                '*-the interest rate can go up by more than 1%
                blnGen1 = True
                blnGen3 = True
                '*-we only have to find 1 so stop looking
                Exit For
            End If
        
        Else
            '*-this is a variable or collared rate and can therefore
            '*-potentially increase by more than 1%
            blnGen1 = True
            blnGen3 = True
            '*-we only have to find 1 so stop looking
            Exit For
        End If
    Next xmlItem

    If blnGen1 Then
        Call vobjCommon.CreateNewElement("GEN1", vxmlNode)
    End If
    
    If blnGen3 Then
        Call vobjCommon.CreateNewElement("GEN3", vxmlNode)
    End If
    
    Set xmlItems = Nothing
    Set xmlItem = Nothing
    Set xmlInterestRate = Nothing
Exit Sub
ErrHandler:
    Set xmlItems = Nothing
    Set xmlItem = Nothing
    Set xmlInterestRate = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       Section7_AddFixedRateElement
'** Created by:     Andy Maggs
'** Date:           15/07/2004
'** Description:    Procedure to add the elements specific to the component with
'**                 Interest Rate type "Fixed"
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlLoanComp - the loan component.
'**                 vxmlNode - the node to add the fixed rate element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_AddFixedRateElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vobjHelper As IntRateTypesHelper, ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "Section7_AddFixedRateElement"
    
    Dim xmlFixed As IXMLDOMNode
    Dim blnHasEndDate As Boolean
    Dim dtEndDate As Date
    Dim intPeriod As Integer
    Dim intSeqNum As Integer
    
    On Error GoTo ErrHandler
    
    '*-create the FIXEDPART element
    Set xmlFixed = vobjCommon.CreateNewElement("FIXEDPART", vxmlNode)
    
    '*-then get the information about the term of this interest rate
    Call GetBasicInterestRateTermData(vobjCommon, vobjHelper.InterestRates, _
            vxmlInterestRate, blnHasEndDate, dtEndDate, intPeriod)
    
    '*-add the UNTILDATE or FORPERIOD elements using this term data
    Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
            intPeriod, xmlFixed)
    
    If vobjCommon.LoanComponents.length > 1 Then
        Call vobjCommon.CreateNewElement("MULTIPART", xmlFixed)
    End If
    
    intSeqNum = xmlGetAttributeAsInteger(vxmlInterestRate, "INTERESTRATETYPESEQUENCENUMBER")
    If (intSeqNum = 1 And vobjHelper.InterestRates.length > 1) _
            Or (intSeqNum > 1 And intPeriod <> -1) Then
        '*-add the UNTILAFTERTHISDATE element
        Call vobjCommon.CreateNewElement("UNTILAFTERTHISDATE", xmlFixed)
    End If
        
    If vobjHelper.FixedRates = vobjHelper.RatesCount Then
        '*-add the CHANGEWITHSTEPS element if all the rate types are fixed
        Call vobjCommon.CreateNewElement("CHANGEWITHSTEPS", xmlFixed)
    End If

    Set xmlFixed = Nothing
Exit Sub
ErrHandler:
    Set xmlFixed = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       Section7_AddCappedCollaredElement
'** Created by:     Andy Maggs
'** Date:           13/07/2004
'** Description:    Procedure to add the elements specific to the component with
'**                 Interest Rate type "Capped" or "Collared" or "Capped&Collared"
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlLoanComp - element containing the loan component details.
'**                 vxmlComponent - the node to add the elements and attributes to
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_AddCappedCollaredElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vobjHelper As IntRateTypesHelper, ByVal vxmlComponent As IXMLDOMNode)
    Const cstrFunctionName As String = "Section7_AddCappedCollaredElement"
    
    Dim xmlCapped As IXMLDOMNode
    Dim xmlInterestRates As IXMLDOMNodeList
    Dim xmlInterestRate As IXMLDOMNode
    Dim dblFloor As Long
    Dim dblCeiling As String
    Dim blnHasEndDate As Boolean
    Dim dtEndDate As Date
    Dim intPeriod As Integer
    Dim eType As MortgageInterestRateType
    Dim dblMinCost As Double
    Dim dblMaxCost As Double
    Dim xmlTemp As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Set xmlCapped = vobjCommon.CreateNewElement("CAPPEDCOLLAREDPART", vxmlComponent)
    If vobjHelper.HasMixedRates Then
        Call vobjCommon.CreateNewElement("CAPPEDPARTOFTERM", xmlCapped)
    Else
        Call vobjCommon.CreateNewElement("CAPPEDTHROUGHOUT", xmlCapped)
    End If
    
    
    '*-get the minimum floor and maximum cap values
    Set xmlInterestRates = vobjHelper.InterestRates
    dblFloor = xmlListHelper.GetMaxDblAttribValue(xmlInterestRates, "FLOOREDRATE")
    dblCeiling = xmlListHelper.GetMaxDblAttribValue(xmlInterestRates, "CEILINGRATE")
    
    'MAR54 add attribute MAXCAPPEDFLOOREDRATE as 2DP
    Call xmlSetAttributeValue(xmlCapped, "MAXCAPPEDFLOOREDRATE", set2DP(CStr(dblCeiling)))
    
    If dblFloor > 0 Then
        Call vobjCommon.CreateNewElement("MINIMUMFLOOR", xmlCapped)
    End If
    
    If dblCeiling > 0 Then
        Call vobjCommon.CreateNewElement("MAXIMUMCAP", xmlCapped)
    End If
    
    '*-get the first interest rate
    Set xmlInterestRate = vobjCommon.GetLoanComponentFirstInterestRate(vobjHelper.LoanComponent)
    eType = GetInterestRateType(xmlInterestRate)
    
    '*-then get the information about the term of this interest rate
    Call GetBasicInterestRateTermData(vobjCommon, xmlInterestRates, xmlInterestRate, _
            blnHasEndDate, dtEndDate, intPeriod)
    
    '*-add the UNTILDATE or FORPERIOD elements using this term data
    Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
            intPeriod, xmlCapped)
    
    If vobjCommon.LoanComponents.length > 1 Then
        Call vobjCommon.CreateNewElement("MULTIPART", xmlCapped)
    End If
    
    Call GetRatePeriodIncreaseCost(vobjHelper.LoanComponent, 1, dblMinCost, dblMaxCost)
    
    'MAR54 - Add check for Capped and Collared to test
    If eType = mrtCappedRate Or eType = mrtCappedAndCollaredRate Then
        Set xmlTemp = vobjCommon.CreateNewElement("NOTEXCEED", xmlCapped)
        '*-add the CALC attribute
        Call xmlSetAttributeValue(xmlTemp, "CALC", set2DP(CStr(dblMaxCost)))
    ElseIf eType = mrtCollaredRate Or eType = mrtCappedAndCollaredRate Then
        Set xmlTemp = vobjCommon.CreateNewElement("NOTLESSTHAN", xmlCapped)
        '*-add the CALC attribute
        Call xmlSetAttributeValue(xmlTemp, "CALC", set2DP(CStr(dblMinCost)))
    End If
    
    If xmlInterestRates.length > 1 Then
        '*-add the UNTILAFTERTHISDATE element
        Call vobjCommon.CreateNewElement("UNTILAFTERTHISDATE", xmlCapped)
    End If
    Set xmlCapped = Nothing
    Set xmlInterestRates = Nothing
    Set xmlInterestRate = Nothing
    Set xmlTemp = Nothing
Exit Sub
ErrHandler:
    Set xmlCapped = Nothing
    Set xmlInterestRates = Nothing
    Set xmlInterestRate = Nothing
    Set xmlTemp = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       Section7_AddRevertsToSVRElement
'** Created by:     Andy Maggs
'** Date:           19/07/2004
'** Description:    Adds the REVERTSTOSVR element to the specified node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vobjHelper - the interest rate types helper object.
'**                 vintLoanItem - the index of the loan component.
'**                 vintNumLoanComponents - the total number of loan components.
'**                 vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_AddRevertsToSVRElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vobjHelper As IntRateTypesHelper, _
        ByVal vintLoanItem As Integer, ByVal vintNumLoanComponents As Integer, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "Section7_AddRevertsToSVRElement"
    
    Dim xmlInterestRate As IXMLDOMNode
    Dim eType As MortgageInterestRateType
    Dim xmlPrevInterestRate As IXMLDOMNode
    Dim ePrevType As MortgageInterestRateType
    Dim xmlReverts As IXMLDOMNode
    Dim intSeqNum As Integer
    Dim xmlTemp As IXMLDOMNode
    Dim strAmount As String
    Dim xmlRateDescription As IXMLDOMNode
    
    On Error GoTo ErrHandler

    '*-get the last interest rate type record
    Set xmlInterestRate = vobjHelper.InterestRates.Item(vobjHelper.RatesCount - 1)
    eType = GetInterestRateType(xmlInterestRate)
    intSeqNum = xmlGetAttributeAsInteger(xmlInterestRate, "INTERESTRATETYPESEQUENCENUMBER")
    '*-get the last but one interest rate type record
    Set xmlPrevInterestRate = vobjHelper.InterestRates.Item(vobjHelper.RatesCount - 2)
    ePrevType = GetInterestRateType(xmlPrevInterestRate)
    
    If eType <> mrtFixedRate Then
        '*-create the REVERTSTOSVR element
        Set xmlReverts = vobjCommon.CreateNewElement("REVERTSTOSVR", vxmlNode)
        '*-add the RATETYPE attribute
        Set xmlRateDescription = xmlInterestRate.selectSingleNode("//BASERATE")
        Call xmlSetAttributeValue(xmlReverts, "RATETYPE", _
                xmlGetAttributeText(xmlRateDescription, "RATEDESCRIPTION"))
'        Call xmlSetAttributeValue(xmlReverts, "RATETYPE", _
'                vobjHelper.InterestRateTypeAsText(eType))
        '*-add the CARDINALNUMBER attribute
        'INR CORE82 Need to convert number to Cardinal before setting attribute
        Call xmlSetAttributeValue(xmlReverts, "CARDINALNUMBER", Section7_CreateCardinal(vintLoanItem))
                    
        Select Case ePrevType
            Case mrtFixedRate
                '*-create the FIXED element
                Call vobjCommon.CreateNewElement("FIXED", xmlReverts)
            Case mrtCappedRate
                '*-create the CAPPED element
                Call vobjCommon.CreateNewElement("CAPPED", xmlReverts)
            Case mrtCollaredRate
                '*-create the COLLARED element
                Call vobjCommon.CreateNewElement("COLLARED", xmlReverts)
            Case mrtCappedAndCollaredRate
                '*-create the CAPPEDANDCOLLARED element
                Call vobjCommon.CreateNewElement("CAPPEDANDCOLLARED", xmlReverts)
        End Select
        
        If vintNumLoanComponents > 1 Then
            '*-add the MULTIPART elememt
            Set xmlTemp = vobjCommon.CreateNewElement("MULTIPART", xmlReverts)
            '*-add the PART attribute
            Call xmlSetAttributeValue(xmlTemp, "PART", CStr(vintLoanItem))
        End If
        
        If eType = mrtStandardVariableRate Or eType = mrtDiscountedRate _
                Or eType = mrtTrackerAbove Then
            '*-add the ONEPERCENTINCREASE element if applicable
            Set xmlTemp = vobjCommon.CreateNewElement("ONEPERCENTINCREASE", xmlReverts)
            '*-add the INCREASEAMOUNT attribute
'SR 21/09/2004 : CORE82
'            strAmount = Get1IncDiffFromLoanCompPaySchedule(vobjHelper.LoanComponent, _
'                    intSeqNum)
            strAmount = GetFirstValid1IncDiffFromLCPayScheduleForLC(vobjHelper.LoanComponent)
'SR 21/09/2004 : CORE82 - End
            Call xmlSetAttributeValue(xmlTemp, "INCREASEAMOUNT", strAmount)
        End If
        
    End If

    Set xmlInterestRate = Nothing
    Set xmlPrevInterestRate = Nothing
    Set xmlReverts = Nothing
    Set xmlTemp = Nothing
    Set xmlRateDescription = Nothing
Exit Sub
ErrHandler:
    Set xmlInterestRate = Nothing
    Set xmlPrevInterestRate = Nothing
    Set xmlReverts = Nothing
    Set xmlTemp = Nothing
    Set xmlRateDescription = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       Section7_CreateMaxMinRate
'** Created by:     Andy Maggs
'** Date:           19/07/2004
'** Description:    Creates the MAXMINRATE element and child elements and
'**                 attributes.
'** Parameters:     vobjCommon - the common data helper.
'**                 vobjHelper - the interest rates type helper.
'**                 vxmlInterestRate - the current interest rate type XML.
'**                 veType - the interest rate type.
'**                 vintItemIndex - the cardinal number of this rate type.
'**                 vxmlNode - the node to create the elements on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_CreateMaxMinRate(ByVal vobjCommon As CommonDataHelper, _
        ByVal vobjHelper As IntRateTypesHelper, ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal veType As MortgageInterestRateType, ByVal vintItemIndex As Integer, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "Section7_CreateMaxMinRate"
    
    Dim xmlMaxMinRate As IXMLDOMNode
    Dim xmlRateType As IXMLDOMNode
    Dim dblMaxRate As Double
    Dim dblMinRate As Double
    Dim blnHasEndDate As Boolean
    Dim dtEndDate As Date
    Dim intPeriod As Integer
    Dim xmlTemp As IXMLDOMNode
    Dim intSeqNum As Integer
    Dim dblMinMonthlyCost As Double
    Dim dblMaxMonthlyCost As Double

    On Error GoTo ErrHandler

    'SR 12/09/2004 : CORE82 - Create MAXMIN element only for interest rate type
    '                Capped, Collared, CappedAndCollared
    If Not (veType = mrtCappedRate Or veType = mrtCollaredRate Or veType = mrtCappedAndCollaredRate) Then
        Exit Sub
    End If
    
    '*-create the MAXMINRATE element
    Set xmlMaxMinRate = vobjCommon.CreateNewElement("MAXMINRATE", vxmlNode)
    '*-create the MAXRATE attribute
    dblMaxRate = xmlGetAttributeAsDouble(vxmlInterestRate, "CEILINGRATE")
    Call xmlSetAttributeValue(xmlMaxMinRate, "MAXRATE", set2DP(CStr(dblMaxRate)))
    '*-create the MINRATE attribute
    dblMinRate = xmlGetAttributeAsDouble(vxmlInterestRate, "FLOOREDRATE")
    Call xmlSetAttributeValue(xmlMaxMinRate, "MINRATE", set2DP(CStr(dblMinRate)))
    '*-create the CARDINALNUMBER attribute
    'INR CORE82 Need to convert number to Cardinal before setting attribute
    Call xmlSetAttributeValue(xmlMaxMinRate, "CARDINALNUMBER", Section7_CreateCardinal(vintItemIndex))
            
    Select Case veType
        Case mrtFixedRate
            '*-create the FIXED element
            Set xmlRateType = vobjCommon.CreateNewElement("FIXED", xmlMaxMinRate)
        Case mrtCappedRate
            '*-create the CAPPED element
            Set xmlRateType = vobjCommon.CreateNewElement("CAPPED", xmlMaxMinRate)
        Case mrtCollaredRate
            '*-create the COLLARED element
            Set xmlRateType = vobjCommon.CreateNewElement("COLLARED", xmlMaxMinRate)
        Case mrtCappedAndCollaredRate
            '*-create the CAPPEDANDCOLLARED element
            Set xmlRateType = vobjCommon.CreateNewElement("CAPPEDANDCOLLARED", xmlMaxMinRate)
    End Select
    
    If veType = mrtCappedRate Or veType = mrtCappedAndCollaredRate Then
        '*-create the MAXIMUMCAP element
        If dblMaxRate > 0 Then
            Call vobjCommon.CreateNewElement("MAXIMUMCAP", xmlMaxMinRate)
        End If
    End If
    
    '*-create the AND element
    'INR AQR CORE82
    If veType = mrtCappedAndCollaredRate Then
        Call vobjCommon.CreateNewElement("AND", xmlMaxMinRate)
    End If
    
    If veType = mrtCollaredRate Or veType = mrtCappedAndCollaredRate Then
        If dblMinRate > 0 Then
            '*-create the MINIMUMFLOOR element
            Call vobjCommon.CreateNewElement("MINIMUMFLOOR", xmlMaxMinRate)
        End If
'INR ARQ CORE82        '*-create the AND element
'        Call vobjCommon.CreateNewElement("AND", xmlMaxMinRate)
    End If
    
    '*-get the information about the term of this interest rate
    Call GetBasicInterestRateTermData(vobjCommon, vobjHelper.InterestRates, _
            vxmlInterestRate, blnHasEndDate, dtEndDate, intPeriod)
    If intPeriod = -1 Then
        '*-create the REMAININGTERM element
        Call vobjCommon.CreateNewElement("REMAININGTERM", xmlMaxMinRate)
    End If
    
    '*-add the UNTILDATE or FORPERIOD elements using this term data
    Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
            intPeriod, xmlMaxMinRate)
            
    ' MAR54 - Add check for Capped and Collared to test
    If veType = mrtCappedRate Or veType = mrtCappedAndCollaredRate Or veType = mrtCollaredRate Then
        intSeqNum = xmlGetAttributeAsInteger(vxmlInterestRate, "INTERESTRATETYPESEQUENCENUMBER")
        Call GetRatePeriodIncreaseCost(vobjHelper.LoanComponent, intSeqNum, _
                dblMinMonthlyCost, dblMaxMonthlyCost)
        If veType = mrtCappedRate Or veType = mrtCappedAndCollaredRate Then
            '*-create the NOTEXCEED element
            Set xmlTemp = vobjCommon.CreateNewElement("NOTEXCEED", xmlMaxMinRate)
            '*-create the CALC attribute
            Call xmlSetAttributeValue(xmlTemp, "CALC", CStr(dblMaxMonthlyCost))
        ElseIf veType = mrtCollaredRate Or veType = mrtCappedAndCollaredRate Then
            '*-create the NOTLESSTHAN element
            Set xmlTemp = vobjCommon.CreateNewElement("NOTLESSTHAN", xmlMaxMinRate)
            '*-create the CALC attribute
            Call xmlSetAttributeValue(xmlTemp, "CALC", CStr(dblMinMonthlyCost))
        End If
    End If
    
    '*-is this the last rate type record?
    If vintItemIndex < vobjHelper.RatesCount Then
        '*-create the UNTILAFTERTHISDATE element
        Call vobjCommon.CreateNewElement("UNTILAFTERTHISDATE", xmlMaxMinRate)
    End If

    Set xmlMaxMinRate = Nothing
    Set xmlRateType = Nothing
    Set xmlTemp = Nothing
Exit Sub
ErrHandler:
    Set xmlMaxMinRate = Nothing
    Set xmlRateType = Nothing
    Set xmlTemp = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       Section7_AddIncreasePercentageElement
'** Created by:     Andy Maggs
'** Date:           19/07/2004
'** Description:    Adds the INCREASEPERCENTAGE element if applicable.
'** Parameters:     vobjCommon - the common data helper.
'**                 vobjHelper - the interest rate types helper object.
'**                 vxmlInterestRate - the current interest rate.
'**                 veType - the type of the current interest rate.
'**                 vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_AddIncreasePercentageElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vobjHelper As IntRateTypesHelper, ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal veType As MortgageInterestRateType, ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "Section7_AddIncreasePercentageElement"
    
    Dim intSeqNum As Integer
    Dim xmlIncrease As IXMLDOMNode
    Dim strAmount As String
    Dim xmlNext As IXMLDOMNode
    Dim eNextType As MortgageInterestRateType
    Dim eCurrentType As MortgageInterestRateType
    Dim xmlCurrent As IXMLDOMNode
    Dim xmlRateDescription As IXMLDOMNode
    Dim strDesc As String
    
    On Error GoTo ErrHandler
    
    intSeqNum = xmlGetAttributeAsInteger(vxmlInterestRate, "INTERESTRATETYPESEQUENCENUMBER")
    Set xmlCurrent = vobjHelper.GetInterestRateWithSequenceNum(intSeqNum)
    
        'CORE82
        '*-Just create the INCREASEPERCENTAGE element if  the
        '*-rate is variable, discount or tracker above
        If Not xmlCurrent Is Nothing Then
            
                eCurrentType = GetInterestRateType(xmlCurrent)
'MAR54 Don't know why this process does not allow CAPPED and COLLARED products
'                If eCurrentType = mrtStandardVariableRate Or eCurrentType = mrtDiscountedRate _
'                        Or eCurrentType = mrtTrackerAbove Then
                 If Not (eCurrentType = mrtFixedRate) Then
                    '*-create the INCREASEPERCENTAGE element
                    Set xmlIncrease = vobjCommon.CreateNewElement("INCREASEPERCENTAGE", vxmlNode)
                    '*-add the INCREASEAMOUNT attribute
                    strAmount = Get1IncDiffFromLoanCompPaySchedule(vobjHelper.LoanComponent, _
                            intSeqNum)
                    Call xmlSetAttributeValue(xmlIncrease, "INCREASEAMOUNT", strAmount)
                    '*-create the RATETYPE attribute
                    Set xmlRateDescription = xmlCurrent.selectSingleNode("//BASERATE")
                    strDesc = xmlGetAttributeText(xmlRateDescription, "RATEDESCRIPTION")
                    Call xmlSetAttributeValue(xmlIncrease, "RATETYPE", _
                     strDesc)
'CORE82                    Call xmlSetAttributeValue(xmlIncrease, "RATETYPE", _
'                            vobjHelper.InterestRateTypeAsText(veType))
                   End If
                            
        End If

    Set xmlIncrease = Nothing
    Set xmlNext = Nothing
    Set xmlCurrent = Nothing
    Set xmlRateDescription = Nothing
Exit Sub
ErrHandler:
    Set xmlIncrease = Nothing
    Set xmlNext = Nothing
    Set xmlCurrent = Nothing
    Set xmlRateDescription = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'SR 21/09/2004 : CORE82
'********************************************************************************
'** Function:       Section7_AddLCBasedIncreasePercentageElement
'** Created by:     Srini Rao
'** Date:           21/09/2004
'** Description:    Adds the INCREASEPERCENTAGE element if applicable (once for each Loan Component).
'** Parameters:     vobjCommon - the common data helper.
'**                 vobjHelper - the interest rate types helper object.
'**                 vxmlLC - the current Loan Component.
'**                 vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub Section7_AddLCBasedIncreasePercentageElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vobjHelper As IntRateTypesHelper, ByVal vxmlLC As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    
    Const cstrFunctionName As String = "Section7_AddLCBasedIncreasePercentageElement"
    
    Dim intSeqNum As Integer
    Dim xmlIncrease As IXMLDOMNode
    Dim strAmount As String
    Dim xmlNext As IXMLDOMNode
    Dim eNextType As MortgageInterestRateType
    Dim eCurrentType As MortgageInterestRateType
    Dim xmlCurrent As IXMLDOMNode
    Dim xmlRateDescription As IXMLDOMNode
    Dim strDesc As String
    
    On Error GoTo ErrHandler
    
    Set xmlCurrent = vobjHelper.GetInterestRateWithSequenceNum(1)
    
    '*-Just create the INCREASEPERCENTAGE element if  the
    '*-rate is variable, discount or tracker above
    If Not xmlCurrent Is Nothing Then
        eCurrentType = GetInterestRateType(xmlCurrent)
        If eCurrentType = mrtStandardVariableRate Or eCurrentType = mrtDiscountedRate _
                Or eCurrentType = mrtTrackerAbove Then
            
            '*-create the INCREASEPERCENTAGE element
            Set xmlIncrease = vobjCommon.CreateNewElement("INCREASEPERCENTAGE", vxmlNode)
            '*-add the INCREASEAMOUNT attribute
            strAmount = GetFirstValid1IncDiffFromLCPayScheduleForLC(vobjHelper.LoanComponent)
            Call xmlSetAttributeValue(xmlIncrease, "INCREASEAMOUNT", strAmount)
            '*-create the RATETYPE attribute
            Set xmlRateDescription = xmlCurrent.selectSingleNode("//BASERATE")
            strDesc = xmlGetAttributeText(xmlRateDescription, "RATEDESCRIPTION")
            Call xmlSetAttributeValue(xmlIncrease, "RATETYPE", _
             strDesc)
        End If
    End If

    Set xmlIncrease = Nothing
    Set xmlNext = Nothing
    Set xmlCurrent = Nothing
    Set xmlRateDescription = Nothing
Exit Sub
ErrHandler:
    Set xmlIncrease = Nothing
    Set xmlNext = Nothing
    Set xmlCurrent = Nothing
    Set xmlRateDescription = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'SR 21/09/2004 : CORE82 - End

'SR 21/09/2004 : CORE82
'****************************************************************************************
' Function:       GetFirstValid1IncDiffFromLCPayScheduleForLC
' Description:    Gets the first valid 1% IncreasedMonthlyCostDiifference for the current Loan component
'Parameters:      vxmlLoanComponent - current loan component
'Return   :       IncreasedMonthlyCostDifference (from LoanComponentPaymentSchedule)
'*****************************************************************************************
Private Function GetFirstValid1IncDiffFromLCPayScheduleForLC(ByVal vxmlLoanComponent As IXMLDOMNode) As String
    Const cstrFunctionName As String = "GetFirstValid1IncDiffFromLCPayScheduleForLC"
    Dim strReturnVal As String
    Dim xmlLCPaySchedule As IXMLDOMNode, xmlLCPayScheduleList As IXMLDOMNodeList
    Dim intCount As Integer, intTotalLCPaySchedules As Integer
    Dim dblTemp As Double
    
    On Error GoTo ErrHandler

    strReturnVal = ""
    
    Set xmlLCPayScheduleList = vxmlLoanComponent.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
    intTotalLCPaySchedules = xmlLCPayScheduleList.length
  
'   MAR44 BC 12 Sep 05
'   For intCount = 1 To intTotalLCPaySchedules - 1 Step 1 'Ignore the first schedule, it is always accrued interest
'    For intCount = 0 To intTotalLCPaySchedules Step 1
    For intCount = 0 To intTotalLCPaySchedules - 1 Step 1
       
        Set xmlLCPaySchedule = xmlLCPayScheduleList.Item(intCount)
        dblTemp = xmlGetAttributeAsDouble(xmlLCPaySchedule, "INCREASEDMONTHLYCOSTDIFFERENCE")
        If dblTemp > 0 Then
            strReturnVal = CStr(dblTemp)
            Exit For
        End If
    Next
        
    GetFirstValid1IncDiffFromLCPayScheduleForLC = strReturnVal
    
    Set xmlLCPaySchedule = Nothing
    Set xmlLCPayScheduleList = Nothing
Exit Function
ErrHandler:
    Set xmlLCPaySchedule = Nothing
    Set xmlLCPayScheduleList = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function
'SR 21/09/2004 : CORE82 - End

'****************************************************************************************
' Function:       Get1IncDiffFromLoanCompPaySchedule
' Description:    Gets 1% IncreasedMonthlyCostDiifference for the current Loan component
'                 and InterestRateType (record)
'Parameters:      vxmlLoanComponent - current loan component
'                 vintIntRateSeqNo - current InterestRate sequence number
'Return   :       IncreasedMonthlyCostDifference (from LoanComponentPaymentSchedule)
'*****************************************************************************************
Private Function Get1IncDiffFromLoanCompPaySchedule(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vintIntRateSeqNo As Integer) As String
    Const cstrFunctionName As String = "Get1IncDiffFromLoanCompPaySchedule"
    Dim strReturnVal As String
    Dim strCondition As String
    Dim xmlLCPaySchedule As IXMLDOMNode
    
    On Error GoTo ErrHandler

    strReturnVal = ""
    
    strCondition = ".//LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=" & CStr(vintIntRateSeqNo) & "]"
    Set xmlLCPaySchedule = vxmlLoanComponent.selectSingleNode(strCondition)
    
    If Not xmlLCPaySchedule Is Nothing Then
        If xmlAttributeValueExists(xmlLCPaySchedule, "INCREASEDMONTHLYCOSTDIFFERENCE") Then
            strReturnVal = CStr(xmlGetAttributeAsDouble(xmlLCPaySchedule, "INCREASEDMONTHLYCOSTDIFFERENCE"))
        End If
    End If
    
    Get1IncDiffFromLoanCompPaySchedule = strReturnVal

    Set xmlLCPaySchedule = Nothing
Exit Function
ErrHandler:
    Set xmlLCPaySchedule = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

''********************************************************************************
''** Function:       GetRatePeriodMonthlyCost
''** Created by:     Srini Rao
''** Date:
''** Description:    Gets the monthly cost for the current interest rate type.
''** Parameters:     vxmlLoanComponent - the current loan component.
''**                 vintIntRateSeqNo - the interest rate sequence number to get
''**                 the monthly cost for.
''** Returns:        Monthly Cost (from LoanComponentPaymentSchedule).
''** Errors:         None Expected
''********************************************************************************
'Private Function GetRatePeriodMonthlyCost(ByVal vxmlLoanComponent As IXMLDOMNode, _
'        ByVal vintIntRateSeqNo As Integer) As String
'    Const cstrFunctionName As String = "GetRatePeriodMonthlyCost"
'    Dim strReturnVal As String
'    Dim strCondition As String
'    Dim xmlLCPaySchedule As IXMLDOMNode
'
'    On Error GoTo ErrHandler
'
'    strReturnVal = ""
'
'    strCondition = ".//LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=" & CStr(vintIntRateSeqNo) & "]"
'    Set xmlLCPaySchedule = vxmlLoanComponent.selectSingleNode(strCondition)
'
'    If Not xmlLCPaySchedule Is Nothing Then
'        If xmlAttributeValueExists(xmlLCPaySchedule, "MONTHLYCOST") Then
'            strReturnVal = CStr(xmlGetAttributeAsDouble(xmlLCPaySchedule, "MONTHLYCOST"))
'        End If
'    End If
'
'    GetRatePeriodMonthlyCost = strReturnVal
'
'Exit Function
'ErrHandler:
'    errCheckError cstrFunctionName, mcstrModuleName
'End Function

'****************************************************************************************
' Function:       GetRatePeriodIncreaseCost
' Description:    Gets the Min and Max monthly cost for the current interest rate type
'Parameters:      xmlLoanComponent - current loan component
'                 vintIntRateSeqNo - current InterestRate sequence number
'Return   :       MaxMonthlyCost and MinMonthlyCost (from LoanComponentPaymentSchedule)
'*****************************************************************************************
Private Sub GetRatePeriodIncreaseCost(ByVal xmlLoanComponent As IXMLDOMNode, _
        ByVal vintIntRateSeqNo As Integer, _
        ByRef rdblLCPSMinMonthlyCost As Double, _
        ByRef rdblLCPSMaxMonthlyCost As Double)
    Const cstrFunctionName As String = "GetRatePeriodIncreaseCost"
    Dim strCondition As String
    Dim xmlLCPaySchedule As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    strCondition = ".//LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=" & CStr(vintIntRateSeqNo) & "]"
    Set xmlLCPaySchedule = xmlLoanComponent.selectSingleNode(strCondition)
    
    If Not xmlLCPaySchedule Is Nothing Then
        rdblLCPSMaxMonthlyCost = xmlGetAttributeAsDouble(xmlLCPaySchedule, "MAXMONTHLYCOST")
        rdblLCPSMinMonthlyCost = xmlGetAttributeAsDouble(xmlLCPaySchedule, "MINMONTHLYCOST")
    End If
    
    Set xmlLCPaySchedule = Nothing
Exit Sub
ErrHandler:
    Set xmlLCPaySchedule = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
                                         
'********************************************************************************
'** Function:       BuildSection7ACommon
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section7A
'**                 (Total Borrowing) element.
'** Parameters:     vxmlNode - the section7A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildSection7ACommon(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7ACommon"
    Dim dblTotalBorrowing As Double
    Dim dblTotalLCAmount As Double
    Dim dblTotalOSBalance As Double
    Dim dblMonthlyPayment As Double
    Dim xmlItems As IXMLDOMNodeList
    Dim xmlLCPaymentSChedules As IXMLDOMNodeList
    Dim xmlLoanComps As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlLCPaymentSChedule As IXMLDOMNode
    Dim xmlLoanComp As IXMLDOMNode
    Dim xmlPartNode As IXMLDOMNode
'PB 30/11/2006 EP2_139
    Dim blnInterestOnly As Boolean
    'EP2_1450
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlIntRateTypeCurrent As IXMLDOMNode
    Dim xmlIntRateTypeNext As IXMLDOMNode
    Dim dblExistingLoanPayment As Double
    Dim intCount As Integer
    Dim rateEndDate As String
    Dim ratePeriodType As String
    Dim ratePeriod As String
    Dim mortgageType5 As String
    Dim eType As MortgageInterestRateType
    Dim futureMonthlyPayment As Double
    
    On Error GoTo ErrHandler

    '----------------------------------------------------
    'Total Borrowing
    '-----------------------------------------------------
    'Total Loan Component Amount
    dblTotalLCAmount = 0
    Set xmlLoanComps = vobjCommon.LoanComponents
    For Each xmlLoanComp In xmlLoanComps
        dblTotalLCAmount = dblTotalLCAmount + xmlGetAttributeAsDouble(xmlLoanComp, "LOANAMOUNT")
        
        
    Next xmlLoanComp
'PB 30/11/2006 EP2_139
    'EP2_1450 also take into account whether there is a repayment vehicle
    If blnInterestOnly Then
        vobjCommon.CreateNewElement "INTERESTONLYNOREPAYMENTVEHICLE", vxmlNode
    End If
                
    'Total Outstanding Balance
    Set xmlItems = vobjCommon.Data.selectNodes(".//ACCOUNT/MORTGAGEACCOUNT/MORTGAGELOAN")
    dblTotalOSBalance = 0
    
    For Each xmlItem In xmlItems
        dblTotalOSBalance = xmlGetAttributeAsDouble(xmlItem, "OUTSTANDINGBALANCE")
        'EP2_1450
        dblExistingLoanPayment = xmlGetAttributeAsDouble(xmlItem, "MONTHLYREPAYMENT")
    Next xmlItem
    
    dblTotalBorrowing = dblTotalLCAmount + dblTotalOSBalance
    
    xmlSetAttributeValue vxmlNode, "TOTALBORROWING", dblTotalBorrowing
    
    '----------------------------------------------------
    'Initial Monthly Payment
    '-----------------------------------------------------
    dblMonthlyPayment = 0
    
    Set xmlItems = vobjCommon.LoanComponents
    If xmlItems.length > 1 Then
        mortgageType5 = "a part of your additional borrowing"
    Else
        mortgageType5 = "your additional borrowing"
    End If
 
    
    'EP2_1450
    For Each xmlItem In xmlItems
        Dim intRateTypeSeq As String
        Dim selectString As String
        
        Set xmlLCPaymentSChedules = xmlItem.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
        
        For intCount = 0 To xmlLCPaymentSChedules.length - 1 Step 1
           
            Set xmlLCPaymentSChedule = xmlLCPaymentSChedules.Item(intCount)
            dblMonthlyPayment = dblExistingLoanPayment + xmlGetAttributeAsDouble(xmlLCPaymentSChedule, "MONTHLYCOST")
            intRateTypeSeq = xmlGetAttributeText(xmlLCPaymentSChedule, "INTERESTRATETYPESEQUENCENUMBER")
            
            selectString = ".//MORTGAGEPRODUCT/INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER = '" + intRateTypeSeq + "']"
            Set xmlIntRateTypeCurrent = xmlItem.selectSingleNode(selectString)
            Dim nextSeq As String
            nextSeq = intRateTypeSeq + 1
            If nextSeq <= xmlLCPaymentSChedules.length Then
                selectString = ".//MORTGAGEPRODUCT/INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER = '" + nextSeq + "']"
                Set xmlIntRateTypeNext = xmlItem.selectSingleNode(selectString)
                futureMonthlyPayment = dblExistingLoanPayment + xmlGetAttributeAsDouble(xmlLCPaymentSChedules.Item(intCount + 1), "MONTHLYCOST")
            End If
            ratePeriod = xmlGetAttributeText(xmlIntRateTypeCurrent, "INTERESTRATEPERIOD")

            eType = GetInterestRateType(xmlIntRateTypeCurrent)
 
            If eType = mrtFixedRate Then
                'FIXED
                ratePeriodType = "Fixed"
            Else
                'VARIABLE  eType = mrtStandardVariableRate  Or eType = mrtDiscountedRate
                ratePeriodType = "Variable"
            End If

            If intCount = 0 Then
                xmlSetAttributeValue vxmlNode, "INITIALMONTHLYPAYMENT", dblMonthlyPayment
            End If
            
            If Not xmlIntRateTypeNext Is Nothing Then
                Set xmlPartNode = vobjCommon.CreateNewElement("PART", vxmlNode)
                'If we have an EndDate
                If xmlGetAttributeAsDate(xmlIntRateTypeCurrent, "INTERESTRATEENDDATE") Then
                    rateEndDate = xmlGetAttributeAsDate(xmlIntRateTypeCurrent, "INTERESTRATEENDDATE")
                    Set xmlTempNode = vobjCommon.CreateNewElement("RATEENDDATE", xmlPartNode)
                    xmlSetAttributeValue xmlTempNode, "RATEENDDATE", rateEndDate
                ElseIf xmlGetAttributeText(xmlIntRateTypeCurrent, "INTERESTRATEPERIOD") Then
                    Set xmlTempNode = vobjCommon.CreateNewElement("RATEPERIOD", xmlPartNode)
                    ratePeriod = xmlGetAttributeText(xmlIntRateTypeCurrent, "INTERESTRATEPERIOD")
                    xmlSetAttributeValue xmlTempNode, "PERIOD", ratePeriod
                End If
                xmlSetAttributeValue xmlPartNode, "PERIODDESCRIPTION", ratePeriodType
                xmlSetAttributeValue xmlPartNode, "MORTGAGEOFFERTYPE5", mortgageType5
                xmlSetAttributeValue xmlPartNode, "NEWAMOUNT", futureMonthlyPayment
            End If
                
            Set xmlIntRateTypeNext = Nothing
        Next intCount
        
    Next xmlItem
        
    '----------------------------------------------------
    'Interest Rate Period
    '-----------------------------------------------------

    Set xmlItems = Nothing
    Set xmlLCPaymentSChedules = Nothing
    Set xmlLoanComps = Nothing
    Set xmlItem = Nothing
    Set xmlLCPaymentSChedule = Nothing
    Set xmlLoanComp = Nothing
    Set xmlPartNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlIntRateTypeCurrent = Nothing
    Set xmlIntRateTypeNext = Nothing
    
Exit Sub
ErrHandler:
    Set xmlItems = Nothing
    Set xmlLCPaymentSChedules = Nothing
    Set xmlLoanComps = Nothing
    Set xmlItem = Nothing
    Set xmlLCPaymentSChedule = Nothing
    Set xmlLoanComp = Nothing
    Set xmlPartNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlIntRateTypeCurrent = Nothing
    Set xmlIntRateTypeNext = Nothing

    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection11Common
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section11
'**                 (What happens if you want to make overpayments?) element.
'** Parameters:     vxmlNode - the section11 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildSection11Common(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildSection11Common"
    
    On Error GoTo ErrHandler
    
    Dim xmlOverPayments As IXMLDOMNode 'EP2_1629
    'EP2_139
    If vobjCommon.IsFlexible Then
        'PB 05/03/2007 EP2_1629
        Set xmlOverPayments = vobjCommon.CreateNewElement("OVERPAYMENTS", vxmlNode)
        'EP2_2429 Always have section2 now
'        If vobjCommon.IsTransferOfEquity Or vobjCommon.IsProductSwitch Then
'            xmlSetAttributeValue xmlOverPayments, "SECTIONNUMBER", "9"
'        Else
            xmlSetAttributeValue xmlOverPayments, "SECTIONNUMBER", "10"
'        End If
    Else
        Call vobjCommon.CreateNewElement("NOOVERPAYMENTS", vxmlNode)
    End If
    
Exit Sub
ErrHandler:
    
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'*****************************************************************************************************
'** Function:       AddMortgageComponentsDetails
'** Created by:     Srini Rao
'** Date:           04/06/2004
'** Description:    Adds all MortgageComponent's data (Rates, Insurance and restrictions)
'                   Called From LifeTime KFI - Section 5 and Standard Mortgage KFI  - Section 4
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - The node to add the information to.
'** Returns:        N/A
'** Errors:         None Expected
'****************************************************************************************************
Public Sub AddMortgageComponentsDetails(ByVal vobjCommon As CommonDataHelper, _
                                        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddMortgageComponentsDetails"
    Dim xmlItem As IXMLDOMNode
    Dim xmlFT As IXMLDOMNode
    Dim xmlRate As IXMLDOMNode
    Dim xmlItems As IXMLDOMNodeList
    Dim xmlMainComp As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    Dim xmlProduct As IXMLDOMNode
    Dim xmlLTV As IXMLDOMNode
    Dim xmlPurpose As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    ' PB 05/06/2006 EP561/MAR1590 begin
    Dim xmlComp2 As IXMLDOMNode
    ' PB EP561/MAR1590 End
    Dim xmlMortProd As IXMLDOMNode
    Dim xmlMonthsNode As IXMLDOMNode 'BC MAR1685 01/05/2006
    Dim dblPercent As Double
    Dim dblMaxPercent As Double
    Dim intLowestSeqNum As Integer
    Dim intPart As Integer
    Dim blnHasMoreRateTypes As Boolean
    Dim xmlLC As IXMLDOMNode, dblResolvedRate As Double ' SR 22/09/2004 : CORE82
    'BBG1587
    Dim xmlSpecialDeals As IXMLDOMNode
    'EP983 - reverse MAR1071
    'Dim strMaxPercent As String                         ' MAR1071

    On Error GoTo ErrHandler

     '*-add one of SINGLECOMPONENT or MULTICOMPONENT elements
    Set xmlMainComp = vobjCommon.AddComponentsTypeElement(vxmlNode)

    '*-create an element to temporarily store elements to be added to the
    '*-document later
    Set xmlTemp = vobjCommon.Document.createElement("TEMP")

    '*-get the loan components
    Set xmlItems = vobjCommon.LoanComponents
    If xmlItems.length = 1 Then
        '*-get the single loan component
        Set xmlItem = xmlItems.Item(0)

        '*-add the PRODUCT element
        Set xmlProduct = vobjCommon.CreateNewElement("PRODUCT", xmlMainComp)
        Call AddProductNameAttribute(xmlItem, xmlProduct, vobjCommon)
        
        ' SR 22/09/2004 : CORE82
        Set xmlLC = vobjCommon.SingleLoanComponent
        If xmlAttributeValueExists(xmlLC, "RESOLVEDRATE") Then
            dblResolvedRate = xmlGetAttributeAsDouble(xmlLC, "RESOLVEDRATE")
            xmlSetAttributeValue xmlProduct, "RESOLVEDRATE", set2DP(CStr(dblResolvedRate))
        End If
        ' SR 22/09/2004 : CORE82 - end

        '*-add the mortgage rate period elements
        Call AddRatePeriodElements(vobjCommon, xmlItem, xmlProduct, False, _
                 False)

        '*-add any Insurance elements
        Call AddTiedInsuranceElements(vobjCommon, xmlTemp)

        '*-get the mortgage product for this loan component and set the
        '*-maximum percentage
        Set xmlMortProd = xmlItem.selectSingleNode("MORTGAGEPRODUCT")
        dblMaxPercent = xmlGetAttributeAsDouble(xmlMortProd, "MAXIMUMLTV")
        'EP2_1449
        If Not vobjCommon.IsAdditionalBorrowing() Then
            vobjCommon.CreateNewElement "NOTADDITIONAL", xmlMainComp
        End If
    
    Else
        '*-get the lowest component sequence number and deduct 1 because we are going
        '*-to use this value to ensure that the PART attribute starts at 1
        intLowestSeqNum = GetMinIntAttribValue(xmlItems, "LOANCOMPONENTSEQUENCENUMBER") - 1

        '*-for each component, add a COMPONENT and an FT element
        For Each xmlItem In xmlItems
            Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", xmlMainComp)

            '*-add the INITIALRATE attribute
            Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlItem)
            Call AddRateAttribute(vobjCommon, xmlItem, xmlRate, xmlComponent, "INITIALRATE")
            '*-add the COMPONENTREPAYMENTMETHOD attribute
            Call AddRepaymentMethodAttribute(xmlItem, xmlComponent, "COMPONENTREPAYMENTMETHOD")
            
            'BC MAR1685 Begin
            If Not (xmlGetAttributeAsInteger(xmlItem, "TERMINMONTHS") = 0) Then
                Set xmlMonthsNode = vobjCommon.CreateNewElement("MONTHS", xmlComponent)
                '*-add the TERMINMONTHS attribute
                Call xmlCopyAttributeValue(xmlItem, xmlMonthsNode, "TERMINMONTHS", "TERMINMONTHS")
            End If
            'BC MAR1685 End
            
            '*-add the TERMINYEARS attribute
            Call xmlCopyAttributeValue(xmlItem, xmlComponent, "TERMINYEARS", "TERMINYEARS")
            '*-add the LOANCOMPONENTAMOUNT attribute
            Call xmlCopyAttributeValue(xmlItem, xmlComponent, "LOANAMOUNT", "LOANCOMPONENTAMOUNT")
            '*-add the PART attribute
            intPart = xmlGetAttributeAsInteger(xmlItem, "LOANCOMPONENTSEQUENCENUMBER") - intLowestSeqNum
            ' PB 05/06/2006 EP561/MAR1590
            'Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intPart))
            ' PB EP561/MAR1590 End

            '*-add the PRODUCT element
            Set xmlProduct = vobjCommon.CreateNewElement("PRODUCT", xmlComponent)
            Call AddProductNameAttribute(xmlItem, xmlProduct, vobjCommon)
            '*-add the first mortgage rate period element
            
            Call xmlSetAttributeValue(xmlItem, "LOANCOMPONENTSEQUENCENUMBER", intPart) 'BC MAR1736
            Call AddRatePeriodElements(vobjCommon, xmlItem, xmlProduct, True, False, _
                    blnHasMoreRateTypes)

            '*-add the subsequent rate periods to the FT element if any
            If blnHasMoreRateTypes Then
                'MAR88 - BC 05 Oct Create FT Element moved inside this 'If'
                Set xmlFT = vobjCommon.CreateNewElement("FT", xmlMainComp)
                Call AddRatePeriodElements(vobjCommon, xmlItem, xmlFT, False, True)
            End If

            '*-add any Insurance elements
            Call AddTiedInsuranceElements(vobjCommon, xmlTemp)

            '*-get the mortgage product for this loan component and update the
            '*-maximum percentage if it is larger
            Set xmlMortProd = xmlItem.selectSingleNode("MORTGAGEPRODUCT")
            dblPercent = xmlGetAttributeAsDouble(xmlMortProd, "MAXIMUMLTV")
            If dblMaxPercent < dblPercent Then
                dblMaxPercent = dblPercent
            End If
            
            'PB 06/06/06 EP651/MAR1590
            If (InStr(1, xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD"), "3", vbTextCompare)) Then
                Set xmlComp2 = xmlComponent.cloneNode(True)
                xmlMainComp.appendChild xmlComp2
                Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intPart) + "(a)")
                Call xmlSetAttributeValue(xmlComponent, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlItem, "CAPITALANDINTERESTELEMENT"))
                'PB 12/06/2006 EP730/MAR1831 Begin
                'Call xmlSetAttributeValue(xmlComponent, "COMPONENTREPAYMENTMETHOD", "Repayment")
                Call xmlSetAttributeValue(xmlComponent, "COMPONENTREPAYMENTMETHOD", "Repayment (Capital and Interest)")
                'PB EP730/MAR1831 End
            
                Call xmlSetAttributeValue(xmlComp2, "PART", CStr(intPart) + "(b)")
                Call xmlSetAttributeValue(xmlComp2, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlItem, "INTERESTONLYELEMENT"))
                Call xmlSetAttributeValue(xmlComp2, "COMPONENTREPAYMENTMETHOD", "Interest Only")
            Else
                Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intPart))
            End If
            'PB EP651/MAR1590 End
            'EP2_1449
            If Not vobjCommon.IsAdditionalBorrowing() Then
                vobjCommon.CreateNewElement "NOTADDITIONAL", xmlComponent
            End If

        Next xmlItem
    End If

    '*-now add any INSURANCE elements to the document
    For Each xmlItem In xmlTemp.childNodes
        If xmlItem.baseName = "INSURANCE" Then
            Call vxmlNode.appendChild(xmlItem)
        End If
    Next xmlItem

    '*-add the LTV element
'    Set xmlLTV = vobjCommon.CreateNewElement("LTV", vxmlNode)
    '*-add the MAXPERCENT attribute
    ' EP983 - reverse MAR1071
    'MAR1071 Use the Global Parameter for Maximum LTV percentage
'    Set xmlItem = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "MaximumLTVallowed" & Chr$(34) & "]")
'    strMaxPercent = xmlGetAttributeText(xmlItem, "PERCENTAGE")
'    Call xmlSetAttributeValue(xmlLTV, "MAXPERCENT", strMaxPercent)
    'EP2_583
     Call xmlSetAttributeValue(vxmlNode, "MAXLTV", CStr(dblMaxPercent))
    ' EP983 - end
'    If vobjCommon.LoanPurposeText = "PURPOSEPURCHASE" Then
'        '*-add the PURPOSEPURCHASE element
'        Set xmlPurpose = vobjCommon.AddLoanPurposeElement(xmlLTV)
'    End If

    '*-add the RESTRICTIONS element
    Call AddRestrictionsElements(vobjCommon, xmlItems, intLowestSeqNum, vxmlNode, dblMaxPercent)

    '*-add the CATSTANDARD element
    If IsCatStandardMortgage(vobjCommon.Data) Then
        Call vobjCommon.CreateNewElement("CATSTANDARD", vxmlNode)
    End If

    'BBG1587*-add the SPECIALDEALS elements
    'ASSUMPTION THAT THERE IS ONLY A SINGLE LOANCOMPONENT
    If IsSpecialDeals(vobjCommon) Then
        Set xmlSpecialDeals = vobjCommon.CreateNewElement("SPECIALDEALS", vxmlNode)
        AddSpecialDealElements vobjCommon, vobjCommon.SingleLoanComponent, xmlSpecialDeals
    End If

    Set xmlItem = Nothing
    Set xmlFT = Nothing
    Set xmlRate = Nothing
    Set xmlItems = Nothing
    Set xmlMainComp = Nothing
    Set xmlTemp = Nothing
    Set xmlProduct = Nothing
    Set xmlLTV = Nothing
    Set xmlPurpose = Nothing
    Set xmlComponent = Nothing
    Set xmlMortProd = Nothing
    Set xmlLC = Nothing
    Set xmlSpecialDeals = Nothing
    'PB 06/06/06 EP651/MAR1590
    Set xmlComp2 = Nothing
    'PB EP651/MAR1590 End
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlFT = Nothing
    Set xmlRate = Nothing
    Set xmlItems = Nothing
    Set xmlMainComp = Nothing
    Set xmlTemp = Nothing
    Set xmlProduct = Nothing
    Set xmlLTV = Nothing
    Set xmlPurpose = Nothing
    Set xmlComponent = Nothing
    Set xmlMortProd = Nothing
    Set xmlLC = Nothing
    Set xmlSpecialDeals = Nothing
    'PB 06/06/06 EP651/MAR1590
    Set xmlComp2 = Nothing
    'PB EP651/MAR1590 End
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'EP2_583 Clean this up
'*****************************************************************************************************
'** Function:       AddServicesProvided
'** Created by:     Srini Rao
'** Date:           04/06/2004
'** Description:    Adds the details of service provided
'                   Called From LifeTime KFI-Section 2 and Standard Mortgage KFI-Section 2
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - The node to add the information to.
'** Returns:        N/A
'** Errors:         None Expected
'****************************************************************************************************
Public Sub AddServicesProvided(ByVal vobjCommon As CommonDataHelper, _
                               ByVal vxmlNode As IXMLDOMNode, _
                               Optional ByVal vblnCallFromLifeTime As Boolean = False)
    Const cstrFunctionName As String = "AddServicesProvided"
    Dim xmlNewNode As IXMLDOMNode
    Dim strIntermediaryContactName As String
    Dim strIntermediaryCompanyName As String
    Dim strMORTGAGETYPE1 As String
    Dim strMORTGAGETYPE3 As String

    On Error GoTo ErrHandler
    
    'EP2_1627
    gblnSection2 = True
    '
    If vobjCommon.GetMainMortgageTypeGroup = "F" Then
        'Additional borrowing
        strMORTGAGETYPE1 = "additional borrowing"
        strMORTGAGETYPE3 = "mortgage for the additional borrowing"
    Else
        strMORTGAGETYPE1 = "mortgage"
        strMORTGAGETYPE3 = "mortgage"
    End If
    
    'Use the intermediary firm name
    If vobjCommon.IsIntroducedByIntermediary(strIntermediaryContactName, False, strIntermediaryCompanyName, True) Then
        '*-add the INTERMEDIARY element
        Set xmlNewNode = vobjCommon.CreateNewElement("INTERMEDIARY", vxmlNode)
        'add the INTERMEDIARYNAME attribute
        Call xmlSetAttributeValue(xmlNewNode, "INTERMEDIARYNAME", strIntermediaryCompanyName)
        Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE1", strMORTGAGETYPE1)
        Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE3", strMORTGAGETYPE3)
        '*-add a RECOMMENDS or NOTRECOMMENDS element to the INTERMEDIARY element
        Call AddRecommendsNotRecommendsElement(vobjCommon, xmlNewNode)
    Else
        '*-or add the LENDER element
        Set xmlNewNode = vobjCommon.CreateNewElement("LENDER", vxmlNode)
        Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE1", strMORTGAGETYPE1)
        Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE3", strMORTGAGETYPE3)
        '*-add a RECOMMENDS or NOTRECOMMENDS element to the LENDER element
        Call AddRecommendsNotRecommendsElement(vobjCommon, xmlNewNode)
    End If

    Set xmlNewNode = Nothing

Exit Sub
ErrHandler:
    Set xmlNewNode = Nothing

    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'********************************************************************************
'** Function:       Section7_CreateCardinal
'** Created by:     Ian Ross
'** Date:           07/09/2004
'** Description:    Creates the Cardinal Number
'**
'** Parameters:     vintItemIndex - the cardinal number of this rate type.
'**
'** Returns:        Cardinal Number
'** Errors:         None Expected
'********************************************************************************

Private Function Section7_CreateCardinal(ByVal vintItemIndex As Integer) As String

    Const cstrFunctionName As String = "Section7_CreateCardinal"
    
    Dim intLastNum As Integer
    Dim strCardinal As String

    On Error GoTo ErrHandler
    
    strCardinal = CStr(vintItemIndex)
    If (vintItemIndex >= 10 And vintItemIndex <= 20) Then
        strCardinal = strCardinal + " th"
    Else
        intLastNum = Right(strCardinal, 1)
                
        Select Case intLastNum
            Case 1
                strCardinal = strCardinal + "st"
            Case 2
                strCardinal = strCardinal + "nd"
            Case 3
                strCardinal = strCardinal + "rd"
            Case 4, 5, 6, 7, 8, 9, 0
                strCardinal = strCardinal + "th"
        End Select
    End If
    
    Section7_CreateCardinal = strCardinal

Exit Function
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

Private Function GetInterestRateTypeText(eType As MortgageInterestRateType)
    '
    Select Case eType
        Case mrtStandardVariableRate
            GetInterestRateTypeText = "standard variable rate"
        Case mrtTrackerAbove
            GetInterestRateTypeText = "tracker rate"
        Case mrtDiscountedRate
            GetInterestRateTypeText = "discounted rate"
        Case mrtFixedRate
            GetInterestRateTypeText = "fixed rate"
        Case mrtCappedAndCollaredRate
            GetInterestRateTypeText = "capped and collared rate"
        Case mrtCollaredRate
            GetInterestRateTypeText = "collared rate"
        Case mrtCappedRate
            GetInterestRateTypeText = "capped rate"
    End Select
    '
End Function
