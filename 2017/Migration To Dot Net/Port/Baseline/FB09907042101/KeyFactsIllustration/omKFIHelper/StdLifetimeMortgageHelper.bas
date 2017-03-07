Attribute VB_Name = "StdLifetimeMortgageHelper"
'********************************************************************************
'** Module:         StdLifetimeMortgageHelper
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    This module contains functions/procedures that are common to
'**                 lifetime mortgage KFI documents (KFI and Offer).
'********************************************************************************
Option Explicit

Private Const mcstrModuleName As String = "StdLifetimeMortgageHelper"

'********************************************************************************
'** Function:       BuildLifeTimeSection3
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeSection3(ByVal vobjCommon As CommonDataHelper, _
                                 ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildLifeTimeSection3"

    On Error GoTo ErrHandler

    xmlSetAttributeValue vxmlNode, "EARLYREPAYMENTCHARGESSECTION", "13"
    
    If vobjCommon.IsAdditionalBorrowing() Then
        Call vobjCommon.CreateNewElement("ADDITIONALBORROWING", vxmlNode)
    Else
        Call vobjCommon.CreateNewElement("NOTADDITIONALBORROWING", vxmlNode)
    End If

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeSection5
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifetimeSection5(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildLifetimeSection5"
    Dim xmlNode As IXMLDOMNode
    Dim xmlNode2 As IXMLDOMNode
    Dim intNoOfApplicants As Integer, dblDrawDownAmount As Double
    
    On Error GoTo ErrHandler
    
    'SR 12/09/2004 : CORE82 - Add TERMYEARS and TERMMONTHS
    AddMortgageTermAttributes vobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS"
    
    '* - Add ROLLUP element to SECTION5 node
    Set xmlNode = vobjCommon.CreateNewElement("ROLLUP", vxmlNode)
    intNoOfApplicants = GetNumberOfApplicants(vobjCommon)
    If intNoOfApplicants = 1 Then 'SR 10/09/2004 : CORE82
        Set xmlNode2 = vobjCommon.CreateNewElement("FIRSTDEATH", xmlNode)
    Else
        If intNoOfApplicants > 1 Then 'SR 10/09/2004 : CORE82
            Set xmlNode2 = vobjCommon.CreateNewElement("LASTDEATH", xmlNode)
        End If
    End If
    
    Set xmlNode2 = vobjCommon.CreateNewElement("LUMPSUM", xmlNode)
    
    '* - Add DRAWDOWN element to SECTION5 node
    'SR 10/09/2004 : CORE82 - Create drawdown only when MortgageSubQuote.DrawDown > 0
    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrMORTGAGESUBQUOTE_PATH)
    If Not xmlNode Is Nothing Then
        If xmlGetAttributeAsDouble(xmlNode, "DRAWDOWN") > 0 Then
            Set xmlNode = vobjCommon.CreateNewElement("DRAWDOWN", vxmlNode)
            Set xmlNode2 = vobjCommon.CreateNewElement("ONREQUEST", xmlNode)
        End If
    End If 'SR 10/09/2004 : CORE82 - End
    
    '* - Add LUMPSUM element to SECTION5 node
    Set xmlNode = vobjCommon.CreateNewElement("LUMPSUM", vxmlNode) 'SR 10/09/2004 : CORE82
    
    '* - Add data of all the Loan Components
    AddMortgageComponentsDetails vobjCommon, vxmlNode
    
    Set xmlNode = Nothing
    Set xmlNode2 = Nothing
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    Set xmlNode2 = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeSection6
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeSection6(ByVal vobjCommon As CommonDataHelper, _
                                 ByVal vxmlNode As IXMLDOMNode, _
                                 ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection6"
    Dim xmlNode As IXMLDOMNode

    On Error GoTo ErrHandler

    Set xmlNode = vobjCommon.CreateNewElement("LUMPSUM", vxmlNode)
    xmlSetAttributeValue xmlNode, "AMOUNT", vobjCommon.LoanAmount
    If vblnIsOfferDocument Then
        Call AddMortgageTypeAttribute(vobjCommon, xmlNode)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("OTHERBENEFITS", vxmlNode) 'SR 11/09/2004 : CORE82
    
    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeSection7
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeSection7(ByVal vobjCommon As CommonDataHelper, _
                             ByVal vxmlNode As IXMLDOMNode, _
                             ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection7"
    Dim xmlNode As IXMLDOMNode
    
    On Error GoTo ErrHandler

    Set xmlNode = vobjCommon.CreateNewElement("REPOSSESSION", vxmlNode)
    '* TODO - hard-coded for now; this needs to be removed
    xmlSetAttributeValue xmlNode, "REPOSSESSIONPERIOD", 12 'SR 15/10/2004 : BBG1623
    If vblnIsOfferDocument Then
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("NEGATIVEEQUITY", vxmlNode)
    
    Set xmlNode = vobjCommon.CreateNewElement("MOVEHOME", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlNode)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("MOVEINTOCARE", vxmlNode)
    xmlSetAttributeValue xmlNode, "EARLYREPAYMENTCHARGESSECTION", "13"
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlNode)
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)
    End If
     
    Set xmlNode = vobjCommon.CreateNewElement("OTHERTENANT", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("GOODREPAIR", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlNode)
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("BENEFITSPOSITION", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlNode)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("ANOTHERMORTGAGE", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlNode = vobjCommon.CreateNewElement("DISCONTINUEDRAWDOWN", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeSection8
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeSection8(ByVal vobjCommon As CommonDataHelper, _
                                 ByVal vxmlNode As IXMLDOMNode, _
                                 ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection8"
                             
    Dim xmlNoRepayments As IXMLDOMNode, xmlRatePeriod As IXMLDOMNode 'SR 12/09/2004 : CORE82
    Dim xmlRollUp As IXMLDOMNode, xmlfees As IXMLDOMNode  'SR 20/09/2004 : CORE82
    Dim xmlLC As IXMLDOMNode
    Dim xmlLCBalanceScheduleList As IXMLDOMNodeList
    Dim xmlLCBalanceSchedule As IXMLDOMNode
    Dim intCount As Integer
    Dim intNumSchedules As Integer
    Dim dblOpeningBalance As Double
    Dim dblClosingBalance As Double
    Dim dblIntOnlyAmount As Double
    Dim dblFeeAmount As Double
    Dim adblArrayOfBalance() As Double
    Dim dblInterestCharged As Double 'SR 22/09/2004
    
    On Error GoTo ErrHandler
    
    Set xmlNoRepayments = vobjCommon.CreateNewElement("NOREPAYMENTS", vxmlNode)
    xmlSetAttributeValue xmlNoRepayments, "INTERESTCALCULATEDFREQUENCY", "monthly" 'SR 11/09/2004 : CORE82
    xmlSetAttributeValue xmlNoRepayments, "STANDARDVARIABLERATE", vobjCommon.StandardVariableRate 'SR 12/09/2004 : CORE82
    '*- Add attribute 'TERMINYEARS'
    AddMortgageTermAttributes vobjCommon, xmlNoRepayments, "TERMYEARS", "TERMMONTHS" 'SR 11/09/2004 : CORE82
    If vblnIsOfferDocument Then
        Call AddMortgageTypeAttribute(vobjCommon, xmlNoRepayments)
    End If
    
    Set xmlRatePeriod = vobjCommon.CreateNewElement("RATEPERIOD", xmlNoRepayments) 'SR 12/09/2004 : CORE82
    
    '*-TODO - assuming that there is only one LoanComponent.
    '*       - modify code for multiple loan components
    Set xmlLC = vobjCommon.SingleLoanComponent
    Set xmlLCBalanceScheduleList = xmlLC.selectNodes(".//LOANCOMPONENTBALANCESCHEDULE")
    
    intNumSchedules = xmlLCBalanceScheduleList.length
    ReDim adblArrayOfBalance(intNumSchedules)
    For intCount = 0 To intNumSchedules - 1 Step 1
    Set xmlLCBalanceSchedule = xmlLCBalanceScheduleList.Item(intCount)
        adblArrayOfBalance(intCount) = xmlGetAttributeAsDouble(xmlLCBalanceSchedule, "BALANCE")
    Next intCount
    
'IK 01/12/2004 E2EM00003125, E2EM00003126
    If vobjCommon.HasDrawDown Then
        adblArrayOfBalance(0) = vobjCommon.TotalLoanAmountLessDrawDown
    Else
        adblArrayOfBalance(0) = vobjCommon.TotalLoanAmount
    End If
'IK 01/12/2004 E2EM00003125, E2EM00003126
    For intCount = 1 To intNumSchedules - 1 Step 1
        Set xmlRollUp = vobjCommon.CreateNewElement("ROLLUP", xmlRatePeriod) 'SR 12/09/2004 : CORE82
            
        dblOpeningBalance = adblArrayOfBalance(intCount - 1)
        xmlSetAttributeValue xmlRollUp, "YEAR", intCount
        xmlSetAttributeValue xmlRollUp, "OPENINGBALANCE", set2DP(CStr(dblOpeningBalance))
'IK 01/12/2004 E2EM00003125, E2EM00003126
        If intCount = 1 Then
            If vobjCommon.HasDrawDown Then
                xmlSetAttributeValue xmlRollUp, "PAYMENT", set2DP(CStr(vobjCommon.DrawDownAmount))
'IK 03/12/2004 E2EM00003125
                dblOpeningBalance = vobjCommon.TotalLoanAmount
'IK 03/12/2004 E2EM00003125 ends
            Else
                xmlSetAttributeValue xmlRollUp, "PAYMENT", "0.00"
            End If
        Else
            xmlSetAttributeValue xmlRollUp, "PAYMENT", "0.00"
        End If
'IK 01/12/2004 E2EM00003125, E2EM00003126 ends
        
        Set xmlfees = vobjCommon.CreateNewElement("FEES", xmlRollUp)
        If intCount = 1 Then
            dblFeeAmount = 0
            dblClosingBalance = adblArrayOfBalance(intCount)
            dblInterestCharged = dblClosingBalance - dblOpeningBalance - dblFeeAmount
        ElseIf intCount = intNumSchedules - 1 Then
            dblFeeAmount = vobjCommon.MortgageCompletionFee
            dblClosingBalance = adblArrayOfBalance(intCount) + dblFeeAmount
            dblInterestCharged = adblArrayOfBalance(intCount) - dblOpeningBalance
        Else
            dblFeeAmount = 0
            dblClosingBalance = adblArrayOfBalance(intCount)
            dblInterestCharged = adblArrayOfBalance(intCount) - dblOpeningBalance
        End If

        xmlSetAttributeValue xmlfees, "FEEAMOUNT", set2DP(CStr(dblFeeAmount))
        xmlSetAttributeValue xmlRollUp, "INTERESTAMOUNT", set2DP(CStr(dblInterestCharged))
        xmlSetAttributeValue xmlRollUp, "CLOSINGBALANCE", set2DP(CStr(dblClosingBalance))
            
    Next intCount

    Set xmlNoRepayments = Nothing
    Set xmlRatePeriod = Nothing
    Set xmlRollUp = Nothing
    Set xmlfees = Nothing
    Set xmlLCBalanceScheduleList = Nothing
    Set xmlLCBalanceSchedule = Nothing
    Set xmlLC = Nothing
Exit Sub
ErrHandler:
    Set xmlNoRepayments = Nothing
    Set xmlRatePeriod = Nothing
    Set xmlRollUp = Nothing
    Set xmlfees = Nothing
    Set xmlLCBalanceScheduleList = Nothing
    Set xmlLCBalanceSchedule = Nothing
    Set xmlLC = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddFeeInfoToRollupNode
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Private Sub AddFeeInfoToRollupNode(ByVal vobjCommon As CommonDataHelper, _
                    ByVal vxmlRollUp As IXMLDOMNode, ByRef rastrFees() As String)
    Const cstrFunctionName As String = "AddFeeInfoToRollupNode"
    Dim intArrLength As Integer
    Dim intCount As Integer
    Dim xmlfees As IXMLDOMNode
    Dim xmlMore As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    intArrLength = UBound(rastrFees)
    For intCount = 0 To intArrLength Step 1
        If rastrFees(intCount) <> "" Then
            Set xmlfees = vobjCommon.CreateNewElement("FEES", vxmlRollUp)
            xmlSetAttributeValue xmlfees, "FEEAMOUNT", rastrFees(intCount)
        End If
        
        If intCount <> intArrLength Then
             Set xmlMore = vobjCommon.CreateNewElement("MORE", xmlfees)
        End If
    Next intCount

    Set xmlfees = Nothing
    Set xmlMore = Nothing
Exit Sub
ErrHandler:
    Set xmlfees = Nothing
    Set xmlMore = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeContactsSection
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeContactsSection(ByVal vobjCommon As CommonDataHelper, ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildLifeTimeContactsSection"
    Dim blnIsIntermediary As Boolean
    Dim strContactName As String

    On Error GoTo ErrHandler
    
    blnIsIntermediary = vobjCommon.IsIntroducedByIntermediary(strContactName, False)
    If blnIsIntermediary Then ' Add Intermediary details
        AddIntermediaryContactAttributes vobjCommon, vxmlNode
    Else ' Add Lender details
        AddLenderContactAttributes vobjCommon.Data, vxmlNode
    End If

    '*-add the mandatory REFERENCE attribute
    Set vxmlNode = vobjCommon.CreateNewElement("NONDISPOSABLE", vxmlNode)
    Call AddReferenceAttribute(vobjCommon.Data, vxmlNode)
    
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeSection15
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeSection15(ByVal vobjCommon As CommonDataHelper, _
                                 ByVal vxmlNode As IXMLDOMNode, _
                                 ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection15"
    Dim dblTotBalance As Double
    Dim dblFeeAmount As Double
    
    On Error GoTo ErrHandler
    
    Call AddMortgageTermAttributes(vobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS") 'SR 12/09/2004 : CORE82
    '*-add the mandatory APR attribute
    Call AddAPRAttribute(vobjCommon, vxmlNode)
    '*-add the mandatory TOTALREPAYMENTAMOUNT attribute
    'E2EM00002933 Add Sealing and deeds release fees to the total balance
    dblFeeAmount = vobjCommon.MortgageCompletionFee
    dblTotBalance = GetTotalCapIntBalance(vobjCommon) + dblFeeAmount
    xmlSetAttributeValue vxmlNode, "TOTALREPAYMENTAMOUNT", dblTotBalance
    '*-add the mandatory TERMININYEARS attribute
    AddMortgageTermAttributes vobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS"
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, vxmlNode)
    End If
    
    '*-add the VARIABLE element
    If vobjCommon.IsVariableRate Then  'SR 02/12/2004 : E2EM00003123
        vobjCommon.CreateNewElement "VARIABLE", vxmlNode
    End If 'SR 02/12/2004 : E2EM00003123

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetTotalCapIntBalance
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Function GetTotalCapIntBalance(ByVal vobjCommon As CommonDataHelper) As Double
    Const cstrFunctionName As String = "GetTotalCapIntBalance"
    Dim xmlLCList As IXMLDOMNodeList
    Dim xmlLCBalanceSchedList As IXMLDOMNodeList
    Dim xmlLC As IXMLDOMNode
    Dim xmlLCBalanceSched As IXMLDOMNode
    Dim dblFinalCapIntBalance As Double
    Dim dblTemp As Double
    Dim intCount As Integer
    
    On Error GoTo ErrHandler
    
    Set xmlLCList = vobjCommon.LoanComponents
    For Each xmlLC In xmlLCList
        Set xmlLCBalanceSchedList = xmlLC.selectNodes(".//LOANCOMPONENTBALANCESCHEDULE")
        intCount = xmlLCBalanceSchedList.length
        
        If intCount > 0 Then
            Set xmlLCBalanceSched = xmlLCBalanceSchedList.Item(intCount - 1)
            dblTemp = xmlGetAttributeAsDouble(xmlLCBalanceSched, "BALANCE")  'SR 12/09/2004 : CORE82
            dblFinalCapIntBalance = dblFinalCapIntBalance + dblTemp
        End If
    Next xmlLC
    
    GetTotalCapIntBalance = dblFinalCapIntBalance

    Set xmlLCList = Nothing
    Set xmlLCBalanceSchedList = Nothing
    Set xmlLC = Nothing
    Set xmlLCBalanceSched = Nothing
Exit Function
ErrHandler:
    Set xmlLCList = Nothing
    Set xmlLCBalanceSchedList = Nothing
    Set xmlLC = Nothing
    Set xmlLCBalanceSched = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildLifeTimeSection10
'** Created by:     Srini Rao
'** Date:
'** Description:
'** Parameters:
'** Returns:
'********************************************************************************
Public Sub BuildLifeTimeSection10(ByVal vobjCommon As CommonDataHelper, _
                                 ByVal vxmlNode As IXMLDOMNode, _
                                 ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection10"
    Dim dblEstValLessOnePercent As Double
    Dim dblEstValPlusOnePercent As Double
    Dim blnIsValuationUsed As Boolean
    Dim lngEstimatedPrice As Long
    Dim xmlRollUp As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    PurchasePriceOrEstimatedValue vobjCommon.Data, True, blnIsValuationUsed, lngEstimatedPrice 'SR 11/09/2004 : CORE82
    
    dblEstValLessOnePercent = CDbl(lngEstimatedPrice) * (0.99 ^ vobjCommon.TermInYears) 'SR 11/09/2004 : CORE82
    dblEstValPlusOnePercent = CDbl(lngEstimatedPrice) * (1.01 ^ vobjCommon.TermInYears) 'SR 11/09/2004 : CORE82
    
    xmlSetAttributeValue vxmlNode, "ESTIMATEDVALUELESSONEPERCENT", CLng(dblEstValLessOnePercent) 'SR 20/09/2004 : CORE82
    xmlSetAttributeValue vxmlNode, "ESTIMATEDVALUEPLUSONEPERCENT", CLng(dblEstValPlusOnePercent) 'SR 20/09/2004 : CORE82
    xmlSetAttributeValue vxmlNode, "TERMYEARS", vobjCommon.TermInYears
    
    xmlSetAttributeValue vxmlNode, "ESTIMATEDVALUE", lngEstimatedPrice
    
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, vxmlNode)
    End If
    
    Set xmlRollUp = vobjCommon.CreateNewElement("ROLLUP", vxmlNode)
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlRollUp)
        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlRollUp, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlRollUp = Nothing
Exit Sub
ErrHandler:
    Set xmlRollUp = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildLifeTimeSection9
'** Created by:     Andy Maggs
'** Date:
'** Description:    Builds section 9 () for both the lifetime mortgage KFI and
'**                 Offer document.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section 9 node to build.
'**                 vblnIsOfferDocument - true if it is an offer document, else
'**                 false.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildLifeTimeSection9(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection9"
    Dim xmlRollUp As IXMLDOMNode
    Dim xmlComp As IXMLDOMNode
    Dim xmlNoPayments As IXMLDOMNode
    Dim xmlLCList As IXMLDOMNodeList
    Dim xmlLC As IXMLDOMNode
    Dim xmlAddBorrow As IXMLDOMNode
    Dim dblLastClosingBalance As Double
    Dim xmlRates As IXMLDOMNodeList
    Dim xmlRate As IXMLDOMNode
    Dim xmlRatePeriod As IXMLDOMNode
    Dim intPart As Integer
    
    On Error GoTo ErrHandler
             
    '*-get the list of loan components
    Set xmlLCList = vobjCommon.LoanComponents
    
    If vobjCommon.IsAdditionalBorrowing Then
        '*-create the ADDITIONALBORROWING element
        Set xmlAddBorrow = vobjCommon.CreateNewElement("ADDITIONALBORROWING", vxmlNode)
        '*-create the TOTALBORROWINGAMOUNT attribute
        Call xmlSetAttributeValue(xmlAddBorrow, "TOTALBORROWINGAMOUNT", _
                vobjCommon.LoanAmountRepayable)
        If vblnIsOfferDocument Then
            '*-create the PROVIDER attribute - offer document only
            Call xmlSetAttributeValue(xmlAddBorrow, "PROVIDER", vobjCommon.Provider)
        End If
        
        '*-create the ROLLUP element
        Set xmlRollUp = vobjCommon.CreateNewElement("ROLLUP", xmlAddBorrow)
        '*-create the TOTALPROJECTEDBORROWINGAMOUNT attribute
        dblLastClosingBalance = GetLastBalScheduleClosingBalance(vobjCommon)
        Call xmlSetAttributeValue(xmlRollUp, "TOTALPROJECTEDBORROWINGAMOUNT", dblLastClosingBalance)
        '*-create the TERMYEARS attribute
        Call xmlSetAttributeValue(xmlRollUp, "TERMYEARS", vobjCommon.TermInYears)
        
        intPart = 0
        For Each xmlLC In xmlLCList
            intPart = intPart + 1
            '*-get the list of interest rates
            Set xmlRates = vobjCommon.GetLoanComponentInterestRates(xmlLC)
            For Each xmlRate In xmlRates
                '*-create the RATEPERIOD element
                Set xmlRatePeriod = vobjCommon.CreateNewElement("RATEPERIOD", xmlAddBorrow)
                '*-create the ROLLEDUP element
                Set xmlRollUp = vobjCommon.CreateNewElement("ROLLEDUP", xmlRatePeriod)
                '*-create the PART attribute
                Call xmlSetAttributeValue(xmlRollUp, "PART", CStr(intPart))
                '*-create the RATETYPE attribute
                Call xmlCopyAttribute(xmlRate, xmlRollUp, "RATETYPE")
                '*-create the PERIOD attribute
                Call AddPeriodAttribute(xmlRate, "INTERESTRATEPERIOD", _
                        "INTERESTRATEENDDATE", vobjCommon.ExpectedCompletionDate, _
                        "PERIOD", xmlRollUp, "") 'BC MAR1721
            Next xmlRate
        Next xmlLC
    End If
    
    Set xmlRollUp = vobjCommon.CreateNewElement("ROLLUP", vxmlNode)
    
    Set xmlLCList = vobjCommon.LoanComponents
    For Each xmlLC In xmlLCList
        'Set xmlComp = vobjCommon.CreateNewElement("COMPONENT", xmlRollUp) 'SR 12/09/2004 : CORE82
        Set xmlNoPayments = vobjCommon.CreateNewElement("NOPAYMENTS", xmlRollUp) 'SR 12/09/2004 : CORE82
        
        Call AddProductDescription(vobjCommon, xmlLC, xmlNoPayments)
    Next xmlLC

    Set xmlRollUp = Nothing
    Set xmlComp = Nothing
    Set xmlNoPayments = Nothing
    Set xmlLCList = Nothing
    Set xmlLC = Nothing
    Set xmlAddBorrow = Nothing
    Set xmlRates = Nothing
    Set xmlRate = Nothing
    Set xmlRatePeriod = Nothing
Exit Sub
ErrHandler:
    Set xmlRollUp = Nothing
    Set xmlComp = Nothing
    Set xmlNoPayments = Nothing
    Set xmlLCList = Nothing
    Set xmlLC = Nothing
    Set xmlAddBorrow = Nothing
    Set xmlRates = Nothing
    Set xmlRate = Nothing
    Set xmlRatePeriod = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddProductDescription
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Adds the product description to the no payments element.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlLC - the loan component.
'**                 vxmlNoPayments - the no payments element to add the product
'**                 description to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub AddProductDescription(ByVal vobjCommon As CommonDataHelper, _
                                  ByVal vxmlLC As IXMLDOMNode, _
                                  ByVal vxmlNoPayments As IXMLDOMNode)
    Const cstrFunctionName As String = "AddProductDescription"
    Dim xmlMortProdLang As IXMLDOMNode
    Dim xmlIfIncrease As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlIntRateList As IXMLDOMNodeList
    Dim xmlIntRate As IXMLDOMNode
    Dim xmlRateDesc As IXMLDOMNode
    Dim eRateType As MortgageInterestRateType
    Dim strProductName As String
    Dim strRateDesc As String
    Dim blnFixed As Boolean
    Dim blnVariable As Boolean
    Dim blnDiscounted As Boolean
    Dim blnCapped As Boolean
    Dim blnCollared As Boolean
    Dim blnTrackerabove As Boolean
    Dim dblCappedRate As Double
    Dim dblCollaredRate As Double
    
    On Error GoTo ErrHandler
    
    Set xmlMortProdLang = vxmlLC.selectSingleNode(".//MORTGAGEPRODUCT/MORTGAGEPRODUCTLANGUAGE")
    If Not xmlMortProdLang Is Nothing Then
        strProductName = xmlGetAttributeText(xmlMortProdLang, "PRODUCTNAME")
    End If
    
    Set xmlIntRateList = vxmlLC.selectNodes(".//MORTGAGEPRODUCT/INTERESTRATETYPE")
    For Each xmlIntRate In xmlIntRateList
        eRateType = GetInterestRateType(xmlIntRate)
        Select Case eRateType
            Case mrtStandardVariableRate
               blnVariable = True
            Case mrtDiscountedRate
                blnDiscounted = True
            Case mrtTrackerAbove
                blnTrackerabove = True
            Case mrtFixedRate
                blnFixed = True
            Case mrtCappedRate
                blnCapped = True
                dblCappedRate = xmlGetAttributeAsDouble(xmlIntRate, "CEILINGRATE")
            Case mrtCollaredRate
                blnCollared = True
                dblCollaredRate = xmlGetAttributeAsDouble(xmlIntRate, "FLOOREDRATE")
            Case mrtCappedAndCollaredRate
                blnCapped = True
                blnCollared = True
                dblCappedRate = xmlGetAttributeAsDouble(xmlIntRate, "CEILINGRATE")
                dblCollaredRate = xmlGetAttributeAsDouble(xmlIntRate, "FLOOREDRATE")
        End Select
    Next xmlIntRate
    
    If blnFixed And Not (blnVariable Or blnDiscounted) Then
        strRateDesc = "FIXEDFORWHOLEOFTERM"
        Set xmlRateDesc = vobjCommon.CreateNewElement(strRateDesc, vxmlNoPayments)
    ElseIf blnFixed And (blnVariable Or blnDiscounted) Then
        strRateDesc = "FIXEDFORPARTOFTERM"
        Set xmlRateDesc = vobjCommon.CreateNewElement(strRateDesc, vxmlNoPayments)
    ElseIf blnCapped And Not (blnVariable Or blnDiscounted Or blnFixed Or blnTrackerabove Or blnCollared) Then
        strRateDesc = "CAPPEDFORWHOLEOFTERM"
        Set xmlRateDesc = vobjCommon.CreateNewElement(strRateDesc, vxmlNoPayments)
        xmlSetAttributeValue xmlRateDesc, "CAPRATE", dblCappedRate
    ElseIf blnCollared And Not (blnVariable Or blnDiscounted Or blnFixed Or blnTrackerabove Or blnCapped) Then
        strRateDesc = "COLLAREDFORWHOLEOFTERM"
        Set xmlRateDesc = vobjCommon.CreateNewElement(strRateDesc, vxmlNoPayments)
        xmlSetAttributeValue xmlRateDesc, "FLOORRATE", dblCollaredRate
    ElseIf blnCollared And blnCapped And Not (blnVariable Or blnDiscounted Or blnFixed Or blnTrackerabove) Then
        strRateDesc = "CAPPEDANDCOLLAREDFORWHOLEOFTERM"
        Set xmlRateDesc = vobjCommon.CreateNewElement(strRateDesc, vxmlNoPayments)
        xmlSetAttributeValue xmlRateDesc, "CAPRATE", dblCappedRate
        xmlSetAttributeValue xmlRateDesc, "FLOORRATE", dblCollaredRate
    ElseIf blnCapped And (blnVariable Or blnDiscounted Or blnFixed Or blnTrackerabove Or blnCapped) Then
        strRateDesc = "CAPPEDFORPARTOFTERM"
        Set xmlRateDesc = vobjCommon.CreateNewElement(strRateDesc, vxmlNoPayments)
        xmlSetAttributeValue xmlRateDesc, "CAPRATE", dblCappedRate
    End If
    
    ' Create IFINCREASE Node
    If blnFixed And Not (blnVariable Or blnDiscounted) Then
        'N/A
    Else
        If blnVariable Or blnDiscounted Or blnFixed Or blnCapped Then
            Set xmlIfIncrease = vobjCommon.CreateNewElement("IFINCREASE", vxmlNoPayments)
            If blnFixed Then
                Set xmlNode = vobjCommon.CreateNewElement("FIXEDBECOMESVARIABLE", xmlIfIncrease)
            End If
        
            If blnCapped Then
                Set xmlNode = vobjCommon.CreateNewElement("CAPPEDBECOMESVARIABLE", xmlIfIncrease)
            End If
        End If
    End If
    
    ' Create NEWRATE Node
    'SR 12/09/2004 : CORE82
    If blnFixed And Not (blnVariable Or blnDiscounted) Then
        'N/A
    Else
        Set xmlIntRate = xmlIntRateList.Item(0)
        Set xmlNode = vobjCommon.CreateNewElement("NEWRATE", vxmlNoPayments)
        xmlSetAttributeValue xmlNode, "NEWINTERESTRATE", vobjCommon.StandardVariableRate + 1
        xmlSetAttributeValue xmlNode, "TERMYEARS", vobjCommon.TermInYears
        xmlSetAttributeValue xmlNode, "AMOUNTOWING", xmlGetAttributeAsDouble(vxmlLC, "FINALBALANCEPLUSONEPERCENT")
    End If
    'SR 12/09/2004 : CORE82 - End
    Set xmlMortProdLang = Nothing
    Set xmlIfIncrease = Nothing
    Set xmlNode = Nothing
    Set xmlIntRateList = Nothing
    Set xmlIntRate = Nothing
    Set xmlRateDesc = Nothing
Exit Sub
ErrHandler:
    Set xmlMortProdLang = Nothing
    Set xmlIfIncrease = Nothing
    Set xmlNode = Nothing
    Set xmlIntRateList = Nothing
    Set xmlIntRate = Nothing
    Set xmlRateDesc = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetLastBalScheduleClosingBalance
'** Created by:     Andy Maggs
'** Date:           14/06/2004
'** Description:    Gets the closing balance from the last balance schedule item.
'** Parameters:     vobjCommon - the common data helper.
'** Returns:        The closing balance value.
'** Errors:         None Expected
'********************************************************************************
Private Function GetLastBalScheduleClosingBalance(ByVal vobjCommon As CommonDataHelper) As Double
    Const cstrFunctionName As String = "GetLastBalScheduleClosingBalance"
    Dim xmlLC As IXMLDOMNode
    Dim xmlLCBalanceScheduleList As IXMLDOMNodeList
    Dim xmlLCBalanceSchedule As IXMLDOMNode
    Dim dblOpeningBalance As Double
    Dim dblIntOnlyAmount As Double

    On Error GoTo ErrHandler

    'TODO must deal with multiple loan components
    Set xmlLC = vobjCommon.SingleLoanComponent
    Set xmlLCBalanceScheduleList = xmlLC.selectNodes(".//LOANCOMPONENTBALANCESCHEDULE")
    
    Set xmlLCBalanceSchedule = xmlLCBalanceScheduleList.Item(xmlLCBalanceScheduleList.length - 1)
    dblOpeningBalance = xmlGetAttributeAsDouble(xmlLCBalanceSchedule, "BALANCE")
    dblIntOnlyAmount = xmlGetAttributeAsDouble(xmlLCBalanceSchedule, "INTONLYBALANCE")
    
    'TODO is this logic correct? - spec needs updating if it is
    GetLastBalScheduleClosingBalance = dblOpeningBalance + dblIntOnlyAmount

    Set xmlLC = Nothing
    Set xmlLCBalanceScheduleList = Nothing
    Set xmlLCBalanceSchedule = Nothing
Exit Function
ErrHandler:
    Set xmlLC = Nothing
    Set xmlLCBalanceScheduleList = Nothing
    Set xmlLCBalanceSchedule = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildLifeTimeSection13
'** Created by:     Andy Maggs
'** Date:           16/06/2004
'** Description:    Builds the common lifetime section 13 (What happens if you
'**                 do not want this mortgage any more?).
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to build.
'**                 vblnIsOfferDocument - whether this is an offer document or
'**                 not.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildLifeTimeSection13(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildLifeTimeSection13"
    Dim xmlCharges As IXMLDOMNodeList
    Dim xmlCharge As IXMLDOMNode
    Dim xmlAmount As IXMLDOMAttribute
    Dim xmlInitRelease As IXMLDOMNode
    Dim xmlSummary As IXMLDOMNode
    Dim xmlRollUp As IXMLDOMNode
    Dim xmlEarlyRepayment As IXMLDOMNode  'SR 11/10/2004 : BBG1602
    
    On Error GoTo ErrHandler

    Call BuildMortgageNoMoreSection(vobjCommon, vxmlNode, vblnIsOfferDocument, True)
    
    '*-post-process to add the additional elements/attributes required for
    '*-lifetime mortgages
    '*-get the list of repayment charges
    Set xmlCharges = vxmlNode.selectNodes("//REPAYMENTCHARGES")
    For Each xmlCharge In xmlCharges
        '*-get the loan amount attribute
        Set xmlAmount = xmlCharge.Attributes.getNamedItem("LOANAMOUNT")
        '*-and remove it
        Call xmlCharge.Attributes.removeNamedItem("LOANAMOUNT")
        
        '*-create the INITIALRELEASEOFLUMPSUM element
        Set xmlInitRelease = vobjCommon.CreateNewElement("INITIALRELEASEOFLUMPSUM", xmlCharge)
        '*-add the loan amount attribute to it
        Call xmlInitRelease.Attributes.setNamedItem(xmlAmount)
    Next xmlCharge
    
    '*-get the SUMMARY element
    Set xmlSummary = vxmlNode.selectSingleNode("//SUMMARY")
    If Not xmlSummary Is Nothing Then
        '*-add the INCLUDESROLLUP element
        'Set xmlRollUp = vobjCommon.CreateNewElement("INCLUDESROLLUP", xmlSummary)  'SR 11/10/2004 : BBG1602
        '*-add the TERMYEARS attribute
        'Call xmlSetAttributeValue(xmlRollUp, "TERMYEARS", CStr(vobjCommon.TermInYears)) 'SR 11/10/2004 : BBG1602
        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, xmlRollUp)
        End If
    End If
    
    Set xmlCharges = Nothing
    Set xmlCharge = Nothing
    Set xmlAmount = Nothing
    Set xmlInitRelease = Nothing
    Set xmlSummary = Nothing
    Set xmlRollUp = Nothing
    Set xmlEarlyRepayment = Nothing
Exit Sub
ErrHandler:
    Set xmlCharges = Nothing
    Set xmlCharge = Nothing
    Set xmlAmount = Nothing
    Set xmlInitRelease = Nothing
    Set xmlSummary = Nothing
    Set xmlRollUp = Nothing
    Set xmlEarlyRepayment = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'SR 12/09/2004 : CORE82
Public Sub SetSVR(ByVal vxmlNode As IXMLDOMNode, vobjCommon As CommonDataHelper)
    Dim xmlLoanComponent As IXMLDOMNode, xmlInterestRate As IXMLDOMNode
    Dim eType As MortgageInterestRateType
    Dim dblRate As Double, dblTempRate As Double, dblBaseRate As Double
    
    Set xmlLoanComponent = vobjCommon.Data.selectSingleNode(gcstrLOANCOMPONENT_PATH & "[@LOANCOMPONENTSEQUENCENUMBER=1]")
    Set xmlInterestRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlLoanComponent)
    eType = GetInterestRateType(xmlInterestRate)
    dblTempRate = xmlGetAttributeAsDouble(xmlInterestRate, "RATE")
    
    If eType = mrtStandardVariableRate Then
        dblRate = xmlGetAttributeAsDouble(xmlLoanComponent, "RESOLVEDRATE")
    Else
        If eType = mrtFixedRate Then
            dblRate = dblTempRate
        Else
            dblBaseRate = GetBaseInterestRate(vobjCommon, xmlInterestRate, "")
            dblRate = dblBaseRate - dblTempRate
        End If
    End If
    
    xmlSetAttributeValue vxmlNode, "STANDARDVARIABLERATE", dblRate
    vobjCommon.StandardVariableRate = dblRate
    
    Set xmlLoanComponent = Nothing
    Set xmlInterestRate = Nothing
End Sub
'SR 12/09/2004 : CORE82 - End
