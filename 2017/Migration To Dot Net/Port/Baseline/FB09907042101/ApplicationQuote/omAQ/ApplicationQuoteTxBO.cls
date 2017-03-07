VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ApplicationQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationQuoteTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Transaction object for Application Quote
'
'Dependencies:  List any other dependent components
'               ApplicationQuoteBO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    22/03/00    Created
'JLD    15/05/00    SYS0715 Add DEPOSIT to ResetMortgageSubQuote()
'APS    02/06/00    Creating an MTS object as New is not allowed
'APS    04/06/00    Use CUSTOMERA and CUSTOMERB highest earning total net monthly income
'                   amounts, add them together to form the total monthly
'                   income on the quotation
'PSC    30/06/00    SYS1012 Amend the way outstanding payments is calculated and change LTV
'                   to be a double
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'SR     17/10/00    SYS0883
'BG     05/12/00    SYS0883 changed CalculateOutgoingDetails to select LIFEPOLICY node as
'                   opposed to LIFEPOLICIES
'PF     21/03/01    SYS1679 changed CalculateOutgoingDetails to cater for no RegularOutgoings
'PSC    19/03/01    SYS2053 Add RemodelQuotation
'PSC    03/04/01    SYS2236 Amend RemodelQuotation to cater for single applicant cases
'MC     20/08/01    SYS2250 Add Rate Change batch functionality
'DC     07/12/01    SYS0977
'MC     10/12/01    SYS3440. Fixes to Rate Change process
'STB    18/04/02    SYS3948 Check calcs response in RemodelQuotation.
'STB    30/04/02    SYS4506 Overflow error, Amount Request is a double.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IApplicationQuoteTxBO
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================

Private Const cstrAPPLICATION_NAME = "APPLICATION"
Private Const cstrAPPLICATIONQUOTE_NAME = "APPLICATIONQUOTE"    'SYS2250

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXMLAssist As XMLAssist             ' Remove if not required
Private m_objConvertAssist As ConvertAssist     ' BMIDS986

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'PSC    11/06/2002  BMIDS00043  Get Completion Date to be used in RemodelQuotation and
'                               RecostQuotationForRateChange
'PSC    17/06/2002  BMIDS00043  Inherit RecostQuotationForRateChange from core as change has
'                               been made in CostModelAssistBO
'PSC    17/06/2002  BMIDS00043  Late bind omRotBO due to build problems
'PSC    11/07/2002  BMIDS00062  Amend RemodelQuotation to use composite calcs
'                               Amend RecostQuotationforrateChange to use composite calcs
'DPF    31/07/2002  BMIDS00266  Have removed call to save routine for life cover
'MV     12/08/2002  BMIDS00311  Core Ref : SYS1705 Removed SafeRef ;
'                               Modified IApplicationQuoteTxBO_CalculateMortgageCosts
'MDC    13/08/2002  BMIDS00244  Replace xmlTempRequest with xmlTempRequest.xml for call to
'                               RunCompositeRateCalcAndSaveResults string interface
'GHun   30/08/2002  BMIDS00390  Temporary solution to CalculateAffordability error
'MV     16/09/2002  BMIDS00432  Removed LifeSubQuote Functionality fromn IApplicationQuoteTxBO_RemodelQuotation at Stage18
'MDC    17/09/2002  BMIDS00048  RCWP1 BM026 Rate Change
'DPF    11/11/2002  BMIDS00904  Have removed call to Life Cover calculations witin IApplicationQuoteTXBO_ReCostMortgageComponentsAndLifeCover
'MO     22/11/2002  BMIDS01061  Modified remodel quotation to bring it inline with changes to cost modelling for the BMIDS project
'DPF    25/11/2002  BMIDS01066  Amended RemodelQuotation to re-calculate drawdown amounts
'MDC    18/12/2002  BM0197      Prevent One Off Costs being added to loan twice when Remodelling a quote
'MDC    31/12/2002  BM0211      Rate Change fixes
'MDC    09/01/2003  BM0230      Failed Batch Runs
'GHun   02/07/2003  BMIDS591    CC040 Amended RemodelQuotation
'GHun   09/09/2003  BMIDS622    Restructure RateChange
'GHun   05/11/2003  BMIDS622    Amended RecostQuotationForRateChange to comment out some dead variables
'GHun   22/04/2004  BMIDS736    CORE49 Alpha+ integration
'DRC    20/05/2004  BMIDS767    Added UpdateQuotationForAddedCosts
'DRC    25/06/2004  BMIDS763    Altered RemodelQuotation,UpdateQuotationForAddedCosts & RecostQuotationForRateChange
'                               to include ApplicationDate in call to CalcCharges
'INR    05/07/2004  BMIDS774    Use ApplicationFactFind EstimatedCompletion Date instead of ROT Completion date.
'GHun   09/07/2004  BMIDS767    Amended RemodelQuotation to handle multiple OneOffCost validation types
'GHun   14/07/2004  BMIDS793    Amended RemodelQuotation and UpdateQuotationForAddedCosts to ignore drawdown = 0, and updated RecostQuotationForRateChange
'GHun   16/07/2004  BMIDS767    Amended UpdateQuotationForAddedCosts to save TotalLoanAmount
'GHun   04/08/2004  BMIDS834    Amended RecostQuotationForRateChange to add fees to loan
'GHun   23/08/2004  BMIDS856    Minor cleanup & amended RecostQuotationForRateChange
'JD     24/08/2004  BMIDS854    Change RemodelQuotation and recostQuotationForRateChange to include APR adhoc oneoffcosts
'SR     26/08/2004  BMIDS815    New methods 'RefreshMortgageAccountData',
'                               'GetAndSavePortedStepAndPeriodFromMortgageAccount'
'SR     26/08/2004  BMIDS815    Modified method 'Remodel Quotation'
'GHun   31/08/2004  BMIDS854    Changed RemodelQuotation
'SR     02/09/2004  BMIDS815    Modified IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount
'GHun   07/09/2004  BMIDS815    Changed GetAndSavePortedStepAndPeriodFromMortgageAccount & RemodelQuotation
'GHun   08/09/2004  BMIDS815    Changed RemodelQuotation
'GHun   09/09/2004  BMIDS815    Changed RemodelQuotation
'JD     14/09/2004  BMIDS866    in UpdateQuotationForAddedCosts, TotalLoanAmount is used in CalcMig so should contain the AmountRequested amount
'GHun   15/09/2004  BMIDS866    Changed RecostQuotationForRateChange
'GHun   15/09/2004  BMIDS854    Changed RemodelQuotation & UpdateQuotationForAddedCosts
'GHun   24/09/2004  BMIDS891    Changed RemodelQuotation & RecostQuotationForRateChange
'HMA    29/09/2004  BMIDS881    Changed RemodelQuotation and UpdateQuotationForAddedCosts to calculate costs correctly.
'HMA    08/03/2005  BMIDS986    Changed RemodelQuotation, RecostQuotationForRateChange for MIG calculation.
'HMA    15/03/2005  BMIDS986    Changed UpdateQuotationForAddedCosts for MIG calculation.
'JD     23/03/2005  BMIDS975    RefreshMortgageAccountData - added refreshing of term from ICBS
'JD     30/03/2005  BMIDS975    Changed warning message 7028 to show value of global parameter.
' JD    04/04/2005 BMIDS977     Added CreateAdHocFeeAndRemodel
'HMA    06/04/2005  BMIDS977    Changes to RemodelQuotation
'                               Added ProcessAllFeesForRemodelQuote
'HMA    11/04/2005  BMIDS977    Changes to RecostQuotationForRateChange
'HMA    18/04/2005  BMIDS977    Further changes to RemodelQuotation and ProcessAllFeesForRemodelQuote
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        AQR         Description
'GHun   15/08/2005  MAR43       Changed GetAndSavePortedStepAndPeriodFromMortgageAccount
'HMA    24/08/2005  MAR28       Changed CreateAdHocFeeAndRemodel to add RefundAmount to OneOffCost.
'GHun   06/10/2005  MAR46       Changed GetAndSavePortedStepAndPeriodFromMortgageAccount for porting
'                               Changed RefreshMortgageAccountData (WP3)
'JD     22/03/2006  MAR1061     ResetMortgageSubquote - pass purchase price to omCM
'GHun   10/07/2006  MAR1893     Commented out redundant LifeSubQuote code
'------------------------------------------------------------------------------------------
'Epsom Specific History
'
'Prog   Date        AQR         Description
'PSC    09/11/2006  EP2_41      Amend CalculateMortgageCosts to return costs
'AShaw  29/12/2006  EP2_56      New Item added to IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount
'GHun   26/01/2006  EP2_974     Merged EP1241 Duplicate row error on one off cost during remodelquotation
'PSC    12/03/2007  EP2_1611    Amend GetAndSavePortedStepAndPeriodFromMortgageAccount to match loancomponent and
'                               mortgagloan on MORTGAGELOANGUID
'PSC    15/03/2007  EP2_1956    Amend RemodelQuotation to copy Mortgage Introducer Fees
'------------------------------------------------------------------------------------------


Private Function IApplicationQuoteTxBO_CalculateAffordability( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAffordabilityVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_CalculateAffordability"

    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlHighestEarners As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode, xmlNewRequest As IXMLDOMNode, xmlTableNode As IXMLDOMNode
   
    Dim strResponse As String
    
    Dim objIAppQuoteBO As IApplicationQuoteBO
    'Dim objLifeSubQuoteBO As LifeSubQuoteBO
    Dim objQuotationBO As QuotationBO
    Dim objConvertAssist As ConvertAssist
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIAppQuoteBO = New ApplicationQuoteBO
        'Set objLifeSubQuoteBO = New LifeSubQuoteBO
        Set objQuotationBO = New QuotationBO
    Else
        Set objIAppQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")
        'Set objLifeSubQuoteBO = _
        '    m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".LifeSubQuoteBO")
        Set objQuotationBO = _
            m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    End If
    
    '-------------------------------------------------------------------------
    'Get the highest earners
    '-------------------------------------------------------------------------
    Set xmlHighestEarners = objIAppQuoteBO.GetHighestEarners(vxmlRequest)
    m_objErrAssist.CheckResponse xmlHighestEarners, xmlResponseElem, True
    
    Dim dblTotalRegOutgoings As Double
    Dim dblTotalBankCreditCards As Double, dblTotalLoansLiabilities As Double
    Dim dblTotalLifeProducts As Double, dblTotalArrearsHistory As Double
    Dim dblTotalBankruptcyHistory As Double, dblTotalCCJHistory As Double
    Dim dblMonthlyMortgagePayments As Double, dblMortgageRelatedInsurance As Double
    
    ' For each of the highest earners, get the outgoing details and add them
    Call CalculateOutGoingDetails(vxmlRequest, xmlHighestEarners, _
                                  dblTotalRegOutgoings, dblTotalBankCreditCards, _
                                  dblTotalLoansLiabilities, dblTotalLifeProducts, _
                                  dblTotalArrearsHistory, dblTotalBankruptcyHistory, _
                                  dblTotalCCJHistory)

    '-------------------------------------------------------------------------------------
    'Calculate the Monthly Mortgage Payment and Mortgage related insurance
    '-------------------------------------------------------------------------------------
    Call CalculateMonthlyPayments(vxmlRequest, xmlHighestEarners, _
                                  dblMonthlyMortgagePayments, dblMortgageRelatedInsurance)

    '-----------------------------------------------------------------------
    'Update Quotation - Build the Request before calling Update method
    '-----------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableNode = xmlNewRequest.appendChild(xmlOut.createElement("QUOTATION"))
    xmlNewRequest.appendChild xmlTableNode
        
    xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    xmlTableNode.appendChild vxmlRequest.selectSingleNode(".//QUOTATIONNUMBER").cloneNode(True)
        
    Set xmlNode = xmlOut.createElement("MORTGAGERELATEDINSURANCE")
    xmlNode.Text = dblMortgageRelatedInsurance
    xmlTableNode.appendChild xmlNode
    xmlResponseElem.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlOut.createElement("OTHEROUTGOINGS")
    xmlNode.Text = dblTotalRegOutgoings + dblTotalBankCreditCards + dblTotalLifeProducts + _
                   dblTotalArrearsHistory + dblTotalBankruptcyHistory + dblTotalCCJHistory
    xmlTableNode.appendChild xmlNode
    xmlResponseElem.appendChild xmlNode.cloneNode(True)
    
    ' APS 04/06/00 - Use CUSTOMERA and CUSTOMERB highest earning total net monthly income
    ' amounts, add them together to form the total monthly income on the quotation
    Dim xmlIncomeList As IXMLDOMNodeList
    Dim xmlIncome As IXMLDOMNode
    Dim lngTotalIncome As Long
    Set objConvertAssist = New ConvertAssist
    
    Set xmlIncomeList = xmlHighestEarners.selectNodes(".//INCOMESUMMARY/TOTALNETMONTHLYINCOME")
    For Each xmlIncome In xmlIncomeList
        lngTotalIncome = lngTotalIncome + objConvertAssist.CSafeLng(xmlIncome.Text)
    Next
    Set xmlNode = xmlOut.createElement("TOTALMONTHLYINCOME")
    xmlNode.Text = CStr(lngTotalIncome)
    xmlTableNode.appendChild xmlNode
    xmlResponseElem.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlOut.createElement("LOANSANDLIABILITIES")
    xmlNode.Text = dblTotalLoansLiabilities
    xmlTableNode.appendChild xmlNode
    xmlResponseElem.appendChild xmlNode.cloneNode(True)
    
    Set xmlNode = xmlOut.createElement("MONTHLYMORTGAGEPAYMENTS")
    xmlNode.Text = dblMonthlyMortgagePayments
    xmlTableNode.appendChild xmlNode
    xmlResponseElem.appendChild xmlNode.cloneNode(True)
    
    strResponse = objQuotationBO.Update(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
       
    Set IApplicationQuoteTxBO_CalculateAffordability = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateAffordabilityExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlHighestEarners = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableNode = Nothing
    Set objIAppQuoteBO = Nothing
    'Set objLifeSubQuoteBO = Nothing
    Set objQuotationBO = Nothing
    Set objConvertAssist = Nothing
        
    Exit Function
CalculateAffordabilityVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_CalculateAffordability = _
                            m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CalculateAffordabilityExit
End Function

Private Sub CalculateOutGoingDetails(ByVal vxmlRequest As IXMLDOMNode, _
                                     ByVal xmlHighestEarners As IXMLDOMNode, _
                                     ByRef dblTotalRegOutgoings As Double, _
                                     ByRef dblTotalBankCreditCards As Double, _
                                     ByRef dblTotalLoansLiabilities As Double, _
                                     ByRef dblTotalLifeProducts As Double, _
                                     ByRef dblTotalArrearsHistory As Double, _
                                     ByRef dblTotalBankruptcyHistory As Double, _
                                     ByRef dblTotalCCJHistory As Double)
'-------------------------------------------------------------------------------------------------
'Description :
'Called from : CalculateAffordability
'-------------------------------------------------------------------------------------------------
On Error GoTo CalculateVbErr
    
    Const strFunctionName As String = "CalculateOutGoingDetails"
        
    Dim objConvertAssist As ConvertAssist
    Dim objCFBO As CustomerFinancialBO
    Dim objICFDO As ICustomerFinancialDO
    
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlTableNode As IXMLDOMNode, xmlNewRequest As IXMLDOMNode
    Dim xmlNodeCustomerA As IXMLDOMNode, xmlNodeCustomerB As IXMLDOMNode, xmlNode2 As IXMLDOMNode
    Dim xmlListNode As IXMLDOMNode
    
    Dim strResponse As String
    
    If m_objContext Is Nothing Then
        Set objCFBO = New CustomerFinancialBO
    Else
        Set objCFBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    End If
    '-----------------------------------------------------------------------------------
    ' Find Total of Monthly RegularOutGoings (for CustomerA and CustomerB)
    '----------------------------------------------------------------------------------
    Set xmlNodeCustomerA = xmlHighestEarners.selectSingleNode(".//EMPLOYMENTANDINCOME/" & "CUSTOMERA")
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlListNode = xmlNewRequest.appendChild(xmlOut.createElement("REGULAROUTGOINGSLIST"))
    Set xmlTableNode = xmlListNode.appendChild(xmlOut.createElement("REGULAROUTGOINGS"))
    xmlTableNode.appendChild xmlNodeCustomerA.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
    xmlTableNode.appendChild xmlNodeCustomerA.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
        
    ' If CustomerB exists add the respective details to Request
    Set xmlNodeCustomerB = xmlHighestEarners.selectSingleNode(".//EMPLOYMENTANDINCOME/" & "CUSTOMERB")
    If Not xmlNodeCustomerB Is Nothing Then
        Set xmlTableNode = xmlListNode.appendChild(xmlOut.createElement("REGULAROUTGOINGS"))
        xmlTableNode.appendChild xmlNodeCustomerB.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
        xmlTableNode.appendChild xmlNodeCustomerB.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
    End If
    
    strResponse = objCFBO.FindRegularOutgoingsList(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
    
    Dim iPayFrequency As Integer, strAmount As String
    Dim strFrequency As String
    
    Set objConvertAssist = New ConvertAssist
    
    For Each xmlNode2 In xmlDoc.documentElement.selectNodes(".//REGULAROUTGOINGS")
        ' PF 21/03/01   SYS1679 - Start
        strFrequency = m_objXMLAssist.GetNodeValue(xmlNode2, ".//PAYMENTFREQUENCY", True, False)
        iPayFrequency = objConvertAssist.CSafeInt(strFrequency)
        ' PF 21/03/01   SYS1679 - End
        strAmount = xmlNode2.selectSingleNode(".//AMOUNT").Text
        dblTotalRegOutgoings = dblTotalRegOutgoings _
                               + (iPayFrequency * objConvertAssist.CSafeDbl(strAmount) / 12)
    Next xmlNode2
    
    '------------------------------------------------------------------------------------------
    ' Find the sum of Outgoings for both the customers
    '------------------------------------------------------------------------------------------
    Call m_objXMLAssist.ChangeNodeName(xmlNewRequest, "REGULAROUTGOINGSLIST", "CUSTOMERLIST")
    Call m_objXMLAssist.ChangeNodeName(xmlNewRequest, "REGULAROUTGOINGS", "CUSTOMER")
    
    'BMIDS00390 The 3 lines below are commented out temporarily until FindOutgoingsTotalList is fixed
    'strResponse = objCFBO.FindOutgoingsTotalList(xmlNewRequest.xml)
    'Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    'm_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
    'BMIDS00390 End
    
    ' Bank Credit Cards
    For Each xmlNode2 In xmlDoc.selectNodes(".//BANKCREDITCARD")
        dblTotalBankCreditCards = dblTotalBankCreditCards + _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlNode2, "AVERAGEMONTHLYREPAYMENT"))
    Next xmlNode2
    
    ' Total Loans & Liabilities
    For Each xmlNode2 In xmlDoc.selectNodes(".//LOANSLIABILITIES")
        dblTotalLoansLiabilities = dblTotalLoansLiabilities + _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlNode2, "MONTHLYREPAYMENT"))
    Next xmlNode2
    'BG SYS0883 changed to select LIFEPOLICY node as opposed to LIFEPOLICIES
    ' Total Life Policies
    For Each xmlNode2 In xmlDoc.selectNodes(".//LIFEPOLICY")
        dblTotalLifeProducts = dblTotalLifeProducts + _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlNode2, "MONTHLYPREMIUM"))
    Next xmlNode2
    
    ' Total Arrears History
    For Each xmlNode2 In xmlDoc.selectNodes(".//ARREARSHISTORY")
        dblTotalArrearsHistory = dblTotalArrearsHistory + _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlNode2, "MONTHLYREPAYMENT"))
    Next xmlNode2
    
    ' Total Bankruptcy History
    For Each xmlNode2 In xmlDoc.selectNodes(".//BANKRUPTCYHISTORY")
        dblTotalBankruptcyHistory = dblTotalBankruptcyHistory + _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlNode2, "MONTHLYREPAYMENT"))
    Next xmlNode2
    
    ' Total CCJHistory
    For Each xmlNode2 In xmlDoc.selectNodes(".//CCJHISTORY")
        dblTotalCCJHistory = dblTotalCCJHistory + _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlNode2, "MONTHLYREPAYMENT"))
    Next xmlNode2
        
    Set objCFBO = Nothing
    Set xmlOut = Nothing
    Set xmlTableNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlNodeCustomerA = Nothing
    Set xmlNodeCustomerB = Nothing
    Set xmlNode2 = Nothing
    Set objConvertAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

CalculateVbErr:
    Set objCFBO = Nothing
    Set objICFDO = Nothing
    Set xmlOut = Nothing
    Set xmlTableNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlNodeCustomerA = Nothing
    Set xmlNodeCustomerB = Nothing
    Set xmlNode2 = Nothing
    Set objConvertAssist = Nothing
        
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub CalculateMonthlyPayments(ByVal vxmlRequest As IXMLDOMNode, _
                                     ByVal xmlHighestEarners As IXMLDOMNode, _
                                     ByRef dblMonthlyMortgagePayment As Double, _
                                     ByRef dblMortgageRelatedInsurance As Double)
'-------------------------------------------------------------------------------------------
'Description :
'Called From : CalculateAffordability
'--------------------------------------------------------------------------------------------
On Error GoTo CalculateVbErr

    Const strFunctionName As String = "CalculateMonthlyPayments"

    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlListNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40, xmlOut As FreeThreadedDOMDocument40
    
    Dim objApplicationBO As ApplicationBO
    Dim objCFBO As CustomerFinancialBO
    'Dim objLifeSubQuoteBO As LifeSubQuoteBO
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    Dim objBCSubQuoteBO As BuildingsAndContentsSubQuoteBO
    
    Dim objConvertAssist As ConvertAssist
    
    Dim dblUnredeemedMonthlyPayment As Double
    Dim strResponse As String
    
    If m_objContext Is Nothing Then
        Set objCFBO = New CustomerFinancialBO
        Set objMortgageSubQuoteBO = New MortgageSubQuoteBO
        'Set objLifeSubQuoteBO = New LifeSubQuoteBO
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
        Set objBCSubQuoteBO = New BuildingsAndContentsSubQuoteBO
        Set objApplicationBO = New ApplicationBO
    Else
        Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
        Set objMortgageSubQuoteBO = _
                m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        'Set objLifeSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".LifeSubQuoteBO")
        Set objPPSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".PaymentProtectionSubQuoteBO")
        Set objBCSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".BuildingsAndContentsSubQuoteBO")
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    ' Find Monthly Mortgage Payments
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableNode = xmlNewRequest.appendChild(xmlOut.createElement("MORTGAGESUBQUOTE"))
    Set xmlNode = xmlOut.createElement("APPLICATIONNUMBER")
    xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    xmlTableNode.appendChild xmlNode

    Set xmlNode = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    xmlTableNode.appendChild xmlNode

    Set xmlNode = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
    xmlNode.Text = vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").Text
    xmlTableNode.appendChild xmlNode
    
    strResponse = objMortgageSubQuoteBO.GetData(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
    
    Set objConvertAssist = New ConvertAssist
    dblMonthlyMortgagePayment = objConvertAssist.CSafeDbl(m_objXMLAssist.GetElementText(xmlDoc.documentElement, ".//TOTALGROSSMONTHLYCOST"))
    
    If objApplicationBO.IsFurtherAdvance( _
                vxmlRequest.selectSingleNode(".//TYPEOFAPPLICATION").Text) Then
        ' Find MortgageAccounts details for Highest Earning Custmomers
        
        'PSC 30/06/00 SYS1012 - Start
        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlListNode = xmlNewRequest.appendChild(xmlOut.createElement("CUSTOMERLIST"))
        
        ' Add the details of CustomerA, if they exist
        Set xmlNode = xmlHighestEarners.selectSingleNode(".//EMPLOYMENTANDINCOME/" & "CUSTOMERA")
        If Not xmlNode Is Nothing Then
            Set xmlTableNode = xmlListNode.appendChild(xmlOut.createElement("CUSTOMER"))
            xmlTableNode.appendChild xmlNode.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
            xmlTableNode.appendChild xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
        End If
        
        ' Add the details of CustomerB, if they exist
        Set xmlNode = xmlHighestEarners.selectSingleNode(".//EMPLOYMENTANDINCOME/" & "CUSTOMERB")
        If Not xmlNode Is Nothing Then
            Set xmlTableNode = xmlListNode.appendChild(xmlOut.createElement("CUSTOMER"))
            xmlTableNode.appendChild xmlNode.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
            xmlTableNode.appendChild xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
        End If
                
        strResponse = objCFBO.GetLoansNotRedeemedTotals(xmlNewRequest.xml)
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
          
        dblUnredeemedMonthlyPayment = _
            objConvertAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "PAYMENTSONLOANSNOTREDEEMED"))
        'PSC 30/06/00 SYS1012 - End

        dblMonthlyMortgagePayment = dblMonthlyMortgagePayment + dblUnredeemedMonthlyPayment
    End If
    
    'MAR1893 GHun No longer used
    ''-----------------------------------------------------------------------------
    ''Get the cost of any new life insurance from the current life sub-quote
    ''-----------------------------------------------------------------------------
    'Dim strLifeSubQuoteNumber As String
    '
    'strLifeSubQuoteNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "LIFESUBQUOTENUMBER")
    'If Len(strLifeSubQuoteNumber) > 0 Then
    '    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    '    Set xmlTableNode = xmlDoc.createElement("LIFESUBQUOTE")
    '    xmlNewRequest.appendChild xmlTableNode
    '
    '    Set xmlNode = xmlOut.createElement("APPLICATIONNUMBER")
    '    xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    '    xmlTableNode.appendChild xmlNode
    '
    '    Set xmlNode = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    '    xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    '    xmlTableNode.appendChild xmlNode
    '
    '    Set xmlNode = xmlOut.createElement("LIFESUBQUOTENUMBER")
    '    xmlNode.Text = strLifeSubQuoteNumber
    '    xmlTableNode.appendChild xmlNode
    '
    '    strResponse = objLifeSubQuoteBO.GetData(xmlNewRequest.xml)
    '    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    '    m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
    '
    '    dblMortgageRelatedInsurance = dblMortgageRelatedInsurance + xmlDoc.selectSingleNode(".//TOTALLIFEMONTHLYCOST").Text
    'End If
    'MAR1893 End
    
    '-----------------------------------------------------------------------------
    'Get the cost of any payment protection insurance from the current PPSubQuote
    '-----------------------------------------------------------------------------
    Dim strPPSubQuoteNumber As String
    
    strPPSubQuoteNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "PPSUBQUOTENUMBER")
    If Len(strPPSubQuoteNumber) > 0 Then
        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableNode = xmlDoc.createElement("PAYMENTPROTECTIONSUBQUOTE")
        xmlNewRequest.appendChild xmlTableNode
        
        Set xmlNode = xmlOut.createElement("APPLICATIONNUMBER")
        xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
        xmlTableNode.appendChild xmlNode

        Set xmlNode = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
        xmlTableNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("PPSUBQUOTENUMBER")
        xmlNode.Text = strPPSubQuoteNumber
        xmlTableNode.appendChild xmlNode
        
        strResponse = objPPSubQuoteBO.GetSubQuoteDetails(xmlNewRequest.xml)
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
        
        dblMortgageRelatedInsurance = dblMortgageRelatedInsurance + xmlDoc.selectSingleNode(".//TOTALPPMONTHLYCOST").Text
    End If
    
    '-----------------------------------------------------------------------------------
    'Get the cost of any Buildings & contents insurance from the current B&C SubQuote
    '-----------------------------------------------------------------------------------
    Dim strBCSubQuoteNumber As String
    
    strBCSubQuoteNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "BCSUBQUOTENUMBER")
    If Len(strBCSubQuoteNumber) > 0 Then
        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableNode = xmlDoc.createElement("BCSUBQUOTEDETAILSNOTES")
        xmlNewRequest.appendChild xmlTableNode
        
        Set xmlNode = xmlOut.createElement("APPLICATIONNUMBER")
        xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
        xmlTableNode.appendChild xmlNode

        Set xmlNode = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
        xmlTableNode.appendChild xmlNode
        
        Set xmlNode = xmlOut.createElement("BCSUBQUOTENUMBER")
        xmlNode.Text = strBCSubQuoteNumber
        xmlTableNode.appendChild xmlNode
        
        strResponse = objBCSubQuoteBO.GetSubQuoteDetails(xmlNewRequest.xml)
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
        
        dblMortgageRelatedInsurance = dblMortgageRelatedInsurance + xmlDoc.selectSingleNode(".//TOTALBCMONTHLYCOST").Text
        
    End If
        
CalculateExit:
    Set xmlNewRequest = Nothing
    Set xmlTableNode = Nothing
    Set xmlNode = Nothing
    Set xmlListNode = Nothing
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set objApplicationBO = Nothing
    Set objCFBO = Nothing
    Set objMortgageSubQuoteBO = Nothing
    'Set objLifeSubQuoteBO = Nothing
    Set objPPSubQuoteBO = Nothing
    Set objBCSubQuoteBO = Nothing
    Set objConvertAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub
CalculateVbErr:
    Set xmlNewRequest = Nothing
    Set xmlTableNode = Nothing
    Set xmlNode = Nothing
    Set xmlListNode = Nothing
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set objApplicationBO = Nothing
    Set objCFBO = Nothing
    Set objMortgageSubQuoteBO = Nothing
    'Set objLifeSubQuoteBO = Nothing
    Set objPPSubQuoteBO = Nothing
    Set objBCSubQuoteBO = Nothing
    Set objConvertAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IApplicationQuoteTxBO_CalculateAndSaveBCSubQuote( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAndSaveBCSubQuoteVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_CalculateAndSaveBCSubQuote"

    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement, xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode, xmlTableElem As IXMLDOMNode
    Dim xmlTableListElem As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlNewProperty As IXMLDOMNode
    Dim objNewPropertyBO As NewPropertyBO
    Dim objApplicationBO As ApplicationBO
    Dim objCFBO As CustomerFinancialBO
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Dim strResponse As String   ', strDOB As String
    Dim blnBankruptcyAppl1 As Boolean, blnBankruptcyAppl2 As Boolean
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objNewPropertyBO = New NewPropertyBO
        Set objApplicationBO = New ApplicationBO
        Set objCFBO = New CustomerFinancialBO
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objNewPropertyBO = _
            m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyBO")
        Set objApplicationBO = _
            m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objCFBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
        Set objICostModelAssistBO = _
            m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If

    ' Fetch input parameters into local variables
    Dim strCustNumber1 As String, strCustNumber2 As String
    Dim strCustVersionNumber1 As String, strCustVersionNumber2 As String
    Dim strApplNumber As String, strAFFNumber As String
    
    strCustNumber1 = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustVersionNumber1 = vxmlRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    
    Set xmlNode = vxmlRequest.selectNodes(".//CUSTOMER").Item(1)
    If Not xmlNode Is Nothing Then
        strCustNumber2 = xmlNode.selectSingleNode(".//CUSTOMERNUMBER").Text
        strCustVersionNumber2 = xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    End If
    
    strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    
    '---------------------------------------------------------------------------
    ' Get the details of newPropertyData
    '---------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("NEWPROPERTY")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlTableElem.appendChild xmlElement
        
    strResponse = objNewPropertyBO.GetNewPropertyGeneral(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set xmlNewProperty = xmlDoc.documentElement.firstChild
    
    '------------------------------------------------------------------------------
    ' Get the B&C Declaration record
    '------------------------------------------------------------------------------
'    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlTableElem = xmlOut.createElement("APPLICATION")
'    xmlNewRequest.appendChild xmlTableElem
'
'    xmlTableElem.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
'    xmlTableElem.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
'
'    Set xmlElement = xmlOut.createElement("DECLARATIONTYPE")
'    xmlElement.Text = "1"
'    xmlTableElem.appendChild xmlElement
'
'    strResponse = objApplicationBO.FindDeclarationList(xmlNewRequest.xml)
'    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    
    '----------------------------------------------------------------------------
    'Retrieve bankruptcy details for applicant 1 and 2
    '----------------------------------------------------------------------------
    Dim lngErrNumber As Long, strPattern As String
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableListElem = xmlOut.createElement("BANKRUPTCYHISTORYLIST")
    xmlNewRequest.appendChild xmlTableListElem
    
    Set xmlTableElem = xmlOut.createElement("BANKRUPTCYHISTORY")
    xmlTableListElem.appendChild xmlTableElem
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
    xmlElement.Text = strCustNumber1
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = strCustVersionNumber1
    xmlTableElem.appendChild xmlElement
    
    If Len(strCustNumber2) > 0 Then
        Set xmlTableElem = xmlOut.createElement("BANKRUPTCYHISTORY")
        xmlTableListElem.appendChild xmlTableElem
        
        Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
        xmlElement.Text = strCustNumber2
        xmlTableElem.appendChild xmlElement
    
        Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlElement.Text = strCustVersionNumber2
        xmlTableElem.appendChild xmlElement
    End If
    
    strResponse = objCFBO.FindBankruptcyHistoryList(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    lngErrNumber = m_objErrAssist.CheckResponse(xmlDoc.documentElement, xmlResponseElem, False)
    If lngErrNumber <> 0 Then
        If m_objErrAssist.GetOmigaErrorNumber(lngErrNumber) = oeRecordNotFound Then
            blnBankruptcyAppl1 = False
            blnBankruptcyAppl2 = False
        Else
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem
        End If
    Else
        ' Set the bankruptcy flag for customer 1
        strPattern = ".//BANKRUPTCYHISTORY[CUSTOMERNUMBER = " & strCustNumber1 & "]"
        Set xmlNode = xmlDoc.documentElement.selectSingleNode(strPattern)
        
        If Not xmlNode Is Nothing Then
            blnBankruptcyAppl1 = True
        End If
        
        ' Set the bankruptcy flag for customer 2, if existing
        If Len(strCustVersionNumber2) > 0 Then
            strPattern = ".//BANKRUPTCYHISTORY[CUSTOMERNUMBER = " & strCustNumber2 & "]"
            Set xmlNode = xmlDoc.documentElement.selectSingleNode(strPattern)
            If Not xmlNode Is Nothing Then
                blnBankruptcyAppl2 = True
            End If
        End If
    End If

    Set xmlNode = objICostModelAssistBO.CalculateAndSaveBCSubQuote( _
                       vxmlRequest, strCustNumber1, strCustNumber2, strCustVersionNumber1, _
                        strApplNumber, strAFFNumber, "COSTMODELLLING", blnBankruptcyAppl1, _
                        blnBankruptcyAppl2, "Y", xmlNewProperty)
                        
    m_objErrAssist.CheckResponse xmlNode, xmlResponseElem, True
    
    Set IApplicationQuoteTxBO_CalculateAndSaveBCSubQuote = xmlNode
    
CalculateAndSaveBCSubQuoteExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set xmlTableListElem = Nothing
    Set xmlNode = Nothing
    Set xmlNewProperty = Nothing
    Set objNewPropertyBO = Nothing
    Set objApplicationBO = Nothing
    Set objCFBO = Nothing
    Set objICostModelAssistBO = Nothing

   Exit Function
CalculateAndSaveBCSubQuoteVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_CalculateAndSaveBCSubQuote = _
                            m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit
    
End Function

Private Function IApplicationQuoteTxBO_CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement, _
                                                              xmlLoanCompositionNode As IXMLDOMNode, _
                                                              xmlAppNo As IXMLDOMNode, _
                                                              xmlAppFactFindNo As IXMLDOMNode, _
                                                              xmlLifeSubQuoteNumber As IXMLDOMNode, _
                                                              xmlMortgageSubQuoteNumber As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_CalculateMortgageCosts"

    'Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlMortgageCostsResp As IXMLDOMNode
    'Dim xmlDOMDoc As FreeThreadedDOMDocument40
    
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objIAppQuoteBO As IApplicationQuoteBO
    
    Dim blnCalcFeesOnly As Boolean ' PSC 10/11/2006 EP2_41
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' APS 02/06/00 - Creating an MTS object as New is not allowed
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objIAppQuoteBO = New ApplicationQuoteBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".CostModelAssistBO")
        Set objIAppQuoteBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    End If

    ' Start transaction
    Set xmlMortgageCostsResp = objICostModelAssistBO.CalculateMortgageCosts(vxmlRequest, _
                                                                            xmlLoanCompositionNode, _
                                                                            xmlAppNo, _
                                                                            xmlAppFactFindNo, _
                                                                            xmlLifeSubQuoteNumber, _
                                                                            xmlMortgageSubQuoteNumber)
                                                                            
    m_objErrAssist.CheckResponse xmlMortgageCostsResp, xmlResponseElem, True
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'BMIDS00266 - DPF 31/07/02 - THIS BIT TAKEN OUT AS WE DON'T OFFER LIFE COVER
    ' Calculate and save Life Cover Costs
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Dim xmlLifeCoverReq As IXMLDOMNode
    'Dim xmlLifeCoverResp As IXMLDOMNode
    
    'Set xmlLifeCoverReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    'Set xmlDOMDoc = xmlLifeCoverReq.ownerDocument
    
    'm_objXMLAssist.AttachResponseData xmlLifeCoverReq, xmlMortgageCostsResp
    
    ' Call back...
    'Set xmlLifeCoverResp = objIAppQuoteBO.CalculateAndSaveLifeCoverCosts(xmlLifeCoverReq)
    'm_objErrAssist.CheckResponse xmlLifeCoverResp, xmlResponseElem, True
    ' End transaction
    
    ' PSC 10/11/2006 EP2_41 - Start
    blnCalcFeesOnly = xmlGetNodeAsBoolean(vxmlRequest, "MORTGAGECOSTS/CALCULATEFEESONLY")
    
    If blnCalcFeesOnly Then
        xmlResponseElem.appendChild xmlGetMandatoryNode(xmlMortgageCostsResp, "ONEOFFCOSTLIST")
        xmlResponseElem.appendChild xmlGetMandatoryNode(xmlMortgageCostsResp, "MORTGAGELENDER")
    End If
    ' PSC 10/11/2006 EP2_41 - End

    Set IApplicationQuoteTxBO_CalculateMortgageCosts = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_CalculateMortgageCosts = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit

End Function


Private Function IApplicationQuoteTxBO_CreateNewQuotation( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_CreateNewQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim strNewQuotationNumber As String
    Dim xmlTableElem As IXMLDOMNode, xmlElement As IXMLDOMNode
    
    Dim objICostModelAssistDO As ICostModelAssistDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem

    If m_objContext Is Nothing Then
        Set objICostModelAssistDO = New CostModelAssistDO
    Else
        Set objICostModelAssistDO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".CostModelAssistDO")
    End If

    ' Call the function in CostModelAssist and add the Quote number to Response
    strNewQuotationNumber = objICostModelAssistDO.CreateNewQuotation(vxmlRequest)
    
    ' Add the new quotation number to Response
    Set xmlTableElem = xmlOut.createElement("QUOTATION")
    Set xmlElement = xmlOut.createElement("NEWQUOTATIONNUMBER")
    xmlElement.Text = strNewQuotationNumber
    xmlResponseElem.appendChild xmlElement
    
    Set IApplicationQuoteTxBO_CreateNewQuotation = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:
    Set objICostModelAssistDO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    
    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_CreateNewQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IApplicationQuoteTxBO_RecommendQuotation( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RecommendQuotationVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_RecommendQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMNode, xmlElement As IXMLDOMNode
    Dim xmlApplNumberElem As IXMLDOMNode, xmlAFFNumberElem As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim objApplicationBO As ApplicationBO
    Dim objQuotationBO As QuotationBO
        
    Dim strResponse As String
    Dim strActiveQuoteNumber As String, strRecomQuoteNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
        Set objQuotationBO = New QuotationBO
    Else
        Set objApplicationBO = _
                m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objQuotationBO = _
                m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    End If
    '-----------------------------------------------
    ' Check Quotaion is not already recommended
    '------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    
    Dim strTag As String
    strTag = "APPLICATIONNUMBER"
    Set xmlApplNumberElem = xmlOut.createElement(strTag)
    xmlApplNumberElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlApplNumberElem
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlAFFNumberElem = xmlOut.createElement(strTag)
    xmlAFFNumberElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlAFFNumberElem
    
    strResponse = objApplicationBO.GetApplicationData(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    strActiveQuoteNumber = xmlDoc.selectSingleNode(".//ACTIVEQUOTENUMBER").Text
    strRecomQuoteNumber = xmlDoc.selectSingleNode(".//RECOMMENDEDQUOTENUMBER").Text
    
    If strActiveQuoteNumber = strRecomQuoteNumber Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 200
    End If
    
    '-------------------------------------------------------------
    'Check Quotation is affordable (Build the REQUEST first)
    '-------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("QUOTATION")
    xmlNewRequest.appendChild xmlTableElem
    
    xmlTableElem.appendChild xmlApplNumberElem
    xmlTableElem.appendChild xmlAFFNumberElem
    
    strTag = "QUOTATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = strActiveQuoteNumber
    xmlTableElem.appendChild xmlElement
    
    strResponse = objQuotationBO.IsQuotationAffordable(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Dim strIsQuoteAffordble As String
    strIsQuoteAffordble = xmlDoc.selectSingleNode(".//ISQUOTATIONAFFORDABLE").Text
    If strIsQuoteAffordble = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 201
    End If
    
    '--------------------------------------------------------------------
    ' Set Quotation as 'Recommended' - Update Table 'ApplicationFactFind'
    '--------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATIONFACTFIND")
    xmlNewRequest.appendChild xmlTableElem
    
    xmlTableElem.appendChild xmlApplNumberElem
    xmlTableElem.appendChild xmlAFFNumberElem
    
    strTag = "RECOMMENDEDQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = strActiveQuoteNumber
    xmlTableElem.appendChild xmlElement

    strResponse = objApplicationBO.UpdateApplicationFactFind(xmlNewRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    
    Set IApplicationQuoteTxBO_RecommendQuotation = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecommendQuotationExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlNewRequest = Nothing
    Set xmlApplNumberElem = Nothing
    Set xmlAFFNumberElem = Nothing
    Set xmlDoc = Nothing
    Set objApplicationBO = Nothing
    Set objQuotationBO = Nothing
    
    Exit Function
    
RecommendQuotationVbErr:
     If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_RecommendQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume RecommendQuotationExit

End Function

Private Function IApplicationQuoteTxBO_RecostMortgageComponentAndLifeCover( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RecostMortgageComponentAndLifeCoverErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_RecostMortgageComponentAndLifeCover"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlRootElement As IXMLDOMElement
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = "APPLICATIONQUOTE" Then
        Set xmlRootElement = vxmlRequest
    Else
        Set xmlRootElement = vxmlRequest.getElementsByTagName("APPLICATIONQUOTE").Item(0)
    End If

    If xmlRootElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                  "APPLICATIONQUOTE" & " tag not found"
    End If

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    Set xmlTempResponseNode = _
            objICostModelAssistBO.RecostMortgageComponentAndLifeCover(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    '--------------------------------------------------------------------------------
    ' Re-cost the associated Life Cover
    ' DPF - 11/11/2002 - remove life cover calls
    '--------------------------------------------------------------------------------
    'Set xmlTempResponseNode = RecostLifeCover(vxmlRequest)
    'm_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    'm_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
        
    Set IApplicationQuoteTxBO_RecostMortgageComponentAndLifeCover = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RecostMortgageComponentAndLifeCoverExit:
    
    Set xmlRootElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICostModelAssistBO = Nothing
    
    Exit Function

RecostMortgageComponentAndLifeCoverErr:
     If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_RecostMortgageComponentAndLifeCover = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit
    
End Function

Private Function RecostLifeCover(ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMElement
'--------------------------------------------------------------------------------------------
'Description : Recosts the loan component for ApplicationQuote
'              Called from - RecostMortgageComponentAndLifeCover
'--------------------------------------------------------------------------------------------
On Error GoTo RecostVbErr
    
    Const strFunctionName As String = "RecostLifeCover"
    
    Dim objApplicationBO As ApplicationBO
    Dim objAppQuoteBO As ApplicationQuoteBO
    Dim objLifeCoverBO As LifeCoverBO
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCustRoleElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMNode, xmlCustNode As IXMLDOMNode
    
    Dim strCust1Number As String, strCust1VersionNumber As String
    Dim strCust2Number As String, strCust2VersionNumber As String
    Dim strApplNumber As String, strAFFNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    '-------------------------------------------------------------------------
    ' Get Main Applicants and build input for 'BuildLifeCoverCalculationInput'
    '-------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    
    Dim strTag As String, strResponse As String
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    strApplNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlElement.Text = strApplNumber
    xmlTableElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    strAFFNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlElement.Text = strAFFNumber
    xmlTableElem.appendChild xmlElement
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    strResponse = objApplicationBO.FindMainApplicants(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set xmlCustRoleElem = xmlDoc.selectNodes("//CUSTOMERROLE").Item(0)
    strCust1Number = xmlCustRoleElem.selectSingleNode("CUSTOMERNUMBER").Text
    strCust1VersionNumber = xmlCustRoleElem.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
    
    Set xmlCustRoleElem = xmlDoc.selectNodes("CUSTOMERROLE").Item(1)
    If Not xmlCustRoleElem Is Nothing Then
        strCust2Number = xmlCustRoleElem.selectSingleNode("CUSTOMERNUMBER").Text
        strCust2VersionNumber = xmlCustRoleElem.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
    End If
    
    '-----------------------------------------------------------------------------------
    'Build Life Cover Calculations Input Interface
    '-----------------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("LIFECOVERCOSTS")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlListElem = vxmlRequest.selectSingleNode(".//LOANCOMPONENTLIST").cloneNode(True)
    xmlTableElem.appendChild xmlListElem
    
    Set xmlListElem = xmlOut.createElement("CUSTOMERLIST")
    xmlTableElem.appendChild xmlListElem
    
    Set xmlCustNode = xmlOut.createElement("CUSTOMER")
    xmlListElem.appendChild xmlCustNode
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
    xmlElement.Text = strCust1Number
    xmlCustNode.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = strCust1VersionNumber
    xmlCustNode.appendChild xmlElement
    
    If Len(strCust2Number) > 0 Then
        Set xmlCustNode = xmlOut.createElement("CUSTOMER")
        xmlListElem.appendChild xmlCustNode
    
        Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
        xmlElement.Text = strCust2Number
        xmlCustNode.appendChild xmlElement
    
        Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlElement.Text = strCust2VersionNumber
        xmlCustNode.appendChild xmlElement
    End If
    
    If m_objContext Is Nothing Then
        Set objAppQuoteBO = New ApplicationQuoteBO
    Else
        Set objAppQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")
    End If
    
    strResponse = objAppQuoteBO.BuildLifeCoverCalcInputInterface(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    '---------------------------------------------------------------------------
    'Calculate Life Cover Premiums
    '----------------------------------------------------------------------------
    If m_objContext Is Nothing Then
        Set objLifeCoverBO = New LifeCoverBO
    Else
        Set objLifeCoverBO = m_objContext.CreateInstance(gstrLIFECOVER_COMPONENT & ".LifeCoverBO")
    End If
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    m_objXMLAssist.AttachResponseData xmlNewRequest, xmlDoc.documentElement
    
    strResponse = objLifeCoverBO.CalculatePremiums(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set RecostLifeCover = xmlDoc.documentElement
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostVbExit:
    Set objApplicationBO = Nothing
    Set objAppQuoteBO = Nothing
    Set objLifeCoverBO = Nothing
    
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlCustRoleElem = Nothing
    Set xmlListElem = Nothing
    Set xmlCustNode = Nothing
    
    Exit Function
RecostVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set RecostLifeCover = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostVbExit

End Function


Private Function IApplicationQuoteTxBO_ReinstateQuotation( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ReinstateQuotationVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_ReinstateQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    'Dim strQuotationNumber As String
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    'SYS0977 DC 07/12/2001
    Set xmlTempResponseNode = objICostModelAssistBO.ReinstateQuotation(vxmlRequest, "COST MODELLING")
    'SYS0977 End
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteTxBO_ReinstateQuotation = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ReinstateQuotationExit:
    Set objICostModelAssistBO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

ReinstateQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_ReinstateQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume ReinstateQuotationExit

End Function
Private Function IApplicationQuoteTxBO_CreateAdHocFeeAndRemodel( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateAdHocFeeAndRemodelVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_CreateAdHocFeeAndRemodel"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTmpNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlInAppNumberNode As IXMLDOMNode
    Dim xmlInAppFFNumberNode As IXMLDOMNode
    Dim xmlQuotationNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlMSQNode As IXMLDOMNode
    Dim strResponse As String
    Dim objIAppQuoteBO As IApplicationQuoteBO
    Dim objMortSubQuoteBO As MortgageSubQuoteBO
    Dim objPProcBO As PaymentProcessingBO
    Dim objComboDO As ComboDO
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    Dim strCostType As String
    Dim strValidationType As String
    Dim strPattern As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIAppQuoteBO = New ApplicationQuoteBO
        Set objMortSubQuoteBO = New MortgageSubQuoteBO
    Else
        Set objIAppQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")
        Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
    Set xmlInAppNumberNode = vxmlRequest.selectSingleNode("//APPLICATIONNUMBER")
    Set xmlInAppFFNumberNode = vxmlRequest.selectSingleNode("//APPLICATIONFACTFINDNUMBER")
    
    'Get the active or accepted quotation
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    xmlTableElem.appendChild xmlInAppNumberNode.cloneNode(True)
    xmlTableElem.appendChild xmlInAppFFNumberNode.cloneNode(True)
    
    Set xmlTempResponseNode = objIAppQuoteBO.GetAcceptedOrActiveQuoteData(xmlNewRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    Set xmlQuotationNode = xmlTempResponseNode.selectSingleNode(".//MORTGAGESUBQUOTE")
    
    'Create fee in one off costs
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    Set xmlMSQNode = xmlNewRequest.appendChild(xmlTableElem)
    Set xmlTableElem = xmlOut.createElement("ONEOFFCOST")
    xmlMSQNode.appendChild xmlTableElem
    xmlTableElem.appendChild xmlInAppNumberNode.cloneNode(True)
    xmlTableElem.appendChild xmlInAppFFNumberNode.cloneNode(True)
    xmlTableElem.appendChild xmlQuotationNode.selectSingleNode("MORTGAGESUBQUOTENUMBER").cloneNode(True)
    'FeeType passed in is the ValueId. CreateOneOffCosts requires the validationType
    ' Get combo data for one off costs
    Set xmlRootElement = xmlOut.createElement("LIST")
    Set xmlElement = xmlOut.createElement("LISTNAME")
    xmlElement.Text = "OneOffCost"
    xmlRootElement.appendChild xmlElement
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    strResponse = objComboDO.GetComboList(xmlRootElement.xml)
    Set xmlComboDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    strCostType = m_objXMLAssist.GetElementText(vxmlRequest, "//FEETYPE")
    strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE[. != 'APR']"
    strValidationType = m_objXMLAssist.GetNodeValue(xmlComboDoc, strPattern)
    
    Set xmlTmpNode = xmlTableElem.appendChild(xmlOut.createElement("IDENTIFIER"))
    xmlTmpNode.Text = strValidationType
    Set xmlTmpNode = xmlTableElem.appendChild(vxmlRequest.selectSingleNode("//AMOUNT").cloneNode(True))
    Set xmlTmpNode = xmlTableElem.appendChild(vxmlRequest.selectSingleNode("//REFUNDAMOUNT").cloneNode(True))  ' MAR28
    Set xmlTmpNode = xmlTableElem.appendChild(vxmlRequest.selectSingleNode("//ADDTOLOAN").cloneNode(True))
    Set xmlTmpNode = xmlTableElem.appendChild(xmlOut.createElement("ADHOCIND"))
    xmlTmpNode.Text = "1"
    
    strResponse = objMortSubQuoteBO.CreateOneOffCosts(xmlNewRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    'Remodel quotation
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    xmlTableElem.appendChild xmlInAppNumberNode.cloneNode(True)
    xmlTableElem.appendChild xmlInAppFFNumberNode.cloneNode(True)
    xmlTableElem.appendChild xmlQuotationNode.cloneNode(True)
    xmlTableElem.appendChild vxmlRequest.selectSingleNode("//CALLEDFROM").cloneNode(True)
        
    Set xmlResponseNode = IApplicationQuoteTxBO_RemodelQuotation(xmlNewRequest)
    m_objErrAssist.CheckResponse xmlResponseNode, xmlResponseElem, True
    
    
    Set IApplicationQuoteTxBO_CreateAdHocFeeAndRemodel = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateAdHocFeeAndRemodelExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlInAppNumberNode = Nothing
    Set xmlInAppFFNumberNode = Nothing
    Set xmlQuotationNode = Nothing
    Set xmlElement = Nothing
    Set xmlRootElement = Nothing
    Set objIAppQuoteBO = Nothing
    Set objMortSubQuoteBO = Nothing
    Set xmlMSQNode = Nothing
    Set xmlTmpNode = Nothing
    Set objComboDO = Nothing
    Set xmlComboDoc = Nothing
    
    Exit Function

CreateAdHocFeeAndRemodelVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_CreateAdHocFeeAndRemodel = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateAdHocFeeAndRemodelExit

End Function

Private Function IApplicationQuoteTxBO_ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_ResetMortgageSubQuote"

    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMNode
    Dim objIAppQuoteBO As IApplicationQuoteBO
    Dim objIApplicationDO As IApplicationDO
    Dim objMSQBO As MortgageSubQuoteBO
    Dim objConvertAssist As ConvertAssist
    Set objConvertAssist = New ConvertAssist
    Dim strAmtRequested As String
    Dim lngDeposit As Long
    
    Dim dblLTV As Double
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '----------------------------------------------------
    ' Recalculate LTV using the new amount requested
    '----------------------------------------------------
    If m_objContext Is Nothing Then
        Set objIAppQuoteBO = New ApplicationQuoteBO
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIAppQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
    End If
    
    Set xmlNewRequest = vxmlRequest.cloneNode(True)
        
    m_objXMLAssist.ChangeNodeName xmlNewRequest, "APPLICATIONQUOTE", "LTV"
        
    'Note : ApplicationQuoteBO.CalcCostModelLTV is assuming the Request to be attached to FreeThreadedDOMDocument40.
    '       (in Search Pattern in the method CostModelAssistBO.CalcCostModelLTV)
    Set xmlDoc = m_objXMLAssist.load(xmlNewRequest.xml, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = objIAppQuoteBO.CalcCostModelLTV(xmlDoc.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' PSC 30/06/00 SYS1012
    dblLTV = objConvertAssist.CSafeDbl(m_objXMLAssist.GetMandatoryElementText(xmlTempResponseNode, ".//LTV"))
    
    ' attach LTV Element to Response
    xmlResponseElem.appendChild xmlTempResponseNode.selectSingleNode(".//LTV")
    
    '-------------------------------------------------------------------------------
    ' Recalculate Deposit           15/05/00 JLD SYS0715
    '-------------------------------------------------------------------------------
    Set xmlTempResponseNode = objIApplicationDO.GetApplicationData(vxmlRequest)
    Set xmlTempNode = xmlTempResponseNode.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE")
    lngDeposit = 0
    If Not xmlTempNode Is Nothing Then
        strAmtRequested = vxmlRequest.selectSingleNode(".//AMOUNTREQUESTED").Text
        
        'SYS4506 - Overflow error, Amount Request is a double.
        If objConvertAssist.CSafeLng(xmlTempNode.Text) > objConvertAssist.CSafeDbl(strAmtRequested) Then
        'SYS4506 - End.
            lngDeposit = objConvertAssist.CSafeLng(xmlTempNode.Text) - objConvertAssist.CSafeLng(strAmtRequested)
        End If
    End If
            
    '------------------------------------------------------------------------------
    ' Reset the mortgage sub-quote
    '-------------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//AMOUNTREQUESTED").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    'MAR1061 pass purchaseprice too
    Set xmlElement = vxmlRequest.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("LTV")
    ' PSC 30/06/00 SYS1012
    xmlElement.Text = dblLTV
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("DEPOSIT")
    xmlElement.Text = lngDeposit
    xmlTableElem.appendChild xmlElement
    
    If m_objContext Is Nothing Then
        Set objMSQBO = New MortgageSubQuoteBO
    Else
        Set objMSQBO = _
            m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
    Dim strResponse As String
    strResponse = objMSQBO.ResetMortgageSubQuote(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set IApplicationQuoteTxBO_ResetMortgageSubQuote = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ResetMortgageSubQuoteExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set objIAppQuoteBO = Nothing
    Set objMSQBO = Nothing
    Set objIApplicationDO = Nothing
    Set objConvertAssist = Nothing
    
    Exit Function

ResetMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit
End Function

Private Function IApplicationQuoteTxBO_RemodelQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Remodels the repayment costs and APR to ensure that the latest rates
'               have been applied prior to issuing the offer
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------

' BMIDS977  CC089  Many changes - old code not preserved for clarity
'                  See Version 60 for code prior to CC089.

On Error GoTo RemodelQuotationVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_RemodelQuotation"
        
    Dim xmlRequestElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponse As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlAppNo As IXMLDOMElement
    Dim xmlMSQNo As IXMLDOMElement
    Dim xmlAppFFNo As IXMLDOMElement
    Dim xmlMortSubQuoteNo As IXMLDOMElement
    'Dim xmlLifeSubQuoteNo As IXMLDOMElement
    Dim xmlCost As IXMLDOMElement
    Dim xmlMtgeCostList As IXMLDOMNodeList
    Dim xmlMtgeCost As IXMLDOMElement
    Dim xmlMtgeSubQuote As IXMLDOMElement
    Dim xmlComponentList As IXMLDOMNodeList
    Dim xmlComponent As IXMLDOMElement
    Dim xmlMortgageProduct As IXMLDOMElement
    Dim xmlOneOffCosts As IXMLDOMElement
    Dim xmlCalcsRequest As IXMLDOMElement
    Dim xmlIllusRequest As IXMLDOMElement
    Dim xmlCustomerList As IXMLDOMElement
    Dim xmlComponentListElement As IXMLDOMElement
    Dim xmlCustomersDoc As FreeThreadedDOMDocument40
    Dim xmlMtgeOneOffCostsDoc As FreeThreadedDOMDocument40
    Dim xmlAppDetails As IXMLDOMElement
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    Dim xmlLoanCompListNode As IXMLDOMNode

    Dim objApplicationBO As ApplicationBO
    Dim objAppManagerBO As ApplicationManagerBO
    Dim objQuotationBO As QuotationBO
    Dim objIApplicationQuoteBO As IApplicationQuoteBO
    Dim objMortSubQuoteBO As MortgageSubQuoteBO
    Dim objComboDO As ComboDO
    Dim objConvAssist As ConvertAssist
    
    ' PSC 15/03/2007 EP2_1956 - Start
    Dim xmlCopyFeeRequest As IXMLDOMNode
    Dim xmlMortgageIntroducerFee As IXMLDOMNode
    ' PSC 15/03/2007 EP2_1956 - End
    
    Dim strResponse As String
    Dim strAcceptedQuoteNo As String
    Dim lngTotalLoanAmt As Long
    Dim strApplicationType As String
    Dim strOtherSysAccountNumber As String
    Dim lngAmountRequested As Long
    Dim dblLTV As Double
    Dim blnFurtherAdvance As Boolean
    Dim xmlMainNode As IXMLDOMElement
    Dim xmlCustomerRoleList As IXMLDOMNodeList
    Dim xmlQuotationDoc As FreeThreadedDOMDocument40
    Dim strCostType As String
    Dim strPattern As String
    Dim strValidationType As String
    Dim strType As String
    
    Dim objIApplClassDef As IApplicationClassDef
    Dim xmlApplClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlAFFData As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement 'Input for fetching AFFData
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlAFFElement As IXMLDOMElement
    Dim objIDOAssist As IDOAssist
    Dim xmlCompDateElem As IXMLDOMElement
    
    Dim xmlValidationType As IXMLDOMNode
    Dim xmlDrawDown As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim strValue As String
    
    Dim objICostModelAssistDO   As ICostModelAssistDO
    
    Dim lngTotalIncentives           As Long
    Dim strCalledFrom                As String
    Dim bApplyRebateAddition         As Boolean
    Dim lngCharge                    As Long
    Dim lngTotalNonAPRCharges        As Long
    Dim xmlMtgeOneOffCostList        As IXMLDOMNode
    Dim xmlMtgeOneOffCostsForCalc    As IXMLDOMNode
    Dim xmlNewRequest                As IXMLDOMElement

    Set objIApplClassDef = New ApplicationClassDef
    Set xmlApplClassDefDoc = New FreeThreadedDOMDocument40
    xmlApplClassDefDoc.validateOnParse = False
    xmlApplClassDefDoc.setProperty "NewParser", True
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objMortSubQuoteBO = New MortgageSubQuoteBO
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO") ' SR 06/09/2004 : BMIDS815
    End If
    
    Set objConvAssist = New ConvertAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrAPPLICATION_NAME Then
        Set xmlRequestElement = vxmlRequest
    Else
        Set xmlRequestElement = vxmlRequest.getElementsByTagName(cstrAPPLICATION_NAME).Item(0)
    End If
    
    If xmlRequestElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrAPPLICATION_NAME & " tag not found"
    End If
    
    Set xmlAppNo = m_objXMLAssist.GetNode(xmlRequestElement, "APPLICATIONNUMBER")
    Set xmlAppFFNo = m_objXMLAssist.GetNode(xmlRequestElement, "APPLICATIONFACTFINDNUMBER")

    strAcceptedQuoteNo = m_objXMLAssist.GetNodeValue(xmlRequestElement, "ACCEPTEDQUOTENUMBER", False, False)
    strApplicationType = m_objXMLAssist.GetNodeValue(xmlRequestElement, "TYPEOFAPPLICATION", False, False)
    strOtherSysAccountNumber = m_objXMLAssist.GetNodeValue(xmlRequestElement, "OTHERSYSTEMACCOUNTNUMBER", False, False)
                                                   
    strCalledFrom = m_objXMLAssist.GetNodeValue(xmlRequestElement, "CALLEDFROM", False, False)     ' BMIDS977
                                                   
'Stage1: Get customers for the application
    
    Set objAppManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
   
    strResponse = objAppManagerBO.FindCustomersForApplication(vxmlRequest.xml)
    Set xmlCustomersDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlCustomersDoc.documentElement, xmlResponseElem, True
    
    Set xmlCustomerRoleList = xmlCustomersDoc.selectNodes(".//CUSTOMERROLE")
    If xmlCustomerRoleList.length > 0 Then
        Set xmlCustomerList = xmlOut.createElement("CUSTOMERLIST")
        
        For Each xmlTempNode In xmlCustomerRoleList
            Set xmlTableElem = xmlOut.createElement("CUSTOMER")
            Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
            xmlElement.Text = xmlTempNode.selectSingleNode(".//CUSTOMERNUMBER").Text
            xmlTableElem.appendChild xmlElement
            
            Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
            xmlElement.Text = xmlTempNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
            xmlTableElem.appendChild xmlElement
            
            Set xmlElement = xmlOut.createElement("OTHERSYSTEMCUSTOMERNUMBER")
            xmlElement.Text = xmlTempNode.selectSingleNode(".//OTHERSYSTEMCUSTOMERNUMBER").Text
            xmlTableElem.appendChild xmlElement
            
            xmlCustomerList.appendChild xmlTableElem
        Next xmlTempNode
    End If
    
'Stage2: Get the accepted quote if it has not been passed in

    If Len(strAcceptedQuoteNo) = 0 Or Len(strApplicationType) = 0 Then
        Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlRootElement = xmlOut.createElement("APPLICATIONFACTFIND")
        xmlTempRequest.appendChild xmlRootElement
        xmlRootElement.appendChild xmlAppNo.cloneNode(True)
        xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
        
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        
        strResponse = objApplicationBO.GetApplicationFactFindData(xmlTempRequest.xml)
        
        Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
        ' Set objApplicationBO = Nothing
        If Len(strAcceptedQuoteNo) = 0 Then
            strAcceptedQuoteNo = m_objXMLAssist.GetNodeValue(xmlTempResponseDoc, ".//ACCEPTEDQUOTENUMBER")
        End If
        
        If Len(strApplicationType) = 0 Then
            strApplicationType = m_objXMLAssist.GetNodeValue(xmlTempResponseDoc, ".//TYPEOFAPPLICATION")
        End If
    End If
    
    If Len(strOtherSysAccountNumber) = 0 Then
        Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlRootElement = xmlOut.createElement("APPLICATION")
        xmlTempRequest.appendChild xmlRootElement
        xmlRootElement.appendChild xmlAppNo.cloneNode(True)
        
        strResponse = objApplicationBO.GetData(xmlTempRequest.xml)
        Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
        
        strOtherSysAccountNumber = m_objXMLAssist.GetTagValue(xmlTempResponseDoc.documentElement, "OTHERSYSTEMACCOUNTNUMBER")
    End If

    ' Get the accepted quotation data
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlRootElement = xmlOut.createElement("QUOTATION")
    xmlTempRequest.appendChild xmlRootElement
    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
    Set xmlElement = xmlOut.createElement("QUOTATIONNUMBER")
    xmlElement.Text = strAcceptedQuoteNo
    xmlRootElement.appendChild xmlElement
    
    Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    
    strResponse = objQuotationBO.GetData(xmlTempRequest.xml)
    Set xmlQuotationDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlQuotationDoc.documentElement, xmlResponseElem, True

'Stage3: Refresh account data if applicable.

    ' When remodel is called from PP010 or PP035 (application fee screens) Stage 3 will be skipped.

    If ((strCalledFrom <> "PP010") And (strCalledFrom <> "PP035")) Then
    
        Set xmlMSQNo = xmlQuotationDoc.selectSingleNode(".//MORTGAGESUBQUOTENUMBER")
    
        Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlRootElement = xmlOut.createElement("MORTGAGESUBQUOTE")
        xmlTempRequest.appendChild xmlRootElement
        xmlRootElement.appendChild xmlAppNo.cloneNode(True)
        xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
        xmlRootElement.appendChild xmlMSQNo.cloneNode(True)
    
        strResponse = objMortSubQuoteBO.GetData(xmlTempRequest.xml)
        Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
        lngAmountRequested = objConvAssist.CSafeLng( _
                        m_objXMLAssist.GetTagValue(xmlTempResponseDoc.documentElement, "AMOUNTREQUESTED"))
        dblLTV = objConvAssist.CSafeDbl(m_objXMLAssist.GetTagValue(xmlTempResponseDoc.documentElement, "LTV"))
        
        blnFurtherAdvance = IsValidationType("TypeOfMortgage", strApplicationType, "GA")
            Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        If blnFurtherAdvance Then
            Set xmlMainNode = xmlOut.createElement("REFRESHMORTGAGEACCOUNTDATA")
        Else
            Set xmlMainNode = xmlOut.createElement("GETANDSAVEPORTEDSTEPANDPERIODFROMMORTGAGEACCOUNT")
        End If
        
        xmlTempRequest.appendChild xmlMainNode
        
        Set xmlElement = xmlOut.createElement("PORTINGINDICATOR")
        xmlElement.Text = IIf(blnFurtherAdvance, "0", "1")
        xmlMainNode.appendChild xmlElement
    
        Set xmlTableElem = xmlOut.createElement("APPLICATION")
        xmlMainNode.appendChild xmlTableElem
        xmlTableElem.appendChild xmlAppNo.cloneNode(True)
        xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
    
        Set xmlElement = xmlOut.createElement("TYPEOFAPPLICATION")
        xmlElement.Text = strApplicationType
        xmlTableElem.appendChild xmlElement
    
        Set xmlElement = xmlOut.createElement("BMACCOUNTNUMBER")
        xmlElement.Text = strOtherSysAccountNumber
        xmlTableElem.appendChild xmlElement
            
        ' Use Customer List set up above
        If xmlCustomerRoleList.length > 0 Then
            xmlMainNode.appendChild xmlCustomerList
        End If
            
        Set xmlTableElem = xmlOut.createElement("MORTGAGESUBQUOTE")
        xmlMainNode.appendChild xmlTableElem
        xmlTableElem.appendChild xmlAppNo.cloneNode(True)
        xmlTableElem.appendChild xmlAppFFNo.cloneNode(True)
    
        xmlTableElem.appendChild xmlMSQNo.cloneNode(True)
    
        Set xmlElement = xmlOut.createElement("AMOUNTREQUESTED")
        xmlElement.Text = lngAmountRequested
        xmlTableElem.appendChild xmlElement
    
        Set xmlElement = xmlOut.createElement("LTV")
        xmlElement.Text = dblLTV
        xmlTableElem.appendChild xmlElement
           
        If blnFurtherAdvance Then
            Set xmlTempNode = IApplicationQuoteTxBO_RefreshMortgageAccountData(xmlTempRequest)
            m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
            strValue = m_objXMLAssist.GetTagValue(xmlTempNode, "LTVCHANGED")
            If strValue = "1" Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7020
           
                GoTo Warning7020Exit
            End If
        Else
            Set xmlTempNode = IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount(xmlTempRequest)
            m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
        End If
    End If
    
'Stage4: Create the new mortgage and life sub quotes based on accepted quotation
    
    'Set the operation to Remodel so that AdHoc fees and Incentives are saved correctly when
    'the new subquote is created.
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    xmlTempRequest.setAttribute "OPERATION", "Remodel"
    
    Set xmlRootElement = xmlQuotationDoc.selectSingleNode(".//QUOTATION").cloneNode(True)
    xmlTempRequest.appendChild xmlRootElement
    
    Set xmlElement = xmlOut.createElement("CONTEXT")
    xmlElement.Text = "Cost Modelling"
    xmlRootElement.appendChild xmlElement
        
    strResponse = objQuotationBO.CreateNewMortgageLifeSubQuotes(xmlTempRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    Set xmlMortSubQuoteNo = m_objXMLAssist.GetNode(xmlTempResponseDoc, ".//MORTGAGESUBQUOTENUMBER")
    'Set xmlLifeSubQuoteNo = m_objXMLAssist.GetNode(xmlTempResponseDoc, ".//LIFESUBQUOTENUMBER")
    
    ' PSC 15/03/2007 EP2_1956 - Start
    Set xmlCopyFeeRequest = xmlGetRequestNode(vxmlRequest)
    Set xmlMortgageIntroducerFee = xmlOut.createElement("MORTGAGEINTRODUCERFEE")
    xmlCopyFeeRequest.appendChild xmlMortgageIntroducerFee
    Set xmlTempNode = xmlOut.createElement("APPLICATIONNUMBER")
    xmlTempNode.Text = xmlAppNo.Text
    xmlMortgageIntroducerFee.appendChild xmlTempNode
    Set xmlTempNode = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
    xmlTempNode.Text = xmlGetMandatoryNodeText(xmlQuotationDoc, "RESPONSE/QUOTATION/MORTGAGESUBQUOTENUMBER")
    xmlMortgageIntroducerFee.appendChild xmlTempNode
    Set xmlTempNode = xmlOut.createElement("NEWMORTGAGESUBQUOTENUMBER")
    xmlTempNode.Text = xmlMortSubQuoteNo.Text
    xmlMortgageIntroducerFee.appendChild xmlTempNode
    
    strResponse = objMortSubQuoteBO.CopyIntroducerFees(xmlCopyFeeRequest.xml)
    errCheckXMLResponse strResponse, True, xmlResponseElem
    ' PSC 15/03/2007 EP2_1956 - End


'Stage5: Get all fees to be used in calculations.
    
    ' xmlMtgeOneOffCostList holds a list of ALL fees in the current MortgageOneOffCosts table.
    ' xmlMtgeOneOffCostsForCalc holds
    '       1) ApplicationFeeType : fees not added to loan but included in APR
    '       2) MortgageOneOffCost : fees added to loan
     
    ' Get current MortgageOneOffCosts
     
    m_objXMLAssist.ChangeNodeName xmlTempRequest, "QUOTATION", "ONEOFFCOST"
    strResponse = objMortSubQuoteBO.GetOneOffCostsDetails(xmlTempRequest.xml)
    Set xmlMtgeOneOffCostsDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlMtgeOneOffCostsDoc.documentElement, xmlResponseElem, True
   
    ' Process One Off Costs to create the list required for use in the calculations
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    xmlNewRequest.appendChild xmlAppNo.cloneNode(True)
       
    Set xmlMtgeOneOffCostList = xmlMtgeOneOffCostsDoc.selectSingleNode(".//MORTGAGEONEOFFCOSTLIST").cloneNode(True)
    
    'Save INCENTIVES value for later use
    strValue = m_objXMLAssist.GetNodeValue(xmlMtgeOneOffCostsDoc.documentElement, ".//TOTALINCENTIVES", False)
    lngTotalIncentives = objConvAssist.CSafeLng(strValue)
    
    ' Remove any costs with an AMOUNT of 0
    Set xmlMtgeCostList = xmlMtgeOneOffCostList.selectNodes(".//MORTGAGEONEOFFCOST")
    For Each xmlMtgeCost In xmlMtgeCostList

        lngCharge = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//AMOUNT")

        If (lngCharge = 0) Then
            xmlMtgeOneOffCostList.removeChild xmlMtgeCost
        End If

    Next
     
    xmlNewRequest.appendChild xmlMtgeOneOffCostList.cloneNode(True)
        
    Set xmlTempNode = IApplicationQuoteTxBO_ProcessAllFeesForRemodelQuote(xmlNewRequest)
    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
        
    Set xmlMtgeOneOffCostsForCalc = xmlTempNode.selectSingleNode(".//MORTGAGEONEOFFCOSTLIST").cloneNode(True)
    
    lngTotalNonAPRCharges = m_objXMLAssist.GetNodeValue(xmlTempNode, ".//NONAPRCHARGES")
    
    ' Create a new copy of the one off costs for the mortgage subquote just created
     
    ' Get combo data for one off costs
    Set xmlRootElement = xmlOut.createElement("LIST")
    Set xmlElement = xmlOut.createElement("LISTNAME")
    xmlElement.Text = "OneOffCost"
    xmlRootElement.appendChild xmlElement
       
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")

    strResponse = objComboDO.GetComboList(xmlRootElement.xml)
    Set xmlComboDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
     
    ' Manipulate both lists of fees
    ' Remove AppNumber, AppFactFindNumber, MortgageSubQuoteNumber and OneOffCostsequenceNumber
    ' from each One Off Cost.
    ' Add IDENTIFIER node to each One Off cost
    
    Set xmlMtgeCostList = xmlMtgeOneOffCostsForCalc.selectNodes(".//MORTGAGEONEOFFCOST")
    For Each xmlMtgeCost In xmlMtgeCostList
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//APPLICATIONNUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//APPLICATIONFACTFINDNUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//MORTGAGESUBQUOTENUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
    
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//ONEOFFCOSTSEQUENCENUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
    
        Set xmlRootElement = xmlOut.createElement("IDENTIFIER")
        
        'Need to put the ValidationType into the Identifier
        strCostType = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
        
        ' Get all the Validation Types for this OneOffCost
        strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE"

        ' Choose the first validation type which is not APR or END
        For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
            strType = xmlValidationType.Text
            If (strType <> "APR") And (strType <> "END") Then
                strValidationType = strType
                Exit For ' get out of this loop
            End If
        Next
 
        xmlRootElement.Text = strValidationType
        xmlMtgeCost.appendChild xmlRootElement
   
    Next
    
    Set xmlMtgeCostList = xmlMtgeOneOffCostList.selectNodes(".//MORTGAGEONEOFFCOST")
    For Each xmlMtgeCost In xmlMtgeCostList
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//APPLICATIONNUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//APPLICATIONFACTFINDNUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//MORTGAGESUBQUOTENUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
    
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//ONEOFFCOSTSEQUENCENUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
  
    
        Set xmlRootElement = xmlOut.createElement("IDENTIFIER")
        
        'Need to put the ValidationType into the Identifier
        strCostType = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
        
        ' Get all the Validation Types for this OneOffCost
        strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE"

        ' Choose the first validation type which is not APR or END
        For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
            strType = xmlValidationType.Text
            If (strType <> "APR") And (strType <> "END") Then
                strValidationType = strType
                Exit For ' get out of this loop
            End If
        Next
        
        xmlRootElement.Text = strValidationType
        xmlMtgeCost.appendChild xmlRootElement
   
        ' EP1241 MORTGAGEONEOFFCOSTTYPE shouldnt go to CreateOneOffCost or it will get used
        ' in the get highest sequence routine and cause a duplicate row insert.
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
        ' EP1241 End
    
    Next
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlRootElement = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlNewRequest.appendChild xmlRootElement
    
    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
    xmlRootElement.appendChild xmlMortSubQuoteNo.cloneNode(True)
    
    Set xmlElement = xmlMtgeOneOffCostList.cloneNode(True)
    m_objXMLAssist.ChangeNodeName xmlElement, "MORTGAGEONEOFFCOSTLIST", "ONEOFFCOSTLIST"
    m_objXMLAssist.ChangeNodeName xmlElement, "MORTGAGEONEOFFCOST", "ONEOFFCOST"
    
    ' The AdHoc costs will have been created when the new mortgage sub quote was created
    ' in CreateNewMortgageLifeSubQuotes above.
    ' Thus, when creating the one off costs for the new subquote now, do not re-create
    ' the adHoc costs.
    
    Set xmlMtgeCostList = xmlElement.selectNodes(".//ONEOFFCOST")
    For Each xmlMtgeCost In xmlMtgeCostList
     
        If (m_objXMLAssist.GetNodeValue(xmlMtgeCost, "ADHOCIND", True, False) = "1") Then
            xmlElement.removeChild xmlMtgeCost
        End If
    Next
      
    xmlNewRequest.appendChild xmlElement
    
    strResponse = objMortSubQuoteBO.CreateOneOffCosts(xmlNewRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    
'Stage6: Get all the necessary Application and Loan details.

    ' Get loan composition details for the new sub quotes
    m_objXMLAssist.ChangeNodeName xmlTempRequest, "ONEOFFCOST", "LOANCOMPOSITION"
    Set xmlRootElement = m_objXMLAssist.GetNode(xmlTempRequest, ".//LOANCOMPOSITION")
    Set xmlElement = m_objXMLAssist.GetNode(xmlRootElement, ".//MORTGAGESUBQUOTENUMBER")
    xmlRootElement.replaceChild xmlMortSubQuoteNo.cloneNode(True), xmlElement
    'Set xmlElement = m_objXMLAssist.GetNode(xmlRootElement, ".//LIFESUBQUOTENUMBER")
    'xmlRootElement.replaceChild xmlLifeSubQuoteNo.cloneNode(True), xmlElement

    Set objIApplicationQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")

    Set xmlAppDetails = objIApplicationQuoteBO.GetLoanCompositionDetails(xmlTempRequest)
    m_objErrAssist.CheckResponse xmlAppDetails, xmlResponseElem, True
    Set xmlMtgeSubQuote = m_objXMLAssist.GetNode(xmlAppDetails, ".//MORTGAGESUBQUOTE")

    Set objAppManagerBO = Nothing

'Stage7: Get the loans not redeemed

    ' BMIDS977  Not needed
      
'Stage 8: Calculate total of costs to be added to loan using the costs calculated above.

    Dim lngCostsToBeAdded As Long
    
    lngCostsToBeAdded = 0
    
    ' Get all the costs where AddedToLoan indicator = true
    Set xmlMtgeCostList = xmlMtgeOneOffCostsForCalc.selectNodes(".//MORTGAGEONEOFFCOST[ADDTOLOAN = '1']")
        
    For Each xmlMtgeCost In xmlMtgeCostList
 
        strValue = m_objXMLAssist.GetNodeValue(xmlMtgeCost, "AMOUNT")
        lngCostsToBeAdded = lngCostsToBeAdded + objConvAssist.CSafeLng(strValue)
   
    Next

'Stage 9:  Calculate total charges
    Dim lngTotalCharges As Long
    Dim bAdd As Boolean
    
    lngTotalCharges = 0
    lngCharge = 0

    'For the total charges, use xmlMtgeOneOffCostsForCalc which holds all the APR charges
    'with rebates and additions applied and ADD the nonAPR charges returned from ProcessAllFeesForRemodelQuote
    
    'BMIDS881 MIG should be included in the total charges.
    'Search for costs where validation type <> TID or CBC or CPP
    
    Set xmlMtgeCostList = xmlMtgeOneOffCostsForCalc.selectNodes(".//MORTGAGEONEOFFCOST")
    
    For Each xmlCost In xmlMtgeCostList
        
        ' Convert the cost type to the validation type
        strCostType = m_objXMLAssist.GetNodeValue(xmlCost, ".//MORTGAGEONEOFFCOSTTYPE")
        lngCharge = m_objXMLAssist.GetNodeValue(xmlCost, ".//AMOUNT")
        
        strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE"

        bAdd = True
        
        For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
            
            strValidationType = xmlValidationType.Text
            
            If (strValidationType = "TID") Or _
               (strValidationType = "CBC") Or _
               (strValidationType = "CPP") Then      ' BMIDS881
                
                bAdd = False
               
            End If
        Next
        
        If (bAdd = True) Then
            lngTotalCharges = lngTotalCharges + lngCharge
        End If
    Next
  
    lngTotalCharges = lngTotalCharges + lngTotalNonAPRCharges
  
'Stage 10:  Update the total loan amount on the mortgage subquote

    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//TOTALLOANAMOUNT")
    lngTotalLoanAmt = objConvAssist.CSafeLng(xmlElement.Text)
    lngTotalLoanAmt = lngTotalLoanAmt + lngCostsToBeAdded
    xmlElement.Text = CStr(lngTotalLoanAmt)
    
                   
'Stage 11:  Update the mortgage sub quote with the new charges
    
    Dim lngTotalInitCharges As Long
    Dim lngTotalNetCharges As Long
    Dim lngSealingFee As Long
    Dim lngReleaseFee As Long
    Dim lngDeposit As Long
    Dim lngManualIncentives As Long                ' BMIDS881
    Dim lngIncentives As Long                      ' BMIDS881

 
    strValue = m_objXMLAssist.GetNodeValue(xmlMtgeOneOffCostList, ".//MORTGAGEONEOFFCOST[IDENTIFIER = 'SEA']/AMOUNT", False)
    lngSealingFee = objConvAssist.CSafeLng(strValue)
    
    strValue = m_objXMLAssist.GetNodeValue(xmlMtgeOneOffCostList, ".//MORTGAGEONEOFFCOST[IDENTIFIER = 'DEE']/AMOUNT", False)
    lngReleaseFee = objConvAssist.CSafeLng(strValue)
   
    strValue = m_objXMLAssist.GetNodeValue(xmlMtgeSubQuote, "DEPOSIT", False)
    lngDeposit = objConvAssist.CSafeLng(strValue)
    
    ' BMIDS881 Get Manual Incentives and, if they exist, use these in preference to TotalIncentives
    strValue = m_objXMLAssist.GetNodeValue(xmlMtgeSubQuote, "MANUALINCENTIVEAMOUNT", False)
    lngManualIncentives = objConvAssist.CSafeLng(strValue)
  
    If (lngManualIncentives > 0) Then
        lngIncentives = lngManualIncentives
    Else
        lngIncentives = lngTotalIncentives
    End If

    lngTotalInitCharges = lngTotalCharges - lngSealingFee - lngReleaseFee
    'BMIDS881  Subtract the costs which have been added to the loan when calculating TotalNetCharges
    lngTotalNetCharges = lngTotalCharges + lngDeposit - lngIncentives - lngCostsToBeAdded
    
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, "TOTALCHARGES")
    xmlElement.Text = CStr(lngTotalCharges)
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, "TOTALINITIALCHARGES")
    xmlElement.Text = CStr(lngTotalInitCharges)
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, "TOTALNETCHARGES")
    xmlElement.Text = CStr(lngTotalNetCharges)
    
    Set xmlComponentList = xmlAppDetails.selectNodes(".//LOANCOMPONENT[PORTEDLOAN != '1']")
         
           
'Stage12:

    ' Find main loan component
    Dim intMainLoanNo As Integer
    Dim intCurrentNo As Integer
    Dim strMainLoanNo As String
    
    intMainLoanNo = 9999
    
    For Each xmlComponent In xmlComponentList
        strValue = m_objXMLAssist.GetNodeValue(xmlComponent, ".//LOANCOMPONENTSEQUENCENUMBER")
        intCurrentNo = objConvAssist.CSafeInt(strValue)
        
        If intCurrentNo < intMainLoanNo Then
            intMainLoanNo = intCurrentNo
        End If
        
    Next
    
    strMainLoanNo = CStr(intMainLoanNo)
    
    ' Remove previous request
    For Each xmlElement In xmlTempRequest.childNodes
        xmlTempRequest.removeChild xmlElement
    Next
    
    Set xmlElement = xmlMtgeOneOffCostsForCalc.cloneNode(True)
    m_objXMLAssist.ChangeNodeName xmlElement, "MORTGAGEONEOFFCOSTLIST", "ONEOFFCOSTLIST"
    m_objXMLAssist.ChangeNodeName xmlElement, "MORTGAGEONEOFFCOST", "ONEOFFCOST"

    Set xmlOneOffCosts = xmlElement.cloneNode(True)

    Dim intComponentNo As Integer
    
    ' Add application date into the request
    Set xmlRootElement = xmlOut.createElement("APPLICATION")
    xmlTempRequest.appendChild xmlRootElement
    Set xmlElement = xmlOut.createElement("APPLICATIONDATE")
    xmlElement.Text = CStr(Date)
    xmlRootElement.appendChild xmlElement
        
    ' Set up requests to call the calcs engine and save the results
    Set xmlCalcsRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlIllusRequest = xmlCalcsRequest.cloneNode(True)
    
    'BMIDS774 Need to get the ApplicationFactFind EstimatedCompletionDate
    Set xmlInputElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlTableElem = vxmlRequest.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlInputElem.appendChild xmlTableElem
    
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONNUMBER")
    xmlAFFElement.Text = xmlAppNo.Text
    xmlTableElem.appendChild xmlAFFElement
    
    Set xmlAFFElement = xmlApplClassDefDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlAFFElement.Text = xmlAppFFNo.Text
    xmlTableElem.appendChild xmlAFFElement
    
    Set xmlApplClassDefDoc = objIApplClassDef.LoadApplicationFactFindData()
    Set xmlAFFData = objIDOAssist.GetData(xmlInputElem, xmlApplClassDefDoc)
    
    Set xmlCompDateElem = xmlAFFData.selectSingleNode("ESTIMATEDCOMPLETIONDATE")
    
    Set xmlLoanCompListNode = xmlOut.createElement("LOANCOMPONENTLIST")
    xmlTempRequest.appendChild xmlLoanCompListNode
    
    ' Recalculate each component and save the results
    For intComponentNo = 0 To xmlComponentList.length - 1
        
        Set xmlComponent = xmlComponentList.Item(intComponentNo).cloneNode(True)
            
        If Not xmlCompDateElem Is Nothing Then
            'BMIDS774 If EstimatedCompletionDate is >= Sys Date, use it
            'otherwise don't create it so we get the alpha+ default.
            If DateDiff("d", Now, CSafeDate(xmlCompDateElem.Text)) >= 0 Then
                Set xmlElement = xmlOut.createElement("COMPLETIONDATE")
                xmlElement.Text = xmlCompDateElem.Text
                xmlComponent.appendChild xmlElement
            End If
        End If
        
        strValue = m_objXMLAssist.GetNodeValue(xmlComponent, ".//LOANCOMPONENTSEQUENCENUMBER")
        
        If strValue = strMainLoanNo Then
            
            Dim lngTotalLoanCompAmt As Long
    
            Set xmlElement = m_objXMLAssist.GetNode(xmlComponent, ".//TOTALLOANCOMPONENTAMOUNT")
            lngTotalLoanCompAmt = objConvAssist.CSafeLng(xmlElement.Text) + lngCostsToBeAdded
            xmlElement.Text = CStr(lngTotalLoanCompAmt)
            
            xmlComponent.appendChild xmlOneOffCosts
           
        End If
                      
        ' Add the product info for the component into the request
        Set xmlMortgageProduct = m_objXMLAssist.GetNode(xmlComponent, ".//MORTGAGEPRODUCTDETAILS")
        
        xmlComponent.appendChild xmlMortgageProduct
        
        ' Add the component into the request
        xmlLoanCompListNode.appendChild xmlComponent
                     
    Next
    
    'If a Drawdown amount has been passed in we calculate our monthly costs less drawdown
    Dim strDrawDown As String
    
    'Only Calculate draw down costs if the drawdown exists and is > 0
    Set xmlDrawDown = xmlAppDetails.selectSingleNode("//DRAWDOWN")
    If Not xmlDrawDown Is Nothing Then
        If CSafeDbl(xmlDrawDown.Text) > 0 Then
            Dim xmlDrawDownDoc As FreeThreadedDOMDocument40
            Set xmlDrawDownDoc = New FreeThreadedDOMDocument40
            xmlDrawDownDoc.validateOnParse = False
            xmlDrawDownDoc.setProperty "NewParser", True
            xmlTempRequest.firstChild.appendChild xmlDrawDown.cloneNode(True)
            strDrawDown = objMortSubQuoteBO.CalculateDrawDownCosts(xmlTempRequest.xml)
            Set xmlDrawDownDoc = m_objXMLAssist.load(strDrawDown, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDrawDownDoc.documentElement, xmlResponseElem, True
        Else
            strDrawDown = vbNullString
        End If
    Else
        strDrawDown = ""
    End If
     
' Stage 13:  Calculate and save the loan component costs.
     
    'DPF:  If we have calculated drawdown costs we pass in a slightly amended request
    
    If Len(strDrawDown) = 0 Then
        strResponse = objMortSubQuoteBO.RunCompositeRateCalcAndSaveResults(xmlTempRequest.xml)
    Else
        Set xmlTempRequest = m_objXMLAssist.GetMandatoryNode(xmlDrawDownDoc.documentElement, "REQUEST").cloneNode(True)
        strResponse = objMortSubQuoteBO.RunCompositeRateCalcAndSaveResults(xmlTempRequest.xml)
    End If

    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    Set xmlTempResponseDoc = Nothing
    
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    For Each xmlElement In xmlTempRequest.childNodes
        xmlTempRequest.removeChild xmlElement
    Next
    
' Stage 14:  Update MortgageSubQuote Record
    
    'If we have calculated a monthly cost less drawdown append it to response to GUI
    If Len(strDrawDown) > 0 Then
        strValue = m_objXMLAssist.GetElementText(xmlDrawDownDoc, "//REQUEST/MONTHLYCOSTLESSDRAWDOWN")
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//MONTHLYCOSTLESSDRAWDOWN")
        xmlElement.Text = Format$(strValue, "0.00")
    End If
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALGROSSMONTHLYCOST")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//TOTALGROSSMONTHLYCOST")
    xmlElement.Text = Format$(strValue, "0.00")
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALNETMONTHLYCOST")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//TOTALNETMONTHLYCOST")
    xmlElement.Text = Format$(strValue, "0.00")
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALACCRUEDINTEREST")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//TOTALACCRUEDINTEREST")
    xmlElement.Text = Format$(strValue, "0.00")
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//AMOUNTPERUNITBORROWED")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//AMOUNTPERUNITBORROWED")
    xmlElement.Text = strValue
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//APR")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//APR")
    xmlElement.Text = strValue
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALAMOUNTPAYABLE")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//TOTALAMOUNTPAYABLE")
    xmlElement.Text = strValue
    
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALMORTGAGEPAYMENTS")
    Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeSubQuote, ".//TOTALMORTGAGEPAYMENTS")
    xmlElement.Text = strValue
    
    xmlTempRequest.appendChild xmlMtgeSubQuote.cloneNode(True)
    strResponse = objMortSubQuoteBO.Update(xmlTempRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    ' Get the recalculated loan component data
    strResponse = objMortSubQuoteBO.FindLoanComponentList(xmlTempRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    Set xmlComponentListElement = m_objXMLAssist.GetNode(xmlTempResponseDoc, ".//LOANCOMPONENTLIST")
    
'Stage15: Recalculate total Quotation cost.
  
    ' Get the quotation summary to update the quotation record with total cost
    
    xmlTempRequest.removeChild xmlTempRequest.firstChild
    Set xmlRootElement = xmlOut.createElement("QUOTATION")
    xmlTempRequest.appendChild xmlRootElement
    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
    'xmlRootElement.appendChild xmlLifeSubQuoteNo.cloneNode(True)
    xmlRootElement.appendChild xmlCustomerList
    xmlRootElement.appendChild xmlComponentListElement
    
    Set xmlElement = m_objXMLAssist.GetNode(xmlRootElement, ".//LOANCOMPONENTLIST")
    xmlRootElement.removeChild xmlElement
    Set xmlElement = xmlOut.createElement("ACTIVEQUOTENUMBER")
    xmlElement.Text = strAcceptedQuoteNo
    xmlRootElement.appendChild xmlElement
    
    strResponse = objQuotationBO.GetQuotationDetails(xmlTempRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    ' Calculate total cost
    Dim dblTotalQuotationCost As Double
       
    strValue = m_objXMLAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALNETMONTHLYCOST")
    dblTotalQuotationCost = objConvAssist.CSafeDbl(strValue)
    strValue = m_objXMLAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALLIFEMONTHLYCOST", False, False)
    dblTotalQuotationCost = dblTotalQuotationCost + objConvAssist.CSafeDbl(strValue)
    strValue = m_objXMLAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALBCMONTHLYCOST", False, False)
    dblTotalQuotationCost = dblTotalQuotationCost + objConvAssist.CSafeDbl(strValue)
    strValue = m_objXMLAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALPPMONTHLYCOST", False, False)
    dblTotalQuotationCost = dblTotalQuotationCost + objConvAssist.CSafeDbl(strValue)
    
    Set xmlElement = xmlOut.createElement("TOTALQUOTATIONCOST")
    xmlElement.Text = Format$(dblTotalQuotationCost, "0.00")
    xmlRootElement.appendChild xmlElement
  
'Stage16: Update Quotation with new total cost
    
    m_objXMLAssist.ChangeNodeName xmlTempRequest, "ACTIVEQUOTENUMBER", "QUOTATIONNUMBER"
    strResponse = objQuotationBO.Update(xmlTempRequest.xml)
    Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True

Warning7020Exit:

    Set IApplicationQuoteTxBO_RemodelQuotation = xmlResponseElem
    
    m_objContext.SetComplete
    
RemodelQuotationExit:
    
    Set xmlRequestElement = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlRootElement = Nothing
    Set xmlAppNo = Nothing
    Set xmlAppFFNo = Nothing
    Set xmlMortSubQuoteNo = Nothing
    'Set xmlLifeSubQuoteNo = Nothing
    Set xmlCost = Nothing
    Set xmlMtgeCostList = Nothing
    Set xmlMtgeCost = Nothing
    Set xmlMtgeSubQuote = Nothing
    Set xmlComponentList = Nothing
    Set xmlComponent = Nothing
    Set xmlMortgageProduct = Nothing
    Set xmlOneOffCosts = Nothing
    Set xmlCalcsRequest = Nothing
    Set xmlIllusRequest = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomersDoc = Nothing
    Set xmlMtgeOneOffCostsDoc = Nothing
    Set xmlAppDetails = Nothing
    Set xmlComboDoc = Nothing
    Set objApplicationBO = Nothing
    Set objAppManagerBO = Nothing
    Set objQuotationBO = Nothing
    Set objIApplicationQuoteBO = Nothing
    Set objMortSubQuoteBO = Nothing
    Set objComboDO = Nothing
    Set objConvAssist = Nothing
    Set objApplicationBO = Nothing
    
    Set xmlValidationType = Nothing
    
    Set xmlLoanCompListNode = Nothing

    Set xmlDrawDownDoc = Nothing
    Set xmlAFFData = Nothing
    Set xmlInputElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlAFFElement = Nothing
    Set xmlCompDateElem = Nothing
    Set objIApplClassDef = Nothing
    Set xmlApplClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlDrawDown = Nothing
    
    ' PSC 15/03/2007 EP2_1956 - Start
    Set xmlCopyFeeRequest = Nothing
    Set xmlMortgageIntroducerFee = Nothing
    ' PSC 15/03/2007 EP2_1956 - End

    
    Exit Function

RemodelQuotationVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_RemodelQuotation = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RemodelQuotationExit

End Function
' BMIDS977 Private Function IApplicationQuoteTxBO_UpdateQuotationForAddedCosts removed

'Private Function IApplicationQuoteTxBO_UpdateQuotationForAddedCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:  Remodels the quotation whenever a new application fee type (converted
''               to a one-off cost) is added during the payment processing if the fee is to
''               be added to the loan or has a type which affects the APR calculation.
''               The method is called from Payment Processing.
'' pass:         vxmlRequest  xml Request element
'' return:       xml Response Node
''------------------------------------------------------------------------------------------
'
'On Error GoTo UpdateQuotationForAddedCostsVbErr
'
'    Const strFunctionName As String = "IApplicationQuoteTxBO_UpdateQuotationForAddedCosts"
'
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
'
'    Dim xmlRequestElement As IXMLDOMElement
'    Dim xmlElement As IXMLDOMElement
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlTempRequest As IXMLDOMElement
'    Dim xmlRootElement As IXMLDOMElement
'    Dim xmlTmpNode As IXMLDOMNode 'JD BMIDS866
'
'    ' Request Nodes
'    Dim xmlAppNo As IXMLDOMElement
'    Dim xmlAppFFNo As IXMLDOMElement
'    Dim xmlMSQNo As IXMLDOMElement
'
'    ' Stage 1 - Get Loan Composition details
'    Dim xmlAppDetails As IXMLDOMElement
'    Dim xmlMortSubQuote As IXMLDOMElement
'    Dim xmlComponentList As IXMLDOMNodeList
'    Dim xmlComponent As IXMLDOMElement
'    Dim xmlMortgageProduct As IXMLDOMElement
'
'    ' Stage 4 - Calculate total incentives
'    Dim xmlMtgeOneOffCostsDoc As FreeThreadedDOMDocument40
'    Dim xmlMtgeCostList As IXMLDOMNodeList
'    Dim xmlMtgeCost As IXMLDOMElement
'
'    ' Stage 5 - GetOneOffCosts
'    Dim xmlOneOffCostsDoc As FreeThreadedDOMDocument40
'    Dim xmlOneOffCosts As IXMLDOMElement
'    Dim xmlCostList As IXMLDOMNodeList
'    Dim xmlCost As IXMLDOMElement
'    Dim xmlTempNode As IXMLDOMNode
'
'    ' BMIDS986
'    Dim objAppProcBO As Object
'
'    Dim dblROTPurchasePrice     As Double
'    Dim dblPostWorksValuation   As Double
'    Dim dblPresentValuation     As Double
'    Dim dblAppPurchasePriceOrEstimatedValue  As Double
'    Dim lngRetentionRoads       As Long
'    Dim lngRetentionWorks       As Long
'    Dim strPurchasePrice        As String
'    Dim lngErrorNo              As Long
'    Dim strApplicationType      As String
'
'    Dim xmlEleAppData As IXMLDOMElement
'    Dim xmlLatestValAmtsReqDoc  As New FreeThreadedDOMDocument40
'    Dim xmlLatestValAmtsReq     As IXMLDOMElement
'    Dim xmlLatestValAmtsReqElem As IXMLDOMElement
'    Dim xmlLatestValAmtsRespDoc As New FreeThreadedDOMDocument40
'    Dim xmlTempElement          As IXMLDOMNode
'    Dim xmlTypeOfAppNode        As IXMLDOMNode
'
'    ' BMIDS986 End
'
'
'    ' Stages 6 & 9 Combos
'    Dim xmlComboDoc As FreeThreadedDOMDocument40
'
'    'Stage 8 - Get MIG Cost
'    Dim xmlMIGCost As IXMLDOMElement
'
'    ' Stage 11 - Get completion date
'    Dim xmlROTRequest As IXMLDOMElement
'    Dim xmlROTRoot As IXMLDOMElement
'    Dim xmlCompDate As IXMLDOMAttribute
'    Dim xmlROTResponse As FreeThreadedDOMDocument40
'
'    ' Stage 11 - Set up loan components
'    Dim xmlLoanCompListNode As IXMLDOMNode
'    Dim xmlDrawDownDoc As FreeThreadedDOMDocument40
'
'    ' Business Objects
'    Dim objApplicationBO As ApplicationBO
'    Dim objAppManagerBO As ApplicationManagerBO
'    Dim objQuotationBO As QuotationBO
'    Dim objIApplicationQuoteBO As IApplicationQuoteBO
'    Dim objCustomerFinancialBO As CustomerFinancialBO
'    Dim objMortSubQuoteBO As MortgageSubQuoteBO
'    Dim objChargesBO As ChargesBO
'    Dim objMigBO As MIGBO
'    Dim objROTBO As Object
'
'    ' Other Objects
'    Dim objComboDO As ComboDO
'    Dim objConvAssist As ConvertAssist
'
'    ' Response String from Stage 1
'    Dim strResponse As String
'
'    ' General purpose string
'    Dim strValue As String
'
'    ' Working figures
'    'Dim lngTotalLoanAmt As Long
'    Dim lngTotalIncentives As Long
'    Dim lngManualIncentives As Long     'BMIDS881
'    Dim lngIncentives As Long           'BMIDS881
'    Dim lngTotalCharges As Long
'    Dim strAdded As String              'BMIDS881
'    Dim lngChargesAddedToLoan As Long   'BMIDS881
'    Dim lngTotalInitCharges As Long
'    Dim lngTotalNetCharges As Long
'    Dim lngSealingFee As Long
'    Dim lngReleaseFee As Long
'    Dim lngDeposit As Long
'
'    Dim xmlDrawDown As IXMLDOMNode  'BMIDS793 GHun
'
'    Dim strAmtReq As String 'JD BMIDS866
'
'    'BMIDS866 GHun
'    Dim xmlTempResponse As IXMLDOMNode
'    Dim objICostModelAssistDO As ICostModelAssistDO
'    Dim xmlSharePercentage As IXMLDOMNode
'    Dim xmlStaffMember As IXMLDOMNode
'    Dim objApplicationManagerBO As ApplicationManagerBO
'    'BMIDS866 End
'
'    ' CODE
'    ' Set up Response Node
'
'    Set objConvAssist = New ConvertAssist
'    Set xmlOut = New FreeThreadedDOMDocument40
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    'Preliminary checks
'
'    ' Check if the element passed in is the one with the table name else look further
'    ' below the node that is passed in
'    If vxmlRequest.nodeName = cstrAPPLICATION_NAME Then
'        Set xmlRequestElement = vxmlRequest
'    Else
'        Set xmlRequestElement = vxmlRequest.getElementsByTagName(cstrAPPLICATION_NAME).Item(0)
'    End If
'    If xmlRequestElement Is Nothing Then
'        m_objErrAssist.ThrowError TypeName(Me), _
'                                  strFunctionName, _
'                                  oeMissingPrimaryTag, _
'                                  cstrAPPLICATION_NAME & " tag not found"
'    End If
'
'    ' Set up Nodes for the request node, Application Number,
'    '  App Fact Find Number and Mortgage SubQuote Number
'
'    Set xmlAppNo = m_objXmlAssist.GetNode(xmlRequestElement, "APPLICATIONNUMBER")
'    Set xmlAppFFNo = m_objXmlAssist.GetNode(xmlRequestElement, "APPLICATIONFACTFINDNUMBER")
'
'    'BMIDS854 GHun The MortgageSubQuoteNumber passed in comes from context and can sometimes
'    ' be out of date. It is safer to get it from the quotation instead - see below
'    'Set xmlMSQNo = m_objXmlAssist.GetNode(xmlRequestElement, "MORTGAGESUBQUOTENUMBER")
'    'BMIDS854 End
'
'    ' Get the accepted quote  passed in
'    Dim strAcceptedQuoteNo As String
'
'    strAcceptedQuoteNo = m_objXmlAssist.GetNodeValue(xmlRequestElement, "ACCEPTEDQUOTENUMBER", False, False)
'
'' Stage 1 : Get the accepted quotation data
'
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlRootElement = xmlOut.createElement("QUOTATION")
'    xmlTempRequest.appendChild xmlRootElement
'    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
'    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
'    Set xmlElement = xmlOut.createElement("QUOTATIONNUMBER")
'    xmlElement.Text = strAcceptedQuoteNo
'    xmlRootElement.appendChild xmlElement
'    Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
'    strResponse = objQuotationBO.GetData(xmlTempRequest.xml)
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'
'    'BMIDS854 GHun
'    Set xmlMSQNo = m_objXmlAssist.GetMandatoryNode(xmlTempResponseDoc, ".//MORTGAGESUBQUOTENUMBER")
'    'BMIDS854 End
'
'' Stage 2 : Get loan composition details
'
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlRootElement = xmlOut.createElement("LOANCOMPOSITION")
'    xmlTempRequest.appendChild xmlRootElement
'    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
'    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
'    xmlRootElement.appendChild xmlMSQNo.cloneNode(True)
'    Set objIApplicationQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")
'    Set xmlAppDetails = objIApplicationQuoteBO.GetLoanCompositionDetails(xmlTempRequest)
'    m_objErrAssist.CheckResponse xmlAppDetails, xmlResponseElem, True
'    Set xmlMortSubQuote = m_objXmlAssist.GetNode(xmlAppDetails, ".//MORTGAGESUBQUOTE")
'
''Stage 3 : Get the total of Loans Not Redeemed
'
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlRootElement = xmlOut.createElement("MORTGAGEACCOUNT")
'    xmlTempRequest.appendChild xmlRootElement
'    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
'    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
'    Set objCustomerFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
'    strResponse = objCustomerFinancialBO.GetMtgAcctDataForFurtherAdvance(xmlTempRequest.xml)
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'    Dim strLoansNotRedeemed As String
'    Dim strTotalOriginalMIGAmount As String
'    Dim strOriginalMonthlyRepayment As String
'    strLoansNotRedeemed = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALOUTSTANDINGBALANCE", True, False)
'    strTotalOriginalMIGAmount = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALORIGINALMIGAMOUNT", True, False)
'    strOriginalMonthlyRepayment = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALMONTHLYREPAYMENT", True, False)
'    If Len(strLoansNotRedeemed) = 0 Then
'        strLoansNotRedeemed = "0"
'    End If
'
''Stage 4 : Calculate Total Incentives
'
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlRootElement = xmlOut.createElement("ONEOFFCOST")
'    xmlTempRequest.appendChild xmlRootElement
'    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
'    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
'    xmlRootElement.appendChild xmlMSQNo.cloneNode(True)
'    Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
'    strResponse = objMortSubQuoteBO.GetOneOffCostsDetails(xmlTempRequest.xml)
'    Set xmlMtgeOneOffCostsDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlMtgeOneOffCostsDoc.documentElement, xmlResponseElem, True
'    strValue = m_objXmlAssist.GetNodeValue(xmlMtgeOneOffCostsDoc, ".//TOTALINCENTIVES")
'    lngTotalIncentives = objConvAssist.CSafeLng(strValue)
'
'    ' BMIDS881 Get Manual Incentives and, if they exist, use these in preference to TotalIncentives
'    strValue = m_objXmlAssist.GetNodeValue(xmlMtgeOneOffCostsDoc, ".//MANUALINCENTIVEAMOUNT", False)
'    lngManualIncentives = objConvAssist.CSafeLng(strValue)
'    If (lngManualIncentives > 0) Then
'        lngIncentives = lngManualIncentives
'    Else
'        lngIncentives = lngTotalIncentives
'    End If
'
'    'BMIDS866 GHun
'    'Get New Property Details
'    Set xmlTempNode = xmlOut.createElement("NEWPROPERTY")
'    xmlTempNode.appendChild xmlAppNo.cloneNode(True)
'    xmlTempNode.appendChild xmlAppFFNo.cloneNode(True)
'    Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
'    Set xmlTempResponse = objICostModelAssistDO.GetValuationTypeAndLocation(xmlTempNode)
'    If Not xmlTempResponse Is Nothing Then
'        Set xmlSharePercentage = xmlTempResponse.selectSingleNode("SHAREDOWNERSHIPDETAILS/SHAREDPERCENTAGE")
'    End If
'
'    'Is member of staff
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlTempNode = xmlOut.createElement("APPLICATION")
'    xmlTempRequest.appendChild xmlTempNode
'    xmlTempNode.appendChild xmlAppNo.cloneNode(True)
'    xmlTempNode.appendChild xmlAppFFNo.cloneNode(True)
'
'    Set objApplicationManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
'    strResponse = objApplicationManagerBO.IsStaffApplication(xmlTempRequest.xml)
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'    Set xmlStaffMember = xmlTempResponseDoc.selectSingleNode(".//MEMBEROFSTAFF")
'    'BMIDS866 End
'
''Stage 5 : Calculate One Off Costs
'
'    ' BMIDS986  Calculate the Purchase Price to use - looking at valuation
'
'    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
'
'    strApplicationType = m_objXmlAssist.GetNodeValue(xmlAppDetails, "./APPLICATIONFACTFIND/TYPEOFAPPLICATION", False, False)
'
'    Set xmlLatestValAmtsReq = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    xmlLatestValAmtsReq.setAttribute "OPERATION", "GETLATESTVALUATIONAMOUNTS"
'    xmlLatestValAmtsReq.setAttribute "ACTION", "GETLATESTVALUATIONAMOUNTS"
'
'    Set xmlLatestValAmtsReqElem = xmlLatestValAmtsReqDoc.createElement("VALUATIONREPVALUATION")
'    xmlLatestValAmtsReqElem.setAttribute "APPLICATIONNUMBER", xmlAppNo.Text
'    xmlLatestValAmtsReqElem.setAttribute "APPLICATIONFACTFINDNUMBER", xmlAppFFNo.Text
'
'    Set xmlEleAppData = m_objXmlAssist.GetNode(xmlAppDetails, ".//APPLICATIONDATE")
'
'    xmlLatestValAmtsReqElem.setAttribute "APPLICATIONDATE", xmlEleAppData.Text
'    xmlLatestValAmtsReq.appendChild xmlLatestValAmtsReqElem
'
'    Set objAppProcBO = m_objContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
'    strResponse = objAppProcBO.omAppProcRequest(xmlLatestValAmtsReq.xml)
'    Set xmlLatestValAmtsRespDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'    lngErrorNo = m_objErrAssist.CheckResponse(xmlLatestValAmtsRespDoc.documentElement, xmlResponseElem, False)
'    ' If not success or record not found re-raise the error
'    If lngErrorNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
'        m_objErrAssist.ReRaiseResponseError xmlLatestValAmtsRespDoc.documentElement
'    End If
'
'    Set xmlTempElement = xmlLatestValAmtsRespDoc.selectSingleNode(".//VALUATIONREPVALUATION")
'    If Not xmlTempElement Is Nothing Then
'
'        dblPostWorksValuation = m_objConvertAssist.CSafeDbl(m_objXmlAssist.GetAttributeValue( _
'                        xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "POSTWORKSVALUATION", False))
'        dblPresentValuation = m_objConvertAssist.CSafeDbl(m_objXmlAssist.GetAttributeValue( _
'                        xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "PRESENTVALUATION", False))
'        lngRetentionRoads = m_objConvertAssist.CSafeDbl(m_objXmlAssist.GetAttributeValue( _
'                        xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "RETENTIONROADS", False))
'        lngRetentionWorks = m_objConvertAssist.CSafeDbl(m_objXmlAssist.GetAttributeValue( _
'                        xmlLatestValAmtsRespDoc.documentElement, "VALUATIONREPVALUATION", "RETENTIONWORKS", False))
'    End If
'
'    dblAppPurchasePriceOrEstimatedValue = m_objXmlAssist.GetNodeValue(xmlAppDetails, ".//PURCHASEPRICEORESTIMATEDVALUE")
'
'    If (lngRetentionRoads > 0 Or lngRetentionWorks > 0) And dblPostWorksValuation > 0 Then
'        strPurchasePrice = CStr(dblPostWorksValuation)
'    Else
'        If objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "N") Then
'            If dblPresentValuation > 0 And dblPresentValuation < dblAppPurchasePriceOrEstimatedValue Then
'                strPurchasePrice = CStr(dblPresentValuation)
'            Else
'                strPurchasePrice = CStr(dblAppPurchasePriceOrEstimatedValue)
'            End If
'        Else
'            If dblPresentValuation > 0 Then
'                strPurchasePrice = CStr(dblPresentValuation)
'            Else
'                strPurchasePrice = CStr(dblAppPurchasePriceOrEstimatedValue)
'            End If
'        End If
'    End If
'
'    ' BMIDS986  End
'
'
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlRootElement = xmlOut.createElement("CHARGES")
'    xmlTempRequest.appendChild xmlRootElement
'    Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//LOANCOMPONENTLIST")
'    xmlRootElement.appendChild xmlElement.cloneNode(True)
'    Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//AMOUNTREQUESTED")
'    strAmtReq = xmlElement.Text
'    xmlRootElement.appendChild xmlElement.cloneNode(True)
'    Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//LTV")
'    xmlRootElement.appendChild xmlElement.cloneNode(True)
'    'JD BMIDS866 TotalLoanAmount is used in CalcMig so should be the AmountRequested amount
'    Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//TOTALLOANAMOUNT")
'    Set xmlTmpNode = xmlRootElement.appendChild(xmlElement.cloneNode(True))
'    xmlTmpNode.Text = strAmtReq
'    Set xmlElement = xmlOut.createElement("PURCHASEPRICE")
'    xmlRootElement.appendChild xmlElement
'    xmlElement.Text = strPurchasePrice                     ' BMIDS986
'    Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//TYPEOFAPPLICATION")
'    xmlRootElement.appendChild xmlElement.cloneNode(True)
'    'BMIDS 763 - Need Application Date for CalcCharges
'    Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//APPLICATIONDATE")
'    xmlRootElement.appendChild xmlElement.cloneNode(True)
'    'BMIDS 763 - End
'    Set xmlElement = xmlOut.createElement("LOANSNOTREDEEMED")
'    xmlElement.Text = strLoansNotRedeemed
'    xmlRootElement.appendChild xmlElement
'    Set xmlElement = xmlOut.createElement("TOTALORIGINALMIGAMOUNT")
'    xmlElement.Text = strTotalOriginalMIGAmount
'    xmlRootElement.appendChild xmlElement
'    Set xmlElement = xmlOut.createElement("CONTEXT")
'    xmlRootElement.appendChild xmlElement
'    xmlElement.Text = "Cost Modelling"
'
'    'BMIDS866 GHun
'    If xmlSharePercentage Is Nothing Then
'        Set xmlSharePercentage = xmlOut.createElement("SHAREDPERCENTAGE")
'        xmlSharePercentage.Text = "100"
'    Else
'        If Len(xmlSharePercentage.Text) = 0 Then
'            xmlSharePercentage.Text = "100"
'        End If
'    End If
'    xmlRootElement.appendChild xmlSharePercentage
'    xmlRootElement.appendChild xmlStaffMember
'    'BMIDS866 End
'
'    Set objChargesBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".ChargesBO")
'    strResponse = objChargesBO.CalcCharges(xmlTempRequest.xml)
'    Set xmlOneOffCostsDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlOneOffCostsDoc.documentElement, xmlResponseElem, True
'
'    Dim xmlOneOffCostsList As IXMLDOMNodeList
'    Set xmlOneOffCosts = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, ".//ONEOFFCOSTLIST")
'    Set xmlOneOffCostsList = xmlOneOffCostsDoc.selectNodes(".//ONEOFFCOST")
'
'
'    ' Stage 6 : Calculate total of costs to be added to loan
'
'    ' Get combo data for one off costs
'    Set xmlRootElement = xmlOut.createElement("LIST")
'    Set xmlElement = xmlOut.createElement("LISTNAME")
'    xmlElement.Text = "OneOffCost"
'    xmlRootElement.appendChild xmlElement
'    'Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")      'BMIDS986
'    strResponse = objComboDO.GetComboList(xmlRootElement.xml)
'    Set xmlComboDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'    ' Calculate one off cost to be added to the loan
'    'Create a new OneOffCost record for the Added cost
'    Dim xmlOneOffCostRequest As IXMLDOMNode
'    Dim xmlAddedCost As IXMLDOMNode
'    Dim strPattern As String
'    Dim strCostType As String
'    Dim strValidationType As String
'
'    Set xmlOneOffCostRequest = vxmlRequest.cloneNode(False)
'    Set xmlRootElement = xmlOut.createElement("MORTGAGESUBQUOTE")
'    xmlOneOffCostRequest.appendChild xmlRootElement
'    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
'    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
'    xmlRootElement.appendChild xmlMSQNo.cloneNode(True)
'    Set xmlAddedCost = vxmlRequest.selectSingleNode(".//ADDEDONEOFFCOST")
'    m_objXmlAssist.ChangeNodeName xmlAddedCost, "ADDEDONEOFFCOST", "ONEOFFCOST"
'
'    xmlRootElement.appendChild xmlAddedCost
'    Set xmlRootElement = xmlOut.createElement("IDENTIFIER")
'    'Need to put the ValidationType into the Identifier
'     strCostType = m_objXmlAssist.GetNodeValue(xmlAddedCost, ".//FEETYPE")
'     strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE[. != 'APR']"
'     strValidationType = m_objXmlAssist.GetNodeValue(xmlComboDoc, strPattern)
'    xmlRootElement.Text = strValidationType
'    xmlAddedCost.appendChild xmlRootElement
'
'    strResponse = objMortSubQuoteBO.CreateOneOffCosts(xmlOneOffCostRequest.xml)
'
'
'    Dim lngCostsToBeAdded As Long
'    'Dim strCostToBeAdded As String
'    lngCostsToBeAdded = 0
'     If Not xmlAddedCost Is Nothing Then
'        If (xmlAddedCost.selectSingleNode("ADDTOLOAN").Text = "1") Then
'            lngCostsToBeAdded = CSafeLng(xmlAddedCost.selectSingleNode("AMOUNT").Text)
'        End If
'    End If
'
'    'get a list of all mortgage costs that are to be included in the APR calcs
'
'
'    'Dim lngAmount As Long
'    'Dim xmlTestAPRCostNode As IXMLDOMNode
'    strPattern = ".//MORTGAGEONEOFFCOST"
'
'    Set xmlMtgeCostList = xmlMtgeOneOffCostsDoc.selectNodes(strPattern)
'    For Each xmlMtgeCost In xmlMtgeCostList
'
'        ' Convert the cost type to the validation type
'        strCostType = m_objXmlAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
'        strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE[. != 'APR']"
'        strValidationType = m_objXmlAssist.GetNodeValue(xmlComboDoc, strPattern)
'
'
'        ' Get the calculated one off cost with the same validation type and copy the parameters over
'        ' ******** Catch TTF
'
'            strPattern = ".//ONEOFFCOST[IDENTIFIER = '" & strValidationType & "']"
'            Set xmlCost = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, strPattern, False)
'            If Not (xmlCost Is Nothing) Then
'
'            'if it's adhoc, we need to create a new OneOffCostElement
'                If (m_objXmlAssist.GetNodeValue(xmlMtgeCost, "ADHOCIND", True, False) = "1") Then
'
'                    Set xmlCost = xmlOneOffCostsDoc.createElement("ONEOFFCOST")
'                    m_objXmlAssist.CopyNode "ADDTOLOAN", "ADDTOLOAN", xmlMtgeCost, xmlCost, False
'                    m_objXmlAssist.CopyNode "AMOUNT", "AMOUNT", xmlMtgeCost, xmlCost, False
'                    Set xmlTempNode = xmlOneOffCostsDoc.createElement("IDENTIFIER")
'                    xmlTempNode.Text = strValidationType
'                    xmlCost.appendChild xmlTempNode
'                    xmlOneOffCosts.appendChild xmlCost
'                Else
'                'if its a calculated cost, copy over the Add to Loan filed from the MortgageOneOff costs record
'                    m_objXmlAssist.CopyNode "ADDTOLOAN", "ADDTOLOAN", xmlMtgeCost, xmlCost, False
'
'                End If
'
'            End If
'    Next
'    'Now must add in the new added cost to the OneOffCostList
'    strValidationType = m_objXmlAssist.GetNode(xmlAddedCost, ".//IDENTIFIER").Text
'    strPattern = ".//ONEOFFCOST[IDENTIFIER = '" & strValidationType & "']"
'    Set xmlCost = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, strPattern, False)
'    If Not (xmlCost Is Nothing) Then
'        Set xmlCost = xmlOneOffCostsDoc.createElement("ONEOFFCOST")
'        m_objXmlAssist.CopyNode "ADDTOLOAN", "ADDTOLOAN", xmlAddedCost, xmlCost, False
'        m_objXmlAssist.CopyNode "AMOUNT", "AMOUNT", xmlAddedCost, xmlCost, False
'        Set xmlTempNode = xmlOneOffCostsDoc.createElement("IDENTIFIER")
'        xmlTempNode.Text = strValidationType
'        xmlCost.appendChild xmlTempNode
'        xmlOneOffCosts.appendChild xmlCost
'
'
'    End If
'
''Stage 7 : Calculate Total Charges
'
'    strPattern = ".//ONEOFFCOST[IDENTIFIER != 'TID' and IDENTIFIER != 'CBC' and IDENTIFIER != 'CPP']"
'    Set xmlCostList = xmlOneOffCostsDoc.selectNodes(strPattern)
'    lngTotalCharges = 0
'    lngChargesAddedToLoan = 0      ' BMIDS881
'    For Each xmlCost In xmlCostList
'
'         lngTotalCharges = lngTotalCharges + objConvAssist.CSafeLng(m_objXmlAssist.GetNodeValue(xmlCost, "AMOUNT"))
'
'         'BMIDS881  Add up any charges which are added to loan
'         If (m_objXmlAssist.GetNodeValue(xmlCost, "ADDTOLOAN", False) = "1") Then
'            lngChargesAddedToLoan = lngChargesAddedToLoan + objConvAssist.CSafeLng(m_objXmlAssist.GetNodeValue(xmlCost, "AMOUNT"))
'         End If
'
'    Next
'
'    strValue = m_objXmlAssist.GetNodeValue(xmlOneOffCostsDoc, ".//ONEOFFCOST[IDENTIFIER = 'SEA']/AMOUNT")
'    lngSealingFee = objConvAssist.CSafeLng(strValue)
'    strValue = m_objXmlAssist.GetNodeValue(xmlOneOffCostsDoc, ".//ONEOFFCOST[IDENTIFIER = 'DEE']/AMOUNT")
'    lngReleaseFee = objConvAssist.CSafeLng(strValue)
'    strValue = m_objXmlAssist.GetNodeValue(xmlMortSubQuote, "DEPOSIT")
'    lngDeposit = objConvAssist.CSafeLng(strValue)
'    lngTotalInitCharges = lngTotalCharges - lngSealingFee - lngReleaseFee
'    lngTotalNetCharges = lngTotalCharges + lngDeposit - lngIncentives - lngChargesAddedToLoan  ' BMIDS881
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, "TOTALCHARGES")
'    xmlElement.Text = CStr(lngTotalCharges)
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, "TOTALINITIALCHARGES")
'    xmlElement.Text = CStr(lngTotalInitCharges)
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, "TOTALNETCHARGES")
'    xmlElement.Text = CStr(lngTotalNetCharges)
'
''Stage 8 : Recalculate MIG if required
'
'''    If lngCostsToBeAdded > 0 Then
'''        ' Set up MIG request
'''        Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'''        Set xmlRootElement = xmlOut.createElement("MIG")
'''        xmlTempRequest.appendChild xmlRootElement
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//MIGSTARTLTV")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//MINIMUMMIGPREMIUM")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//MIGTHRESHOLD")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//IGNOREMIGPREMIUMAMOUNT")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//ADDIPTTOMIGPREMIUM")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        strValue = m_objXmlAssist.GetNodeValue(xmlMortSubQuote, ".//TOTALLOANAMOUNT")
'''        lngTotalLoanAmt = objConvAssist.CSafeLng(strValue)
'''        lngTotalLoanAmt = lngTotalLoanAmt + lngCostsToBeAdded
'''        Set xmlElement = xmlOut.createElement("TOTALLOANAMOUNT")
'''        xmlElement.Text = CStr(lngTotalLoanAmt)
'''        xmlRootElement.appendChild xmlElement
'''
'''        Set xmlElement = xmlOut.createElement("PURCHASEPRICE")
'''        xmlElement.Text = m_objXmlAssist.GetNodeValue(xmlAppDetails, ".//PURCHASEPRICEORESTIMATEDVALUE")
'''        xmlRootElement.appendChild xmlElement
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//LTV")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//TYPEOFAPPLICATION")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = m_objXmlAssist.GetNode(xmlAppDetails, ".//MORTGAGELENDER/RATESET")
'''        xmlRootElement.appendChild xmlElement.cloneNode(True)
'''
'''        Set xmlElement = xmlOut.createElement("LOANSNOTREDEEMED")
'''        xmlElement.Text = strLoansNotRedeemed
'''        xmlRootElement.appendChild xmlElement
'''
'''        Set objMigBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MIGBO")
'''        strResponse = objMigBO.CalcMIG(xmlTempRequest.xml)
'''        Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'''        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'''
'''        Dim lngMigPremium As Long
'''        lngMigPremium = 0
'''
'''        ' Get the MIG one off cost
'''        Set xmlMIGCost = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, ".//ONEOFFCOST[IDENTIFIER = 'MIG']", False)
'''
'''        ' Get the existing MIG premium
'''        If Not xmlMIGCost Is Nothing Then
'''            strValue = m_objXmlAssist.GetNodeValue(xmlMIGCost, "AMOUNT")
'''            lngMigPremium = objConvAssist.CSafeLng(strValue)
'''        End If
'''
'''' Stage 9 : Calculate Revised Total Loan if MIG is to added
'''
'''        Dim strMIGValue As String
'''        strMIGValue = m_objXmlAssist.GetNodeValue(xmlComboDoc, ".//LISTENTRY[.//VALIDATIONTYPE = 'MIG']/VALUEID")
'''        strPattern = ".//MORTGAGEONEOFFCOST[MORTGAGEONEOFFCOSTTYPE = '" & strMIGValue & "' and ADDTOLOAN = '1']"
'''        Set xmlMIGCost = m_objXmlAssist.GetNode(xmlMtgeOneOffCostsDoc, strPattern, False)
'''
'''        ' Add MIG to costs to be added if it is to be added to the loan
'''        lngMigPremium = 0
'''        If Not xmlMIGCost Is Nothing Then
'''            Set xmlElement = m_objXmlAssist.GetNode(xmlMIGCost, ".//AMOUNT")
'''            lngMigPremium = objConvAssist.CSafeLng(xmlElement.Text)
'''        End If
'''
'''        ' Update the TID one off cost
'''        Set xmlCost = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, ".//ONEOFFCOST[IDENTIFIER = 'TID']/AMOUNT", False)
'''
'''        If Not xmlCost Is Nothing Then
'''            Set xmlElement = m_objXmlAssist.GetNode(xmlTempResponseDoc, ".//TOTALINDEMNITY")
'''            xmlCost.Text = xmlElement.Text
'''        End If
'''
'''        If Not xmlMtgeCost Is Nothing Then
'''            lngCostsToBeAdded = lngCostsToBeAdded + lngMigPremium
'''
'''            ' Set new MIG cost to be added to the loan as well
'''            Set xmlCost = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, ".//ONEOFFCOST[IDENTIFIER = 'MIG']")
'''            Set xmlElement = xmlOut.createElement("ADDTOLOAN")
'''            xmlElement.Text = "1"
'''            xmlCost.appendChild xmlElement
'''        End If
'''    End If
'
'' Stage 10 : Save Recalculated costs to the MortgageOneOffCosts table
'
'''    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'''    Set xmlElement = xmlMortSubQuote.cloneNode(True)
'''    xmlTempRequest.appendChild xmlElement
'''    xmlTempRequest.appendChild xmlOneOffCostsDoc.documentElement.cloneNode(True)
'''    strResponse = objMortSubQuoteBO.UpdateOneOffCosts(xmlTempRequest.xml)
'''    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'
'' Stage 11 : Get completion date and set up loan components
'
'''    Set xmlROTRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'''    xmlROTRequest.setAttribute "OPERATION", "GetReportOnTitleData"
'''    Set xmlROTRoot = xmlOut.createElement("REPORTONTITLE")
'''    xmlROTRequest.appendChild xmlROTRoot
'''    xmlROTRoot.setAttribute "APPLICATIONNUMBER", xmlAppNo.Text
'''    xmlROTRoot.setAttribute "APPLICATIONFACTFINDNUMBER", xmlAppFFNo.Text
'''    Set objROTBO = m_objContext.CreateInstance(gstrREPORTONTITLE_COMPONENT & ".omROTBO")
'''    strResponse = objROTBO.OmRotRequest(xmlROTRequest.xml)
'''    Set xmlROTResponse = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'''    Dim lngErrorNo As Long
'''    lngErrorNo = m_objErrAssist.CheckResponse(xmlROTResponse.documentElement, xmlResponseElem, False)
'''
'''    ' If not success or record not found re-raise the error
'''    If lngErrorNo <> 0 And m_objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
'''        m_objErrAssist.ReRaiseResponseError xmlROTResponse.documentElement
'''    End If
'''    Set xmlCompDate = xmlROTResponse.selectSingleNode("RESPONSE/REPORTONTITLE/@COMPLETIONDATE")
'
'    Set xmlTempRequest = vxmlRequest.cloneNode(False)
'    Set xmlLoanCompListNode = xmlOut.createElement("LOANCOMPONENTLIST")
'    xmlTempRequest.appendChild xmlLoanCompListNode
'
'    ' Find main loan component
'    Dim intMainLoanNo As Integer
'    Dim intCurrentNo As Integer
'    Dim strMainLoanNo As String
'    intMainLoanNo = 9999
'    Set xmlComponentList = xmlAppDetails.selectNodes(".//LOANCOMPONENT[PORTEDLOAN != '1']")
'    For Each xmlComponent In xmlComponentList
'        strValue = m_objXmlAssist.GetNodeValue(xmlComponent, ".//LOANCOMPONENTSEQUENCENUMBER")
'        intCurrentNo = objConvAssist.CSafeInt(strValue)
'        If intCurrentNo < intMainLoanNo Then
'            intMainLoanNo = intCurrentNo
'        End If
'    Next
'    strMainLoanNo = CStr(intMainLoanNo)
'
'    ' Recalculate each component and save the results
'    Set xmlElement = m_objXmlAssist.GetNode(xmlOneOffCostsDoc, ".//ONEOFFCOSTLIST")
'    Set xmlOneOffCosts = xmlElement.cloneNode(False)
'
'    Dim intComponentNo As Integer
'
'  ' Go thru the cost list adding costs
'    For Each xmlCost In xmlOneOffCostsList
'
'          xmlOneOffCosts.appendChild xmlCost
'    Next
'    For intComponentNo = 0 To xmlComponentList.length - 1
'        Set xmlComponent = xmlComponentList.Item(intComponentNo).cloneNode(True)
'
'''        If Not xmlCompDate Is Nothing Then
'''            Set xmlElement = xmlOut.createElement("COMPLETIONDATE")
'''            xmlElement.Text = xmlCompDate.Value
'''            xmlComponent.appendChild xmlElement
'''        End If
'
'        strValue = m_objXmlAssist.GetNodeValue(xmlComponent, ".//LOANCOMPONENTSEQUENCENUMBER")
'        If strValue = strMainLoanNo Then
'
'            Dim lngTotalLoanCompAmt As Long
'            Set xmlElement = m_objXmlAssist.GetNode(xmlComponent, ".//TOTALLOANCOMPONENTAMOUNT")
'            lngTotalLoanCompAmt = objConvAssist.CSafeLng(xmlElement.Text) + lngCostsToBeAdded
'            xmlElement.Text = CStr(lngTotalLoanCompAmt)
'            xmlComponent.appendChild xmlOneOffCosts
'        End If
'
'        ' Add the product info for the component into the request
'        Set xmlMortgageProduct = m_objXmlAssist.GetNode(xmlComponent, ".//MORTGAGEPRODUCTDETAILS")
'        xmlComponent.appendChild xmlMortgageProduct
'
'        ' Add the component into the request
'        xmlLoanCompListNode.appendChild xmlComponent
'   Next
'
'    'If a Drawdown amount has been passed in calculate monthly costs less drawdown
'    Dim strDrawDown As String
'    strDrawDown = ""
'    'BMIDS793 GHun Only
'    Set xmlDrawDown = xmlAppDetails.selectSingleNode("//DRAWDOWN")
'    If Not xmlDrawDown Is Nothing Then
'        If CSafeDbl(xmlDrawDown.Text) > 0 Then
'    'BMIDS793 End
'            Set xmlDrawDownDoc = New FreeThreadedDOMDocument40
'            xmlTempRequest.firstChild.appendChild xmlDrawDown.cloneNode(True)
'            strDrawDown = objMortSubQuoteBO.CalculateDrawDownCosts(xmlTempRequest.xml)
'            Set xmlDrawDownDoc = m_objXmlAssist.load(strDrawDown, TypeName(Me), strFunctionName)
'            m_objErrAssist.CheckResponse xmlDrawDownDoc.documentElement, xmlResponseElem, True
'    'BMIDS793 GHun
'        Else
'            strDrawDown = vbNullString
'        End If
'    Else
'        strDrawDown = vbNullString
'    End If
'    'BMIDS793 End
'
''Stage 12 : Do Mortgage Calculation
'
'    If Len(strDrawDown) = 0 Then
'        strResponse = objMortSubQuoteBO.RunCompositeRateCalcAndSaveResults(xmlTempRequest.xml)
'    Else
'        Set xmlTempRequest = m_objXmlAssist.GetMandatoryNode(xmlDrawDownDoc.documentElement, "REQUEST").cloneNode(True)
'        strResponse = objMortSubQuoteBO.RunCompositeRateCalcAndSaveResults(xmlTempRequest.xml)
'    End If
'
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'    For Each xmlElement In xmlTempRequest.childNodes
'        xmlTempRequest.removeChild xmlElement
'    Next
'
'    'If we have calculated a monthly cost less drawdown append it to response to GUI
'    If Len(strDrawDown) > 0 Then
'        strValue = m_objXmlAssist.GetElementText(xmlDrawDownDoc, "//REQUEST/MONTHLYCOSTLESSDRAWDOWN")
'        Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//MONTHLYCOSTLESSDRAWDOWN")
'        xmlElement.Text = Format$(strValue, "0.00")
'    End If
'
'' Stage 13 : Update Mortgage Sub Quote record
'
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALGROSSMONTHLYCOST")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//TOTALGROSSMONTHLYCOST")
'    xmlElement.Text = Format$(strValue, "0.00")
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALNETMONTHLYCOST")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//TOTALNETMONTHLYCOST")
'    xmlElement.Text = Format$(strValue, "0.00")
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALACCRUEDINTEREST")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//TOTALACCRUEDINTEREST")
'    xmlElement.Text = Format$(strValue, "0.00")
'
'    'BMIDS793 GHun
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//AMOUNTPERUNITBORROWED")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//AMOUNTPERUNITBORROWED")
'    xmlElement.Text = strValue
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//APR")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//APR")
'    xmlElement.Text = strValue
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALAMOUNTPAYABLE")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//TOTALAMOUNTPAYABLE")
'    xmlElement.Text = strValue
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlTempResponseDoc, ".//TOTALMORTGAGEPAYMENTS")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//TOTALMORTGAGEPAYMENTS")
'    xmlElement.Text = strValue
'    'BMIDS793 End
'
'    'BMIDS767 GHun Save TotalLoanAmount
'    strValue = m_objXmlAssist.GetMandatoryElementText(xmlMortSubQuote, ".//AMOUNTREQUESTED")
'    Set xmlElement = m_objXmlAssist.GetNode(xmlMortSubQuote, ".//TOTALLOANAMOUNT")
'    xmlElement.Text = CSafeLng(xmlElement.Text) + lngCostsToBeAdded
'    'BMIDS767 End
'
'    xmlTempRequest.appendChild xmlMortSubQuote.cloneNode(True)
'    strResponse = objMortSubQuoteBO.Update(xmlTempRequest.xml)
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'
'' Stage 14 : Add total costs and update the accepted quotation
'
'    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlRootElement = xmlOut.createElement("QUOTATION")
'    xmlTempRequest.appendChild xmlRootElement
'    xmlRootElement.appendChild xmlAppNo.cloneNode(True)
'    xmlRootElement.appendChild xmlAppFFNo.cloneNode(True)
'    'Set xmlElement = m_objXmlAssist.GetNode(xmlRootElement, ".//LOANCOMPONENTLIST")
'    'xmlRootElement.removeChild xmlElement
'    Set xmlElement = xmlOut.createElement("ACTIVEQUOTENUMBER")
'    xmlElement.Text = strAcceptedQuoteNo
'    xmlRootElement.appendChild xmlElement
'    strResponse = objQuotationBO.GetQuotationDetails(xmlTempRequest.xml)
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'
'    ' Calculate total cost
'
'    Dim dblTotalQuotationCost As Double
'
'    strValue = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALNETMONTHLYCOST")
'    dblTotalQuotationCost = objConvAssist.CSafeDbl(strValue)
'    strValue = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALLIFEMONTHLYCOST", False, False)
'    dblTotalQuotationCost = dblTotalQuotationCost + objConvAssist.CSafeDbl(strValue)
'    strValue = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALBCMONTHLYCOST", False, False)
'    dblTotalQuotationCost = dblTotalQuotationCost + objConvAssist.CSafeDbl(strValue)
'    strValue = m_objXmlAssist.GetNodeValue(xmlTempResponseDoc, ".//TOTALPPMONTHLYCOST", False, False)
'    dblTotalQuotationCost = dblTotalQuotationCost + objConvAssist.CSafeDbl(strValue)
'    Set xmlElement = xmlOut.createElement("TOTALQUOTATIONCOST")
'    xmlElement.Text = Format$(dblTotalQuotationCost, "0.00")
'    xmlRootElement.appendChild xmlElement
'
'    m_objXmlAssist.ChangeNodeName xmlTempRequest, "ACTIVEQUOTENUMBER", "QUOTATIONNUMBER"
'    strResponse = objQuotationBO.Update(xmlTempRequest.xml)
'    Set xmlTempResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
'
'    Set IApplicationQuoteTxBO_UpdateQuotationForAddedCosts = xmlResponseElem
'    m_objContext.SetComplete
'
'UpdateQuotationForAddedCostsExit:
'
'    Set xmlOut = Nothing
'    Set xmlTempResponseDoc = Nothing
'
'    Set xmlRequestElement = Nothing
'    Set xmlElement = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlTempRequest = Nothing
'    Set xmlRootElement = Nothing
'
'    Set xmlAppNo = Nothing
'    Set xmlAppFFNo = Nothing
'    Set xmlMSQNo = Nothing
'
'    Set xmlAppDetails = Nothing
'    Set xmlMortSubQuote = Nothing
'    Set xmlComponentList = Nothing
'    Set xmlComponent = Nothing
'    Set xmlMortgageProduct = Nothing
'
'    Set xmlCostList = Nothing
'    Set xmlCost = Nothing
'
'    Set xmlMIGCost = Nothing
'
'    Set xmlMtgeOneOffCostsDoc = Nothing
'    Set xmlMtgeCostList = Nothing
'    Set xmlMtgeCost = Nothing
'
'    Set xmlOneOffCosts = Nothing
'    Set xmlOneOffCostsDoc = Nothing
'    Set xmlComboDoc = Nothing
'    Set xmlLoanCompListNode = Nothing
'    Set xmlROTRequest = Nothing
'    Set xmlROTRoot = Nothing
'    Set xmlCompDate = Nothing
'    Set xmlROTResponse = Nothing
'    Set xmlDrawDownDoc = Nothing
'
'    Set objApplicationBO = Nothing
'    Set objAppManagerBO = Nothing
'    Set objQuotationBO = Nothing
'    Set objIApplicationQuoteBO = Nothing
'    Set objCustomerFinancialBO = Nothing
'    Set objMortSubQuoteBO = Nothing
'    Set objChargesBO = Nothing
'    Set objMigBO = Nothing
'    Set objROTBO = Nothing
'
'    Set objComboDO = Nothing
'    Set objConvAssist = Nothing
'    Set xmlDrawDown = Nothing
'    'BMIDS866 GHun
'    Set xmlTempResponse = Nothing
'    Set objICostModelAssistDO = Nothing
'    Set xmlSharePercentage = Nothing
'    Set xmlStaffMember = Nothing
'    Set objApplicationManagerBO = Nothing
'    'BMIDS866 End
'
'    'BMIDS986
'    Set xmlEleAppData = Nothing
'    Set xmlLatestValAmtsReqDoc = Nothing
'    Set xmlLatestValAmtsReq = Nothing
'    Set xmlLatestValAmtsReqElem = Nothing
'    Set xmlLatestValAmtsRespDoc = Nothing
'    Set xmlTempElement = Nothing
'    Set xmlTypeOfAppNode = Nothing
'
'
'    Exit Function
'
'UpdateQuotationForAddedCostsVbErr:
'
'    If m_objErrAssist.IsWarning = True Then
'        m_objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'
'    m_objErrAssist.AddToErrSource strFunctionName
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'    Set IApplicationQuoteTxBO_UpdateQuotationForAddedCosts = m_objErrAssist.CreateErrorResponseEx()
'
'
'    m_objContext.SetAbort
'
'    ' Go to clean-up section
'    Resume UpdateQuotationForAddedCostsExit
'
'End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXMLAssist = New XMLAssist
    
    ' BMIDS986
    If m_objConvertAssist Is Nothing Then
        Set m_objConvertAssist = New ConvertAssist
    End If

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXMLAssist = Nothing

End Sub

Private Function IApplicationQuoteTxBO_RecostQuotationForRateChange(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Remodels the repayment costs and APR to ensure that the latest rates
'               have been applied prior to issuing the offer
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo RecostQuotationForRateChangeVbErr

' BMIDS977  CC089  Many changes - old code not preserved for clarity
'                  Removed commented out code for clarity
'                  See Version 60 for code prior to CC089.

Const strFunctionName As String = "IApplicationQuoteTxBO_RecostQuotationForRateChange"
    
'Declare Variables
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlDoc As FreeThreadedDOMDocument40

Dim xmlLoanComposition As IXMLDOMNode
Dim xmlRequestElement As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim xmlResponseElem As IXMLDOMElement
Dim xmlAppNo As IXMLDOMElement
Dim xmlAppFFNo As IXMLDOMElement
Dim xmlMortgageSubQuoteNo As IXMLDOMElement
Dim xmlOneOffCosts As IXMLDOMNode
Dim xmlMortCalcOutput As IXMLDOMNode

Dim xmlTempRequest As IXMLDOMElement
Dim xmlTempElement  As IXMLDOMElement
Dim xmlTempNode As IXMLDOMNode
Dim xmlTempResponseNode As IXMLDOMElement

Dim objIApplicationQuoteBO As IApplicationQuoteBO
Dim objICostModelAssistBO As ICostModelAssistBO
Dim objMortSubQuoteBO As MortgageSubQuoteBO
Dim objQuotationBO As QuotationBO
Dim objConvAssist As ConvertAssist
    
Dim strApplicationDate As String
Dim strResponse As String
Dim strMSQTotalGrossMonthlyCost As String

'BMIDS834 GHun
Dim xmlMtgeOneOffCosts  As IXMLDOMNode
Dim xmlMtgeCost         As IXMLDOMElement
Dim xmlCost             As IXMLDOMNode
Dim xmlComboDoc         As FreeThreadedDOMDocument40
Dim objComboDO          As ComboDO
Dim strCostType         As String
Dim strPattern          As String
Dim strValidationType   As String
Dim strType             As String
'BMIDS834 End

Dim strAmountRequested  As String   'BMIDS866 GHun
    
' BMIDS986
Dim objAppProcBO As Object
    
Dim dblROTPurchasePrice     As Double
Dim dblPostWorksValuation   As Double
Dim dblPresentValuation     As Double
Dim dblAppPurchasePriceOrEstimatedValue  As Double
Dim lngRetentionRoads       As Long
Dim lngRetentionWorks       As Long
Dim strPurchasePrice        As String
Dim lngErrorNo              As Long
Dim strApplicationType      As String

Dim xmlEleAppData As IXMLDOMElement
Dim xmlLatestValAmtsReqDoc  As New FreeThreadedDOMDocument40
xmlLatestValAmtsReqDoc.validateOnParse = False
xmlLatestValAmtsReqDoc.setProperty "NewParser", True
Dim xmlLatestValAmtsReq     As IXMLDOMElement
Dim xmlLatestValAmtsReqElem As IXMLDOMElement
Dim xmlLatestValAmtsRespDoc As New FreeThreadedDOMDocument40
xmlLatestValAmtsRespDoc.validateOnParse = False
xmlLatestValAmtsRespDoc.setProperty "NewParser", True
Dim xmlTypeOfAppNode        As IXMLDOMNode
    
' BMIDS986 End

' BMIDS977
Dim lngCharge                  As Long
Dim xmlMtgeCostList            As IXMLDOMNodeList
Dim xmlMtgeOneOffCostsForCalc  As IXMLDOMNode
Dim xmlNewRequest              As IXMLDOMElement
Dim xmlRootElement             As IXMLDOMElement
Dim xmlValidationType          As IXMLDOMNode

   
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrAPPLICATION_NAME Then
        Set xmlRequestElement = vxmlRequest
    Else
        Set xmlRequestElement = vxmlRequest.getElementsByTagName(cstrAPPLICATIONQUOTE_NAME).Item(0)
    End If
    
    If xmlRequestElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrAPPLICATION_NAME & " tag not found"
    End If
    
    Set objConvAssist = New ConvertAssist
    
       
'Stage1: Get One Off Costs -----------------------------------------------------------------------------------------

    Set xmlAppNo = m_objXMLAssist.GetMandatoryNode(xmlRequestElement, "APPLICATIONNUMBER")
    Set xmlAppFFNo = m_objXMLAssist.GetMandatoryNode(xmlRequestElement, "APPLICATIONFACTFINDNUMBER")
    Set xmlMortgageSubQuoteNo = m_objXMLAssist.GetMandatoryNode(xmlRequestElement, "MORTGAGESUBQUOTENUMBER")
    
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempElement = xmlOut.createElement("ONEOFFCOST")
    xmlTempRequest.appendChild xmlTempElement
    xmlTempElement.appendChild xmlAppNo.cloneNode(True)
    xmlTempElement.appendChild xmlAppFFNo.cloneNode(True)
    xmlTempElement.appendChild xmlMortgageSubQuoteNo.cloneNode(True)
    
    Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    strResponse = objMortSubQuoteBO.GetOneOffCostsDetails(xmlTempRequest.xml)
    
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    Set xmlMtgeOneOffCosts = m_objXMLAssist.GetMandatoryNode(xmlDoc.documentElement, "//MORTGAGEONEOFFCOSTLIST")
    Set xmlTempNode = m_objXMLAssist.GetMandatoryNode(xmlDoc, ".//MORTGAGESUBQUOTE")

    'Store the TotalGrossMonthlyCost for later use
    strMSQTotalGrossMonthlyCost = m_objXMLAssist.GetNodeValue(xmlTempNode, "TOTALGROSSMONTHLYCOST")

    ' Process the One Off Costs to include application costs with fee rebates and
    ' additions applied and fees from MortgageOneOffCosts which have been added to loan
    '
    ' Hold results in xmlMtgeOneOffCostsForCalc
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    xmlNewRequest.appendChild xmlAppNo.cloneNode(True)
       
    ' Remove any costs with an AMOUNT of 0
    Set xmlMtgeCostList = xmlMtgeOneOffCosts.selectNodes(".//MORTGAGEONEOFFCOST")
    For Each xmlMtgeCost In xmlMtgeCostList

        lngCharge = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//AMOUNT")

        If (lngCharge = 0) Then
            xmlMtgeOneOffCosts.removeChild xmlMtgeCost
        End If

    Next
     
    xmlNewRequest.appendChild xmlMtgeOneOffCosts.cloneNode(True)
        
    Set xmlTempNode = IApplicationQuoteTxBO_ProcessAllFeesForRemodelQuote(xmlNewRequest)
    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
        
    Set xmlMtgeOneOffCostsForCalc = xmlTempNode.selectSingleNode(".//MORTGAGEONEOFFCOSTLIST").cloneNode(True)

    ' Manipulate this list for use later
    ' Remove AppNumber, AppFactFindNumber, MortgageSubQuoteNumber and OneOffCostsequenceNumber
    ' from each One Off Cost.
    ' Add IDENTIFIER node to each One Off cost
    
    Set xmlTempElement = xmlOut.createElement("LIST")
    Set xmlTempNode = xmlOut.createElement("LISTNAME")
    xmlTempNode.Text = "OneOffCost"
    xmlTempElement.appendChild xmlTempNode
    
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")

    strResponse = objComboDO.GetComboList(xmlTempElement.xml)
    Set xmlComboDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    Set xmlMtgeCostList = xmlMtgeOneOffCostsForCalc.selectNodes(".//MORTGAGEONEOFFCOST")
    For Each xmlMtgeCost In xmlMtgeCostList
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//APPLICATIONNUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//APPLICATIONFACTFINDNUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
     
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//MORTGAGESUBQUOTENUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
    
        Set xmlElement = m_objXMLAssist.GetNode(xmlMtgeCost, ".//ONEOFFCOSTSEQUENCENUMBER", False)
        If Not xmlElement Is Nothing Then
            xmlMtgeCost.removeChild xmlElement
        End If
    
        Set xmlRootElement = xmlOut.createElement("IDENTIFIER")
        
        'Need to put the ValidationType into the Identifier
        strCostType = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
        
        ' Get all the Validation Types for this OneOffCost
        strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE"

        ' Choose the first validation type which is not APR or END
        For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
            strType = xmlValidationType.Text
            If (strType <> "APR") And (strType <> "END") Then
                strValidationType = strType
                Exit For ' get out of this loop
            End If
        Next
        
        xmlRootElement.Text = strValidationType
        xmlMtgeCost.appendChild xmlRootElement
   
    Next

'Stage2: Get loan composition details for the sub quotes ------------------------------------------------------------
    m_objXMLAssist.ChangeNodeName xmlTempRequest, "ONEOFFCOST", "LOANCOMPOSITION"
    
    'BM0230 MDC 08/01/2003 - Check BASERATEAPPLIED indicator to cater for Failed Batch Runs
    'BM0211 MDC 31/12/2002 - Indicate to GetLoanCompositionDetails that this is for a Rate Change
    If xmlGetAttributeAsBoolean(vxmlRequest, "BASERATEAPPLIED") = False Then
        xmlSetAttributeValue xmlTempRequest, "RATECHANGE", "1"
    End If
    'BM0211 MDC 31/12/2002 - End
    
    xmlCopyAttribute vxmlRequest, xmlTempRequest, "BASERATEAPPLIED"
    'BM0230 MDC 08/01/2003 - End
    
    Set objIApplicationQuoteBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteBO")
    Set xmlTempResponseNode = objIApplicationQuoteBO.GetLoanCompositionDetails(xmlTempRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    Set xmlLoanComposition = m_objXMLAssist.GetMandatoryNode(xmlTempResponseNode, "//LOANCOMPONENTLIST")
    

'Stage3: Get the Application Date --------------------------------------------------------
    
    strApplicationDate = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseNode, "//APPLICATIONFACTFIND/APPLICATIONDATE")


'Stage5: Recost each Loan Component...... ------------------------------------------------------------------------------
    
    'Instantiate objects used within the loop
    Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempElement = xmlDoc.createElement("APPLICATIONQUOTE")
    xmlTempElement.appendChild xmlAppNo
    xmlTempElement.appendChild xmlAppFFNo

    Set xmlTempNode = xmlDoc.createElement("CONTEXT")
    xmlTempNode.Text = "COST MODELLING"
    xmlTempElement.appendChild xmlTempNode

    xmlTempElement.appendChild xmlLoanComposition
    
    m_objXMLAssist.ChangeNodeName xmlMtgeOneOffCostsForCalc, "MORTGAGEONEOFFCOSTLIST", "ONEOFFCOSTLIST"
    xmlTempElement.appendChild xmlMtgeOneOffCostsForCalc   ' BMIDS977
        
    Set xmlTempNode = xmlDoc.createElement("INTERESTONLYAMOUNT")
    xmlTempNode.Text = ""
'    xmlTempNode.Text = strInterestOnlyElement
    xmlTempElement.appendChild xmlTempNode

    Set xmlTempNode = xmlDoc.createElement("CAPITALANDINTERESTAMOUNT")
    xmlTempNode.Text = ""
'    xmlTempNode.Text = strCapitalAndInterestElement
    xmlTempElement.appendChild xmlTempNode

    Set xmlTempNode = xmlDoc.createElement("APPLICATIONDATE")
    xmlTempNode.Text = strApplicationDate
    xmlTempElement.appendChild xmlTempNode

    Set xmlTempNode = xmlDoc.createElement("RATECHANGE")
    xmlTempNode.Text = 1
    xmlTempElement.appendChild xmlTempNode

    xmlTempRequest.appendChild xmlTempElement
        
    Set xmlMortCalcOutput = objICostModelAssistBO.RecostMortgageComponentAndLifeCover(xmlTempRequest)
    m_objErrAssist.CheckResponse xmlMortCalcOutput, xmlResponseElem, True
    
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempNode = m_objXMLAssist.GetMandatoryNode(xmlMortCalcOutput, "LOANCOMPONENTLIST")
    xmlTempRequest.appendChild xmlTempNode
    Set xmlTempNode = xmlDoc.createElement("CONTEXT")
    xmlTempNode.Text = "UPDATE"
    xmlTempRequest.appendChild xmlTempNode
    strResponse = objMortSubQuoteBO.SaveCompositeRateResults(xmlTempRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    
    
'Stage 6 - Update the Mortgage Sub Quote ------------------------------------------------------------------------------
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempNode = xmlDoc.createElement("MORTGAGESUBQUOTE")
    xmlTempRequest.appendChild xmlTempNode
    xmlTempNode.appendChild xmlAppNo                'Application Number
    xmlTempNode.appendChild xmlAppFFNo              'Application Fact Find Number
    xmlTempNode.appendChild xmlMortgageSubQuoteNo   'Mortgage Sub Quote Number
    'Total Gross Monthly Cost
    
    Set xmlTempElement = xmlDoc.createElement("TOTALGROSSMONTHLYCOST")
    
    ' PSC 11/07/2002 BMIDS00062
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "TOTALGROSSMONTHLYCOST")
    xmlTempNode.appendChild xmlTempElement
    'Total Net Monthly Cost
    Set xmlTempElement = xmlDoc.createElement("TOTALNETMONTHLYCOST")
    
    ' PSC 11/07/2002 BMIDS00062
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "TOTALNETMONTHLYCOST")
    xmlTempNode.appendChild xmlTempElement
    
    ' PSC 11/07/2002 BMIDS00062 - Start
    'Total Accrued Interest
    
    Set xmlTempElement = xmlDoc.createElement("TOTALACCRUEDINTEREST")
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "TOTALACCRUEDINTEREST")
    xmlTempNode.appendChild xmlTempElement
    ' PSC 11/07/2002 BMIDS00062 - End
    
    'BMIDS793 GHun
    Set xmlTempElement = xmlDoc.createElement("AMOUNTPERUNITBORROWED")
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "AMOUNTPERUNITBORROWED")
    xmlTempNode.appendChild xmlTempElement
    
    Set xmlTempElement = xmlDoc.createElement("APR")
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "APR")
    xmlTempNode.appendChild xmlTempElement
    
    Set xmlTempElement = xmlDoc.createElement("TOTALAMOUNTPAYABLE")
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "TOTALAMOUNTPAYABLE")
    xmlTempNode.appendChild xmlTempElement
    
    Set xmlTempElement = xmlDoc.createElement("TOTALMORTGAGEPAYMENTS")
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "TOTALMORTGAGEPAYMENTS")
    xmlTempNode.appendChild xmlTempElement
    'BMIDS793 End
    
    'Total Gross Monthly Cost
    Set xmlTempElement = xmlDoc.createElement("DATEANDTIMEGENERATED")
    xmlTempElement.Text = Format$(Now, "dd/mm/yyyy hh:mm:ss")
    xmlTempNode.appendChild xmlTempElement
    
    strResponse = objMortSubQuoteBO.Update(xmlTempRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    
'Stage 7 - Update Linked Quotations -----------------------------------------------------------------------------------
    Set xmlTempRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempNode = xmlDoc.createElement("QUOTATION")
    xmlTempRequest.appendChild xmlTempNode
    xmlTempNode.appendChild xmlAppNo
    xmlTempNode.appendChild xmlAppFFNo
    xmlTempNode.appendChild xmlMortgageSubQuoteNo
    'Old MSQ Cost
    Set xmlTempElement = xmlDoc.createElement("OLDMORTGAGESUBQUOTECOST")
    xmlTempElement.Text = strMSQTotalGrossMonthlyCost
    xmlTempNode.appendChild xmlTempElement
    'New MSQ Cost
    Set xmlTempElement = xmlDoc.createElement("NEWMORTGAGESUBQUOTECOST")
    xmlTempElement.Text = m_objXMLAssist.GetMandatoryElementText(xmlMortCalcOutput, "TOTALGROSSMONTHLYCOST")
    xmlTempNode.appendChild xmlTempElement
    
    Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    strResponse = objQuotationBO.UpdateLinkedQuotations(xmlTempRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem

'Stage 8 - Return Success Response ------------------------------------------------------------------------------------
    'If we've got here everything must have worked !
    Set IApplicationQuoteTxBO_RecostQuotationForRateChange = xmlResponseElem
    m_objContext.SetComplete

RecostQuotationForRateChangeExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing

    Set xmlLoanComposition = Nothing
    Set xmlRequestElement = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElem = Nothing
    Set xmlAppNo = Nothing
    Set xmlAppFFNo = Nothing
    Set xmlMortgageSubQuoteNo = Nothing
    Set xmlOneOffCosts = Nothing
    Set xmlMortCalcOutput = Nothing

    Set xmlTempRequest = Nothing
    Set xmlTempElement = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempResponseNode = Nothing

    Set objIApplicationQuoteBO = Nothing
    Set objICostModelAssistBO = Nothing
    Set objMortSubQuoteBO = Nothing
    Set objConvAssist = Nothing
    Set objQuotationBO = Nothing
    'BMIDS834 GHun
    Set xmlMtgeOneOffCosts = Nothing
    Set xmlMtgeCost = Nothing
    Set xmlCost = Nothing
    Set xmlComboDoc = Nothing
    Set objComboDO = Nothing
    'BMIDS834 End
    
    'BMIDS986
    Set xmlEleAppData = Nothing
    Set xmlLatestValAmtsReqDoc = Nothing
    Set xmlLatestValAmtsReq = Nothing
    Set xmlLatestValAmtsReqElem = Nothing
    Set xmlLatestValAmtsRespDoc = Nothing
    Set xmlTypeOfAppNode = Nothing
       
    ' BMIDS977
    Set xmlMtgeCostList = Nothing
    Set xmlMtgeOneOffCostsForCalc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlRootElement = Nothing
    Set xmlValidationType = Nothing
      
    Exit Function

RecostQuotationForRateChangeVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_RecostQuotationForRateChange = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RecostQuotationForRateChangeExit

End Function

Private Function IApplicationQuoteTxBO_RefreshMortgageAccountData( _
                                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RefreshMortgageAccountDataVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_RefreshMortgageAccountData"
    
    Dim strApplicationNumber As String, strAFFNumber As String, strPortingIndicator As String
    Dim strResponse As String, strRAExists As String, strAmountRequested As String
    Dim strLTV As String, strTemp As String, strICBSTerm As String
    Dim dtLastAccDownLoadDate As Date
    Dim blnFurtherAdvanceTermInd As Boolean
    
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement, xmlTable As IXMLDOMElement, xmlElem As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim objApplicationBO As ApplicationBO, objAppManagerBO As ApplicationManagerBO
    Dim objRiskAssessmentBO As RiskAssessmentBO
    Dim objICostModelAssistBO As ICostModelAssistBO, objMSQBO As MortgageSubQuoteBO
    Dim objConvertAssist As ConvertAssist
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim strTermInYears As String
    Dim strTermInMonths As String
    Dim xmlLCSQNONode As IXMLDOMNode
    Dim xmlTermYears As IXMLDOMNode
    Dim xmlTermMonths As IXMLDOMNode
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlImportNode As IXMLDOMNode
    Dim nGlobalTermMnths As Integer
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")  'JD BMIDS975
    'MAR46 GHun
    If objGlobalParameterDO.GetCurrentParameterByType("UseAdminGetAccountRefresh", "BOOLEAN") = 1 Then
    'MAR46 End
    
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objAppManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
        Set objRiskAssessmentBO = m_objContext.CreateInstance(gstrRISK_ASSESSMENT & ".RiskAssessmentBO")
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objMSQBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objConvertAssist = New ConvertAssist
        
        '-----------------------------------------------------------------
        ' Get Application.LastDownLoadDate
        '-----------------------------------------------------------------
        strApplicationNumber = m_objXMLAssist.GetTagValue(vxmlRequest, ".//APPLICATIONNUMBER")
        strAFFNumber = m_objXMLAssist.GetTagValue(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER")
        
        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTable = xmlOut.createElement("APPLICATION")
        xmlNewRequest.appendChild xmlTable
        
        Set xmlElem = xmlOut.createElement("APPLICATIONNUMBER")
        xmlElem.Text = strApplicationNumber
        xmlTable.appendChild xmlElem
        
        Set xmlElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        xmlElem.Text = strAFFNumber
        xmlTable.appendChild xmlElem
        
        'JD BMIDS975 change to call GetApplicationData in order to get the FurtherAdvanceTermInd
        Set xmlElem = xmlOut.createElement("CALCSDATAONLY")
        xmlElem.Text = "1"
        xmlTable.appendChild xmlElem
        
        strResponse = objApplicationBO.GetApplicationData(xmlNewRequest.xml)
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
        Set xmlTable = xmlDoc.selectSingleNode("//APPLICATIONFACTFIND")  'JD BMIDS975
        dtLastAccDownLoadDate = objConvertAssist.CSafeDate(m_objXMLAssist.GetTagValue(xmlTable, "LASTACCOUNTDOWNLOADDATE"))
        blnFurtherAdvanceTermInd = objConvertAssist.CSafeBool(m_objXMLAssist.GetTagValue(xmlTable, "FURTHERADVANCETERMIND")) ' JD BMIDS975
        
        '-----------------------------------------------------------------
        'Refresh Account data if LastAccDownLoadDate is not today
        '-----------------------------------------------------------------
        If FormatDateTime(dtLastAccDownLoadDate, vbShortDate) <> FormatDateTime(Now, vbShortDate) Then
            'JD BMIDS975 First check that if the loancomponents term is going to be
            '           changed that we have a sensible value in global parameter AccRefreshMinMthsTermToApply
            If IsValidationType("TypeOfMortgage", m_objXMLAssist.GetTagValue(vxmlRequest, ".//TYPEOFAPPLICATION"), "ET") And _
                   blnFurtherAdvanceTermInd Then
                On Error GoTo GetGlobalErr
                nGlobalTermMnths = objGlobalParameterDO.GetCurrentParameterByType("AccRefreshMinMthsTermToApply", "AMOUNT")
                On Error GoTo RefreshMortgageAccountDataVbErr
                If nGlobalTermMnths = Null Or nGlobalTermMnths = 0 Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7029
                End If
            End If
            
            'JD BMIDS975 need to pass in the FurtherAdvanceInd to ImportAccountRefresh too
            Set xmlImportNode = vxmlRequest.cloneNode(True)
            Set xmlNode = xmlImportNode.selectSingleNode(".//APPLICATION")
            Set xmlElem = xmlDoc.createElement("FURTHERADVANCETERMIND")
            xmlElem.Text = m_objXMLAssist.GetTagValue(xmlTable, "FURTHERADVANCETERMIND")
            xmlNode.appendChild xmlElem
            'strResponse = objAppManagerBO.ImportAccountRefreshIntoApplication(vxmlRequest.xml)
            strResponse = objAppManagerBO.ImportAccountRefreshIntoApplication(xmlImportNode.xml)
            Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
            
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
            
            'JD BMIDS975 Get the Term returned from the Import
            Set xmlTable = xmlDoc.selectSingleNode("//TERM")
            strICBSTerm = m_objXMLAssist.GetTagValue(xmlTable, "ICBSOUTSTANDINGTERM")
            
            strPortingIndicator = m_objXMLAssist.GetTagValue(vxmlRequest, "PORTINGINDICATOR")
            If strPortingIndicator <> "1" Then
                'Check whether a Risk Assessment was already done for this application
                Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
                Set xmlTable = xmlOut.createElement("RISKASSESSMENT")
                xmlNewRequest.appendChild xmlTable
        
                Set xmlElem = xmlOut.createElement("APPLICATIONNUMBER")
                xmlElem.Text = strApplicationNumber
                xmlTable.appendChild xmlElem
                
                Set xmlElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
                xmlElem.Text = strAFFNumber
                xmlTable.appendChild xmlElem
                
                strResponse = objRiskAssessmentBO.CheckRiskAssessmentExistsForApplication(xmlNewRequest.xml)
                Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
                m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
                
                strRAExists = m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "RISKASSESSMENTEXISTS")
                If strRAExists = "1" Then 'Risk Assessment exists
                    'Calculate LTV again
                    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
                    Set xmlTable = xmlOut.createElement("LTV")
                    xmlNewRequest.appendChild xmlTable
                    
                    Set xmlElem = xmlOut.createElement("APPLICATIONNUMBER")
                    xmlElem.Text = strApplicationNumber
                    xmlTable.appendChild xmlElem
                
                    Set xmlElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
                    xmlElem.Text = strAFFNumber
                    xmlTable.appendChild xmlElem
                    
                    Set xmlElem = vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTE")
                    strAmountRequested = m_objXMLAssist.GetTagValue(xmlElem, "AMOUNTREQUESTED")
                    strLTV = m_objXMLAssist.GetTagValue(xmlElem, "LTV")  ' Current LTV (passed in)
                    
                    Set xmlElem = xmlOut.createElement("AMOUNTREQUESTED")
                    xmlElem.Text = strAmountRequested
                    xmlTable.appendChild xmlElem
                                                
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//CUSTOMERLIST").cloneNode(True)
                    
                    Set xmlTempResponse = objICostModelAssistBO.CalcCostModelLTV(xmlNewRequest, "COST MODELLING")
                    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
                    
                    strTemp = m_objXMLAssist.GetTagValue(xmlTempResponse, "LTV")
                    If strTemp <> strLTV Then 'LTV changed
                        Set xmlElem = xmlOut.createElement("LTVCHANGED")
                        xmlElem.Text = "1"
                        xmlResponseElem.appendChild xmlElem
                        
                        Set xmlElem = xmlOut.createElement("LTV")
                        xmlElem.Text = strTemp
                        xmlResponseElem.appendChild xmlElem
                        
                        'Update the MortgageSubQuote with new LTV
                        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
                        Set xmlTable = xmlOut.createElement("MORTGAGESUBQUOTE")
                        xmlNewRequest.appendChild xmlTable
                        
                        Set xmlElem = xmlOut.createElement("APPLICATIONNUMBER")
                        xmlElem.Text = strApplicationNumber
                        xmlTable.appendChild xmlElem
                        
                        Set xmlElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
                        xmlElem.Text = strAFFNumber
                        xmlTable.appendChild xmlElem
                        
                        Set xmlElem = m_objXMLAssist.GetMandatoryNode(vxmlRequest, ".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
                        xmlTable.appendChild xmlElem
                        
                        Set xmlElem = m_objXMLAssist.GetMandatoryNode(vxmlRequest, ".//LTV").cloneNode(True)
                        xmlElem.Text = strTemp
                        xmlTable.appendChild xmlElem
                        
                        strResponse = objMSQBO.Update(xmlNewRequest.xml)
                        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
                        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
                    Else ' LTV has not changed
                        Set xmlElem = xmlOut.createElement("LTVCHANGED")
                        xmlElem.Text = "0"
                        xmlResponseElem.appendChild xmlElem
                    End If
                End If
                'BMIDS975
                'update the loancomponent with the ICBS term if the type of mortgage is ET and the application has a FurtherAdvanceTermInd
                If IsValidationType("TypeOfMortgage", m_objXMLAssist.GetTagValue(vxmlRequest, ".//TYPEOFAPPLICATION"), "ET") And _
                   blnFurtherAdvanceTermInd Then
                    
                    'Check that the ICBS term is sensible. Set to global value if not.
                    Dim bShowWarning As Boolean
                    bShowWarning = False
                    If strICBSTerm = "" Then
                        strICBSTerm = CStr(nGlobalTermMnths)
                        'show warning to user
                        bShowWarning = True
                    End If
                    
                    If CInt(strICBSTerm) < nGlobalTermMnths Then
                        strICBSTerm = CStr(nGlobalTermMnths)
                        'show warning to user
                        bShowWarning = True
                    End If
                    
                    'change the returned term into years and months
                    strTermInYears = CStr(Int(CInt(strICBSTerm) / 12))
                    strTermInMonths = CStr(CInt(strICBSTerm) Mod 12)
                    
                    If bShowWarning Then
                        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7028, "", strTermInYears, strTermInMonths
                    End If
                    
                    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
                    Set xmlTable = xmlOut.createElement("MORTGAGESUBQUOTE")
                    xmlNewRequest.appendChild xmlTable
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
                    
                    strResponse = objMSQBO.FindLoanComponentList(xmlNewRequest.xml)
                    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
                    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
                    
                    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
                    Set xmlTable = xmlOut.createElement("LOANCOMPONENT")
                    xmlNewRequest.appendChild xmlTable
                    
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
                    xmlTable.appendChild vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
                    Set xmlLCSQNONode = xmlTable.appendChild(xmlOut.createElement("LOANCOMPONENTSEQUENCENUMBER"))
                    Set xmlTermYears = xmlTable.appendChild(xmlOut.createElement("TERMINYEARS"))
                    Set xmlTermMonths = xmlTable.appendChild(xmlOut.createElement("TERMINMONTHS"))
                    
                    Set xmlLoanComponentList = xmlDoc.selectNodes(".//LOANCOMPONENT")
                    For Each xmlNode In xmlLoanComponentList
                        xmlLCSQNONode.Text = m_objXMLAssist.GetTagValue(xmlNode, "LOANCOMPONENTSEQUENCENUMBER")
                        xmlTermYears.Text = strTermInYears
                        xmlTermMonths.Text = strTermInMonths
                        strResponse = objMSQBO.UpdateLoanComponent(xmlNewRequest.xml)
                        Set xmlOut = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
                        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
                    Next xmlNode
                        
                End If
            End If
        End If
    
    End If 'MAR46 GHun
    
    Set IApplicationQuoteTxBO_RefreshMortgageAccountData = xmlResponseElem
    
    m_objContext.SetComplete
    
RefreshMortgageAccountDataExit:
              
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlElem = Nothing
    Set xmlResponseElem = Nothing
    Set objApplicationBO = Nothing
    Set objAppManagerBO = Nothing
    Set objRiskAssessmentBO = Nothing
    Set objICostModelAssistBO = Nothing
    Set objMSQBO = Nothing
    Set objConvertAssist = Nothing
    Set xmlLCSQNONode = Nothing
    Set xmlTermYears = Nothing
    Set xmlTermMonths = Nothing
    Set xmlImportNode = Nothing
    Set objGlobalParameterDO = Nothing
    
    Exit Function

GetGlobalErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If
    
RefreshMortgageAccountDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_RefreshMortgageAccountData = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume RefreshMortgageAccountDataExit
End Function

Private Function IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount( _
                                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetAndSavePortedStepAndPeriodFromMortgageAccountVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount"

    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode, xmlTable As IXMLDOMElement, xmlElem As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlLCNumber As IXMLDOMNode, xmlProductStep As IXMLDOMNode
    Dim xmlRemStepDuration As IXMLDOMNode, xmlProdStartDate As IXMLDOMNode
    
    Dim objMSQBO As MortgageSubQuoteBO, objCFBO As CustomerFinancialBO
    Dim strResponse As String, strProductStep As String, strRemainingStepDuration As String
    Dim strLCNumber As String
    Dim xmlNode As IXMLDOMNode
    Dim xmlHighNode As IXMLDOMNode
    'MAR46 GHun
    Dim strTypeOfMortgage           As String
    Dim blnIsTOE                    As Boolean
    Dim blnIsPorting                As Boolean
    Dim blnIsProductSwitchRetained  As Boolean
    Dim xmlUpdatedLC                As IXMLDOMNode
    Dim xmlCurrentRateExpDate       As IXMLDOMElement
    'MAR46 End
    
    Dim strMortgageLoanGuid As String   ' PSC 12/03/2007 EP2_1611
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objMSQBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTable = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlNewRequest.appendChild xmlTable
    
    Set xmlElem = xmlOut.createElement("PORTINGDATAONLY")
    xmlElem.Text = "1"
    xmlTable.appendChild xmlElem
    
    xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    xmlTable.appendChild vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
    
    strResponse = objMSQBO.FindLoanComponentList(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    'MAR46 GHun Check if this is a transfer of equity, porting or product switch retained
    blnIsPorting = False
    blnIsProductSwitchRetained = False
    
    strTypeOfMortgage = xmlGetMandatoryNodeText(vxmlRequest, "..//TYPEOFAPPLICATION")
    blnIsTOE = IsValidationType("TypeOfMortgage", strTypeOfMortgage, "TOE")
    
    If blnIsTOE Then
        Set xmlLoanComponentList = xmlDoc.selectNodes(".//LOANCOMPONENT")
    Else
        Set xmlLoanComponentList = xmlDoc.selectNodes(".//LOANCOMPONENT[MANUALPORTEDLOANIND='1']")
        If xmlLoanComponentList.length > 0 Then 'If at least one LoanComponent is ported
            blnIsPorting = True
        Else
            Set xmlLoanComponentList = xmlDoc.selectNodes(".//LOANCOMPONENT[PRODUCTSWITCHRETAINPRODUCTIND='1']")
            If xmlLoanComponentList.length > 0 Then
                blnIsProductSwitchRetained = True
            End If
        End If
    End If
    
    If (blnIsPorting Or blnIsTOE Or blnIsProductSwitchRetained) Then
    'MAR46 End
        Set xmlNewRequest = vxmlRequest.cloneNode(True)
        Set xmlElem = xmlNewRequest.selectSingleNode(".//PORTINGINDICATOR")
        If xmlElem Is Nothing Then
            Set xmlElem = xmlOut.createElement("PORTINGINDICATOR")
            Set xmlTable = xmlNewRequest.selectSingleNode(".//GETANDSAVEPORTEDSTEPANDPERIODFROMMORTGAGEACCOUNT")
            xmlTable.appendChild xmlElem
        End If
        xmlElem.Text = "1"
        'xmlNewRequest.appendChild xmlElem
        
        Set xmlTempResponse = IApplicationQuoteTxBO_RefreshMortgageAccountData(xmlNewRequest)
        m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
        
        ' Now get PortedStep and Period from Account
        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTable = xmlOut.createElement("APPLICATION")
        xmlNewRequest.appendChild xmlTable
        
        xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
        xmlTable.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
        xmlTable.appendChild vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
        'MAR46 GHun Pass TOE flag
        Set xmlNode = xmlDoc.createElement("TOEIND")
        xmlNode.Text = Abs(blnIsTOE)
        xmlTable.appendChild xmlNode
        'MAR46 End
        
        strResponse = objCFBO.GetPortedStepAndPeriodFromMortgageAccount(xmlNewRequest.xml)
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
        'Now update the Ported Loan Components with new values fetched above
        xmlChangeNodeName xmlTable, "APPLICATION", "LOANCOMPONENT"
                
        For Each xmlElem In xmlLoanComponentList
            Set xmlUpdatedLC = xmlNewRequest.cloneNode(True)
            Set xmlTable = xmlGetMandatoryNode(xmlUpdatedLC, "LOANCOMPONENT")
            
            strLCNumber = xmlGetMandatoryNodeText(xmlElem, "LOANCOMPONENTSEQUENCENUMBER")
            Set xmlLCNumber = xmlOut.createElement("LOANCOMPONENTSEQUENCENUMBER")
            xmlTable.appendChild xmlLCNumber
            xmlLCNumber.Text = strLCNumber
            
            ' PSC 12/03/2007 EP2_1611 - Start
            strMortgageLoanGuid = xmlGetMandatoryNodeText(xmlElem, "MORTGAGELOANGUID")
        
            'The stored proc returns the nodes sorted by highest PORTEDSTEP, lowest REMAININGSTEPDURATION (ASC)
            Set xmlNode = xmlDoc.selectSingleNode("RESPONSE/MORTGAGELOAN[MORTGAGELOANGUID='" & strMortgageLoanGuid & "']")
             ' PSC 12/03/2007 EP2_1611 - End
           
            If xmlNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7022
            End If
            
            strProductStep = xmlGetNodeText(xmlNode, "PRODUCTSTEP")
            If Len(strProductStep) = 0 Then
                'MAR46 GHun  Product Step is now optional
                'm_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7023
                Set xmlProductStep = xmlOut.createElement("PRODUCTSTEP")
                xmlTable.appendChild xmlProductStep
                xmlProductStep.Text = strProductStep
                'MAR46 End
            End If
            
            strRemainingStepDuration = xmlGetNodeText(xmlNode, "REMAININGSTEPDURATION")
            ' Use the lowest remaining step duration that is non zero. If there are none then zero is acceptable
            If strRemainingStepDuration = "0" Then
                Set xmlHighNode = xmlDoc.selectSingleNode("RESPONSE/MORTGAGELOAN[LOANCOMPONENTSEQUENCENUMBER='" _
                    & strLCNumber & "' and PRODUCTSTEP='" & strProductStep _
                    & "' and REMAININGSTEPDURATION > '0']")     'MAR43 GHun Changed &gt; to >
                
                If Not xmlHighNode Is Nothing Then
                    Set xmlNode = xmlHighNode
                End If
            End If
            
            'MAR46 GHun RemainingStepDuration is now optional
            strRemainingStepDuration = xmlGetNodeText(xmlNode, "REMAININGSTEPDURATION")
            If Len(strRemainingStepDuration) > 0 Then
                Set xmlRemStepDuration = xmlOut.createElement("REMAININGSTEPDURATION")
                xmlTable.appendChild xmlRemStepDuration
                xmlRemStepDuration.Text = xmlGetNodeText(xmlNode, "REMAININGSTEPDURATION")
            End If
            'MAR46 End
            
            Set xmlProdStartDate = xmlOut.createElement("PORTEDPRODUCTSTARTDATE")
            xmlTable.appendChild xmlProdStartDate
            xmlProdStartDate.Text = xmlGetNodeText(xmlNode, "PRODUCTSTARTDATE")
            
            'MAR46 GHun Add new CurrentRateExpiryDate
            Set xmlCurrentRateExpDate = xmlOut.createElement("CURRENTRATEEXPIRYDATE")
            xmlTable.appendChild xmlCurrentRateExpDate
            xmlCurrentRateExpDate.Text = xmlGetNodeText(xmlNode, "CURRENTRATEEXPIRYDATE")
            'MAR46 End
            
            strResponse = objMSQBO.UpdateLoanComponent(xmlUpdatedLC.xml)
            Set xmlOut = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
            
        Next xmlElem
    End If
    
    m_objContext.SetComplete
    
    Set IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount = xmlResponseElem
    
GetAndSavePortedStepAndPeriodFromMortgageAccountExit:
    
    Set objMSQBO = Nothing
    Set objCFBO = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTable = Nothing
    Set xmlElem = Nothing
    Set xmlTempResponse = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlLCNumber = Nothing
    Set xmlNode = Nothing
    Set xmlHighNode = Nothing
    Set xmlProductStep = Nothing
    Set xmlRemStepDuration = Nothing
    Set xmlProdStartDate = Nothing
    Set xmlUpdatedLC = Nothing
    Set xmlCurrentRateExpDate = Nothing

    Exit Function

GetAndSavePortedStepAndPeriodFromMortgageAccountVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_GetAndSavePortedStepAndPeriodFromMortgageAccount = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume GetAndSavePortedStepAndPeriodFromMortgageAccountExit

End Function

Private Function IApplicationQuoteTxBO_ProcessAllFeesForRemodelQuote( _
                                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ProcessAllFeesForRemodelQuoteVbErr

    Const strFunctionName As String = "IApplicationQuoteTxBO_ProcessAllFeesForRemodelQuote"

    Dim xmlRequestElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    
    Dim strApplicationNumber As String
    Dim strResponse As String
    Dim strValidationType As String
    Dim blnIsAddedtoLoan As Boolean
    Dim blnAPR As Boolean
    Dim blnAdd As Boolean
    Dim strCostType As String
    Dim strAmount As String
    Dim strAdHocInd As String
    Dim strPattern As String
    Dim bApplyRebateAddition As Boolean
    
    Dim objPProcBO As PaymentProcessingBO
    Dim objComboDO As ComboDO
    Dim objGlobalParameterDO As GlobalParameterDO

    Dim lngErrNo As Long
    Dim xmlMtgeCost As IXMLDOMNode
    Dim xmlMtgeOneOffCostList As IXMLDOMNodeList
    Dim xmlMtgeOneOffCostListNode As IXMLDOMNode
    Dim xmlApplicationFeeTypeListNode As IXMLDOMNode
    Dim xmlApplicationFeeTypeList As IXMLDOMNodeList
    Dim xmlAppFeeType As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlMtgeOneOffCostsForCalc As IXMLDOMNode
    Dim xmlManipulatedApplicationFeeType As IXMLDOMNode
    Dim xmlManipulatedApplicationFeeTypeList As IXMLDOMNodeList
    Dim xmlNewCostElement As IXMLDOMElement
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlValidationType As IXMLDOMNode
    Dim xmlCost As IXMLDOMNode
    Dim lngNonAPRCharges As Long
   
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Step 1: Get One Off Cost combo Validation Types
    
    Set xmlRootElement = xmlOut.createElement("LIST")
    Set xmlElement = xmlOut.createElement("LISTNAME")
    xmlElement.Text = "OneOffCost"
    xmlRootElement.appendChild xmlElement
       
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
       
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")

    strResponse = objComboDO.GetComboList(xmlRootElement.xml)
    Set xmlComboDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    Set objComboDO = Nothing

    ' Step 2 Get Application Fee Type list

    ' Call the Payment Processing method to return the application fee types so that
    ' they can be compared to the One Off Costs list.
    
    Set objPProcBO = m_objContext.CreateInstance(gstrPAYMENTPROCESSING & ".PaymentProcessingBO")
    
    strApplicationNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    
    Set xmlRequestElement = xmlDoc.createElement("REQUEST")
   
    xmlRequestElement.setAttribute "OPERATION", "FindFeeTypeList"
    Set xmlElement = xmlDoc.createElement("APPLICATIONFEETYPE")
    xmlElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlRequestElement.appendChild xmlElement
    strResponse = objPProcBO.omPayProcRequest(xmlRequestElement.xml)
        
    'Check response for errors. Ignore record not found
    xmlDoc.loadXML strResponse
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
        
    'If no error...
    If lngErrNo = 0 Or lngErrNo = oeRecordNotFound Then
        
        Set xmlMtgeOneOffCostListNode = vxmlRequest.selectSingleNode(".//MORTGAGEONEOFFCOSTLIST")
        Set xmlMtgeOneOffCostList = vxmlRequest.selectNodes(".//MORTGAGEONEOFFCOST")
        
        lngNonAPRCharges = 0
        
        For Each xmlMtgeCost In xmlMtgeOneOffCostList
    
            ' Convert the cost type to the validation type
            strCostType = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
            strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE"
            blnAPR = False
            
            For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
                If (xmlValidationType.Text = "APR") Then
                    blnAPR = True
                End If
            Next
            
            'If this cost type does not have an APR validation type, remove it
            'but keep a running total of the non-APR charges
            
            If (blnAPR = False) Then
            
                'NonAPRCharges must exclude TotalIndemnity (TID), Compulsory BC (CBC) and Compulsory PP (CPP)
                blnAdd = True
                For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
                    strValidationType = xmlValidationType.Text
                    If (strValidationType = "TID") Or _
                       (strValidationType = "CBC") Or _
                       (strValidationType = "CPP") Then
                
                        blnAdd = False
                    End If
                Next
            
                If (blnAdd = True) Then
                    lngNonAPRCharges = lngNonAPRCharges + m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//AMOUNT")
                End If
                
                xmlMtgeOneOffCostListNode.removeChild xmlMtgeCost
            End If
           
        Next
            
        Set xmlMtgeOneOffCostList = xmlMtgeOneOffCostListNode.selectNodes(".//MORTGAGEONEOFFCOST")
            
        Set xmlApplicationFeeTypeListNode = xmlDoc.selectSingleNode(".//RESPONSE")
        Set xmlApplicationFeeTypeList = xmlDoc.selectNodes(".//APPLICATIONFEETYPE")
            
        For Each xmlAppFeeType In xmlApplicationFeeTypeList
    
            ' Convert the cost type to the validation type
            strCostType = m_objXMLAssist.GetAttributeValue(xmlAppFeeType, "APPLICATIONFEETYPE", "FEETYPE")
            strPattern = ".//LISTENTRY[VALUEID = '" & strCostType & "']//VALIDATIONTYPE"
            blnAPR = False

            For Each xmlValidationType In xmlComboDoc.selectNodes(strPattern)
                If (xmlValidationType.Text = "APR") Then
                    blnAPR = True
                End If
            Next
            
            'If this cost type does not have an APR validation type, remove it
            If (blnAPR = False) Then
                xmlApplicationFeeTypeListNode.removeChild xmlAppFeeType
            End If
        Next
    
        Set xmlApplicationFeeTypeList = xmlApplicationFeeTypeListNode.selectNodes(".//APPLICATIONFEETYPE")
    
    '   Now we have : xmlMtgeOneOffCostList excluding any NON APR fee types
    '                 xmlApplicationFeeTypeList excluding any NON APR fee types
    
    
    'Step 3
    
    ' Now generate the correct list of Mortgage One Off Costs for calcs.
    
        Set xmlMtgeOneOffCostsForCalc = xmlOut.createElement("MORTGAGEONEOFFCOSTLIST")
        xmlResponseElem.appendChild xmlMtgeOneOffCostsForCalc
    
        For Each xmlMtgeCost In xmlMtgeOneOffCostList
    
            ' Get Mortgage One Off Cost records that are added to loan and do not exist in App Fee Type.
            blnIsAddedtoLoan = (m_objXMLAssist.GetNodeValue(xmlMtgeCost, "ADDTOLOAN", False, False) = "1")

            If (blnIsAddedtoLoan = True) Then
                xmlMtgeOneOffCostsForCalc.appendChild xmlMtgeCost
                    
            Else
                ' Not added to loan. Is it an application cost?
                    
                strCostType = m_objXMLAssist.GetNodeValue(xmlMtgeCost, ".//MORTGAGEONEOFFCOSTTYPE")
                strPattern = ".//APPLICATIONFEETYPE[@FEETYPE = " & strCostType & "]"
            
                Set xmlCost = xmlApplicationFeeTypeListNode.selectSingleNode(strPattern)
                If xmlCost Is Nothing Then
                    ' Fee exists in One Off Costs but not in Application Fees
                    ' and is included in APR so add to list for calcs.
        
                    xmlMtgeOneOffCostsForCalc.appendChild xmlMtgeCost
                End If
            End If
        Next
    
        ' Step 4
        ' If rebates/additions are to be applied (indicated by Global Parameter)
        ' call method to apply fee rebates and additions to Application Fee Type List.
        
        bApplyRebateAddition = objGlobalParameterDO.GetCurrentParameterByType("UseFeeRebateAdditionInCalcs", "BOOLEAN")
    
        If (bApplyRebateAddition = True) Then
               
            Set xmlRequestElement = xmlDoc.createElement("REQUEST")
   
            xmlRequestElement.setAttribute "OPERATION", "ApplyApplicationFeeRebatesAndAdditions"
    
            For Each xmlAppFeeType In xmlApplicationFeeTypeList
                xmlRequestElement.appendChild xmlAppFeeType
            Next
    
            strResponse = objPProcBO.omPayProcRequest(xmlRequestElement.xml)
        
            xmlDoc.loadXML strResponse
    
            Set xmlManipulatedApplicationFeeTypeList = xmlDoc.selectNodes(".//APPLICATIONFEETYPE")
        Else
            Set xmlManipulatedApplicationFeeTypeList = xmlApplicationFeeTypeListNode.selectNodes(".//APPLICATIONFEETYPE")
        
        End If
        
    
        ' Step 5
        ' Merge the Mortgage One Off costs for calcs (from Step 3) with the
        ' manipulated application fee types.
    
        For Each xmlManipulatedApplicationFeeType In xmlManipulatedApplicationFeeTypeList
        
            Set xmlNewCostElement = xmlDoc.createElement("MORTGAGEONEOFFCOST")
        
            strCostType = m_objXMLAssist.GetAttributeValue(xmlManipulatedApplicationFeeType, "APPLICATIONFEETYPE", "FEETYPE")
            strAmount = m_objXMLAssist.GetAttributeValue(xmlManipulatedApplicationFeeType, "APPLICATIONFEETYPE", "AMOUNT")
            strAdHocInd = m_objXMLAssist.GetAttributeValue(xmlManipulatedApplicationFeeType, "APPLICATIONFEETYPE", "ADHOCIND")
        
            Set xmlElement = xmlDoc.createElement("MORTGAGEONEOFFCOSTTYPE")
            xmlElement.Text = strCostType
            xmlNewCostElement.appendChild xmlElement
        
            Set xmlElement = xmlDoc.createElement("AMOUNT")
            xmlElement.Text = strAmount
            xmlNewCostElement.appendChild xmlElement
        
            Set xmlElement = xmlDoc.createElement("ADDTOLOAN")
            xmlElement.Text = "0"
            xmlNewCostElement.appendChild xmlElement
            
            Set xmlElement = xmlDoc.createElement("ADHOCIND")
            xmlElement.Text = strAdHocInd
            xmlNewCostElement.appendChild xmlElement

            xmlMtgeOneOffCostsForCalc.appendChild xmlNewCostElement
        Next
    
    End If
   
    ' Step 6
    ' Add the calculated non-APR charges to the response element
    Set xmlElement = xmlDoc.createElement("NONAPRCHARGES")
    xmlElement.Text = lngNonAPRCharges
    xmlResponseElem.appendChild xmlElement
    
    ' Return the list of all fees to the calling method
  
    Set IApplicationQuoteTxBO_ProcessAllFeesForRemodelQuote = xmlResponseElem
    
ProcessAllFeesForRemodelQuoteExit:
    
    m_objContext.SetComplete
    
    Set objPProcBO = Nothing
    Set objGlobalParameterDO = Nothing
       
    Set xmlRequestElement = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlMtgeCost = Nothing
    Set xmlMtgeOneOffCostList = Nothing
    Set xmlMtgeOneOffCostListNode = Nothing
    Set xmlApplicationFeeTypeListNode = Nothing
    Set xmlApplicationFeeTypeList = Nothing
    Set xmlAppFeeType = Nothing
    Set xmlResponseElem = Nothing
    Set xmlMtgeOneOffCostsForCalc = Nothing
    Set xmlManipulatedApplicationFeeType = Nothing
    Set xmlManipulatedApplicationFeeTypeList = Nothing
    Set xmlNewCostElement = Nothing
    Set xmlRootElement = Nothing
    Set xmlValidationType = Nothing
    Set xmlCost = Nothing
  
    Exit Function

ProcessAllFeesForRemodelQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteTxBO_ProcessAllFeesForRemodelQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume ProcessAllFeesForRemodelQuoteExit

End Function



