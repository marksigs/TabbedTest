VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "IncomeCalcsTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Workfile:      IncomeCalcsTxBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Income Calculations Transactional Business Object which 'requires
'               transactions'.

'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransaction
'-------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
'
' MDC   01/11/2002  BMIDS00654 Maximum Borrowing Calculations
' MDC   08/11/2002  BMIDS00875 Selection of Loan component for multi component loans
' MDC   19/12/02    BM0206 Add option to throw error in method GetAcceptedOrActiveQuoteData
' GHun  02/07/2003  BMIDS591 CC040 Changed CalculateMaxBorrowing
'-------------------------------------------------------------------------------------------
'
'MARS History:
'Prog   Date        AQR     Description
'MF     15/09/2005  MAR30   Changed RunIcomeCalculation to reads global parameter to determine whether a tolerance
'                           check is required to be called. RunToleranceCheck is a new function.
'JD     27/09/2005  MAR30   Changes for WP10
'GHun   22/10/2005  MAR271  Changed RunToleranceCheck and CalculateMaximumBorrowing
'MV     25/10/2005  MAR282  Amended IIncomeCalcsTxBO_RunIncomeCalculation
'MV     26/10/2005  MAR282  Amended IIncomeCalcsTxBO_RunIncomeCalculation Removed unwanted line
'MV     31/10/2005  MAR282  Amended IIncomeCalcsTxBO_RunIncomeCalculation()
'MV     31/10/2005  MAR282  Amended IIncomeCalcsTxBO_RunIncomeCalculation()
'IK     03/11/2005  MAR510  Amended IIncomeCalcsTxBO_RunIncomeCalculation(), ensure correct results returned
'Maha T 14/11/2005  MAR303  Amended IIncomeCalcsTxBO_CalculateMaximumBorrowing(), get customer 2 income details
'Maha T 15/11/2005  MAR272  Fixed Error on Tolerance
'PJO    18/11/2005  MAR440  Changed CalculateMaximumBorrowing and TunTolerance check
'                           to ignore Record Not Found
' JD    20/12/2005  MAR886  Added second tolerance check to RunIncomeTolerance for Decline task.
'JD     25/01/2006  MAR1040 Changed CalculateMaximumBorrowing to calculate using both declared and confirmed income
'JD     08/02/2006  MAR1040 Use Net Confirmed amount in max borrowing calc
'JD     24/02/2006  MAR1314 Changed income multiple calc.
'PJO    03/03/2006  MAR1354 Allow for partial confirmation
'HMA    13/03/2006  MAR1415 Changes to CalculateMaximumBorrowing.
'JD     15/03/2006  MAR1415 CalculateMaximumBorrowing - set flag to save confirmed values for single app
'JD     16/03/2006  MAR1415 CalculateMaximumBorrowing - deal with case were 1 app on joint app has no earnings.
'DRC    25/03/2006  MAR1445 Alternative MaxBorrowing in Tolerance check
'DRC    27/03/2006  MAR1445 Revisited - trap division by zero ( max borrowing)
'HMA    04/04/2006  MAR1546/1570 Further changes to RunIncomeCalculation and RunToleranceCheck.
'PE     10/04/2006  MAR1546 IncomeCalcs failing on DC160. (Because there is no quote)
'PE     10/04/2006  MAR1546 IncomeCalcs failing on DC160.
'PSC    08/06/2006  MAR1863 Correct call to GetCurrentStage
'JD     09/06/2006  MAR1853 RunToleranceCheck - if tolerance not exceeded check if there is an outstanding tolerance exceeded task and set it to not applicable.
'JD     13/06/06    MAR688  Changes for performance improvement.
'GHun   19/06/2006  MAR688  Changed GetMaxBorrowingValues
'JD     20/06/2006  MAR1876 RunToleranceCheck - if a tolerance referral task is required, make any existing tolerance decline tasks N/A and vice versa.
'GHun   10/07/2006  MAR1891 Changed to reduce locking and increase performance
'GHun   19/07/2006  MAR1891 Changed TMIncUndRefIDExistsIncomplete to add missing quotes in XPath
'-------------------------------------------------------------------------------------------

Implements IIncomeCalcsTxBO
Implements ObjectControl

Private m_objXMLAssist As XMLAssist
Private m_objErrAssist As ErrAssist
Private m_objContext As ObjectContext

Private Const oeNoAcceptedOrActiveSubQuote = 239

Private Function IIncomeCalcsTxBO_CalculateAllowableIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAllowableIncomeVbErr
    
    Const strFunctionName As String = "IIncomeCalcsTxBO_CalculateAllowableIncome"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objApplicationQuoteBO As ApplicationQuoteBO
    
    Dim strResponse As String
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objApplicationQuoteBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")

    strResponse = objApplicationQuoteBO.CalcAllowableInc(vxmlRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    Set IIncomeCalcsTxBO_CalculateAllowableIncome = xmlTempResponseDoc.documentElement
    m_objContext.SetComplete
    
CalculateAllowableIncomeExit:
    Set objApplicationQuoteBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseDoc = Nothing
    
    Exit Function

CalculateAllowableIncomeVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsTxBO_CalculateAllowableIncome = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CalculateAllowableIncomeExit

End Function

Private Function IIncomeCalcsTxBO_CalculateMaximumBorrowing(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateMaximumBorrowingVbErr
    
    Const strFunctionName As String = "IIncomeCalcsTxBO_CalculateMaximumBorrowing"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    'Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlMortgageSubQuote As IXMLDOMNode
    Dim xmlIncomeMultipliers As IXMLDOMNode
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlAllowableIncome As IXMLDOMNode
    Dim xmlCustomer1 As IXMLDOMNode
    Dim xmlCustomer2 As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlResponseHeadingNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute   'BM0206 MDC 19/12/2002
    Dim xmlMSQNode As IXMLDOMElement 'JD
    
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    'Dim objParameterDO As GlobalParameterDO
    Dim objCustomerFinancialBO As CustomerFinancialBO
    
    Dim strResponse As String
    'Dim strIncomeMultiplierCode As String
    Dim dblMultiplier As Double
    Dim dblMultiplierSingle As Double 'JD
    Dim dblHighestMultiplier As Double
    Dim dblLowestMultiplier As Double
    Dim dblIncome1 As Double
    Dim dblIncome2 As Double
    Dim dblMaxBorrowing As Double
    Dim dblMaxBorrowing1 As Double
    Dim dblMaxBorrowing2 As Double
    Dim lngAmountRequested As Long
    Dim strIncomeMultipleType As String
    Dim strCustomerNumber1 As String
    Dim strCustomerNumber2 As String
    Dim strCustomerVersionNumber1 As String
    Dim strCustomerVersionNumber2 As String
    Dim lngTotalLoanAmount As Long
    Dim lngBalancesOutstanding As Long
    'Dim blnCust1HighestEarner As Boolean
    Dim dblTempIncome As Double
    Dim bUsedConfirmedIncome As Boolean  'JD MAR30
    'JD MAR1040 start
    Dim dblConfirmedIncome1 As Double
    Dim dblConfirmedIncome2 As Double
    Dim dblConfirmedMaxBorrowing As Double
    Dim dblConfirmedMultiplier As Double
    Dim strConfirmedIncomeMultipleType As String
    Dim bFoundBalances As Boolean
    'JD MAR1040 end
    Dim bNoMortgageSubquoteForApp As Boolean 'JD MAR688
    Dim strMortgageSubquoteNumber As String 'JD MAR688
    '==================================================================================
    ' Initialise
    '==================================================================================
    bUsedConfirmedIncome = False  'JD MAR30
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
'    Set objParameterBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    'Set objParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    Set objCustomerFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL_COMPONENT & ".CustomerFinancialBO")
    
    '==================================================================================
    ' Determine Allowable Income and Income Multiple
    '==================================================================================
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequest, "INCOMECALCULATION/APPLICATION").cloneNode(True)
    'PJO 18/11/2005 - This is now set as default
    'strIncomeMultiplierCode = objParameterDO.GetCurrentParameterByType("DefaultIncomeMultiplierCode", "STRING")
    
    'MAR1891 GHun GetMaxBorrowingValues is now called from the BO and the results are included in the request
    ''JD use new storedproc to get data
    'Dim strMaxBorrowingValues As String
    'strMaxBorrowingValues = GetMaxBorrowingValues(xmlGetNodeText(xmlApplicationNode, "APPLICATIONNUMBER"), _
    '                        xmlGetNodeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER"))
    'Set xmlTempDoc = xmlLoad(strMaxBorrowingValues, strFunctionName)
    'Set xmlTempNode = xmlTempDoc.selectSingleNode("RESPONSE/MAXBORROWINGVALUES")
    Set xmlTempNode = vxmlRequest.selectSingleNode("MAXBORROWINGVALUES")
    'MAR1891 GHun
    
    If Not xmlTempNode Is Nothing Then
        bNoMortgageSubquoteForApp = False
        strMortgageSubquoteNumber = m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "MORTGAGESUBQUOTENUMBER")
        lngAmountRequested = CSafeLng(m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "AMOUNTREQUESTED"))
        'strIncomeMultiplierCode = m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "INCOMEMULTIPLIERCODE")
        dblHighestMultiplier = CSafeDbl(m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "HIGHESTEARNERMULTIPLE"))
        dblLowestMultiplier = CSafeDbl(m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "LOWESTEARNERMULTIPLE"))
        dblMultiplier = CSafeDbl(m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "JOINTINCOMEMULTIPLE"))
        dblMultiplierSingle = CSafeDbl(m_objXMLAssist.GetMandatoryAttribute(xmlTempNode, "SINGLEINCOMEMULTIPLE"))
    Else
        bNoMortgageSubquoteForApp = True
    End If

    'Calculate Allowable Income
    'MAR1891 GHun Retrieve CalcAllowableInc response
    Set xmlAllowableIncome = vxmlRequest.selectSingleNode("CALCALLOWABLEINC/RESPONSE")

    'JD MAR30 see is confirmed income was used
    'Dim strConfirmedAllowableIncome As String
    'MAR272 GHun bUsedConfirmedIncome should only be true if CONFIRMEDALLOWABLEINCOME = 1
    'bUsedConfirmedIncome = m_objXMLAssist.GetElementText(xmlAllowableIncome, ".//CONFIRMEDALLOWABLEINCOME", strConfirmedAllowableIncome)
    'strConfirmedAllowableIncome = m_objXMLAssist.GetElementText(xmlAllowableIncome, ".//CONFIRMEDALLOWABLEINCOME")
    'JD MAR1040 bUsedConfirmedIncome = (strConfirmedAllowableIncome = 1)
    'PJO MAR1354 03/03/2006 - This is now set later on
    'If strConfirmedAllowableIncome <> "" Then
    '    bUsedConfirmedIncome = True
    'End If
    'MAR272 End

    '==================================================================================
    ' Calculate the initial Maximum Borrowing
    '==================================================================================
    
    'Is this a single or joint application?
    If vxmlRequest.selectNodes("INCOMECALCULATION/CUSTOMERLIST/CUSTOMER").length > 1 Then
        'Joint application (only use first two applicants)
        If (Not bNoMortgageSubquoteForApp) Then 'MAR1546
            Set xmlCustomer1 = xmlGetMandatoryNode(vxmlRequest, "INCOMECALCULATION/CUSTOMERLIST/CUSTOMER[1]")
            Set xmlCustomer2 = xmlGetMandatoryNode(vxmlRequest, "INCOMECALCULATION/CUSTOMERLIST/CUSTOMER[2]")   '' MAR303 (Changed CUSTOMER[1] to CUSTOMER[2])
            strCustomerNumber1 = xmlGetMandatoryNodeText(xmlCustomer1, "CUSTOMERNUMBER")
            strCustomerVersionNumber1 = xmlGetMandatoryNodeText(xmlCustomer1, "CUSTOMERVERSIONNUMBER")
            strCustomerNumber2 = xmlGetMandatoryNodeText(xmlCustomer2, "CUSTOMERNUMBER")
            strCustomerVersionNumber2 = xmlGetMandatoryNodeText(xmlCustomer2, "CUSTOMERVERSIONNUMBER")
            
            'Find New Annual Allowable Incomes
            dblIncome1 = m_objXMLAssist.GetMandatoryElementText(xmlAllowableIncome, _
                ".//LENDERDETAILS/CUSTOMERLIST/CUSTOMER[CUSTOMERNUMBER='" & strCustomerNumber1 & _
                "' and CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber1 & _
                "']/NETANNUALALLOWABLEINCOME")
            dblIncome2 = m_objXMLAssist.GetMandatoryElementText(xmlAllowableIncome, _
                ".//LENDERDETAILS/CUSTOMERLIST/CUSTOMER[CUSTOMERNUMBER='" & strCustomerNumber2 & _
                "' and CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber2 & _
                "']/NETANNUALALLOWABLEINCOME")
            'JD MAR1040 get confirmed income if present
            dblConfirmedIncome1 = CSafeDbl(m_objXMLAssist.GetElementText(xmlAllowableIncome, _
                ".//LENDERDETAILS/CUSTOMERLIST/CUSTOMER[CUSTOMERNUMBER='" & strCustomerNumber1 & _
                "' and CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber1 & _
                "']/NETCONFIRMEDALLOWABLEINCOME")) 'JD use Net value
            dblConfirmedIncome2 = CSafeDbl(m_objXMLAssist.GetElementText(xmlAllowableIncome, _
                ".//LENDERDETAILS/CUSTOMERLIST/CUSTOMER[CUSTOMERNUMBER='" & strCustomerNumber2 & _
                "' and CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber2 & _
                "']/NETCONFIRMEDALLOWABLEINCOME")) 'JD use Net value
        End If
        
        'PJO MAR1354 03/03/2006 - Set boolean if both incomes confirmed
        If dblConfirmedIncome1 > 0 And dblConfirmedIncome2 > 0 Then
            bUsedConfirmedIncome = True
        ElseIf ((dblConfirmedIncome1 > 0 Or dblConfirmedIncome2 > 0) And _
                 (dblIncome1 = 0 Or dblIncome2 = 0)) Then
            'JD MAR1415
            'calculate confirmed max borrowing if we have one confirmed income
            'and one declared income which is zero (ie one applicant doesn't have an income)
            'In this case we assume one applicant doesn't earn so will never have a confirmed income
            bUsedConfirmedIncome = True
        End If
       
        'Find max amount using indivdual incomes
        'JD MAR1040 declared
        'PJO MAR1354 03/03/2006 allow for part confirmed, part declared
        ' We use the confirmed income in the unconfirmed calc unless both confirmed
        ' This section rewritten to allow for this
        If bUsedConfirmedIncome = False Then
            Dim dblCalcIncome1 As Double
            Dim dblCalcIncome2 As Double
            If dblConfirmedIncome1 > 0 Then
                dblCalcIncome1 = dblConfirmedIncome1
            Else
                dblCalcIncome1 = dblIncome1
            End If
            If dblConfirmedIncome2 > 0 Then
                dblCalcIncome2 = dblConfirmedIncome2
            Else
                dblCalcIncome2 = dblIncome2
            End If

            If dblCalcIncome1 < dblCalcIncome2 Then
                'Swap income values over
                dblTempIncome = dblCalcIncome1
                dblCalcIncome1 = dblCalcIncome2
                dblCalcIncome2 = dblTempIncome
            End If
            dblMaxBorrowing1 = (dblCalcIncome1 * dblHighestMultiplier) + _
                               (dblCalcIncome2 * dblLowestMultiplier)
            
            'Find max amount using joint income
            dblMaxBorrowing2 = (dblCalcIncome1 + dblCalcIncome2) * dblMultiplier
        
            'Calculate the best type of multiplier to use
            If dblMaxBorrowing1 > dblMaxBorrowing2 Then
                dblMaxBorrowing = dblMaxBorrowing1
                'JD MAR1314 strIncomeMultipleType = "S"
                strIncomeMultipleType = "H" 'JD MAR1314 H = Highest Lowest
            Else
                dblMaxBorrowing = dblMaxBorrowing2
                strIncomeMultipleType = "J"
            End If
                
        'JD MAR1040 confirmed
        Else        'If bUsedConfirmedIncome = True Then
            'Find max amount using indivdual incomes
            If dblConfirmedIncome1 < dblConfirmedIncome2 Then
                'Swap income values over
                dblTempIncome = dblConfirmedIncome1
                dblConfirmedIncome1 = dblConfirmedIncome2
                dblConfirmedIncome2 = dblTempIncome
            End If
            dblMaxBorrowing1 = (dblConfirmedIncome1 * dblHighestMultiplier) + (dblConfirmedIncome2 * dblLowestMultiplier)
                
            'Find max amount using joint income
            dblMaxBorrowing2 = (dblConfirmedIncome1 + dblConfirmedIncome2) * dblMultiplier
            
            'Calculate the best type of multiplier to use
            If dblMaxBorrowing1 > dblMaxBorrowing2 Then
                dblConfirmedMaxBorrowing = dblMaxBorrowing1
                'JD MAR1314 strConfirmedIncomeMultipleType = "S"
                strConfirmedIncomeMultipleType = "H" 'JD MAR1314 H = Highest Lowest
            Else
                dblConfirmedMaxBorrowing = dblMaxBorrowing2
                strConfirmedIncomeMultipleType = "J"
            End If
            
            'MAR1415 Update the declared fields too
            If dblIncome1 < dblIncome2 Then
                'Swap income values over
                dblTempIncome = dblIncome1
                dblIncome1 = dblIncome2
                dblIncome2 = dblTempIncome
            End If
            dblMaxBorrowing1 = (dblIncome1 * dblHighestMultiplier) + _
                               (dblIncome2 * dblLowestMultiplier)
            
            'Find max amount using joint income
            dblMaxBorrowing2 = (dblIncome1 + dblIncome2) * dblMultiplier
        
            'Calculate the best type of multiplier to use
            If dblMaxBorrowing1 > dblMaxBorrowing2 Then
                dblMaxBorrowing = dblMaxBorrowing1
                strIncomeMultipleType = "H" ' H = Highest Lowest
            Else
                dblMaxBorrowing = dblMaxBorrowing2
                strIncomeMultipleType = "J"
            End If
  
        End If
        
    Else
        'Single applicant
        'dblMultiplier = CSafeDbl(m_objXMLAssist.GetMandatoryElementText(xmlIncomeMultipliers, _
        '                                                    ".//SINGLEINCOMEMULTIPLE"))
        dblMultiplier = dblMultiplierSingle

        dblIncome1 = CSafeDbl(m_objXMLAssist.GetMandatoryElementText(xmlAllowableIncome, _
                                                            ".//NETANNUALALLOWABLEINCOME"))
        'JD MAR1040 get confirmed too if present
        dblConfirmedIncome1 = CSafeDbl(m_objXMLAssist.GetElementText(xmlAllowableIncome, _
                                                            ".//NETCONFIRMEDALLOWABLEINCOME"))
        dblIncome2 = 0
        dblConfirmedIncome2 = 0
        
        dblMaxBorrowing = dblIncome1 * dblMultiplier
        strIncomeMultipleType = "S"
        
        ' PJO MAR1354 03/03/2006 - Boolean only valid if joint case
        'If bUsedConfirmedIncome = True Then
        If dblConfirmedIncome1 > 0 Then
            dblConfirmedMaxBorrowing = dblConfirmedIncome1 * dblMultiplier
            strConfirmedIncomeMultipleType = "S"
            bUsedConfirmedIncome = True 'JD MAR1415 set it here so that the values get saved.
        End If
    
    End If
    
    
    '==================================================================================
    ' Calculate the Income Multiplier(s) in relation to the requested loan amount
    '==================================================================================
    bFoundBalances = False
    If dblMaxBorrowing > 0 Then
        
        'Find outstanding balances on BMBS mortgages not being redeemed
        Set xmlTempRequestNode = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        m_objXMLAssist.CopyNode "APPLICATION", "MORTGAGEACCOUNT", xmlApplicationNode, xmlTempRequestNode, True
        'BMIDS591 GHun 02/07/2003
        m_objXMLAssist.SetMandatoryChildText xmlTempRequestNode.selectSingleNode("MORTGAGEACCOUNT"), "GETBMONLYNOTTOBEREDEEMEDLOANS", "0"
        'BMIDS591 End
        strResponse = objCustomerFinancialBO.GetMtgAcctDataForFurtherAdvance(xmlTempRequestNode.xml)
        Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
        
        lngBalancesOutstanding = CSafeLng( _
                m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc.documentElement, _
                "MORTGAGEACCOUNT/TOTALOUTSTANDINGBALANCE"))
        lngTotalLoanAmount = lngAmountRequested + lngBalancesOutstanding
        bFoundBalances = True
        
        ' PJO MAR1354 03/03/2006 - as we are dividing by the sum, we should check for the sum
        If dblIncome1 + dblIncome2 = 0 Then
            'Raise error to avoid division by zero
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Income = 0"
        End If
        
        'JD MAR1314
'        If strIncomeMultipleType = "S" Then
'            dblMultiplier = (lngTotalLoanAmount - dblIncome2) / dblIncome1
'        Else
'            dblMultiplier = lngTotalLoanAmount / (dblIncome1 + dblIncome2)
'        End If
        dblMultiplier = lngTotalLoanAmount / (dblIncome1 + dblIncome2)
        
    Else
        'Cannot derive income multiple
        dblMultiplier = 0
        strIncomeMultipleType = "U" 'Unknown
        'JD MAR1040 if max borrowing is -ve then set as 0
        dblMaxBorrowing = 0
    End If
    
    'JD MAR1040 do confirmed as well
    If dblConfirmedMaxBorrowing > 0 Then
        
        If bFoundBalances = False Then
            'Find outstanding balances on BMBS mortgages not being redeemed
            Set xmlTempRequestNode = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
            m_objXMLAssist.CopyNode "APPLICATION", "MORTGAGEACCOUNT", xmlApplicationNode, xmlTempRequestNode, True
            'BMIDS591 GHun 02/07/2003
            m_objXMLAssist.SetMandatoryChildText xmlTempRequestNode.selectSingleNode("MORTGAGEACCOUNT"), "GETBMONLYNOTTOBEREDEEMEDLOANS", "0"
            'BMIDS591 End
            strResponse = objCustomerFinancialBO.GetMtgAcctDataForFurtherAdvance(xmlTempRequestNode.xml)
            Set xmlTempResponseDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
            
            lngBalancesOutstanding = CSafeLng( _
                    m_objXMLAssist.GetMandatoryElementText(xmlTempResponseDoc.documentElement, _
                    "MORTGAGEACCOUNT/TOTALOUTSTANDINGBALANCE"))
            lngTotalLoanAmount = lngAmountRequested + lngBalancesOutstanding
        End If
        
        ' PJO MAR1354 03/03/2006 - as we are dividing by the sum, we should check for the sum
        'If dblConfirmedIncome1 = 0 Then
        If dblConfirmedIncome1 + dblConfirmedIncome2 = 0 Then
            'Raise error to avoid division by zero
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Income = 0"
        End If
        
        'JD MAR1314
'        If strConfirmedIncomeMultipleType = "S" Then
'            dblConfirmedMultiplier = (lngTotalLoanAmount - dblConfirmedIncome2) / dblConfirmedIncome1
'        Else
'            dblConfirmedMultiplier = lngTotalLoanAmount / (dblConfirmedIncome1 + dblConfirmedIncome2)
'        End If
        dblConfirmedMultiplier = lngTotalLoanAmount / (dblConfirmedIncome1 + dblConfirmedIncome2)
        
    Else
        'Cannot derive income multiple
        dblConfirmedMultiplier = 0
        strConfirmedIncomeMultipleType = "U" 'Unknown
        'JD MAR1040 if max borrowing is -ve then set as 0
        dblConfirmedMaxBorrowing = 0
    End If
    
    
    '==================================================================================
    ' Calculate the final Maximum Borrowing & save values
    '==================================================================================
    ' JD MAR30 do  not reduce the max borrowing amount by the outstanding balances
    'dblMaxBorrowing = dblMaxBorrowing - lngBalancesOutstanding + 0.5
    
    If (Not bNoMortgageSubquoteForApp) Then 'MAR1546
        'Save Mortgage Sub Quote
        Set xmlTempRequestNode = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlMSQNode = xmlOut.createElement("MORTGAGESUBQUOTE")
        m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "APPLICATIONNUMBER", m_objXMLAssist.GetElementText(xmlApplicationNode, "APPLICATIONNUMBER")
        m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "APPLICATIONFACTFINDNUMBER", m_objXMLAssist.GetElementText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
        m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "MORTGAGESUBQUOTENUMBER", strMortgageSubquoteNumber
        xmlTempRequestNode.appendChild xmlMSQNode
       
        'xmlTempRequestNode.appendChild xmlMortgageSubQuote
        'INCOMEMULTIPLIERTYPE, INCOMEMULTIPLE & MAXIMUMBORROWINGAMOUNT
        'JD MAR30 save different values dependanton UsedConfirmedIncome
        
        m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "INCOMEMULTIPLIERTYPE", strIncomeMultipleType
        m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "INCOMEMULTIPLE", Format$(dblMultiplier, "##0.00")
        m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "MAXIMUMBORROWINGAMOUNT", CStr(CSafeLng(dblMaxBorrowing))
        'MAR1415 save confirmed if we have confirmed income
        'If bUsedConfirmedIncome = True Then
        If (dblConfirmedIncome1 > 0 Or dblConfirmedIncome2 > 0) Then
            m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "CONFIRMEDCALCULATEDINCMULTIPLIERTYPE", strConfirmedIncomeMultipleType
            m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "CONFIRMEDCALCULATEDINCMULTIPLE", Format$(dblConfirmedMultiplier, "##0.00")
            m_objXMLAssist.SetMandatoryChildText xmlMSQNode, "CONFIRMEDMAXBORROWING", CStr(CSafeLng(dblConfirmedMaxBorrowing))
        End If
        
        strResponse = objMortgageSubQuoteBO.Update(xmlTempRequestNode.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    End If
    
    '==================================================================================
    ' Formulate the response
    '==================================================================================
    Set xmlResponseHeadingNode = xmlOut.createElement("INCOMECALCULATION")
    xmlResponseElem.appendChild xmlResponseHeadingNode
    Set xmlTableNode = xmlOut.createElement("ALLOWABLEINCOME")
    xmlResponseHeadingNode.appendChild xmlTableNode
    Set xmlTempNode = m_objXMLAssist.GetMandatoryNode(xmlAllowableIncome, "LENDERLIST")
    xmlTableNode.appendChild xmlTempNode
    Set xmlTableNode = xmlOut.createElement("MAXIMUMBORROWING")
    xmlResponseHeadingNode.appendChild xmlTableNode
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "INCOME1", CStr(dblIncome1)
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "INCOME2", CStr(dblIncome2)
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "INCOMEMULTIPLIERTYPE", strIncomeMultipleType
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "INCOMEMULTIPLE", Format$(dblMultiplier, "##0.00")
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "MAXIMUMBORROWINGAMOUNT", CStr(CSafeLng(dblMaxBorrowing))
    'JD MAR1040 confirmed
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "CONFIRMEDINCOME1", CStr(dblConfirmedIncome1)
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "CONFIRMEDINCOME2", CStr(dblConfirmedIncome2)
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "CONFIRMEDINCOMEMULTIPLIERTYPE", strConfirmedIncomeMultipleType
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "CONFIRMEDINCOMEMULTIPLE", Format$(dblConfirmedMultiplier, "##0.00")
    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "CONFIRMEDMAXIMUMBORROWINGAMOUNT", CStr(CSafeLng(dblConfirmedMaxBorrowing))

    m_objXMLAssist.SetMandatoryChildText xmlTableNode, "AMOUNTREQUESTED", CStr(lngAmountRequested)
    
    Set IIncomeCalcsTxBO_CalculateMaximumBorrowing = xmlResponseElem
    m_objContext.SetComplete
    
CalculateMaximumBorrowingExit:
    Set xmlOut = Nothing
    Set xmlTempResponseDoc = Nothing
    'Set xmlTempDoc = Nothing
    Set xmlTempRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlMortgageSubQuote = Nothing
    Set xmlIncomeMultipliers = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlAllowableIncome = Nothing
    Set xmlCustomer1 = Nothing
    Set xmlCustomer2 = Nothing
    Set xmlApplicationNode = Nothing
    Set objMortgageSubQuoteBO = Nothing
    'Set objParameterDO = Nothing
    Set objCustomerFinancialBO = Nothing
    Set xmlResponseHeadingNode = Nothing
    Set xmlAttrib = Nothing 'BM0206 MDC 19/12/2002
    
    Exit Function

CalculateMaximumBorrowingVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsTxBO_CalculateMaximumBorrowing = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CalculateMaximumBorrowingExit

End Function

Private Function IIncomeCalcsTxBO_RunIncomeCalculation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

    On Error GoTo RunIncomeCalculationVbErr
    
    Const strFunctionName As String = "IIncomeCalcsTxBO_RunIncomeCalculation"

    Dim xmlICNode As IXMLDOMNode               ' MAR1570
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    
    Dim xmlMaxBorrowingNode As IXMLDOMNode     ' MAR1570
    
    Dim objParameterDO As GlobalParameterDO
    
    Dim lngErr As Long
    Dim strValue As String
    
    Set xmlOut = xmlCreateDOMObject()
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlICNode = m_objXMLAssist.GetMandatoryNode(vxmlRequest, "INCOMECALCULATION")   ' MAR1570
    strValue = m_objXMLAssist.GetAttributeFromNode(xmlICNode, "CALCULATEMAXBORROWING")  ' MAR1570
    If strValue = "1" Then
        Set xmlResponseNode = IIncomeCalcsTxBO_CalculateMaximumBorrowing(vxmlRequest)
        lngErr = m_objErrAssist.CheckResponse(xmlResponseNode, xmlResponseElem, False)
        If lngErr <> 0 Then
            lngErr = m_objErrAssist.GetOmigaErrorNumber(lngErr)
            If lngErr <> oeRecordNotFound Then
                If lngErr = oeNoAcceptedOrActiveSubQuote Or lngErr = oeRecordNotFound Then
                    'Sub Quote not found so just do Allowable Income
                    'MAR1891 GHun Return CalcAllowableInc response
                    Set xmlResponseNode = vxmlRequest.selectSingleNode("CALCALLOWABLEINC/RESPONSE")
                Else
                    'Re raise error
                    m_objErrAssist.CheckResponse xmlResponseNode, xmlResponseElem, True
                End If
            End If
        Else
            'MAR1570 Save Maximum Borrowing information to send to Tolerance check.
            Set xmlMaxBorrowingNode = xmlResponseNode.selectSingleNode("//MAXIMUMBORROWING")
            'Peter Edney - 11/04/2006
            'Not cloning the node resulted in stripping the node from the response.
            'xmlICNode.appendChild xmlMaxBorrowingNode
            xmlICNode.appendChild xmlMaxBorrowingNode.cloneNode(True)
        End If
    Else
        'MAR1891 GHun Return CalcAllowableInc response
        Set xmlResponseNode = vxmlRequest.selectSingleNode("CALCALLOWABLEINC/RESPONSE")
    End If
    
    'IK 03/11/2005 MAR510 - do not return result of RunToleranceCheck
    Set IIncomeCalcsTxBO_RunIncomeCalculation = xmlResponseNode
    
    '~ MF 15/09/2005 WP10 MAR30.
    '~ Check global parameter to see if tolerance check calc is required
    If lngErr <> oeRecordNotFound Then
        Set objParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        If objParameterDO.GetCurrentParameterByType("IncMultToleranceChk", "BOOLEAN") Then
        
            Set xmlResponseNode = IIncomeCalcsTxBO_RunToleranceCheck(vxmlRequest)
            m_objErrAssist.CheckResponse xmlResponseNode, xmlResponseElem, True
        End If
    End If
    
    m_objContext.SetComplete
    
RunIncomeCalculationExit:
    Set xmlICNode = Nothing               ' MAR1570
    Set xmlMaxBorrowingNode = Nothing     ' MAR1570
    Set xmlResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    
    Exit Function

RunIncomeCalculationVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsTxBO_RunIncomeCalculation = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RunIncomeCalculationExit

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~ Created      22/09/2005
'~ Prog         MFoster
'~ AQR          MAR30
'~ Description  Determines if the Tolerance has been exceeded. If so then the task
'~              "exceeded tolerance task" is added if it does not already exist in an
'~              incomplete State.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function IIncomeCalcsTxBO_RunToleranceCheck(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo IIncomeCalcsTxBO_RunToleranceCheckVbErr
    
    Const strFunctionName As String = "IIncomeCalcsTxBO_RunToleranceCheck"
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim strResponse As String
    Dim strCaseTaskDetailsResponse As String 'JD MAR1876
    Dim sConfCalcIncMultType As String
    Dim dCalcIncMultiple As Double
    Dim dProdIncMultiple As Double
    Dim dTolerance As Double
    Dim objParameterDO As GlobalParameterDO
    Dim objTMBO As OmTmBO
    Dim strApplicationNumber As String
    'Dim strAppFFNumber As String
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim sTaskID As String
    Dim sOtherTaskId As String 'JD MAR1876
    Dim objMsgTmBO As MsgTmBO
    Dim sActivityID As String

    Dim dTolerance2 As Double  'JD 20/12/2005 MAR886
    Dim dblAmountRequested As Double 'drc 25/03/2006 MAR1445
    Dim dblMaxBorrowing As Double 'drc 25/03/2006 MAR1445
    Dim sToleranceBasis As String 'drc 25/03/2006 MAR1445
    
    'JD MAR1853
    Dim xmlThisCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskElem As IXMLDOMElement
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode 'JD MAR688
    
    Dim strTMIncUndRefID As String  'MAR1891 GHun
    Dim xmlTempRequest2 As IXMLDOMElement   'MAR1891 GHun

    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '~ Read global parameter for income multiplier tolerance
    Set objParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    'MAR272 GHun IncomeMultTolerance is a percentage not an amount
    dTolerance = objParameterDO.GetCurrentParameterByType("IncomeMultTolerance", "PERCENTAGE")
    dTolerance2 = objParameterDO.GetCurrentParameterByType("IncomeMultToDecline", "PERCENTAGE") 'MAR886
    Set xmlTempRequest = vxmlRequest.cloneNode(True)
    '~ Reposition parameters accordingly
    Set xmlNode = vxmlRequest.selectSingleNode("INCOMECALCULATION/APPLICATION")
    Set xmlRequestNode = xmlGetRequestNode(vxmlRequest)
    Call xmlRequestNode.appendChild(xmlNode)
    
    '~ Read App no & App FF no
    Set xmlTempResponseDoc = xmlLoad(xmlRequestNode.xml, strFunctionName)
    If Not xmlTempResponseDoc.selectSingleNode("REQUEST/APPLICATION/APPLICATIONNUMBER") Is Nothing Then
        strApplicationNumber = xmlTempResponseDoc.selectSingleNode("REQUEST/APPLICATION/APPLICATIONNUMBER").Text
    End If
    'If Not xmlTempResponseDoc.selectSingleNode("REQUEST/APPLICATION/APPLICATIONFACTFINDNUMBER") Is Nothing Then
    '    strAppFFNumber = xmlTempResponseDoc.selectSingleNode("REQUEST/APPLICATION/APPLICATIONFACTFINDNUMBER").Text
    'End If
    If Not xmlTempResponseDoc.documentElement.Attributes.getNamedItem("ACTIVITYID") Is Nothing Then
        sActivityID = xmlTempResponseDoc.documentElement.Attributes.getNamedItem("ACTIVITYID").Text
    Else
        sActivityID = "10"
    End If
    Set xmlTempRequest2 = xmlRequestNode.cloneNode(True)
    
    strTMIncUndRefID = objParameterDO.GetCurrentParameterByType("TMIncUndRefID", "STRING")  'MAR1891 GHun
    
    'MAR1445 - tolerance based in income multiple or maximum borrowing
    sToleranceBasis = objParameterDO.GetCurrentParameterByType("IncomeTolerenceMethod", "STRING")
    sToleranceBasis = UCase$(sToleranceBasis)
    
    If sToleranceBasis <> "MAXBORROW" Then
        
        'Calculated Income multiple vs product Income multiple
        
        'MAR1891 GHun GetMaxBorrowingValues is now called from the BO and the results are included in the request
        'JD MAR688 use stored procedure to get all values
        'Dim strMaxBorrowingValues As String
        'strMaxBorrowingValues = GetMaxBorrowingValues(strApplicationNumber, strAppFFNumber)
        'Set xmlTempResponseDoc = xmlLoad(strMaxBorrowingValues, strFunctionName)    'MAR1891 GHun
        'Set xmlTempNode = xmlTempResponseDoc.selectSingleNode("RESPONSE/MAXBORROWINGVALUES")
        Set xmlTempNode = vxmlRequest.selectSingleNode("MAXBORROWINGVALUES")
        'MAR1891 End
        
        If Not xmlTempNode Is Nothing Then
            sConfCalcIncMultType = xmlGetAttributeText(xmlTempNode, "CONFIRMEDCALCULATEDINCMULTIPLIERTYPE")
            If Len(sConfCalcIncMultType) = 0 Or sConfCalcIncMultType = "U" Then
                sConfCalcIncMultType = m_objXMLAssist.GetAttributeValue(xmlTempNode, "MAXBORROWINGVALUES", "INCOMEMULTIPLIERTYPE")
                dCalcIncMultiple = Val(m_objXMLAssist.GetAttributeValue(xmlTempNode, "MAXBORROWINGVALUES", "INCOMEMULTIPLE"))
            Else
                dCalcIncMultiple = Val(m_objXMLAssist.GetAttributeValue(xmlTempNode, "MAXBORROWINGVALUES", "CONFIRMEDCALCULATEDINCMULTIPLE"))
            End If
            ' pick up the correct multiplier
            Select Case sConfCalcIncMultType
                Case "S"
                    dProdIncMultiple = xmlGetAttributeAsDouble(xmlTempNode, "SINGLEINCOMEMULTIPLE")
                Case "J"
                    dProdIncMultiple = xmlGetAttributeAsDouble(xmlTempNode, "JOINTINCOMEMULTIPLE")
                Case Else
                    dProdIncMultiple = 0
            End Select
            '~ Check tolerance level, avoid division by zero error
            If dProdIncMultiple > 0 Then
                If dCalcIncMultiple > dProdIncMultiple Then
                    
                    If (dCalcIncMultiple - dProdIncMultiple) * 100 / dProdIncMultiple > dTolerance Then
                        If (dCalcIncMultiple - dProdIncMultiple) * 100 / dProdIncMultiple > dTolerance2 Then  'MAR886
                            sTaskID = objParameterDO.GetCurrentParameterByType("TMIncomeDeclineID", "STRING")
                            sOtherTaskId = strTMIncUndRefID 'MAR1891 GHun
                        Else
                            sTaskID = strTMIncUndRefID  'MAR1891 GHun
                            sOtherTaskId = objParameterDO.GetCurrentParameterByType("TMIncomeDeclineID", "STRING")
                        End If
                    End If
                 End If
            End If
        End If

    Else
    
        'MAR1445 Tolerance based on Max Borrowing
        'MAR1570 Use the Maximimum Borrowing information passed in the request.
        
        dblAmountRequested = CDbl(vxmlRequest.selectSingleNode(".//AMOUNTREQUESTED").Text)
        
        'MAR1570 Use Confirmed Maximum Borrowing amount if present.
        dblMaxBorrowing = 0
        If Not (vxmlRequest.selectSingleNode(".//CONFIRMEDMAXIMUMBORROWINGAMOUNT") Is Nothing) Then
           dblMaxBorrowing = CDbl(vxmlRequest.selectSingleNode(".//CONFIRMEDMAXIMUMBORROWINGAMOUNT").Text)
        End If
        
        If (dblMaxBorrowing = 0) Then
            If Not (vxmlRequest.selectSingleNode(".//MAXIMUMBORROWINGAMOUNT") Is Nothing) Then
                dblMaxBorrowing = CDbl(vxmlRequest.selectSingleNode(".//MAXIMUMBORROWINGAMOUNT").Text)
            End If
        End If
        
        '~ Check tolerance level, avoid division by zero error if MaxBorrowing  = 0
        If dblMaxBorrowing > 0 Then
          
            If dblAmountRequested > dblMaxBorrowing Then
                    
                If (dblAmountRequested - dblMaxBorrowing) * 100 / dblMaxBorrowing > dTolerance Then
                    If (dblAmountRequested - dblMaxBorrowing) * 100 / dblMaxBorrowing > dTolerance2 Then
                        sTaskID = objParameterDO.GetCurrentParameterByType("TMIncomeDeclineID", "STRING")
                        sOtherTaskId = strTMIncUndRefID  'MAR1891 GHun
                    Else
                        sTaskID = strTMIncUndRefID   'MAR1891 GHun
                        sOtherTaskId = objParameterDO.GetCurrentParameterByType("TMIncomeDeclineID", "STRING")
                    End If
                End If
            End If
        End If
        
    End If
    
    'JD MAR1853 Need to set any outstanding tolerance fail tasks to Not Applicable
    ' if it passed the test this time.
    '~ Get the case task details
    Set objMsgTmBO = m_objContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    'MAR1891 GHun
    Set xmlTempResponseDoc = xmlCreateDOMObject()
    xmlTempResponseDoc.appendChild xmlTempRequest2
    'MAR1891 End
    Call xmlSetAttributeValue(xmlTempResponseDoc.documentElement, "OPERATION", "GetCurrentStage")
    Set xmlNode = xmlTempResponseDoc.createElement("CASEACTIVITY")
    Call xmlSetAttributeValue(xmlNode, "CASEID", strApplicationNumber)
    ' PSC 08/06/2006 MAR1863
    Call xmlSetAttributeValue(xmlNode, "SOURCEAPPLICATION", "Omiga")
    Call xmlSetAttributeValue(xmlNode, "ACTIVITYINSTANCE", "1")
    Call xmlSetAttributeValue(xmlNode, "ACTIVITYID", sActivityID)
    Call xmlTempResponseDoc.documentElement.appendChild(xmlNode)
    strCaseTaskDetailsResponse = objMsgTmBO.TmRequest(xmlTempResponseDoc.xml)
    Set xmlTempResponseDoc = xmlLoad(strCaseTaskDetailsResponse, strFunctionName)
    
    If Not xmlTempResponseDoc.documentElement.selectSingleNode("CASESTAGE") Is Nothing Then
        If Not xmlTempResponseDoc.documentElement.selectSingleNode("CASESTAGE").Attributes.getNamedItem("CASESTAGESEQUENCENO") Is Nothing Then
       
            If Len(sTaskID) > 0 Then
                '~ Check the existing tasks, don't add another if
                '~ its already there and not complete
                If Not TMIncUndRefIDExistsIncomplete(xmlTempResponseDoc, sTaskID) Then
                    '~ Create Adhoc Case Task: Exceeded Tolerance
                    Set objTMBO = m_objContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".OmTmBO")
                    Set xmlRequestNode = xmlGetRequestNode(vxmlRequest)
                    Call xmlSetAttributeValue(xmlRequestNode, "OPERATION", "CREATEADHOCCASETASK")
                    Set xmlNode = xmlRequestNode.ownerDocument.createElement("APPLICATION")
                    Set xmlAppNode = xmlRequestNode.appendChild(xmlNode)
                    Call xmlSetAttributeValue(xmlAppNode, "APPLICATIONPRIORITY", "")
                    Set xmlNode = xmlRequestNode.ownerDocument.createElement("CASETASK")
                    Set xmlCaseTaskNode = xmlRequestNode.appendChild(xmlNode)
                    Call xmlSetAttributeValue(xmlCaseTaskNode, "SOURCEAPPLICATION", "Omiga")
                    Call xmlSetAttributeValue(xmlCaseTaskNode, "CASEID", strApplicationNumber)
                    
                    Call xmlSetAttributeValue(xmlCaseTaskNode, "ACTIVITYID", sActivityID)
                    Call xmlSetAttributeValue(xmlCaseTaskNode, "ACTIVITYINSTANCE", "1")
                    Call xmlSetAttributeValue(xmlCaseTaskNode, "CASESTAGESEQUENCENO", _
                    xmlTempResponseDoc.documentElement.selectSingleNode("CASESTAGE").Attributes.getNamedItem("CASESTAGESEQUENCENO").Text)
                    If Not xmlTempResponseDoc.documentElement.selectSingleNode("CASESTAGE").Attributes.getNamedItem("STAGEID") Is Nothing Then
                        Call xmlSetAttributeValue(xmlCaseTaskNode, "STAGEID", _
                        xmlTempResponseDoc.documentElement.selectSingleNode("CASESTAGE").Attributes.getNamedItem("STAGEID").Text)
                        Call xmlSetAttributeValue(xmlCaseTaskNode, "TASKID", sTaskID)
                        strResponse = objTMBO.OmTmRequest(xmlRequestNode.xml)
                        Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
                                                
                        'Check error
                        m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
                    End If

                End If
                'MAR1876 check existance of the other income tolerance task
                Set xmlTempResponseDoc = xmlLoad(strCaseTaskDetailsResponse, strFunctionName)
                If TMIncUndRefIDExistsIncomplete(xmlTempResponseDoc, sOtherTaskId) Then
                    'set it to not applicable
                    Set xmlCaseTaskNode = xmlTempResponseDoc.selectSingleNode("RESPONSE/CASESTAGE/CASETASK[@TASKID='" + sOtherTaskId + "' and @TASKSTATUS='10']")
                    If Not xmlCaseTaskNode Is Nothing Then
                        Set xmlThisCaseTaskDoc = xmlCreateDOMObject()
                        
                        Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
                        Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
                        xmlThisCaseTaskElem.setAttribute "OPERATION", "UpdateCaseTask"
                        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERID"
                        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "UNITID"
                        xmlSetAttributeValue xmlThisCaseTaskElem, "USERAUTHORITYLEVEL", "99" 'logged on user may not have authority to change it
                        Set xmlCaseTaskNode = xmlCaseTaskNode.cloneNode(True) ' PSC 05/05/2006 MAR1593
                        xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "30" 'Not Applicable
                        xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
                                               
                        ' run Tm UpdateCaseTask request
                        Set xmlOut = xmlLoad(objMsgTmBO.TmRequest(xmlThisCaseTaskDoc.xml), strFunctionName)
                        Debug.Print "UpdateCaseTask response:"
                        Debug.Print xmlOut.xml
                        ' parse response
                        errCheckXMLResponseNode xmlOut.documentElement, , True
                    End If
                End If
            Else
                'Tolerance was not exceeded so check if we need to set any task Not Applicable
                Dim NumOfTasks As Integer
                NumOfTasks = 0
                sTaskID = objParameterDO.GetCurrentParameterByType("TMIncomeDeclineID", "STRING")
                Do While NumOfTasks < 2
                    If TMIncUndRefIDExistsIncomplete(xmlTempResponseDoc, sTaskID) Then
                        'Set this task not applicable
                        Set xmlCaseTaskNode = xmlTempResponseDoc.selectSingleNode("RESPONSE/CASESTAGE/CASETASK[@TASKID='" + sTaskID + "' and @TASKSTATUS='10']")
                        If Not xmlCaseTaskNode Is Nothing Then
                            Set xmlThisCaseTaskDoc = xmlCreateDOMObject()
                            
                            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
                            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
                            xmlThisCaseTaskElem.setAttribute "OPERATION", "UpdateCaseTask"
                            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERID"
                            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "UNITID"
                            xmlSetAttributeValue xmlThisCaseTaskElem, "USERAUTHORITYLEVEL", "99" 'logged on user may not have authority to change it
                            Set xmlCaseTaskNode = xmlCaseTaskNode.cloneNode(True) ' PSC 05/05/2006 MAR1593
                            xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "30" 'Not Applicable
                            xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
                                                        
                            ' run Tm UpdateCaseTask request
                            Set xmlOut = xmlLoad(objMsgTmBO.TmRequest(xmlThisCaseTaskDoc.xml), strFunctionName)
                            Debug.Print "UpdateCaseTask response:"
                            Debug.Print xmlOut.xml
                            ' parse response
                            errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True
                        End If
                    End If
                    NumOfTasks = NumOfTasks + 1
                    sTaskID = strTMIncUndRefID  'MAR1891 GHun
                Loop
                
            End If
        End If
    End If

    Set IIncomeCalcsTxBO_RunToleranceCheck = xmlResponseElem
    m_objContext.SetComplete

IIncomeCalcsTxBO_RunToleranceCheckExit:
    '~ Clean up
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlOut = Nothing
    Set objParameterDO = Nothing
    Set objTMBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set objMsgTmBO = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlTempRequest2 = Nothing

    Exit Function

IIncomeCalcsTxBO_RunToleranceCheckVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsTxBO_RunToleranceCheck = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume IIncomeCalcsTxBO_RunToleranceCheckExit

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~ Created      22/09/2005
'~ Prog         MFoster
'~ AQR          MAR30
'~ Description  Reads the CaseTask nodes and determines if the "exceeded tolerance task"
'~              exists and has not been completed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function TMIncUndRefIDExistsIncomplete(ByRef rxmlxmlDoc As FreeThreadedDOMDocument40, _
                                     ByVal v_sTaskID As String) As Boolean

    On Error GoTo TMIncUndRefIDExistsErr
    
    Dim xmlCaseTask As IXMLDOMElement
    Dim sStatus As String
    Dim bRet As Boolean
    
    bRet = False
    
    'MAR1891 GHun reduce number of nodes to check by using XPATH
    'MAR1891 GHun Add quotes around the value of v_sTaskID
    For Each xmlCaseTask In rxmlxmlDoc.documentElement.selectNodes("CASESTAGE/CASETASK[@TASKID='" & v_sTaskID & "']")
        sStatus = xmlCaseTask.getAttribute("TASKSTATUS")
        bRet = IsValidationTypeInValidationList("TaskStatus", "I", Val(sStatus))
        If bRet Then
            Exit For
        End If
    Next
    
TMIncUndRefIDExistsExit:

    TMIncUndRefIDExistsIncomplete = bRet
    Exit Function

TMIncUndRefIDExistsErr:
    '~ Indicate false so another one will be created
    bRet = False
    Resume TMIncUndRefIDExistsExit
    
End Function

Private Sub ObjectControl_Activate()

    Set m_objXMLAssist = New XMLAssist
    Set m_objErrAssist = New ErrAssist
    Set m_objContext = GetObjectContext()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = True

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objXMLAssist = Nothing
    Set m_objErrAssist = Nothing
    Set m_objContext = Nothing
    
End Sub
