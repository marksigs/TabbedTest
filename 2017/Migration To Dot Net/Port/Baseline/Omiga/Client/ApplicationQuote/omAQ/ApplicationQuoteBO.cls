VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApplicationQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationQuoteBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Component to link cost modelling to AiP
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' MCS   08/11/99 Created
' SR    10-12-99 Added new method CheckApplicantPPEligibility, CalculaterecommendedPPCover,
'                CalculateAndSavePPCosts, CalculateAndSaveLifeCoverCosts
' DJP   10-01-00 Added AcceptQuotation
' DJP   17-02-00 Added GetLoanCompositionDetails
' DJP   22/02/00 Added GetDefaultsForNewLoanComponent
' DJP   28/02/00 Added CreateFirstMortgageLifeSubQuotes
' DJP   28/02/00 Added FindMortgageProducts
' DJP   28/02/00 Updated FindMortgageProducts to call CostModelAssist common method
' DJP   29/02/00 Added SetComplete to FindMortgageProducts
' DJP   29/02/00 Updated GetDefaultsForNewLoanComponent to call CostModelAssistBO
' DJP   16/03/00 Added BuildLifeCoverCalcInputInterface, CalculateAndSaveLifeCoverCosts,
'                CalculateMortgageCosts, CalcCostModelLTV
' JLD   11/04/00 Added ValidateMortgageSubQuote
' APS   12/04/00 Amended GetValidateQuotationDetails as we shouldn't be calling
'                IInterface methods on BO objects
' SR    25/04/00 Copied the method 'GetHighestEarners' from 'ApplicationBO'
' SR    28/04/00 New Method CalculateAffordability
' APS   31/05/00 SYS0796 Handle record not found and assume a single relationship
' APS   01/06/00 SYS0799 Changed tag name from ALLOWABLEANNUALINCOME TO TOTALGROSSEARNEDINCOME
' MS    08/07/00 SYS0824 UK residency NOT a criterion for PP Eligibility
' MS    15/06/00 SYS0914 Accepted Quote NOT Recommended
' MS    20/06/00 Trap errors raised
' DM    27-06-00 Added return response checking to this method
' PSC   30/06/00 SYS1012 - Get LoansNotRedeemed from CustomerFinancial
' MS    07/07/00 Dont append data again
' DM    13/07/00 SYS0993 Wrong XML tags being used. Changed MORTGAGECOVERAMOUNT to MORTGAGEAMOUNT
'                and COVERTYPEID to COVERTYPE and also PRODUCTCODE to PPPRODUCTCODE.
'MS     14/07/00    SYS0968 IsSelfEmployed  changed to IsMainSelfEmployed
' DM    02/08/00 SYS0993 Wrong XML tag being used for SPLITAMOUNT, change from SPLITCOVER.
'DJP    27/09/00 SYS1505 - Comment out change for sys0977.
'PSC    19/03/01 SYS2053 Add RemodelQuotation
'GD     29/06/01 SYS2391 Changes made for xml parser v3
'DRC    10/08/01 SYS0977 Completed & renabled so that LTV will be recalculated and the
'                        mortgagesubquote record updated.
'MC     20/08/01 SYS2250 Add Rate Change batch functionality
'DRC    10/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'DRC    09/11/01 changes from 10/08/01 for  SYS0977 Commented out
'GD     16/11/01 SYS2970 getelementsbytagname() doesn't like a pattern as a parameter
'                       - replaced with selectNodes
'JR     03/12/01 SYS2459 - re-applied Omiplus24 telephone changes to IApplicationQuoteBO_GetApplicationSummary
'SR     14/12/01 SYS3322 - modified private method CreateApplicationFeeType (now it is a function instead of Sub)
'                        - modified IApplicationQuoteBO_AcceptQuotation (modified call to CreateApplicationFeeType)
'JLD    18/12/01 SYS3392 - Add GetAcceptedQuoteData
'DRC    29/01/02 SYS3518 - Bug fix in CreateApplicationFeeTypes
'DRC    29/04/02 SYS4505 - Warning message in GetHighestEarners if no Earners found
'CL     07/04/02 SYS4510 - Modifications to class initialise
'CL     10/05/02 SYS4510 - Remove class initialize & class terminate
'MEVA   13/05/02 SYS0649 - Fix for Joint Applications
'SG     28/05/02 SYS4647 - MSMS to Core integration
'LH     31/01/07 EP2_1130 - Changed CalculateApproximateMonthlyCost() logic according to spec change.
'LH     01/02/07 EP2_1079 - If no quotation exists then exit cleanly with approximate cost = zero
'LH     01/02/07 EP2_1130 - Fixed CalculateApproximateMonthlyCost()
'LH     05/02/07 EP2_1232 - Fix to CalculateApproximateMonthlyCost() to make REPAYMENTVEHICLEMONTHLYCOST optional not mandatory
'LH     06/02/07 EP2_1079 - Fix to CalculateApproximateMonthlyCost() to exclude rounding to 2 DP's in Interest Only calculation
'LH     10/02/07 EP2_1079 - If no quotation exists then exit cleanly with approximate cost = zero
'-----------------------------------------------------------------------------------------------------------------
Option Explicit

Implements IApplicationQuoteBO
Implements ObjectControl

'Private Const cstrTABLE_NAME = "APPLICATIONQUOTE"
'Private Const cstrCUSTOMER_COMP = "omCust"
Private Const cstrPPSUBQUOTE_COMP = "omCM"
'Private Const cstrLIFESUBQUOTE_COMP = "omCM"
'Private Const cstrAPPLICATION_COMP = "omAPP"

'Private Const cstrQuotationCostModel = "Cost Modelling"
'Private Const cstrQuotationQQ = "Quick Quote"

'Private Enum QuotationType
'    enumQuickQuote = 1
'    enumCostModel = 2
'End Enum

Private Const cstrHusband As String = "H"
Private Const cstrWife As String = "W"
Private Const cstrUnmarried As String = "U"

' Private member variables
Private m_objErrAssist As ErrAssist
Private m_objXMLAssist As XMLAssist
Private m_objContext As ObjectContext
'SG 28/05/02 SYS4767 Function added

'------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date     Description
' GD    08/07/02 BMIDS00165 - Added new method GetAcceptedOrActiveQuoteData
' MV    12/08/02 BMIDS00311 - Core Ref : SYS1705 Removed SafeRef ;
'                Modified IApplicationQuoteBO_CalculateMortgageCosts
' GD    30/08/02 BMIDS00256 - Recalculation of LTV needs to take place if Purchase Price has changed - Changed GetValidatedQuotationDetails
' MV    03/09/02 BMIDS00353 - Modified IApplicationQuoteBO_GetLoanCompositionDetails()
' GD    06/09/02 BMIDS00256 - Further changes to GetValidatedQuotationDetails
' SA    12/09/02 BMIDS00365 - Need to return RecommendedQuoteNumber and AcceptedQuoteNumber from GetValidatedQuotationDetails
' MV    11/10/02 BMIDS00590 - Modified IApplicationQuoteBO_GetLoanCompositionDetails
'                             Amended IApplicationQuoteBO_ProcessAddedOneOffCosts
'AW     21/10/02 BMIDS00653 - BM029, Amended Added CalculateOutGoings for Net Allowable Income Calcs
'DPF    29/10/02 BMIDS00757 - Append Valuation Type & Location to response
'MDC    06/11/02 BMIDS00653 - Get EmployersRef for Allowable Income calculation.
'MDC    07/11/02 BMIDS00865 - Include Retired Income in Allowable Income calculation
'MV     06/12/02 BMIDs0145 -  Amended IApplicationQuoteBO_GetAcceptedOrActiveQuoteData()
'MDC    19/12/02 BM0206 Add option to throw error in method GetAcceptedOrActiveQuoteData
'SR     02/01/03 BM0209 - Modified methods GetLoanCompositionDetails, AcceptQuotation
'GD     28/01/03 BM0287 - LTV not being calc'd if a B&C quotes exists.
'MV     03/02/03 BM0221  - Added new function ValidateUserMandateLevel()and IApplicationQuoteBO_ValidateUserMandateLevel
'GHun   10/09/03 BMIDS622 - Restructure RateChange
'DRC    20/05/04 BMIDS767 - Added new method UpdateQuotationForAddedCosts & Changed CreateApplicationFeeTypes
'MC     17/06/04 BMIDS763   -   APPLICATIONDATE retrieved from DB and appended to the Request.
'MC     07/07/04 BMIDS767   - Overflow error fixed in CreateApplicationFeeTypes()
'SR     13/07/04 BMIDS767   - Modified method CreateApplicationFeeTypes
'SR     14/07/04 BMIDS767   - Modified method CreateApplicationFeeTypes
'SR     14/07/04 BMIDS767   - Modified method CreateApplicationFeeTypes
'SR     15/07/04 BMIDS795   - Modified method CreateApplicationFeeTypes
'JD     20/07/04 BMIDS763   - validateManualAdjustmentPercent needs to pass application info to getFullProductDetails
'GHun   20/08/04 BMIDS856   - Minor cleanup and amended RecostQuotationForRateChange
'SR     26/08/04 BMIDS815   New methods 'RefreshMortgageAccountData',
'                           'GetAndSavePortedStepAndPeriodFromMortgageAccount'
'KRW    17/11/04 BMIDS946   More descriptive error message displayed when trying to approve an application with out an accepted Quote
'HMA    21/02/05 BMIDS959   - Correct ValidateManualAdjustmentPercent
' JD    04/04/2005 BMIDS977 Added CreateAdHocFeeAndRemodel
' HMA   12/04/2005 BMIDS977 Remove Interactive warning message from CreateApplicationFeeTypes
'                           Removed method UpdateQuotationForAddedCosts
'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'GHun   15/07/2005  MAR7 Integrate local printing. Changed GetAcceptedOrActiveQuoteData
'HMA    09/09/2005  MAR28 Update CreateApplicationFeeTypes
'DRC    03/11/2005  MAR377 Update CreateApplicationFeeTypes to enable deletion of refund Val Fees
'JD     24/01/2006  MAR1040 pass both declared and confirmed info in BuildRequestForCalculateAllowableIncome
'JD     09/03/2006  MAR1061 GetLoanCompositionDetails - don't pass purchaseprice as it is picked up from MSQ now.
'PSC    04/05/2006  MAR1497 CreateApplicationFeeTypes - Correct error handling an processing of RFV fee payments
'------------------------------------------------------------------------------------------
'EPSOM History:
'
'Prog   Date        Description
'IK     24/04/2006  EP397 - fix GetApplicationSummary
'IK     05/05/2006  EP502 - fix GetApplicationSummary
'PE     14/07/2006  EP974/MAR688 - Mars Merge
'PE     26/07/2006  EP1038 Mars merge/MAR1893 Commented out redundant LifeSubQuote code
'PE     10/07/2006  EP1038 Mars merge/MAR1891 Minor tidy up
'HMA    27/07/2006  EP1042 Correction to GetApplicationSummary.
'LH     31/07/2006  EP1042 Further correction to GetApplicationSummary.
'PSC    09/11/2006  EP2_41 Amend CalculateMortgageCosts to return costs
'MCh    05/12/2006  EP2_128 Add CalculateApproximateMonthlyCost
'PSC    05/12/2006  EP2_249 Add GetIntroducerProcFeeDetails
'IK     23/01/2007  EP2_929 fix CalculateApproximateMonthlyCost
'PSC    19/02/2007  EP2_1488 Amend ValidateManualAdjustmentPercent to use ORIGINALLTV
'------------------------------------------------------------------------------------------

'BMIDS856 GHun ObjectControl_Activate should be used instead
'Private Sub Class_Initialize()
'    If m_objErrAssist Is Nothing Then
'        Set m_objErrAssist = New ErrAssist
'    End If
'
'    If m_objXmlAssist Is Nothing Then
'        Set m_objXmlAssist = New XMLAssist
'    End If
'
'    If m_objConvertAssist Is Nothing Then
'        Set m_objConvertAssist = New ConvertAssist
'    End If
'
'End Sub
'BMIDS856 End

'BMIDS856 GHun ObjectControl_Deactivate should be used instead
'Private Sub Class_Terminate()
'    Set m_objErrAssist = Nothing
'    Set m_objXmlAssist = Nothing
'    Set m_objContext = Nothing
'End Sub
'BMIDS856 End

Private Function IApplicationQuoteBO_CalcAllowableInc(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalcAllowableIncVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_CalcAllowableInc"

    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement, xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode, xmlNodeList As IXMLDOMNodeList, xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement, xmlElem2 As IXMLDOMElement, xmlLendersElem As IXMLDOMElement
    
    Dim objAllowableIncomeBO As AllowableIncomeBO
    Dim objCustomerBO As CustomerBO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objMortgageLenderBO As MortgageLenderBO
    Dim strOrganisationId As String
    
    Dim strResponse As String
    Dim strDefaultLenderCode As String
    Dim strOrganisations As String
    Dim strLenderCode As String
    
    Dim xmlNode5 As IXMLDOMNode

    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objAllowableIncomeBO = New AllowableIncomeBO
        Set objCustomerBO = New CustomerBO
        Set objGlobalParameterDO = New GlobalParameterDO
        Set objMortgageLenderBO = New MortgageLenderBO
    Else
        Set objAllowableIncomeBO = m_objContext.CreateInstance(App.Title & ".AllowableIncomeBO")
        Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objMortgageLenderBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".MortgageLenderBO")
    End If
        
    ' Organisation
    strDefaultLenderCode = objGlobalParameterDO.GetCurrentParameterByType("DefaultLenderCode", "STRING")

    ' Peter Edney - EP974/MAR688 - Mars Merge
    Dim strApplicationNumber As String
    strApplicationNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text

    Dim objOrganisationDoc As FreeThreadedDOMDocument40

    'Get all the lender names and income multiples for each lender
    strOrganisations = objMortgageLenderBO.FindLenderNames("<MORTGAGELENDER></MORTGAGELENDER>")
    Set objOrganisationDoc = m_objXMLAssist.load(strOrganisations, TypeName(Me), strFunctionName)
      
    m_objErrAssist.CheckResponse objOrganisationDoc.documentElement, xmlResponseElem, True
    
    '
    ' Loop through for our Lender Code
    '
    Dim xmlOrganisationNodeList As IXMLDOMNodeList

    Set xmlOrganisationNodeList = objOrganisationDoc.selectNodes("/RESPONSE/MORTGAGELENDERLIST/MORTGAGELENDER")

    For Each xmlNode5 In xmlOrganisationNodeList

        strLenderCode = m_objXMLAssist.GetTagValue(xmlNode5, "LENDERCODE")

        If (strLenderCode = strDefaultLenderCode) Then
            strOrganisationId = m_objXMLAssist.GetTagValue(xmlNode5, "ORGANISATIONID")
        End If

    Next
    
    Set xmlLendersElem = xmlOut.createElement("LENDERLIST")
    xmlResponseElem.appendChild xmlLendersElem
    
    Set xmlElement = xmlOut.createElement("LENDERDETAILS")
    Set xmlElem2 = xmlOut.createElement("ORGANISATIONID")
    xmlElem2.Text = strOrganisationId
    xmlElement.appendChild xmlElem2
        
    'Dim nloop As Integer
    'Dim dblIncome As Double, dblNetIncome As Double
    
    'nloop = 0
    'Dim xmlTaxDetails As IXMLDOMElement

    ' Peter Edney - EP974/MAR688 - Mars Merge
    Dim xmlOutDoc As FreeThreadedDOMDocument40
'    Debug.Print vxmlRequest.xml
'    Set xmlTempResponseNode = BuildRequestForCalculateAllowableIncome(vxmlRequest)
    'JD MAR688 call stored procedure to get allowable income inputs
    Dim strCalcAllowableIncInputs As String
    Dim strCalcOutgoingsInputs As String
    strCalcAllowableIncInputs = GetCalcAllowableIncomeInput(strDefaultLenderCode, strApplicationNumber)
    strCalcOutgoingsInputs = GetCalcOutgoingsInput(strApplicationNumber)
    
    ' Peter Edney - EP974/MAR688 - Mars Merge
    'BMIDS00653 MDC 28/10/2002
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'    xmlNewRequest.appendChild xmlElement
'    Set xmlNode = m_objXMLAssist.GetMandatoryNode(xmlTempResponseNode, "CUSTOMERLIST")
'    xmlNewRequest.appendChild xmlNode.cloneNode(True)
    strResponse = objAllowableIncomeBO.CalculateAllowableIncome(strCalcAllowableIncInputs, strCalcOutgoingsInputs)
            
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    'dblIncome = Val(xmlDoc.selectSingleNode(".//ANNUALALLOWABLEINCOME").Text)
    '   AW  21/10/2002  BMIDS00653
    'dblNetIncome = Val(xmlDoc.selectSingleNode(".//NETANNUALALLOWABLEINCOME").Text)
    xmlLendersElem.appendChild xmlDoc.selectSingleNode("//LENDERDETAILS")
   
'    Set xmlNodeList = xmlTempResponseNode.selectNodes(".//CUSTOMER")
'    Debug.Print xmlTempResponseNode.xml
'    If xmlNodeList.length > 0 Then
'        For Each xmlNode In xmlNodeList
'            Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'
'            xmlNewRequest.appendChild xmlElement
'            xmlNewRequest.appendChild xmlNode.cloneNode(True)
'            Debug.Print xmlNewRequest.xml
'            Debug.Print xmlTempResponseNode.xml
'            strResponse = objAllowableIncomeBO.CalcAllowableInc(xmlNewRequest.xml)
'            Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'
'            dblIncome = Val(xmlDoc.selectSingleNode(".//ANNUALALLOWABLEINCOME").Text)
'            '   AW  21/10/2002  BMIDS00653
'            dblNetIncome = Val(xmlDoc.selectSingleNode(".//NETANNUALALLOWABLEINCOME").Text)
'
'            xmlLendersElem.appendChild xmlDoc.selectSingleNode("//LENDERDETAILS")
'
'            '
'            ' INSERT or UPDATE for the INCOMESUMMARY table
'            '
'
'            Dim rst As ADODB.Recordset, _
'                cmd As New ADODB.Command
'
'            Dim xmlRootNode As IXMLDOMNode
'
'            Dim objIADOAssist As IADOAssist, _
'                objSQLAssist  As New SQLAssist
'
'            Dim strCustomerNumber           As String, _
'                strCustomerVersionNumber    As String, _
'                strEmploymentSequenceNumber As String
'            Dim strStartDate      As String, _
'                strAccountantGUID As String
'
'            Dim strSQL As String
'
'            '
'            ' Initialise
'            '
'            strCustomerNumber = ""
'            strCustomerVersionNumber = ""
'
'            Set objIADOAssist = New ADOAssist
'
'            '
'            ' Main
'            '
'            strCustomerNumber = xmlNodeList.Item(nloop).selectSingleNode("CUSTOMERNUMBER").Text
'            strCustomerVersionNumber = xmlNodeList.Item(nloop).selectSingleNode("CUSTOMERVERSIONNUMBER").Text
'
'            '
'            ' Build and execute the SQL query to get the start date
'            '
'            strSQL = "SELECT CUSTOMERNUMBER FROM INCOMESUMMARY WHERE "
'            strSQL = strSQL + "CUSTOMERNUMBER='" & strCustomerNumber & "' AND "
'            strSQL = strSQL + "CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "'"
'
'            cmd.ActiveConnection = objIADOAssist.GetConnStr
'            cmd.CommandText = strSQL
'
'            Set rst = cmd.Execute(1, , adCmdText)
'
'            '
'            ' Add the value in the recordset to the input XML document
'            ' Update the ALLOWABLEINCOME table
'            '
'            Set xmlTaxDetails = xmlOutDoc.createElement("REQUEST")
'
'            'Set xmlTaxDetails = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
'            xmlTaxDetails.setAttribute "OPERATION", "UPDATE"
'            Debug.Print xmlTaxDetails.xml
'            Set xmlNode1 = xmlOutDoc.createElement("INCOMESUMMARY")
'
'            Set xmlNode2 = xmlOutDoc.createElement("CUSTOMERNUMBER")
'            xmlNode2.Text = strCustomerNumber
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("CUSTOMERVERSIONNUMBER")
'            xmlNode2.Text = strCustomerVersionNumber
'            xmlNode1.appendChild xmlNode2
'            Debug.Print xmlTaxDetails.xml
'            Set xmlNode2 = xmlOutDoc.createElement("ALLOWABLEANNUALINCOME")
'            xmlNode2.Text = dblIncome
'            xmlNode1.appendChild xmlNode2
'            Debug.Print xmlNode1.xml
'            '   AW  21/10/2002  BMIDS00653
'            Set xmlNode2 = xmlOutDoc.createElement("NETALLOWABLEANNUALINCOME")
'            xmlNode2.Text = dblNetIncome
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("ALLOWABLEMONTHLYINCOME")
'            xmlNode2.Text = dblIncome / 12
'            xmlNode1.appendChild xmlNode2
'            Debug.Print xmlTaxDetails.xml
'            Set xmlNode2 = xmlOutDoc.createElement("NONUKTAXDETAILS")
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("TAXOFFICE")
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("TAXREFERENCENUMBER")
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("TAXEDOUTSIDEUKINDICATOR")
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("UKTAXPAYERINDICATOR")
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("HIGHESTTAXRATE")
'            xmlNode1.appendChild xmlNode2
'            Debug.Print xmlNode1.xml
'            xmlTaxDetails.appendChild xmlNode1
'
'            Set xmlNode1 = xmlOutDoc.createElement("CUSTOMERVERSION")
'
'            Set xmlNode2 = xmlOutDoc.createElement("CUSTOMERNUMBER")
'            xmlNode2.Text = strCustomerNumber
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("CUSTOMERVERSIONNUMBER")
'            xmlNode2.Text = strCustomerVersionNumber
'            xmlNode1.appendChild xmlNode2
'
'            Set xmlNode2 = xmlOutDoc.createElement("NATIONALINSURANCENUMBER")
'            xmlNode1.appendChild xmlNode2
'
'            xmlTaxDetails.appendChild xmlNode1
'            Debug.Print xmlTaxDetails.xml
'            '
'            ' Select the Action
'            '
'
'            If rst.EOF Then
'                '
'                ' Change action from UPDATE to INSERT
'                '
'                xmlTaxDetails.Attributes(0).Text = "CREATE"
'                'xmlTaxDetails.Attributes(6).Text = "CREATE"
'            End If
'
'            '
'            ' Perform the Update
'            '
'            strOrganisations = objCustomerBO.SaveTaxDetails(xmlTaxDetails.xml)
'            Debug.Print strOrganisations
'
'    ' ==============================================================================
'            nloop = nloop + 1
'
'        Next xmlNode
'    End If
    'BMIDS00653 MDC 28/10/2002 - End
        
    Set IApplicationQuoteBO_CalcAllowableInc = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalcAllowableIncExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    Set xmlElem2 = Nothing
    Set xmlLendersElem = Nothing
    Set objAllowableIncomeBO = Nothing
    Set objCustomerBO = Nothing
    Set objGlobalParameterDO = Nothing
    Set objMortgageLenderBO = Nothing
    
    Exit Function

CalcAllowableIncVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CalcAllowableInc = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAllowableIncExit

End Function

' Peter Edney - EP974/MAR688 - Mars Merge
Private Function GetCalcAllowableIncomeInput(ByVal vstrLenderCode As String, ByVal vstrApplicationNumber As String) As String
    Const cstrFunctionName As String = "GetCalcAllowableIncomeInput"
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    
On Error GoTo GetCalcAllowableIncomeInputExit
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set adoStream = New ADODB.Stream
    adoStream.Open
       
    Set adoCommand = New ADODB.Command
    With adoCommand
        .CommandText = "USP_GETALLOWABLEINCINPUT"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@p_LenderCode", adBSTR, adParamInput, Len(vstrLenderCode), vstrLenderCode)
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, Len(vstrApplicationNumber), vstrApplicationNumber)
        Set .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
       
    adoConnection.Close
    
    GetCalcAllowableIncomeInput = adoStream.ReadText()
    
    adoStream.Close
    
GetCalcAllowableIncomeInputExit:
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

' Peter Edney - EP974/MAR688 - Mars Merge
Private Function GetCalcOutgoingsInput(ByVal vstrApplicationNumber As String) As String
    Const cstrFunctionName As String = "GetCalcOutgoingsInput"
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    
On Error GoTo GetCalcOutgoingsInputExit
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set adoStream = New ADODB.Stream
    adoStream.Open
       
    Set adoCommand = New ADODB.Command
    With adoCommand
        .CommandText = "usp_GetCalcOutgoingsInput"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, Len(vstrApplicationNumber), vstrApplicationNumber)
        Set .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
       
    adoConnection.Close
    
    GetCalcOutgoingsInput = adoStream.ReadText()
    
    adoStream.Close
    
GetCalcOutgoingsInputExit:
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

'EP1038/MAR1891 No longer used
'Private Function BuildRequestForCalculateAllowableIncome(ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMNode
'
'On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'
'    Const strFunctionName As String = "BuildRequestForCalculateAllowableIncome"
'
'    Dim xmlDoc As FreeThreadedDOMDocument40, xmlOut As FreeThreadedDOMDocument40, xmlResponseElem As IXMLDOMElement
'    Dim xmlNewRequest As IXMLDOMElement, xmlTableElem As IXMLDOMElement
'    Dim xmlElem As IXMLDOMElement, xmlElem2 As IXMLDOMNode, xmlListElem As IXMLDOMElement, xmlElemList As IXMLDOMNodeList
'    Dim xmlCustomer As IXMLDOMNode, xmlCustomerList As IXMLDOMNodeList
'    Dim xmlEmployment As IXMLDOMNode, xmlEmploymentList As IXMLDOMNodeList
'    Dim xmlCustomerRoles As IXMLDOMNode
'    Dim xmlOutCustomerElem As IXMLDOMElement, xmlOutCustomerListElem As IXMLDOMElement
'    Dim xmlOutElem As IXMLDOMElement, xmlOutListElem As IXMLDOMElement
'    Dim xmlOutEmpElem As IXMLDOMElement, xmlOutEmpListElem As IXMLDOMElement
'
'    Dim strResponse As String, strCondition As String, blnRecordNotFound As Boolean
'    'Dim strEmployedID As String
'    Dim strSelfEmployedID As String, strConractId As String
'    Dim strCustomerNumber As String, strCustVersionNumber As String, strEmploymentStatus As String
'    Dim strEmpSeqNo As String   ', strTemp As String
'    Dim lngErrNo As Long
'
'    Dim strEmployed33PctShares As String    'BMIDS00653 MDC 06/11/2002
'
'    Dim objCustomerRoleBO As CustomerRoleBO, objCEBO As CustomerEmploymentBO
'    Dim objComboDO As ComboDO, objAppProcBO As Object
'
'    Set xmlDoc = New FreeThreadedDOMDocument40
'    xmlDoc.validateOnParse = False
'    xmlDoc.setProperty "NewParser", True
'    Set xmlOut = New FreeThreadedDOMDocument40
'    xmlOut.validateOnParse = False
'    xmlOut.setProperty "NewParser", True
'
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    If Not m_objContext Is Nothing Then
'        Set objCustomerRoleBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".CustomerRoleBO")
'        Set objCEBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
'        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
'        Set objAppProcBO = m_objContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
'    Else
'        Set objCustomerRoleBO = New CustomerRoleBO
'        Set objCEBO = New CustomerEmploymentBO
'        Set objComboDO = New ComboDO
'        Set objAppProcBO = CreateObject(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
'    End If
'
'    'strEmployedID = objComboDO.GetFirstComboValueId("EmploymentStatus", "E")
'    strSelfEmployedID = objComboDO.GetFirstComboValueId("EmploymentStatus", "S")
'    strConractId = objComboDO.GetFirstComboValueId("EmploymentStatus", "C")
'    'BMIDS00653 MDC 06/11/2002
'    strEmployed33PctShares = objComboDO.GetFirstComboValueId("EmploymentStatus", "Z33")
'    'BMIDS00653 MDC 06/11/2002 - End
'
'    '----------------------------------------------------------------------
'    'Get RoleType for all the customers attached to this application
'    '----------------------------------------------------------------------
'    ' Build the request
'    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'    Set xmlTableElem = xmlDoc.createElement("CUSTOMERROLE")
'    xmlNewRequest.appendChild xmlTableElem
'    Debug.Print xmlNewRequest.xml
'    xmlTableElem.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
'    xmlTableElem.appendChild vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
'
'    strResponse = objCustomerRoleBO.FindList(xmlNewRequest.xml)
'    Debug.Print strResponse
'    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'    Set xmlCustomerRoles = xmlDoc.documentElement
'
'    '----------------------------------------------------------
'    ' For each customer build the request for CalcAllowableInc
'    '----------------------------------------------------------
'    Set xmlCustomerList = vxmlRequest.selectNodes(".//CUSTOMER")
'
'    Set xmlOutCustomerListElem = xmlDoc.createElement("CUSTOMERLIST")
'    xmlResponseElem.appendChild xmlOutCustomerListElem
'
'    '
'    '
'    '
'    Dim intLoop As Integer
'
'    intLoop = 0
'
'    For Each xmlCustomer In xmlCustomerList
'        ' append CustomerRole data to Response
'        strCustomerNumber = xmlCustomerList.Item(intLoop).selectSingleNode(".//CUSTOMERNUMBER").Text
'        strCustVersionNumber = xmlCustomerList.Item(intLoop).selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
'
'        Set xmlOutCustomerElem = xmlDoc.createElement("CUSTOMER")
'        xmlOutCustomerListElem.appendChild xmlOutCustomerElem
'
'        xmlOutCustomerElem.appendChild xmlCustomer.selectSingleNode("CUSTOMERNUMBER").cloneNode(True)
'        xmlOutCustomerElem.appendChild xmlCustomer.selectSingleNode("CUSTOMERVERSIONNUMBER").cloneNode(True)
'
'        strCondition = ".//CUSTOMERROLE[CUSTOMERNUMBER='" & strCustomerNumber & "' and CUSTOMERVERSIONNUMBER = '" _
'                        & strCustVersionNumber & "']"
'        Set xmlElem = xmlCustomerRoles.selectSingleNode(strCondition)
'        Debug.Print xmlCustomerRoles.xml
'        xmlOutCustomerElem.appendChild xmlElem.selectSingleNode(".//CUSTOMERROLETYPE").cloneNode(True)
'
'        '-----------------------------------------------------------------
'        'Find UnEarnedIncomeList and attach to the response
'        '-----------------------------------------------------------------
'        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'        Set xmlTableElem = xmlDoc.createElement("UNEARNEDINCOME")
'        xmlNewRequest.appendChild xmlTableElem
'        xmlTableElem.appendChild xmlCustomer.selectSingleNode("CUSTOMERNUMBER").cloneNode(True)
'        xmlTableElem.appendChild xmlCustomer.selectSingleNode("CUSTOMERVERSIONNUMBER").cloneNode(True)
'
'        strResponse = objCEBO.FindOtherIncomeList(xmlNewRequest.xml)
'        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'        On Error Resume Next
'        lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
'
'        On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'        If lngErrNo <> 0 Then
'            If lngErrNo <> oeRecordNotFound Then
'                m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'            End If
'        Else
'            Set xmlElemList = xmlDoc.selectNodes(".//UNEARNEDINCOME")
'            If xmlElemList.length > 0 Then
'                Set xmlListElem = xmlDoc.createElement("UNEARNEDINCOMELIST")
'                For Each xmlElem In xmlElemList
'                    Set xmlTableElem = xmlDoc.createElement("UNEARNEDINCOME")
'                    xmlTableElem.appendChild xmlElem.selectSingleNode(".//UNEARNEDINCOMEAMOUNT").cloneNode(True)
'                    xmlTableElem.appendChild xmlElem.selectSingleNode(".//UNEARNEDINCOMETYPE").cloneNode(True)
'                    xmlTableElem.appendChild xmlElem.selectSingleNode(".//PAYMENTFREQUENCY").cloneNode(True)
'                    xmlListElem.appendChild xmlTableElem
'                Next xmlElem
'                xmlOutCustomerElem.appendChild xmlListElem
'            End If
'        End If
'
'        'Find all the employments
'        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'        Set xmlElem = xmlDoc.createElement("EMPLOYMENTLIST")
'        xmlNewRequest.appendChild xmlElem
'        Set xmlTableElem = xmlDoc.createElement("EMPLOYMENT")
'        xmlElem.appendChild xmlTableElem
'
'        xmlTableElem.appendChild xmlCustomer.selectSingleNode(".//CUSTOMERNUMBER").cloneNode(True)
'        xmlTableElem.appendChild xmlCustomer.selectSingleNode(".//CUSTOMERVERSIONNUMBER").cloneNode(True)
'
'        strResponse = objCEBO.FindEmploymentList(xmlNewRequest.xml)
'        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'        On Error Resume Next
'        lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
'
'        On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'
'        blnRecordNotFound = False
'
'        If lngErrNo <> 0 Then
'            If lngErrNo <> oeRecordNotFound Then
'                m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'            Else
'                blnRecordNotFound = True
'            End If
'        End If
'
'        If Not blnRecordNotFound Then
'            Set xmlOutEmpListElem = xmlDoc.createElement("EMPLOYMENTLIST")
'            xmlOutCustomerElem.appendChild xmlOutEmpListElem
'
'            Set xmlEmploymentList = xmlDoc.selectNodes(".//EMPLOYMENT")
'            For Each xmlEmployment In xmlEmploymentList
'                Set xmlOutEmpElem = xmlDoc.createElement("EMPLOYMENT")
'                xmlOutEmpListElem.appendChild xmlOutEmpElem
'
'                strEmploymentStatus = xmlEmployment.selectSingleNode(".//EMPLOYMENTSTATUS").Text
'                strEmpSeqNo = xmlEmployment.selectSingleNode(".//EMPLOYMENTSEQUENCENUMBER").Text
'                xmlOutEmpElem.appendChild xmlEmployment.selectSingleNode(".//EMPLOYMENTSTATUS").cloneNode(True)
'                xmlOutEmpElem.appendChild xmlEmployment.selectSingleNode(".//MAINSTATUS").cloneNode(True)
'
'                'BMIDS00865 MDC 07/11/2002 - Include Retired Income
'                'If strEmploymentStatus = strEmployedID Then
'                If objComboDO.IsItemInValidation("EmploymentStatus", strEmploymentStatus, "E") Then
'                'BMIDS00865 MDC 07/11/2002 - End
'                    '------------------------------------------------------------------------------------
'                    'Get Employer's Reference details and append to the response as ConfirmedEmpEarnings
'                    '------------------------------------------------------------------------------------
'                    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'                    xmlNewRequest.setAttribute "OPERATION", "GETCURREMPLOYERSREF"
'
'                    Set xmlTableElem = xmlDoc.createElement("CURRENTEMPLOYERSREF")
'                    xmlNewRequest.appendChild xmlTableElem
'
'                    xmlSetAttributeValue xmlTableElem, "CUSTOMERNUMBER", strCustomerNumber
'                    xmlSetAttributeValue xmlTableElem, "CUSTOMERVERSIONNUMBER", strCustVersionNumber
'                    xmlSetAttributeValue xmlTableElem, "EMPLOYMENTSEQUENCENUMBER", strEmpSeqNo
'
'                    strResponse = objAppProcBO.omAppProcRequest(xmlNewRequest.xml)
'                    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'                    On Error Resume Next
'                    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
'
'                    On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'                    blnRecordNotFound = False
'                    If lngErrNo <> 0 Then
'                        If lngErrNo <> oeRecordNotFound Then
'                            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'                        Else
'                            blnRecordNotFound = True
'                        End If
'                    End If
'                    If Not blnRecordNotFound Then
'                        'Append ConfirmedEmpEarnings to the response
''                        Set xmlTableElem = xmlDoc.createElement("CONFIRMEDEMPEARNINGS")
''                        xmlOutEmpElem.appendChild xmlTableElem
'
'                        'BMIDS00654 MDC 05/11/2002
''                        Set xmlOutListElem = xmlDoc.createElement("CONFIRMEDEMPEARNINGS")
''                        xmlOutEmpElem.appendChild xmlOutListElem
'                        'BMIDS00654 MDC 05/11/2002 - End
'
'                        'SG 07/02/02 SYS3837
'                        'Amend to confirm to new ConfirmedEarnedIncome data
'                        'OLD CODE:
'                        '
'                        'Set xmlElem2 = xmlDoc.createElement("GROSSBASICINCOME")
'                        'xmlElem2.Text = xmlGetAttributeText(xmlElem, "GROSSBASICINCOME")
'                        'xmlTableElem.appendChild xmlElem2
'
'                        '
'                        'Set xmlElem2 = xmlDoc.createElement("GUARANTEEDOVERTIMEBONUS")
'                        'xmlElem2.Text = xmlGetAttributeText(xmlElem, "GUARANTEEDOVERTIMEBONUS")
'                        'xmlTableElem.appendChild xmlElem2
'                        '
'                        'Set xmlElem2 = xmlDoc.createElement("REGOVERTIMEBONUS")
'                        'xmlElem2.Text = xmlGetAttributeText(xmlElem, "REGOVERTIMEBONUS")
'                        'xmlTableElem.appendChild xmlElem2
'                        '
'                        'Set xmlElem2 = xmlDoc.createElement("PACKAGEDINCOME")
'                        'xmlElem2.Text = xmlGetAttributeText(xmlElem, "PACKAGEDINCOME")
'                        'xmlTableElem.appendChild xmlElem2
'
'                        'BMIDS00654 MDC 05/11/2002
''                        ' Append to the response
''                        Set xmlElemList = xmlDoc.selectNodes("//CONFIRMEDEARNEDINCOME")
''
''                        For Each xmlElem In xmlElemList
''                            Set xmlTableElem = xmlDoc.createElement("CONFIRMEDEARNEDINCOME")
''                            Set xmlElem2 = xmlDoc.createElement("EARNEDINCOMEAMOUNT")
''                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "EARNEDINCOMEAMOUNT")
''                            xmlTableElem.appendChild xmlElem2
''
''                            Set xmlElem2 = xmlDoc.createElement("EARNEDINCOMETYPE")
''                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "EARNEDINCOMETYPE")
''                            xmlTableElem.appendChild xmlElem2
''
''                            Set xmlElem2 = xmlDoc.createElement("PAYMENTFREQUENCYTYPE")
''                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "PAYMENTFREQUENCYTYPE")
''                            xmlTableElem.appendChild xmlElem2
''
''                            xmlOutListElem.appendChild xmlTableElem
''
''                        Next xmlElem
'                        Set xmlElem = xmlDoc.selectSingleNode("//CURRENTEMPLOYERSREF")
'                        xmlOutEmpElem.appendChild xmlElem
'                        'BMIDS00654 MDC 05/11/2002 - End
'                    'Else
'                    End If
'                    'JD MAR1040 send declared and confirmed to calc both
'                        '-----------------------------------------------------------------
'                        ' Append DeclaredEmpEarnings to the response
'                        '-----------------------------------------------------------------
'                        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'                        Set xmlTableElem = xmlDoc.createElement("EARNEDINCOME")
'                        xmlNewRequest.appendChild xmlTableElem
'
'                        Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
'                        xmlElem.Text = strCustomerNumber
'                        xmlTableElem.appendChild xmlElem
'
'                        Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
'                        xmlElem.Text = strCustVersionNumber
'                        xmlTableElem.appendChild xmlElem
'
'                        Set xmlElem = xmlDoc.createElement("EMPLOYMENTSEQUENCENUMBER")
'                        xmlElem.Text = strEmpSeqNo
'                        xmlTableElem.appendChild xmlElem
'
'                        strResponse = objCEBO.FindEarnedIncomeList(xmlNewRequest.xml)
'                        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'                        On Error Resume Next
'                        lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
'
'                        On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'                        blnRecordNotFound = False
'                        If lngErrNo <> 0 Then
'                            If lngErrNo <> oeRecordNotFound Then
'                                m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'                            Else
'                                blnRecordNotFound = True
'                            End If
'                        End If
'
'                        If Not blnRecordNotFound Then
'                            ' Append to the response
'                            Set xmlOutListElem = xmlDoc.createElement("DECLAREDEMPEARNINGSLIST")
'                            xmlOutEmpElem.appendChild xmlOutListElem
'
'                            Set xmlElemList = xmlDoc.selectNodes("//EARNEDINCOME")
'                            For Each xmlElem In xmlElemList
'                                Set xmlTableElem = xmlDoc.createElement("DECLAREDEMPEARNINGS")
'                                Set xmlElem2 = xmlDoc.createElement("EARNEDINCOMEAMOUNT")
'                                xmlElem2.Text = xmlGetNodeText(xmlElem, ".//EARNEDINCOMEAMOUNT")
'                                xmlTableElem.appendChild xmlElem2
'
'                                Set xmlElem2 = xmlDoc.createElement("EARNEDINCOMETYPE")
'                                xmlElem2.Text = xmlGetNodeText(xmlElem, ".//EARNEDINCOMETYPE")
'                                xmlTableElem.appendChild xmlElem2
'
'                                Set xmlElem2 = xmlDoc.createElement("PAYMENTFREQUENCYTYPE")
'                                xmlElem2.Text = xmlGetNodeText(xmlElem, ".//PAYMENTFREQUENCYTYPE")
'                                xmlTableElem.appendChild xmlElem2
'
'                                xmlOutListElem.appendChild xmlTableElem
'                            Next xmlElem
'                        End If
'                   ''JD MAR1040 End If
'
'                'BMIDS00653 MDC 06/11/2002 - Include Employed >33% shares
'                'ElseIf strEmploymentStatus = strSelfEmployedID Or strEmploymentStatus = strConractId Then
'                ElseIf strEmploymentStatus = strSelfEmployedID _
'                        Or strEmploymentStatus = strConractId _
'                        Or strEmploymentStatus = strEmployed33PctShares Then
'                'BMIDS00653 MDC 06/11/2002 - End
'                    '-------------------------------------------------------------------------
'                    'Read AccountantsReference of the customer and append the data to response
'                    '-------------------------------------------------------------------------
'                    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'                    xmlNewRequest.setAttribute "OPERATION", "GETACCOUNTANTSREF"
'
'                    Set xmlTableElem = xmlDoc.createElement("ACCOUNTANT")
'                    xmlNewRequest.appendChild xmlTableElem
'
'                    xmlSetAttributeValue xmlTableElem, "CUSTOMERNUMBER", strCustomerNumber
'                    xmlSetAttributeValue xmlTableElem, "CUSTOMERVERSIONNUMBER", strCustVersionNumber
'                    xmlSetAttributeValue xmlTableElem, "EMPLOYMENTSEQUENCENUMBER", strEmpSeqNo
'
'                    strResponse = objAppProcBO.omAppProcRequest(xmlNewRequest.xml)
'                    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'                    On Error Resume Next
'                    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
'
'                    On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'                    blnRecordNotFound = False
'                    If lngErrNo <> 0 Then
'                        If lngErrNo <> oeRecordNotFound Then
'                            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'                        Else
'                            blnRecordNotFound = True
'                        End If
'                    End If
'
'                    If Not blnRecordNotFound Then
'                        ' Append to the response
'                        Set xmlOutListElem = xmlDoc.createElement("CONFIRMEDSELFEMPEARNINGSLIST")
'                        xmlOutEmpElem.appendChild xmlOutListElem
'
'                        Set xmlElemList = xmlDoc.selectNodes("//ACCOUNTANTREF")
'
'                        For Each xmlElem In xmlElemList
'                            Set xmlTableElem = xmlDoc.createElement("CONFIRMEDSELFEMPEARNINGS")
'                            Set xmlElem2 = xmlDoc.createElement("NETPROFITENTITLEMENT1")
'                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "NETPROFITENTITLEMENT1")
'                            xmlTableElem.appendChild xmlElem2
'
'                            Set xmlElem2 = xmlDoc.createElement("SALARY1")
'                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "SALARY1")
'                            xmlTableElem.appendChild xmlElem2
'
'                            Set xmlElem2 = xmlDoc.createElement("NETPROFITENTITLEMENT2")
'                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "NETPROFITENTITLEMENT2")
'                            xmlTableElem.appendChild xmlElem2
'
'                            Set xmlElem2 = xmlDoc.createElement("SALARY2")
'                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "SALARY2")
'                            xmlTableElem.appendChild xmlElem2
'
'                            Set xmlElem2 = xmlDoc.createElement("NETPROFITENTITLEMENT3")
'                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "NETPROFITENTITLEMENT3")
'                            xmlTableElem.appendChild xmlElem2
'
'                            Set xmlElem2 = xmlDoc.createElement("SALARY3")
'                            xmlElem2.Text = xmlGetAttributeText(xmlElem, "SALARY3")
'                            xmlTableElem.appendChild xmlElem2
'
'                            xmlOutListElem.appendChild xmlTableElem
'                        Next xmlElem
'                    Else
'                        '---------------------------------------------------
'                        'GetNetProfit details and append to the response
'                        '---------------------------------------------------
'                        Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'                        Set xmlTableElem = xmlDoc.createElement("NETPROFIT")
'                        xmlNewRequest.appendChild xmlTableElem
'
'                        Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
'                        xmlElem.Text = strCustomerNumber
'                        xmlTableElem.appendChild xmlElem
'
'                        Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
'                        xmlElem.Text = strCustVersionNumber
'                        xmlTableElem.appendChild xmlElem
'
'                        Set xmlElem = xmlDoc.createElement("EMPLOYMENTSEQUENCENUMBER")
'                        xmlElem.Text = strEmpSeqNo
'                        xmlTableElem.appendChild xmlElem
'
'                        strResponse = objCEBO.GetNetProfitDetails(xmlNewRequest.xml)
'                        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'
'                        On Error Resume Next
'                        lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
'
'                        On Error GoTo BuildRequestForCalculateAllowableIncomeVbErr
'                        blnRecordNotFound = False
'
'                        'BMIDS00653 MDC 06/11/2002
'                        If lngErrNo <> 0 Then
'                            If lngErrNo <> oeRecordNotFound Then
'                                m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
'                            Else
'                                blnRecordNotFound = True
'                            End If
'                        End If
'                        'BMIDS00653 MDC 06/11/2002 - End
'
'                        If Not blnRecordNotFound Then
'                            ' Append to the response
'                            Set xmlOutListElem = xmlDoc.createElement("DECLAREDSELFEMPEARNINGS")
'                            xmlOutEmpElem.appendChild xmlOutListElem
'
'                            Set xmlElemList = xmlDoc.selectNodes("//NETPROFIT")
'                            For Each xmlElem In xmlElemList
'                                Set xmlTableElem = xmlDoc.createElement("NETPROFIT")
'                                Set xmlElem2 = xmlDoc.createElement("YEAR1AMOUNT")
'                                xmlElem2.Text = xmlGetNodeText(xmlElem, ".//YEAR1AMOUNT")
'                                xmlTableElem.appendChild xmlElem2
'
'                                Set xmlElem2 = xmlDoc.createElement("YEAR2AMOUNT")
'                                xmlElem2.Text = xmlGetNodeText(xmlElem, ".//YEAR2AMOUNT")
'                                xmlTableElem.appendChild xmlElem2
'
'                                Set xmlElem2 = xmlDoc.createElement("YEAR3AMOUNT")
'                                xmlElem2.Text = xmlGetNodeText(xmlElem, ".//YEAR3AMOUNT")
'                                xmlTableElem.appendChild xmlElem2
'
'                                xmlOutListElem.appendChild xmlTableElem
'                            Next xmlElem
'                        End If
'                    End If
'                End If
'            Next xmlEmployment
'        End If
'
'        intLoop = intLoop + 1
'
'    Next xmlCustomer
'
'    Set BuildRequestForCalculateAllowableIncome = xmlResponseElem
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'BuildRequestForCalculateAllowableIncomeExit:
'    Set xmlDoc = Nothing
'    Set xmlOut = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlNewRequest = Nothing
'    Set xmlTableElem = Nothing
'    Set xmlElem = Nothing
'    Set xmlListElem = Nothing
'    Set xmlElemList = Nothing
'    Set xmlCustomer = Nothing
'    Set xmlCustomerList = Nothing
'    Set xmlEmployment = Nothing
'    Set xmlEmploymentList = Nothing
'    Set xmlCustomerRoles = Nothing
'    Set xmlOutCustomerElem = Nothing
'    Set xmlOutCustomerListElem = Nothing
'    Set xmlOutElem = Nothing
'    Set xmlOutListElem = Nothing
'    Set xmlOutEmpElem = Nothing
'    Set xmlOutEmpListElem = Nothing
'
'    Set objCustomerRoleBO = Nothing
'    Set objCEBO = Nothing
'    Set objComboDO = Nothing
'    Set objAppProcBO = Nothing
'
'    Exit Function
'
'BuildRequestForCalculateAllowableIncomeVbErr:
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
'
'    Set BuildRequestForCalculateAllowableIncome = m_objErrAssist.CreateErrorResponseEx()
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetAbort
'    End If
'
'    ' go to clean-up section
'    Resume BuildRequestForCalculateAllowableIncomeExit
'
'End Function


Private Function IApplicationQuoteBO_CalculateApproximateMonthlyCost(ByVal vxmlRequest As IXMLDOMElement) As MSXML2.IXMLDOMNode

    On Error GoTo CalculateApproximateMonthlyCostErr
    Const strFunctionName As String = "IApplicationQuoteBO_CalculateApproximateMonthlyCost"
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlQuoteData As IXMLDOMNode
    Dim xmlLoanComposition As IXMLDOMNode
        
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strApplicationNumber As String
    Dim intApplicationFactFindNumber As Integer
    Dim intMortgageSubQuoteNumber As Integer

    Dim objGlobalParameterDO As GlobalParameterDO
    Dim dblAFInterestRate As Double
    Dim intTypeOfApplication As Integer
    
    Dim dblTotalApproximateMonthlyCost As Double
    Dim dblFactor As Double
    Dim intRepaymentMethod As Integer
    Dim dblTotalLoanComponentAmount As Double
    Dim dblRepaymentVehicleMonthlyCost As Double
    Dim intTerminYears As Integer
    Dim intTermInMonths As Integer
    Dim intTotalTermInMonths As Integer
    Dim lngErrNo As Long
    
    '----------------------
    'Default Response XML
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '----------------------
    'Get GlobalParameter data
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    dblAFInterestRate = objGlobalParameterDO.GetCurrentParameterByType("AFInterestRate", "PERCENTAGE")
    
    '-------------------------------------------------------------------------------
    'Step 1 - Start
    '-------------------------------------------------------------------------------
    'Get Accepted/Active Quote data
    
    Set xmlQuoteData = IApplicationQuoteBO_GetAcceptedOrActiveQuoteData(vxmlRequest)
    
    'EP2_1079. If no quotation exists then exit cleanly with approximate cost = zero.
    'Check response for errors. Ignore record not found
    lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlQuoteData, True)
    If lngErrNo = oeRecordNotFound Then
        'set TotalApproximateMonthlyCost = 0 and Exit the method.
        xmlResponseElem.Text = 0
        Set IApplicationQuoteBO_CalculateApproximateMonthlyCost = xmlResponseElem
        GoTo CalculateApproximateMonthlyCostExit
    Else
        m_objErrAssist.CheckResponse xmlQuoteData, xmlResponseElem, True
    End If
    
    '----------------------
    ' Generate XML to pass into GetLoanCompositionDetails()
    
    strApplicationNumber = xmlAssistEx.xmlGetNodeText(vxmlRequest, ".//APPLICATIONNUMBER")
    intApplicationFactFindNumber = xmlAssistEx.xmlGetNodeAsInteger(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER")
    intMortgageSubQuoteNumber = xmlAssistEx.xmlGetNodeAsInteger(xmlQuoteData, ".//MORTGAGESUBQUOTE/MORTGAGESUBQUOTENUMBER")
        
    Set xmlOut = New MSXML2.DOMDocument
    Set xmlRequest = xmlOut.createElement("REQUEST")
    Set xmlLoanComposition = xmlOut.createElement("LOANCOMPOSITION")
    xmlRequest.appendChild xmlLoanComposition
    
    Set xmlNode = xmlOut.createElement("APPLICATIONNUMBER")
    xmlNode.Text = strApplicationNumber
    xmlLoanComposition.appendChild xmlNode
    
    Set xmlNode = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = intApplicationFactFindNumber
    xmlLoanComposition.appendChild xmlNode
    
    Set xmlNode = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
    xmlNode.Text = intMortgageSubQuoteNumber
    xmlLoanComposition.appendChild xmlNode
    
    '----------------------
    'If quotation created, call GetLoanCompositionDetails passing in generated XML request
    If intMortgageSubQuoteNumber <> 0 Then
        xmlOut.loadXML GetLoanCompositionDetails(xmlRequest.xml)
        m_objErrAssist.CheckResponse xmlOut.documentElement, xmlResponseElem, True
    Else
        xmlResponseElem.Text = 0
        Set IApplicationQuoteBO_CalculateApproximateMonthlyCost = xmlResponseElem
        GoTo CalculateApproximateMonthlyCostExit
    End If
    
    '----------------------
    'Iterate Loan Components to calculate TotalApproximateMonthlyCost
    
    dblFactor = 1 / (1 + (dblAFInterestRate / 1200))
    intTypeOfApplication = xmlAssistEx.xmlGetMandatoryNodeAsInteger(xmlOut, "//TYPEOFAPPLICATION")
    
    If Not (IsValidationType("ApplicationType", intTypeOfApplication, "TOE") And _
            IsValidationType("ApplicationType", intTypeOfApplication, "PSW")) Then

        
        For Each xmlNode In xmlOut.selectNodes("//LOANCOMPONENT")
        
            intRepaymentMethod = xmlAssistEx.xmlGetMandatoryNodeAsInteger(xmlNode, "REPAYMENTMETHOD")
            dblTotalLoanComponentAmount = xmlAssistEx.xmlGetMandatoryNodeAsDouble(xmlNode, "TOTALLOANCOMPONENTAMOUNT")
            
            If IsValidationType("RepaymentType", intRepaymentMethod, "C") Then
                intTerminYears = xmlAssistEx.xmlGetMandatoryNodeAsInteger(xmlNode, "//TERMINYEARS")
                intTermInMonths = xmlAssistEx.xmlGetMandatoryNodeAsInteger(xmlNode, "//TERMINMONTHS")
                intTotalTermInMonths = (intTerminYears * 12) + intTermInMonths
                
                dblTotalApproximateMonthlyCost = dblTotalApproximateMonthlyCost + _
                    ((dblTotalLoanComponentAmount * (dblAFInterestRate / 1200)) / _
                    (1 - (dblFactor ^ intTotalTermInMonths)))
                    
                dblTotalApproximateMonthlyCost = Round(dblTotalApproximateMonthlyCost, 2)
            Else
                'EP2_929
                dblRepaymentVehicleMonthlyCost = xmlAssistEx.xmlGetNodeAsDouble(xmlNode, "REPAYMENTVEHICLEMONTHLYCOST")
                
                dblTotalApproximateMonthlyCost = dblTotalApproximateMonthlyCost + _
                    dblTotalLoanComponentAmount * (dblAFInterestRate / 1200) + _
                    dblRepaymentVehicleMonthlyCost
                dblTotalApproximateMonthlyCost = Round(dblTotalApproximateMonthlyCost, 2)
            End If
                    
        Next
    
        xmlResponseElem.Text = dblTotalApproximateMonthlyCost
        
    End If
    
    '-------------------------------------------------------------------------------
    'Set response and exit
    
    Set IApplicationQuoteBO_CalculateApproximateMonthlyCost = xmlResponseElem

CalculateApproximateMonthlyCostExit:

    Set objGlobalParameterDO = Nothing
    Set xmlQuoteData = Nothing
    Set xmlResponseElem = Nothing
    Set xmlLoanComposition = Nothing
    Set xmlOut = Nothing
    Set xmlRequest = Nothing
    Set xmlNode = Nothing
    
    Exit Function

CalculateApproximateMonthlyCostErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CalculateApproximateMonthlyCost = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    Resume CalculateApproximateMonthlyCostExit

End Function

Private Function IApplicationQuoteBO_CalculateAndSaveBCSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAndSaveBCSubQuoteVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_CalculateAndSaveBCSubQuote"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.CalculateAndSaveBCSubQuote(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_CalculateAndSaveBCSubQuote = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateAndSaveBCSubQuoteExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IApplicationQuoteBO_CalculateAndSaveBCSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit

End Function

Private Function IApplicationQuoteBO_CreateAdHocFeeAndRemodel(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo CreateAdHocFeeAndRemodelVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_CreateAdHocFeeAndRemodel"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.CreateAdHocFeeAndRemodel(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_CreateAdHocFeeAndRemodel = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateAdHocFeeAndRemodelExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IApplicationQuoteBO_CreateAdHocFeeAndRemodel = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateAdHocFeeAndRemodelExit
End Function

Private Function IApplicationQuoteBO_GetAcceptedQuoteData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the mortgagesubquote data for the accepted quote.
'
'
' pass:         vstrXMLRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IApplicationQuoteBO_GetAcceptedQuoteData"
    Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objIQuotationBO As IQuotationBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    
    On Error GoTo GetAcceptedQuoteDataErr

    strNodeName = "APPLICATION"
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objIQuotationBO = New QuotationBO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objIQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    End If
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    Dim strAppNo As String
    Dim strFactFindNo As String
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "FACTFINDNUMBER is empty"
    End If
        
    ' Need to build the request to pass to GetApplicationData
    Dim xmlDomAppData As FreeThreadedDOMDocument40
    Dim xmlAppDataItem As IXMLDOMElement
    
    Set xmlAppDataElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    
    m_objXMLAssist.CheckNode xmlAppDataElem, TypeName(Me), strFunctionName
    
    Set xmlDomAppData = New FreeThreadedDOMDocument40
    xmlDomAppData.validateOnParse = False
    xmlDomAppData.setProperty "NewParser", True
    Set xmlDomAppData.documentElement = xmlAppDataElem
    Set xmlAppDataElem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONFACTFIND"))
    
    m_objXMLAssist.CheckNode xmlAppDataElem, TypeName(Me), strFunctionName
    
    ' Add the application number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONNUMBER"))
    m_objXMLAssist.CheckNode xmlAppDataItem, TypeName(Me), strFunctionName
    xmlAppDataItem.Text = strAppNo
    
    ' Add the application fact find number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlAppDataItem.Text = strFactFindNo
    
    ' Get the ApplicationData
    Set xmlAppResponseElem = objIApplicationDO.GetApplicationFactFindData(xmlDomAppData.documentElement)
    Dim strAcceptedQuoteNumber As String
    Dim xmlAppFactFindNode As IXMLDOMNode
    
    Set xmlAppFactFindNode = m_objXMLAssist.GetNode(xmlAppResponseElem, "//APPLICATIONFACTFIND")
    strAcceptedQuoteNumber = m_objXMLAssist.GetTagValue(xmlAppFactFindNode, "ACCEPTEDQUOTENUMBER")
    If Len(strAcceptedQuoteNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Accepted Quote not found"
    End If
    
    ' Get quotation Details
    Dim xmlSubQuoteResp As IXMLDOMElement
    
    Dim xmlDOMSubQuoteReq As FreeThreadedDOMDocument40
    Dim xmlSubQuoteRoot As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlSubQuoteItem As IXMLDOMElement
    
    Set xmlSubQuoteRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlDOMSubQuoteReq = xmlSubQuoteRoot.ownerDocument
    
    Set xmlSubQuoteElem = xmlSubQuoteRoot.appendChild(xmlDOMSubQuoteReq.createElement("QUOTATION"))
    
    ' Now add the application number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONNUMBER"))
    xmlSubQuoteItem.Text = strAppNo
        
    '... and application fact find number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlSubQuoteItem.Text = strFactFindNo
        
    ' accepted quote number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("QUOTATIONNUMBER"))
    xmlSubQuoteItem.Text = strAcceptedQuoteNumber
        
    Set xmlSubQuoteResp = objIQuotationBO.GetData(xmlSubQuoteRoot)
    m_objErrAssist.CheckResponse xmlSubQuoteResp, xmlResponseElem, True
    
    'Get the mortgagesubquote number of the accepted quote
    Dim xmlMSQNode As IXMLDOMNode
    Dim strMSQuoteNumber As String
    
    Set xmlMSQNode = m_objXMLAssist.GetNode(xmlSubQuoteResp, "//QUOTATION")
    strMSQuoteNumber = m_objXMLAssist.GetTagValue(xmlMSQNode, "MORTGAGESUBQUOTENUMBER")
    
    'now get the mortgagesubquote data
    Dim xmlMSubQuoteResp As IXMLDOMElement
    
    Dim xmlDOMMSubQuoteReq As FreeThreadedDOMDocument40
    Dim xmlMSubQuoteRoot As IXMLDOMElement
    Dim xmlMSubQuoteElem As IXMLDOMElement
    Dim xmlMSubQuoteItem As IXMLDOMElement
    
    Set xmlMSubQuoteRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlDOMMSubQuoteReq = xmlMSubQuoteRoot.ownerDocument
    
    Set xmlMSubQuoteElem = xmlMSubQuoteRoot.appendChild(xmlDOMSubQuoteReq.createElement("MORTGAGESUBQUOTE"))
    
    ' Now add the application number
    Set xmlMSubQuoteItem = xmlMSubQuoteElem.appendChild(xmlDOMMSubQuoteReq.createElement("APPLICATIONNUMBER"))
    xmlMSubQuoteItem.Text = strAppNo
        
    '... and application fact find number
    Set xmlMSubQuoteItem = xmlMSubQuoteElem.appendChild(xmlDOMMSubQuoteReq.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlMSubQuoteItem.Text = strFactFindNo
        
    ' mortgagesubquote number
    Set xmlMSubQuoteItem = xmlMSubQuoteElem.appendChild(xmlDOMMSubQuoteReq.createElement("MORTGAGESUBQUOTENUMBER"))
    xmlMSubQuoteItem.Text = strMSQuoteNumber
        
    Set xmlMSubQuoteResp = objIMortgageSubQuoteBO.GetMortgageSubQuoteData(xmlMSubQuoteRoot)
    m_objErrAssist.CheckResponse xmlMSubQuoteResp, xmlResponseElem, True
    
    
    
    
    ' Now build the XML to be returned from this method.
    ' Need to return Mortgage sub quote
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlMSubQuoteResp
            
    Set IApplicationQuoteBO_GetAcceptedQuoteData = xmlResponseElem
GetAcceptedQuoteDataExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set objIApplicationDO = Nothing
    Set objIQuotationBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlDomAppData = Nothing
    Set xmlAppDataItem = Nothing
    Set xmlSubQuoteResp = Nothing
    Set xmlDOMSubQuoteReq = Nothing
    Set xmlSubQuoteRoot = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlSubQuoteItem = Nothing
    Set xmlMSQNode = Nothing
    Set xmlMSubQuoteResp = Nothing
    Set xmlDOMMSubQuoteReq = Nothing
    Set xmlMSubQuoteRoot = Nothing
    Set xmlMSubQuoteElem = Nothing
    Set xmlMSubQuoteItem = Nothing

    Exit Function

GetAcceptedQuoteDataErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetAcceptedQuoteData = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetAcceptedQuoteDataExit
    
End Function

Private Function IApplicationQuoteBO_GetApplicationSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationSummaryVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_GetApplicationSummary"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlTempDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlOutTag As IXMLDOMElement
    Dim xmlTempRequest As FreeThreadedDOMDocument40
    Dim xmlTempRequestElem As IXMLDOMElement
    Dim xmlTempSearchElem As IXMLDOMElement
    Dim xmlTempCustomerElem As IXMLDOMElement
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    'Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNodeList As IXMLDOMNodeList
    Dim xmlTempOutNode As IXMLDOMNode
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String
    Dim strInCustomerNumber As String
    Dim strInCustomerVersionNumber As String
    Dim strTempResponseNode As String
    Dim strValidationGroupName As String
    Dim strValueID As String
    Dim objCustomerBO As CustomerBO
    Dim objNewPropertyBO As NewPropertyBO
    Dim objApplicationBO As ApplicationBO
    Dim objBuildingsandContentsSubQuoteBO As BuildingsAndContentsSubQuoteBO
    Dim objCostModellingBO As QuotationBO
    Dim objApplicationThirdPartyBO As ApplicationThirdPartyBO
    Dim objComboDO As ComboDO
    Dim objIntermediaryBO As IntermediaryBO
    Dim IntTempInt As Integer
    Dim bRecordFound As Boolean
    'JR - SYS2459, Omiplus24
    Dim xmlCTNode As IXMLDOMNode
    'End
    ' IK_EP502_05/05/2006
    Dim xmlAcceptedQuoteNode As IXMLDOMElement
    Dim xmlQuoteNode As IXMLDOMElement
    ' IK_EP502_05/05/2006_ends
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
           
    If vxmlRequest.nodeName = "APPLICATIONDETAILS" Then
       Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("APPLICATIONDETAILS").Item(0)
    End If
    
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONDETAILS tag not found"
    End If
    
    'Need to get the values for these fields out of the XML passed in
    strInApplicationNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER")
    strInApplicationFactFindNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "APPLICATIONFACTFINDNUMBER")
    strInCustomerNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "CUSTOMERNUMBER")
    strInCustomerVersionNumber = m_objXMLAssist.GetTagValue(vxmlRequest, "CUSTOMERVERSIONNUMBER")
           
    'Check the required values are present
    If Len(strInApplicationNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Application Number not found"
    End If
    
    If Len(strInApplicationFactFindNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Application Fact Find Number not found"
    End If
    
    If Len(strInCustomerNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Customer Number not found"
    End If
    
    If Len(strInCustomerVersionNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Customer Version Number not found"
    End If
   
    'Create temporary DOMDocument to use during CustomerBO method call
    Set xmlTempRequest = New FreeThreadedDOMDocument40
    xmlTempRequest.validateOnParse = False
    xmlTempRequest.setProperty "NewParser", True
    Set xmlTempRequestElem = xmlTempRequest.createElement("REQUEST")
    xmlTempRequest.appendChild xmlTempRequestElem
    Set xmlTempSearchElem = xmlTempRequest.createElement("SEARCH")
    xmlTempRequestElem.appendChild xmlTempSearchElem
    Set xmlTempCustomerElem = xmlTempRequest.createElement("CUSTOMER")
    xmlTempSearchElem.appendChild xmlTempCustomerElem
                    
    For Each xmlNode In xmlRequestNode.childNodes
        xmlTempCustomerElem.appendChild xmlNode.cloneNode(True)
    Next
                 
    'Create an instance of Customer.CustomerBO
    Set objCustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
            
    ' Call omCust.CustomerBO.GetCustomerDetails
    strTempResponseNode = objCustomerBO.GetCustomerDetails(xmlTempRequest.xml)
    Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
                
    'Append <CUSTOMER> Tag to outgoing xml string
    Set xmlOutTag = xmlOut.createElement("CUSTOMERNAME")
    Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
    
    Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/TITLE")
    If Not xmlTempNode Is Nothing Then
        xmlOutTag.appendChild xmlTempNode
    End If
    
    Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/FIRSTFORENAME")
    If Not xmlTempNode Is Nothing Then
        xmlOutTag.appendChild xmlTempNode
    End If
    
    Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/SURNAME")
    If Not xmlTempNode Is Nothing Then
        xmlOutTag.appendChild xmlTempNode
    End If
                
    'Get the item with a preferredmethod of contact, if there is one
    Set xmlTempOutNode = xmlTempResponse.selectSingleNode(".//CUSTOMERTELEPHONENUMBER[PREFERREDMETHODOFCONTACT='1']")
            
    'If there isn't one, get the item with the lowest telephonesequencenumber
    If xmlTempOutNode Is Nothing Then
        Set xmlTempNodeList = xmlTempResponse.getElementsByTagName("CUSTOMERTELEPHONENUMBER")
               
        Set xmlTempOutNode = xmlTempNodeList.Item(0)
        Dim intLength As Integer
        intLength = xmlTempNodeList.length
        Dim intIndex As Integer
        Dim intCurrentSeqNo As Integer
        Dim intLowestSeqNo As Integer
        
        For intIndex = 1 To intLength - 1
            Set xmlTempNode = xmlTempNodeList.Item(intIndex)
            
            intCurrentSeqNo = CInt(m_objXMLAssist.GetNodeValue(xmlTempNode, "TELEPHONESEQUENCENUMBER"))
            intLowestSeqNo = CInt(m_objXMLAssist.GetNodeValue(xmlTempOutNode, "TELEPHONESEQUENCENUMBER"))
                        
            If intCurrentSeqNo < intLowestSeqNo Then
                Set xmlTempOutNode = xmlTempNode
            End If
            
        Next
                
    End If
    If Not xmlTempOutNode Is Nothing Then
       'Construct <CUSTOMERCONTACTDETAILS> Element and append the required tags from the node
        Set xmlOutTag = xmlOut.createElement("CUSTOMERCONTACTDETAILS")
        Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
        xmlOutTag.appendChild m_objXMLAssist.GetNode(xmlTempOutNode, "TELEPHONENUMBER")
        xmlOutTag.appendChild m_objXMLAssist.GetNode(xmlTempOutNode, "USAGE")
        
        'JR SYS2459 (Omiplus24) include COUNTRYCODE and AREACODE
        xmlOutTag.appendChild m_objXMLAssist.GetNode(xmlTempOutNode, "COUNTRYCODE")
        xmlOutTag.appendChild m_objXMLAssist.GetNode(xmlTempOutNode, "AREACODE")
    End If
                
    'Create an instance of omApp.Application.NewPropertyBO
    Set objNewPropertyBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyBO")
        
    'Create temporary DOMDocument to use during NewPropertyBO method call
    Set xmlTempRequest = Nothing
    Set xmlTempRequest = New FreeThreadedDOMDocument40
    xmlTempRequest.validateOnParse = False
    xmlTempRequest.setProperty "NewParser", True
    Set xmlTempSearchElem = xmlTempRequest.createElement("NEWPROPERTYADDRESS")
    xmlTempRequest.appendChild xmlTempSearchElem
    
    For Each xmlNode In xmlRequestNode.childNodes
        xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
    Next
        
    'Call omApp.NewPropertyBO.GetNewPropertyAddress and check response
    strTempResponseNode = objNewPropertyBO.GetNewPropertyAddress(xmlTempSearchElem.xml)
    Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
    On Error Resume Next
    m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
    m_objErrAssist.SaveErr
    On Error GoTo GetApplicationSummaryVbErr
    
    bRecordFound = True
    
    'Record Not Found is OK
    If (m_objErrAssist.LastErr <> 0) Then
        If (m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound) Then
            m_objErrAssist.ReRaise
        Else
            bRecordFound = False
        End If
    End If
    
    If bRecordFound = True Then
    
        'Append <PROPERTYADDRESS> Tag to outgoing xml string
        Set xmlOutTag = xmlOut.createElement("PROPERTYADDRESS")
        Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/FLATNUMBER")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/BUILDINGORHOUSENAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/BUILDINGORHOUSENUMBER")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/STREET")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/DISTRICT")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/TOWN")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/COUNTY")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/POSTCODE")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/NEWPROPERTYADDRESS/ADDRESS/COUNTRY")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
    End If

    'Create temporary DOMDocument to use during ApplicationBO method call
    Set xmlTempRequest = Nothing
    Set xmlTempRequest = New FreeThreadedDOMDocument40
    xmlTempRequest.validateOnParse = False
    xmlTempRequest.setProperty "NewParser", True
    Set xmlTempSearchElem = xmlTempRequest.createElement("APPLICATIONFACTFIND")
    xmlTempRequest.appendChild xmlTempSearchElem
    
    For Each xmlNode In xmlRequestNode.childNodes
        xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
    Next
    
    'Create an instance of omApp.ApplicationBO and call GetApplicationFactFindData
    Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    strTempResponseNode = objApplicationBO.GetApplicationFactFindData(xmlTempRequest.xml)
    Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
    
    'Append <TYPEOFAPPLICATIONDETAILS> Tag to outgoing xml string
    Set xmlOutTag = xmlOut.createElement("TYPEOFAPPLICATIONDETAILS")
    Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
    
    Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONFACTFIND/TYPEOFAPPLICATION")
    If Not xmlTempNode Is Nothing Then
        xmlOutTag.appendChild xmlTempNode
    End If
    
    'Append <BUSINESSTYPE> Tag to outgoing xml string
    Set xmlOutTag = xmlOut.createElement("BUSINESSTYPE")
    Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
    
    Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONFACTFIND/DIRECTINDIRECTBUSINESS")
    If Not xmlTempNode Is Nothing Then
        xmlOutTag.appendChild xmlTempNode
        If Len(xmlTempNode.Text) > 0 Then
            strValueID = xmlTempNode.Text
        End If
    End If
        
    ' IK_EP502_05/05/2006
    ' nightmare, temp nodes overwritten etc.
    ' absolute 'kin shambles
    Set xmlAcceptedQuoteNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONFACTFIND/ACCEPTEDQUOTENUMBER")
    If Not xmlAcceptedQuoteNode Is Nothing Then
        ' LH_EP1042_31/07/2006
        If Len(xmlAcceptedQuoteNode.Text) > 0 Then
            bRecordFound = True
        Else
            bRecordFound = False
        End If
    End If
    
'    Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONFACTFIND/ACCEPTEDQUOTENUMBER")
'    If Not xmlTempNode Is Nothing Then
'        If Len(xmlTempNode.Text) > 0 Then
'            'Set tempoutnode to acceptedquotenumber for use in call to
'            'omCM.QuotationBO.GetQuotationSummary later on
'            Set xmlTempOutNode = xmlTempNode
'            m_objXMLAssist.ChangeNodeName xmlTempOutNode, "ACCEPTEDQUOTENUMBER", "QUOTATIONNUMBER"
'            bRecordFound = True
'        Else
'            bRecordFound = False
'        End If
'    End If
            
    If Len(strValueID) > 0 Then
    
        'Create an instance of omBase.ComboBO and call GetFirstComboValidation
        strValidationGroupName = "Direct/Indirect"
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        strTempResponseNode = objComboDO.GetFirstComboValidation(strValidationGroupName, strValueID)
        
        If strTempResponseNode = "I" Then
        
            'Create temporary DOMDocument to use during ApplicationBO method call
            Set xmlTempRequest = Nothing
            Set xmlTempRequest = New FreeThreadedDOMDocument40
            xmlTempRequest.validateOnParse = False
            xmlTempRequest.setProperty "NewParser", True
            Set xmlTempSearchElem = xmlTempRequest.createElement("APPLICATION")
            xmlTempRequest.appendChild xmlTempSearchElem
            
            For Each xmlNode In xmlRequestNode.childNodes
                xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
            Next
            
            'Call omApp.ApplicationBO.GetApplicationData for IntermediaryGuid
            Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
            strTempResponseNode = objApplicationBO.GetApplicationData(xmlTempRequest.xml)
            Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
            Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLATESTDETAILS/APPLICATION/INTERMEDIARYGUID")
            
            If Len(xmlTempNode.Text) > 0 Then
                'Create temporary DOMDocument to use during ApplicationBO method call
                Set xmlTempRequest = Nothing
                Set xmlTempRequest = New FreeThreadedDOMDocument40
                xmlTempRequest.validateOnParse = False
                xmlTempRequest.setProperty "NewParser", True
                Set xmlTempSearchElem = xmlTempRequest.createElement("INTERMEDIARY")
                xmlTempRequest.appendChild xmlTempSearchElem
                xmlTempSearchElem.appendChild xmlTempNode
                
                'Call omIM.IntermediaryBO.GetIndividualIntermediary with intermediaryguid
                Set objIntermediaryBO = m_objContext.CreateInstance(gstrINTERMEDIARY & ".IntermediaryBO")
                strTempResponseNode = objIntermediaryBO.GetIndividualIntermediary(xmlTempRequest.xml)
                Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
                m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
                
                'Append <BUSINESSSOURCE> Tag to outgoing xml string
                Set xmlOutTag = xmlOut.createElement("BUSINESSSOURCE")
                Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
                
                Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/INDIVIDUALINTERMEDIARY/LEADAGENTNAME")
                If Not xmlTempNode Is Nothing Then
                    xmlOutTag.appendChild xmlTempNode
                    ' IK_EP397_24/04/2006 - don't know how this has worked up to now
                    Set xmlTempOutNode = xmlTempNode
                    m_objXMLAssist.ChangeNodeName xmlTempOutNode, "LEADAGENTNAME", "NAME"
                End If
                
                'Append <BUSINESSSOURCECONTACTDETAILS> Tag to outgoing xml string
                Set xmlOutTag = xmlOut.createElement("BUSINESSSOURCECONTACTDETAILS")
                Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
                
'                Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/INDIVIDUALINTERMEDIARY/CONTACTDETAILS/TELEPHONENUMBER")
'                If Not xmlTempNode Is Nothing Then
'                    xmlOutTag.appendChild xmlTempNode
'                End If
'
'                Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/INDIVIDUALINTERMEDIARY/CONTACTDETAILS/TELEPHONEEXTENSIONNUMBER")
'                If Not xmlTempNode Is Nothing Then
'                    xmlOutTag.appendChild xmlTempNode
'                End If
                
                'JR SYS2459 (Omiplus) - Append ContactTelephoneDetails, there may be more than one
                Set xmlTempNodeList = xmlTempResponse.selectNodes("RESPONSE/INDIVIDUALINTERMEDIARY/CONTACTDETAILS/CONTACTTELEPHONEDETAILS")
                For Each xmlCTNode In xmlTempNodeList
                    xmlOutTag.appendChild xmlCTNode
                Next
                Set xmlCTNode = Nothing
            End If
                       
        Else
            'do other stuff for "D's"
            'Create temporary DOMDocument to use during ApplicationBO method call
            Set xmlTempRequest = Nothing
            Set xmlTempRequest = New FreeThreadedDOMDocument40
            xmlTempRequest.validateOnParse = False
            xmlTempRequest.setProperty "NewParser", True
            Set xmlTempSearchElem = xmlTempRequest.createElement("APPLICATION")
            xmlTempRequest.appendChild xmlTempSearchElem
            
            For Each xmlNode In xmlRequestNode.childNodes
                xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
            Next
            
            'Call omApp.ApplicationBO.GetUnitFromApplication
            Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
            strTempResponseNode = objApplicationBO.GetUnitFromApplication(xmlTempRequest.xml)
            Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
                       
            'Append <BUSINESSSOURCE> Tag to outgoing xml string
            Set xmlOutTag = xmlOut.createElement("BUSINESSSOURCE")
            Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
            
            Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/UNITSUMMARY/UNIT/UNITNAME")
            If Not xmlTempNode Is Nothing Then
                xmlOutTag.appendChild xmlTempNode
            End If
            
            'Append <BUSINESSSOURCECONTACTDETAILS> Tag to outgoing xml string
            Set xmlOutTag = xmlOut.createElement("BUSINESSSOURCECONTACTDETAILS")
            Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
            
'            Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/UNITSUMMARY/UNIT/CONTACTDETAILS/TELEPHONENUMBER")
'            If Not xmlTempNode Is Nothing Then
'                xmlOutTag.appendChild xmlTempNode
'            End If
'
'            Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/UNITSUMMARY/UNIT/CONTACTDETAILS/TELEPHONEEXTENSIONNUMBER")
'            If Not xmlTempNode Is Nothing Then
'                xmlOutTag.appendChild xmlTempNode
'            End If
            
            'JR Omiplus - Append ContactTelephoneDetails, there may be more than one
            Set xmlTempNodeList = xmlTempResponse.selectNodes("RESPONSE/UNITSUMMARY/UNIT/CONTACTDETAILS/CONTACTTELEPHONEDETAILS")
            For Each xmlCTNode In xmlTempNodeList
                xmlOutTag.appendChild xmlCTNode
            Next
            Set xmlCTNode = Nothing
        End If
    'Else
     '   m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, _
      '      "No value for Direct/Indirect business type found"
    End If
    
    ' LH_EP1042_31/07/2006
    If bRecordFound = True Then
        'Create temporary DOMDocument to use during method call
        Set xmlTempRequest = Nothing
        Set xmlTempRequest = New FreeThreadedDOMDocument40
        xmlTempRequest.validateOnParse = False
        xmlTempRequest.setProperty "NewParser", True
        Set xmlTempRequestElem = xmlTempRequest.createElement("REQUEST")
        xmlTempRequest.appendChild xmlTempRequestElem
        Set xmlTempSearchElem = xmlTempRequest.createElement("QUOTATION")
        xmlTempRequestElem.appendChild xmlTempSearchElem
        
        'EP1042 The next line causes an error if xmlTempOutNode is null:
        '       xmlTempSearchElem.appendChild xmlTempOutNode
        '       Added an If Not Is Nothing clause to amend this problem.
        '
        If Not xmlTempOutNode Is Nothing Then
            xmlTempSearchElem.appendChild xmlTempOutNode
        End If
        
        For Each xmlNode In xmlRequestNode.childNodes
            If Not xmlNode Is Nothing Then
                xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
            End If
        Next
        
        ' IK_EP502_05/05/2006
        Set xmlQuoteNode = xmlTempRequest.createElement("QUOTATIONNUMBER")
        xmlQuoteNode.Text = xmlAcceptedQuoteNode.Text
        xmlTempSearchElem.appendChild xmlQuoteNode
        ' IK_EP502_05/05/2006_ends
        
        'Create an instance of omCM.QuotationBO and call GetQuotationSummary
        Set objCostModellingBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        strTempResponseNode = objCostModellingBO.GetQuotationSummary(xmlTempRequest.xml)
        Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
        
        'Append <HOMEINSQUOTE> Tag to outgoing xml string
        Set xmlOutTag = xmlOut.createElement("HOMEINSQUOTE")
        Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
        
        'Append BC info
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/QUOTATIONSUMMARY/TOTALBCMONTHLYCOST")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        'Append <PAYMENTPROTECTIONQUOTE> Tag to outgoing xml string
        Set xmlOutTag = xmlOut.createElement("PAYMENTPROTECTIONQUOTE")
        Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
        
        'Append PP info
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/QUOTATIONSUMMARY/TOTALPPMONTHLYCOST")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        'Append <PRODUCTANDLOANDETAILSLIST> Tag to outgoing xml string
        Set xmlTempSearchElem = xmlOut.createElement("PRODUCTANDLOANDETAILSLIST")
        Set xmlTempDataNode = xmlResponseElem.appendChild(xmlTempSearchElem)
                
        'Make a list of loan components
        Set xmlTempNodeList = xmlTempResponse.selectNodes(".//LOANCOMPONENT")
        
        'loop through each loan component and add the relevant nodes
        For Each xmlNode In xmlTempNodeList
            
            'Append <PRODUCTANDLOANDETAILS> Tag to outgoing xml string
            Set xmlOutTag = xmlOut.createElement("PRODUCTANDLOANDETAILS")
            Set xmlTempDataNode = xmlTempSearchElem.appendChild(xmlOutTag)
            
            Set xmlTempNode = xmlNode.selectSingleNode("TOTALLOANCOMPONENTAMOUNT")
            If Not xmlTempNode Is Nothing Then
                m_objXMLAssist.ChangeNodeName xmlTempNode, "TOTALLOANCOMPONENTAMOUNT", "LOANAMOUNT"
                xmlOutTag.appendChild xmlTempNode
            End If
            
            Set xmlTempNode = xmlNode.selectSingleNode("TERMINYEARS")
            If Not xmlTempNode Is Nothing Then
                xmlOutTag.appendChild xmlTempNode
            End If
        
            Set xmlTempNode = xmlNode.selectSingleNode("TERMINMONTHS")
            If Not xmlTempNode Is Nothing Then
                xmlOutTag.appendChild xmlTempNode
            End If
            
            Set xmlTempNode = xmlNode.selectSingleNode("REPAYMENTMETHOD")
            If Not xmlTempNode Is Nothing Then
                xmlOutTag.appendChild xmlTempNode
            End If
            
            Set xmlTempNode = xmlNode.selectSingleNode("NETMONTHLYCOST")
            If Not xmlTempNode Is Nothing Then
                m_objXMLAssist.ChangeNodeName xmlTempNode, "NETMONTHLYCOST", "GROSSMONTHLYCOST"
                xmlOutTag.appendChild xmlTempNode
            End If
        
            'Check to see if loan component is ported
            Set xmlTempNode = xmlNode.selectSingleNode("PORTEDLOAN")
            If Not xmlTempNode Is Nothing Then
                xmlOutTag.appendChild xmlTempNode
                IntTempInt = xmlTempNode.Text
            
            
                'If it is, append the required tags for a Ported loan
                If IntTempInt = 1 Then
                                                    
                    Set xmlTempNode = xmlNode.selectSingleNode("MORTGAGEPRODUCTDESCRIPTION")
                    If Not xmlTempNode Is Nothing Then
                        xmlOutTag.appendChild xmlTempNode
                    End If
                    
                    Set xmlTempNode = xmlNode.selectSingleNode("INTERESTRATE")
                    If Not xmlTempNode Is Nothing Then
                        xmlOutTag.appendChild xmlTempNode
                    End If
                
                    Set xmlTempNode = xmlNode.selectSingleNode("MONTHLYREPAYMENT")
                    If Not xmlTempNode Is Nothing Then
                        xmlOutTag.appendChild xmlTempNode
                    End If
                'Otherwise append the Tags for a non-Ported loan
                Else
                    Set xmlTempNode = xmlNode.selectSingleNode("PRODUCTNAME")
                    If Not xmlTempNode Is Nothing Then
                        xmlOutTag.appendChild xmlTempNode
                    End If
                    Set xmlTempNode = xmlNode.selectSingleNode("WITHDRAWNDATE")
                    If Not xmlTempNode Is Nothing Then
                        xmlOutTag.appendChild xmlTempNode
                    End If
                    
                    Set xmlTempNode = xmlNode.selectSingleNode("RATE")
                    If Not xmlTempNode Is Nothing Then
                        m_objXMLAssist.ChangeNodeName xmlTempNode, "RATE", "INTERESTRATE"
                        xmlOutTag.appendChild xmlTempNode
                    End If
                End If
            End If
            
        Next
    End If
    
    'Create temporary DOMDocument to use during ApplicationThirdPartyBO method call
    Set xmlTempRequest = Nothing
    Set xmlTempRequest = New FreeThreadedDOMDocument40
    xmlTempRequest.validateOnParse = False
    xmlTempRequest.setProperty "NewParser", True
    Set xmlTempSearchElem = xmlTempRequest.createElement("APPLICATIONLEGALREP")
    xmlTempRequest.appendChild xmlTempSearchElem
    
    For Each xmlNode In xmlRequestNode.childNodes
        xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
    Next
      
    'Create an instance of omApp.Application.ApplicationThirdPartyBO and call GetLegalRep
    Set objApplicationThirdPartyBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
    strTempResponseNode = objApplicationThirdPartyBO.GetLegalRep(xmlTempRequest.xml)
    Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
    On Error Resume Next
    m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
    m_objErrAssist.SaveErr
    On Error GoTo GetApplicationSummaryVbErr
    
    bRecordFound = True
    
    'Record Not Found is OK
    If (m_objErrAssist.LastErr <> 0) Then
        If (m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound) Then
            m_objErrAssist.ReRaise
        Else
            bRecordFound = False
        End If
    End If
    
    If bRecordFound = True Then
            
        'Append <LEGALREPS> Tag to outgoing xml string
        Set xmlOutTag = xmlOut.createElement("LEGALREPS")
        Set xmlTempDataNode = xmlResponseElem.appendChild(xmlOutTag)
        
        'Append the desired legal rep tags
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/COMPANYNAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/DXID")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/DXLOCATION")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTFORENAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTSURNAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
'        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/TELEPHONENUMBER")
'        If Not xmlTempNode Is Nothing Then
'            xmlOutTag.appendChild xmlTempNode
'        End If
        
        'JR SYS2459 (Omiplus) - Append ContactTelephoneDetails, there may be more than one
        Set xmlTempNodeList = xmlTempResponse.selectNodes("RESPONSE/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTTELEPHONEDETAILS")
        For Each xmlCTNode In xmlTempNodeList
            xmlOutTag.appendChild xmlCTNode
        Next
        Set xmlCTNode = Nothing
        'End
    End If
    
    'Create temporary DOMDocument to use during ApplicationBO method call
    Set xmlTempRequest = Nothing
    Set xmlTempRequest = New FreeThreadedDOMDocument40
    xmlTempRequest.validateOnParse = False
    xmlTempRequest.setProperty "NewParser", True
    Set xmlTempSearchElem = xmlTempRequest.createElement("APPLICATION")
    xmlTempRequest.appendChild xmlTempSearchElem
    
    For Each xmlNode In xmlRequestNode.childNodes
        xmlTempSearchElem.appendChild xmlNode.cloneNode(True)
    Next
    
    'Call omApp.ApplicationBO.FindMainApplicants to get applicants' customernumber
    'to pass to following call
    strTempResponseNode = objApplicationBO.FindMainApplicants(xmlTempRequest.xml)
    Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
    
           
    'Append <APPLICANTDETAILSLIST> Tag to outgoing xml string
    Set xmlTempCustomerElem = xmlOut.createElement("APPLICANTLIST")
    Set xmlTempDataNode = xmlResponseElem.appendChild(xmlTempCustomerElem)
    
    'Make a list of Applicants
    'Set xmlTempNodeList = xmlTempResponse.getElementsByTagName(".//CUSTOMERROLE")
    'GD SYS2970 16/11/01
    Set xmlTempNodeList = xmlTempResponse.selectNodes(".//CUSTOMERROLE")
    
    'loop through each loan component and add the relevant nodes
    For Each xmlNode In xmlTempNodeList
        
       'Create temporary DOMDocument to use during CustomerBO method call
        Set xmlTempRequest = Nothing
        Set xmlTempRequest = New FreeThreadedDOMDocument40
        xmlTempRequest.validateOnParse = False
        xmlTempRequest.setProperty "NewParser", True
        Set xmlTempSearchElem = xmlTempRequest.createElement("CUSTOMER")
        xmlTempRequest.appendChild xmlTempSearchElem
           
        'Append the CUSTOMERNUMBER tag to the request for this particular
        'iteration through the list as this ir required in the called method
        Set xmlTempNode = xmlNode.selectSingleNode("CUSTOMERNUMBER")
        If Not xmlTempNode Is Nothing Then
            xmlTempSearchElem.appendChild xmlTempNode
        End If
        
        'Append the CUSTOMERVERSIONNUMBER tag to the request for this particular
        'iteration through the list as this ir required in the called method
        Set xmlTempNode = xmlNode.selectSingleNode("CUSTOMERVERSIONNUMBER")
        If Not xmlTempNode Is Nothing Then
            xmlTempSearchElem.appendChild xmlTempNode
        End If
        
        'Call omCust.CustomerBO.GetPersonalDetails for each applicant
        strTempResponseNode = objCustomerBO.GetPersonalDetails(xmlTempRequest.xml)
        Set xmlTempResponse = m_objXMLAssist.load(strTempResponseNode, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlTempResponse.documentElement, xmlResponseElem, True
        
        'Append <APPLICANTDETAILS> Tag to outgoing xml string
        Set xmlOutTag = xmlOut.createElement("APPLICANTDETAILS")
        Set xmlTempDataNode = xmlTempCustomerElem.appendChild(xmlOutTag)
    
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/TITLE")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/FIRSTFORENAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/SECONDFORENAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
        Set xmlTempNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMER/CUSTOMERVERSION/SURNAME")
        If Not xmlTempNode Is Nothing Then
            xmlOutTag.appendChild xmlTempNode
        End If
        
    Next
       
    Set IApplicationQuoteBO_GetApplicationSummary = xmlResponseElem
    
    m_objContext.SetComplete
    
GetApplicationSummaryExit:
    
    Set objCustomerBO = Nothing
    Set objNewPropertyBO = Nothing
    Set objApplicationBO = Nothing
    Set objBuildingsandContentsSubQuoteBO = Nothing
    Set objCostModellingBO = Nothing
    Set objApplicationThirdPartyBO = Nothing
    Set objComboDO = Nothing
    Set objIntermediaryBO = Nothing
    'JR - SYS2459 (Omiplus24)
    Set xmlCTNode = Nothing
    ' IK_EP502_05/05/2006
    Set xmlQuoteNode = Nothing
    

    Exit Function

GetApplicationSummaryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetApplicationSummary = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetApplicationSummaryExit

End Function

Private Function IApplicationQuoteBO_GetHighestEarners( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetHighestEarnersVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_GetHighestEarners"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement

    Dim objCustomerEmploymentBO As CustomerEmploymentBO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlData As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement, xmlTempElem As IXMLDOMElement
    Dim vnt As Variant
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    If m_objContext Is Nothing Then
        Set objCustomerEmploymentBO = New CustomerEmploymentBO
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objCustomerEmploymentBO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
                
    Set xmlRequestNode = vxmlRequest.getElementsByTagName("CUSTOMERLIST").Item(0)
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERLIST tag not found"
    End If

    Dim strXMLData As String
   
    strXMLData = objCustomerEmploymentBO.FindEmploymentAndIncomeDetails(vxmlRequest.xml)
    
    Set xmlData = m_objXMLAssist.load(strXMLData, TypeName(Me), strFunctionName)
    'MS 20/06/00 Trap errors raised
    m_objErrAssist.CheckResponse xmlData.documentElement, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlData.documentElement
    
    'MS 07/07/00 Dont append data again
    'xmlDataNode.appendChild xmlData.documentElement
        
    Dim blnIncludeGuarantors As Boolean
    blnIncludeGuarantors = False
    vnt = objGlobalParameterDO.GetCurrentParameterByType("IncludeGuarantors", "BOOLEAN")

    If Not IsNull(vnt) Then
        blnIncludeGuarantors = CBool(vnt)
    End If
        
    Dim xmlCustomerNodeList As IXMLDOMNodeList
                    
    Set xmlCustomerNodeList = xmlResponseElem.getElementsByTagName("CUSTOMER")
    Dim nloop As Integer
    
    Dim strRoleType As String   ', strCustomerNumber As String, strCustomerVersionNumber As String
    'Dim strSearchPattern As String
    Dim blnIgnore As Boolean

    Dim lngMaximumAmount As Long
    Dim lngLowerAmount As Long
    Dim intMaximumNode As Integer
    Dim intLowerNode As Integer
    
    lngMaximumAmount = 0
    lngLowerAmount = 0
    
    'get all the customers
    For nloop = 0 To (xmlCustomerNodeList.length - 1)
                
        Set xmlElement = xmlCustomerNodeList.Item(nloop)
        strRoleType = m_objXMLAssist.GetTagValue(xmlElement, "CUSTOMERROLETYPE")
        
        blnIgnore = False
        If strRoleType = "1" Or strRoleType = "2" Then
            If strRoleType = "2" And Not blnIncludeGuarantors Then
                'type 2 but we dont want to include them
                blnIgnore = True
            End If
                        
            If Not blnIgnore Then
                'do the calculation
                Dim lngAmount As Long
                
                Set xmlElement = xmlElement.getElementsByTagName("INCOMESUMMARY").Item(0)
                
                '** DIFFERENCE from copied method : use 'AllowableAnnualIncome'
                '** instead of 'TotalGrossEarnedIncome'. The later field does not exist on the table
                
                'APS 01/06/00 - SYS0799 Changed tag name from ALLOWABLEANNUALINCOME TO TOTALGROSSEARNEDINCOME
                lngAmount = CSafeLng(m_objXMLAssist.GetTagValue(xmlElement, "TOTALGROSSEARNEDINCOME"))
                'New value greater than current max
                If lngAmount >= lngMaximumAmount Then
                    'so set lower amount at previous max
                    lngLowerAmount = lngMaximumAmount
                    intLowerNode = intMaximumNode
                    'then set max to current
                    lngMaximumAmount = lngAmount
                    intMaximumNode = nloop
                ElseIf lngAmount >= lngLowerAmount Then
                    'new value greater than previous so set to this value
                    lngLowerAmount = lngAmount
                    intLowerNode = nloop
                End If
            End If
        End If
    Next
    'AQR SYS4505 DRC - Warning message if none found
    If lngMaximumAmount > 0 Then
        Set xmlElement = xmlCustomerNodeList.Item(intMaximumNode)
        m_objXMLAssist.ChangeNodeName xmlElement, "CUSTOMER", "CUSTOMERA"
    Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "No Employment Income has been Input"
    End If

    If lngLowerAmount > 0 Then
        Set xmlElement = xmlCustomerNodeList.Item(intLowerNode)
        m_objXMLAssist.ChangeNodeName xmlElement, "CUSTOMER", "CUSTOMERB"
    End If

    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    Set IApplicationQuoteBO_GetHighestEarners = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetHighestEarnersExit:
    
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objCustomerEmploymentBO = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlData = Nothing
    Set xmlElement = Nothing
    Set xmlTempElem = Nothing

    Exit Function

GetHighestEarnersVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetHighestEarners = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetHighestEarnersExit
End Function

Private Function IApplicationQuoteBO_RecommendQuotation( _
        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RecommendQuotationVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_RecommendQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.RecommendQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_RecommendQuotation = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecommendQuotationExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IApplicationQuoteBO_RecommendQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecommendQuotationExit

End Function

Private Function IApplicationQuoteBO_RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RecostMortgageComponentAndLifeCoverErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_RecostMortgageComponentAndLifeCover"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.RecostMortgageComponentAndLifeCover(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_RecostMortgageComponentAndLifeCover = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostMortgageComponentAndLifeCoverExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IApplicationQuoteBO_RecostMortgageComponentAndLifeCover = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit

End Function

Private Function IApplicationQuoteBO_RecostQuotationForRateChange(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Remodels the repayment costs and APR to ensure that the latest rates
'               have been applied prior to issuing the offer
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo RecostQuotationForRateChangeVbErr
    
Const strFunctionName As String = "IApplicationQuoteBO_RecostQuotationForRateChange"
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
Dim strValue As String
Dim xmlAppQuote As IXMLDOMNode

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    
    'Validate that required elements are present
    Set xmlAppQuote = m_objXMLAssist.GetMandatoryNode(vxmlRequest, "APPLICATIONQUOTE")
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlAppQuote, "APPLICATIONNUMBER")
    'BMIDS856 No need to validate APPLICATIONNUMBER twice
    'strValue = m_objXmlAssist.GetMandatoryElementText(xmlAppQuote, "APPLICATIONNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlAppQuote, "APPLICATIONFACTFINDNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlAppQuote, "QUOTATIONNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlAppQuote, "MORTGAGESUBQUOTENUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(xmlAppQuote, "LOANCOMPONENTS/LOANCOMPONENTSEQUENCENUMBER")
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.RecostQuotationForRateChange(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_RecostQuotationForRateChange = xmlResponseElem
    m_objContext.SetComplete
    
RecostQuotationForRateChangeExit:
    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlAppQuote = Nothing
    
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

    Set IApplicationQuoteBO_RecostQuotationForRateChange = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RecostQuotationForRateChangeExit

End Function

Private Function IApplicationQuoteBO_ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ReinstateQuotationVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_ReinstateQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAppQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIAppQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIAppQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIAppQuoteTxBO.ReinstateQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    ' If every thing is successful, add message 217 as Warning and send it back
    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 217
    Set IApplicationQuoteBO_ReinstateQuotation = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ReinstateQuotationExit:

    Set objIAppQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
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

    Set IApplicationQuoteBO_ReinstateQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateQuotationExit

End Function

Private Function IApplicationQuoteBO_ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ResetMortgageSubQuoteVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_ResetMortgageSubQuote"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAppQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIAppQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIAppQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIAppQuoteTxBO.ResetMortgageSubQuote(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set IApplicationQuoteBO_ResetMortgageSubQuote = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ResetMortgageSubQuoteExit:

    Set objIAppQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IApplicationQuoteBO_ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit
End Function


Private Function IApplicationQuoteBO_UpdateMortgageSubquote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Updates Mortgage Subquote.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGESUBQUOTE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   </MORTGAGESUBQUOTE>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "IApplicationQuoteBO_UpdateMortgageSubquote"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim objICostModelAssistBO As ICostModelAssistBO
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If

    ' call common costModelAssist method
    Set xmlTempResponseNode = objICostModelAssistBO.UpdateMortgageSubquote(vxmlRequest)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    Set IApplicationQuoteBO_UpdateMortgageSubquote = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_UpdateMortgageSubquote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit
End Function

' BMIDS977 IApplicationQuoteBO_UpdateQuotationForAddedCosts removed
'Private Function IApplicationQuoteBO_UpdateQuotationForAddedCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:  Remodels the mortgage when fees are added to the loan
''               invokes appropriate transactioning (Tx) business object
'' pass:         vxmlRequest  xml Request data stream containing data to be persisted
'' return:       xml Response Node
''------------------------------------------------------------------------------------------
'On Error GoTo UpdateQuotationForAddedCostsVbErr
'
'    Const strFunctionName As String = "IApplicationQuoteBO_UpdateQuotationForAddedCosts"
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlTempResponseNode As IXMLDOMNode
'    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
'    Set xmlOut = New FreeThreadedDOMDocument40
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'    Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
'    Set xmlTempResponseNode = objIApplicationQuoteTxBO.UpdateQuotationForAddedCosts(vxmlRequest)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'    Set IApplicationQuoteBO_UpdateQuotationForAddedCosts = xmlResponseElem
'    m_objContext.SetComplete
'
'UpdateQuotationForAddedCostsExit:
'    Set objIApplicationQuoteTxBO = Nothing
'    Set xmlOut = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlTempResponseNode = Nothing
'    Exit Function
'
'UpdateQuotationForAddedCostsVbErr:
'    If m_objErrAssist.IsWarning = True Then
'        m_objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'    m_objErrAssist.AddToErrSource strFunctionName
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'    Set IApplicationQuoteBO_UpdateQuotationForAddedCosts = m_objErrAssist.CreateErrorResponseEx()
'    m_objContext.SetAbort
'
'    ' Go to clean-up section
'    Resume UpdateQuotationForAddedCostsExit
'End Function

Private Function IApplicationQuoteBO_ValidateManualAdjustmentPercent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo ValidateManualAdjustmentVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_ValidateManualAdjustmentPercent"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlCompetency As FreeThreadedDOMDocument40
    Dim xmlCompElem As IXMLDOMElement
    Dim dblUserDecreaseLimitPercent As Double
    Dim dblUserIncreaseLimitPercent As Double
    Dim dblRateAdjustment As Double
    Dim xmlProduct As FreeThreadedDOMDocument40
    Dim xmlProductElem As IXMLDOMElement
    Dim xmlProductNode As IXMLDOMNode
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlKeyElement As IXMLDOMElement
    Dim xmlProdResponse As FreeThreadedDOMDocument40
    Dim xmlProdNode As IXMLDOMNode
    Dim xmlAppElement As IXMLDOMElement
    Dim xmlAppNode As IXMLDOMNode
        
    Dim objIOrganisationDO As IOrganisationDO
    Dim objIMortgageProductBO As MortgageProductBO
    Dim objConvAssist As ConvertAssist
    
    If m_objContext Is Nothing Then
        Set objIOrganisationDO = New OrganisationDO
        Set objIMortgageProductBO = New MortgageProductBO
    Else
        Set objIOrganisationDO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationDO")
        Set objIMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
    End If
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Validate users adjustment limits
    Set xmlCompetency = New FreeThreadedDOMDocument40
    xmlCompetency.validateOnParse = False
    xmlCompetency.setProperty "NewParser", True
    Set xmlCompElem = xmlCompetency.createElement("REQUEST")
    xmlCompElem.setAttribute "USERID", vxmlRequest.getAttribute("USERID")
    xmlCompetency.appendChild xmlCompElem
    Set xmlTempResponseNode = xmlCompElem.appendChild(xmlOut.createElement("USERID"))
    xmlTempResponseNode.Text = vxmlRequest.getAttribute("USERID")
    
    'Set xmlTempResponseNode = objIOrganisationDO.GetCurrentCompetency(xmlCompElem)
    Set xmlTempResponseNode = objIOrganisationDO.ValidateUserMandateLevel(xmlCompElem)
    'm_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set objConvAssist = New ConvertAssist
    dblRateAdjustment = objConvAssist.CSafeDbl(vxmlRequest.getElementsByTagName("MANUALADJUSTMENTPERCENT").Item(0).Text)
    dblUserDecreaseLimitPercent = objConvAssist.CSafeDbl(xmlTempResponseNode.selectSingleNode("USERDECREASELIMITPERCENT").Text)
    dblUserIncreaseLimitPercent = objConvAssist.CSafeDbl(xmlTempResponseNode.selectSingleNode("USERINCREASELIMITPERCENT").Text)
    
    ' JLD MSMS0058
    If dblRateAdjustment < 0 Then
        If Abs(dblRateAdjustment) > dblUserDecreaseLimitPercent Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6001
        End If
    Else
        If Abs(dblRateAdjustment) > dblUserIncreaseLimitPercent Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6001
        End If
    End If
    
    'Validate product parameter levels
    Set xmlProduct = New FreeThreadedDOMDocument40
    xmlProduct.validateOnParse = False
    xmlProduct.setProperty "NewParser", True
    Set xmlProductElem = xmlProduct.createElement("REQUEST")
    xmlProduct.appendChild xmlProductElem
    
    'JD BMIDS763 Application information also required - number, factfind and date
    Set xmlAppElement = xmlProduct.createElement("APPLICATION")
    Set xmlAppNode = xmlProductElem.appendChild(xmlAppElement)
    Set xmlTempElement = xmlProduct.createElement("APPLICATIONNUMBER")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("APPLICATIONNUMBER").Item(0).Text
    xmlAppNode.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(0).Text
    xmlAppNode.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("APPLICATIONDATE")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("APPLICATIONDATE").Item(0).Text
    xmlAppNode.appendChild xmlTempElement
    
    Set xmlKeyElement = xmlProduct.createElement("MORTGAGEPRODUCT")
    xmlProductElem.appendChild xmlKeyElement
    Set xmlTempElement = xmlProduct.createElement("MORTGAGEPRODUCTCODE")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("MORTGAGEPRODUCTCODE").Item(0).Text
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("STARTDATE")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("STARTDATE").Item(0).Text
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("AMOUNTREQUESTED")
    If Not vxmlRequest.getElementsByTagName("AMOUNTREQUESTED").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("AMOUNTREQUESTED").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("PURCHASEPRICE")
    If Not vxmlRequest.getElementsByTagName("PURCHASEPRICE").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("PURCHASEPRICE").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("LTV")
    If Not vxmlRequest.getElementsByTagName("LTV").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("LTV").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("TYPEOFAPPLICATION")
    If Not vxmlRequest.getElementsByTagName("TYPEOFAPPLICATION").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("TYPEOFAPPLICATION").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("LOCATION")
    If Not vxmlRequest.getElementsByTagName("LOCATION").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("LOCATION").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("TYPEOFVALUATION")
    If Not vxmlRequest.getElementsByTagName("TYPEOFVALUATION").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("TYPEOFVALUATION").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("FEETYPE")
    If Not vxmlRequest.getElementsByTagName("FEETYPE").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("FEETYPE").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    
    ' PSC 19/02/2007 EP2_1488 - Start
    Set xmlTempElement = xmlGetNode(vxmlRequest, "VALIDATEDATA/ORIGINALLTV")
    
    If Not xmlTempElement Is Nothing Then
        xmlKeyElement.appendChild xmlTempElement.cloneNode(True)
    End If
    
    Set xmlTempElement = xmlProduct.createElement("APPLICATIONDATE")
    xmlTempElement.Text = xmlGetNodeText(vxmlRequest, "VALIDATEDATA/APPLICATIONDATE")
    xmlKeyElement.appendChild xmlTempElement
    
    Set xmlTempElement = xmlProduct.createElement("APPLICATIONNUMBER")
    xmlTempElement.Text = xmlGetNodeText(vxmlRequest, "VALIDATEDATA/APPLICATIONNUMBER")
    xmlKeyElement.appendChild xmlTempElement
    ' PSC 19/02/2007 EP2_1488 - End
    
    Set xmlProdResponse = New FreeThreadedDOMDocument40
    xmlProdResponse.validateOnParse = False
    xmlProdResponse.setProperty "NewParser", True
    xmlProdResponse.loadXML (objIMortgageProductBO.GetFullProductDetails(xmlProduct.xml))
    m_objErrAssist.CheckResponse xmlProdResponse.documentElement, xmlResponseElem, True
    
    Set xmlProdNode = xmlProdResponse.documentElement.selectSingleNode(".//INTERESTRATETYPELIST/INTERESTRATETYPE[INTERESTRATETYPESEQUENCENUMBER='1']")
    If Not xmlProdNode Is Nothing Then
        If xmlProdNode.selectSingleNode("RATETYPE").Text = "F" Or _
           xmlProdNode.selectSingleNode("RATETYPE").Text = "B" Or _
           xmlProdNode.selectSingleNode("RATETYPE").Text = "D" Then
            Dim dblDiffLowerLimit As Double
            Dim dblDiffUpperLimit As Double
            dblDiffLowerLimit = objConvAssist.CSafeDbl(xmlProdResponse.documentElement.selectSingleNode(".//DIFFERENTIALLOWERLIMIT").Text)
            dblDiffUpperLimit = objConvAssist.CSafeDbl(xmlProdResponse.documentElement.selectSingleNode(".//DIFFERENTIALUPPERLIMIT").Text)
            
            'BMIDS959 Correct limit check
            If dblRateAdjustment < 0 Then
                If Abs(dblRateAdjustment) > dblDiffLowerLimit Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6002
                End If
            Else
                If Abs(dblRateAdjustment) > dblDiffUpperLimit Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6002
                End If
            End If
        Else
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6003
        End If
    End If
    
    ' Check that the resulting resolved rate does not go beneath the decrease limit percent
    Dim dblResolvedRate As Double
    Dim dblProdLimit As Double
    Dim strProdLimit As String
    strProdLimit = xmlProdResponse.documentElement.selectSingleNode(".//MANUALDECREASELIMITPERCENT").Text
    dblProdLimit = objConvAssist.CSafeDbl(strProdLimit)
    dblResolvedRate = objConvAssist.CSafeDbl(vxmlRequest.getElementsByTagName("RESOLVEDRATE").Item(0).Text)
    If dblResolvedRate < dblProdLimit Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6006, "", strProdLimit
    End If
    
    Set IApplicationQuoteBO_ValidateManualAdjustmentPercent = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateManualAdjustmentExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCompetency = Nothing
    Set xmlCompElem = Nothing
    Set xmlProduct = Nothing
    Set xmlProductElem = Nothing
    Set xmlProductNode = Nothing
    Set xmlTempElement = Nothing
    Set xmlKeyElement = Nothing
    Set xmlProdResponse = Nothing
    Set xmlProdNode = Nothing
    Set objConvAssist = Nothing

    Exit Function

ValidateManualAdjustmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_ValidateManualAdjustmentPercent = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateManualAdjustmentExit

End Function

Private Function IApplicationQuoteBO_ValidateMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Decides if the mortgagesubquote is valid or not.
'
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>
'   <BASICQUOTATIONDETAILS>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'       <MORTGAGESUBQUOTENUMBER>7</MORTGAGESUBQUOTENUMBER>
'   </BASICQUOTATIONDETAILS>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS" <VALIDMORTGAGESUBQUOTE>1</VALIDMORTGAGESUBQUOTE>
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IApplicationQuoteBO_ValidateMortgageSubQuote"


    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDomValidData As FreeThreadedDOMDocument40
    Set xmlDomValidData = New FreeThreadedDOMDocument40
    xmlDomValidData.validateOnParse = False
    xmlDomValidData.setProperty "NewParser", True
    Dim xmlDOMCustomer As FreeThreadedDOMDocument40
    Set xmlDOMCustomer = New FreeThreadedDOMDocument40
    xmlDOMCustomer.validateOnParse = False
    xmlDOMCustomer.setProperty "NewParser", True
    Dim xmlDOMValid As FreeThreadedDOMDocument40
    Set xmlDOMValid = New FreeThreadedDOMDocument40
    xmlDOMValid.validateOnParse = False
    xmlDOMValid.setProperty "NewParser", True
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlValidDataElem As IXMLDOMElement
    Dim xmlValidDataResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim objICustFinancialDO As ICustomerFinancialDO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objICostModelAssistDO As ICostModelAssistDO
    Dim objICustEmploymentDO As ICustomerEmploymentDO
'SYS0977   Dim objCostModelAssistBO As ICostModelAssistBO
    
    If m_objContext Is Nothing Then
        Set objICustFinancialDO = New CustomerFinancialDO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objICustEmploymentDO = New CustomerEmploymentDO
 'SYS0977        Set objCostModelAssistBO = New CostModelAssistBO
    Else
        Set objICustEmploymentDO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentDO")
        Set objICustFinancialDO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialDO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
 'SYS0977      Set objCostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & "BASICQUOTATIONDETAILS", True)
    
    Dim strAppNo As String
    Dim strFactFindNo As String
    Dim strMortgageSubQuote As String
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuote = m_objXMLAssist.GetTagValue(xmlRequestElem, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strMortgageSubQuote) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGESUBQUOTENUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONFACTFINDNUMBER is empty"
    End If
        
    ' Need to build the request to pass to GetDataForSubQuoteValidation()
    Dim xmlValidDataItem As IXMLDOMElement
    
    Set xmlValidDataElem = xmlDomValidData.createElement("BASICQUOTATIONDETAILS")
    Set xmlValidDataElem = xmlDomValidData.appendChild(xmlValidDataElem)
    m_objXMLAssist.CheckNode xmlValidDataElem, TypeName(Me), strFunctionName
    
    ' Add the application number
    Set xmlValidDataItem = xmlValidDataElem.appendChild(xmlDomValidData.createElement("APPLICATIONNUMBER"))
    m_objXMLAssist.CheckNode xmlValidDataItem, TypeName(Me), strFunctionName
    xmlValidDataItem.Text = strAppNo
    
    ' Add the application fact find number
    Set xmlValidDataItem = xmlValidDataElem.appendChild(xmlDomValidData.createElement("APPLICATIONFACTFINDNUMBER"))
    m_objXMLAssist.CheckNode xmlValidDataItem, TypeName(Me), strFunctionName
    xmlValidDataItem.Text = strFactFindNo
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the validation data. This returns customer and member of staff details
    Set xmlValidDataResponseElem = objICostModelAssistDO.GetDataForSubQuoteValidation(xmlDomValidData.documentElement)
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Get Self Employed details
    Dim bSelfEmployed As Boolean
    'MS     14/07/00    SYS0968 IsSelfEmployed  changed to IsMainSelfEmployed
    'bSelfEmployed = objICustEmploymentDO.IsSelfEmployed(xmlValidDataResponseElem.cloneNode(True))
    bSelfEmployed = objICustEmploymentDO.IsMainSelfEmployed(xmlValidDataResponseElem.cloneNode(True))
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Compile a customerVersion list for use with findLoansToBePorted
    Dim xmlEmploymentList As IXMLDOMNodeList
    Set xmlEmploymentList = xmlValidDataResponseElem.getElementsByTagName("EMPLOYMENT")
    m_objXMLAssist.CheckNode xmlEmploymentList, TypeName(Me), strFunctionName
    
    If xmlEmploymentList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "List of customers for application is empty"
    End If
    
    Dim xmlCustomerRootNode As IXMLDOMNode
    Set xmlCustomerRootNode = xmlDOMCustomer.createElement("CUSTOMERLIST")
    xmlDOMCustomer.appendChild xmlCustomerRootNode
    
    Dim xmlEmploymentNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlCustInfoNode As IXMLDOMNode
    Dim bIsMemberOfStaff As Boolean
    bIsMemberOfStaff = False
    
    For Each xmlEmploymentNode In xmlEmploymentList
        Set xmlCustomerNode = xmlDOMCustomer.createElement("CUSTOMER")
        Set xmlCustInfoNode = xmlEmploymentNode.selectSingleNode(".//CUSTOMERNUMBER")
        xmlCustomerNode.appendChild xmlCustInfoNode
        Set xmlCustInfoNode = xmlEmploymentNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER")
        xmlCustomerNode.appendChild xmlCustInfoNode
        xmlCustomerRootNode.appendChild xmlCustomerNode
        ' look at the member of staff setting too
        Set xmlCustInfoNode = xmlEmploymentNode.selectSingleNode(".//MEMBEROFSTAFF")
        If xmlCustInfoNode.Text = "1" Then
            bIsMemberOfStaff = True
        End If
    Next
    
    On Error Resume Next
    Dim bPortedLoans As Boolean
    Dim xmlPortedLoansResp As IXMLDOMNode
    Set xmlPortedLoansResp = objICustFinancialDO.FindLoansToBePorted(xmlCustomerRootNode)
    bPortedLoans = True
    
    m_objErrAssist.SaveErr
    On Error GoTo ValidateVbErr
    
    ' Record not found is ok
    If m_objErrAssist.LastErr <> 0 Then
        If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
            m_objErrAssist.ReRaise
        Else
            bPortedLoans = False
        End If
    End If
    'DC  AQR SYS0977
'    Dim lngError As Long
'    Dim xmlTempNode As IXMLDOMNode
'        Dim xmlResponseNode As IXMLDOMNode
'    Set xmlTempNode = objCostModelAssistBO.GetLTV(vxmlRequest, "COST MODELLING")
'    ' Check the response
'        lngError = m_objErrAssist.CheckResponse(xmlTempNode, xmlResponseNode, True)
'
'        ' Create a temporary document
'        Dim xmlDocCalcLTV As DOMDocument
'        Dim xmlDocElement As IXMLDOMElement
'        Set xmlDocCalcLTV = New DOMDocument
'
'        'copy over the request node no children
'        Set xmlDocElement = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'        Set xmlDocCalcLTV = xmlDocElement.ownerDocument
'
'        ' copy over the LTV response node above including children.
'        m_objXMLAssist.CopyNode "LTV", "LTV", xmlTempNode.firstChild, xmlDocCalcLTV.firstChild, True
'        ' copy over the Customerlist from the Request with children
'        m_objXMLAssist.CopyNode "CUSTOMERLIST", "CUSTOMERLIST", xmlCustomerRootNode, xmlDocCalcLTV.firstChild.firstChild, True
'
'        ' Call the CalcCostModelLTV method on this object
'        Dim strResponse As String
'        Dim xmlLTV As DOMDocument
'        Set xmlLTV = New DOMDocument
'
'        strResponse = CalcCostModelLTV(xmlDocCalcLTV.xml)
'        ' Load the string into xml
'        Set xmlLTV = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'        ' Check the response
'        lngError = m_objErrAssist.CheckResponse(xmlLTV.documentElement, xmlResponseNode, True)
'        'DC End AQR SYS0977
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Check that the mortgage subquote is valid.
    Dim xmlValidRootNode As IXMLDOMNode
    Dim xmlValidNode As IXMLDOMNode
    Set xmlValidRootNode = xmlDOMValid.createElement("QUOTATION")
    xmlDOMValid.appendChild xmlValidRootNode
    Set xmlValidNode = xmlDOMValid.createElement("APPLICATIONNUMBER")
    xmlValidNode.Text = strAppNo
    xmlValidRootNode.appendChild xmlValidNode
    Set xmlValidNode = xmlDOMValid.createElement("APPLICATIONFACTFINDNUMBER")
    xmlValidNode.Text = strFactFindNo
    xmlValidRootNode.appendChild xmlValidNode
    Set xmlValidNode = xmlDOMValid.createElement("MORTGAGESUBQUOTENUMBER")
    xmlValidNode.Text = strMortgageSubQuote
    xmlValidRootNode.appendChild xmlValidNode
    Set xmlValidNode = xmlDOMValid.createElement("ISMEMBEROFSTAFF")
    If bIsMemberOfStaff Then
        xmlValidNode.Text = "1"
    Else
        xmlValidNode.Text = "0"
    End If
    xmlValidRootNode.appendChild xmlValidNode
    Set xmlValidNode = xmlDOMValid.createElement("ISSELFEMPLOYED")
    If bSelfEmployed Then
        xmlValidNode.Text = "1"
    Else
        xmlValidNode.Text = "0"
    End If
    xmlValidRootNode.appendChild xmlValidNode
    If bPortedLoans Then
        xmlValidRootNode.appendChild xmlPortedLoansResp
    Else
        xmlValidRootNode.appendChild xmlDOMValid.createElement("LOANSTOBEPORTED")
    End If
    'DC AQR SYS0977    xmlValidRootNode.appendChild xmlLTV
    
    Set xmlTempResponseNode = objIMortgageSubQuoteBO.ValidateSubQuote(xmlDOMValid.documentElement)
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our response, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXMLAssist.WriteXMLToFile xmlTempResponseNode.xml
    
    '
'    ' DC SYS0977 09/08/01
'        Dim bValidMortgageSubQuote As Boolean
'        Dim xmlDoc As DOMDocument
''  NB  This will be false iff the responsenode returns a validXXsubquote flag is "0"
'        bValidMortgageSubQuote = (m_objXMLAssist.GetNodeValue(xmlResponseNode, ".//VALIDMORTGAGESUBQUOTE", False) <> "0")
'
'        If bValidMortgageSubQuote Then
'            If m_objXMLAssist.GetNodeValue(xmlLTV, ".//LTV", False) <> "" Then
'                Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
'                Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
'                Set xmlDocElement = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'                Set xmlDoc = xmlDocElement.ownerDocument
'                m_objXMLAssist.CopyNode "MORTGAGESUBQUOTE", "MORTGAGESUBQUOTE", vxmlRequest, xmlDocElement, True
'                m_objXMLAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", vxmlRequest, xmlDoc.firstChild.firstChild, False
'                m_objXMLAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", vxmlRequest, xmlDoc.firstChild.firstChild, False
'                m_objXMLAssist.CopyNode "LTV", "LTV", xmlLTV, xmlDoc.firstChild.firstChild, True
'                strResponse = objMortgageSubQuoteBO.Update(xmlDoc.xml)
'                Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
'        ' Check the response
'                lngError = m_objErrAssist.CheckResponse(xmlDoc.firstChild, xmlResponseNode, False)
'        ' if the error wasn't Record Not Found then raise it
'                If lngError <> 0 Then
'                    lngError = m_objErrAssist.GetOmigaErrorNumber(lngError)
'                    If lngError <> oeRecordNotFound Then
'                       m_objErrAssist.ThrowError gstrAPPLICATIONQUOTE, strFunctionName, lngError
'                    End If
'                End If
'
'            End If
'        End If
' 'SYS0977
    Set IApplicationQuoteBO_ValidateMortgageSubQuote = xmlTempResponseNode.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlDomValidData = Nothing
    Set xmlDOMCustomer = Nothing
    Set xmlDOMValid = Nothing
    Set objICustFinancialDO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set objICostModelAssistDO = Nothing
    Set objICustEmploymentDO = Nothing


    Exit Function

ValidateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_ValidateMortgageSubQuote = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    ' go to clean-up section
    Resume ValidateExit

End Function







Private Function IApplicationQuoteBO_ValidateUserMandateLevel(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
    
    'On Error GoTo validateUserMandateLevelExit
    On Error GoTo ValidateUserMandateLevelVbErr 'BMIDS946 KRW 17/11/2004
    
    
    Const strFunctionName As String = "IApplicationQuoteBO_ValidateUserMandateLevel"
    
    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlApplication As IXMLDOMElement
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlMortSubQuoteNo As IXMLDOMElement
    Dim xmlAmtRequested As IXMLDOMElement
    Dim xmlLTV As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim strResponse As String
    Dim strQuotationNo As String
    Dim strValidUser As String
    
    Dim objApplicationBO  As ApplicationBO
    Dim objQuotationBO As QuotationBO
    Dim objMortSubQuoteBO  As MortgageSubQuoteBO
    Dim objOrganisationBO  As OrganisationBO
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
'Stage1:

    If vxmlRequestNode.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequestNode
    Else
        Set xmlRequestNode = xmlGetMandatoryNode(xmlRequestNode, "//REQUEST")
    End If
    
    ' Validate request
    xmlCheckMandatoryAttribute xmlRequestNode, "USERID"
    
    ' Validate application data present
    Set xmlApplication = xmlGetMandatoryNode(xmlRequestNode, "APPLICATION")
    
    xmlCheckMandatoryAttribute xmlApplication, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplication, "APPLICATIONFACTFINDNUMBER"
    
    If m_objContext Is Nothing Then
        Set objOrganisationBO = New OrganisationBO
        Set objApplicationBO = New ApplicationBO
        Set objQuotationBO = New QuotationBO
        Set objMortSubQuoteBO = New MortgageSubQuoteBO
    Else
        Set objOrganisationBO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objMortSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
'Stage2:

    'Get ApplicationfactFind Data:

    ' Create request to get the application fact find data
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "APPLICATIONFACTFIND")
    
    ' Get the application fact find data to retrieve accepted quotation number
    strResponse = objApplicationBO.GetApplicationFactFindData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, xmlResponseElem, True
    
    strQuotationNo = xmlGetNodeText(xmlResponseDoc, ".//ACCEPTEDQUOTENUMBER")
    
    'BMIDS946 KRW 17/11/2004 Force error check to display more meaningfull text
    If (Len(strQuotationNo) = 0) Then
      
        Err.Description = "You must have an accepted quote to be able to approve this application. "
        Err.Number = oeMissingElement
        Err.Source = "ValidateUserMandateLevel"
        
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    ' End of BMIDS946 KRW 17/11/2004
    
    strQuotationNo = xmlGetMandatoryNodeText(xmlResponseDoc, ".//ACCEPTEDQUOTENUMBER")
  
   
'Stage3:
    
    'Get Quotation Data
        
    'Create request to get the accepted quotation data
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "QUOTATION")

    Set xmlNode = xmlGetMandatoryNode(xmlTempRequestDoc, ".//QUOTATION")
    Set xmlElement = xmlTempRequestDoc.createElement("QUOTATIONNUMBER")
    xmlElement.Text = strQuotationNo
    xmlNode.appendChild xmlElement
    
    ' Get the quotation data to retrieve the mortgage sub quote number
    strResponse = objQuotationBO.GetData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, xmlResponseElem, True
   
    Set xmlMortSubQuoteNo = xmlGetMandatoryNode(xmlResponseDoc, ".//MORTGAGESUBQUOTENUMBER")

'Stage4:
    
    'Get MortgageSubQuote Data
    
    ' Create request to get the mortgage sub quote
    Set xmlTempRequestDoc = xmlCreateElementRequestFromNode(xmlRequestNode, "APPLICATION", True, "MORTGAGESUBQUOTE")

    Set xmlNode = xmlGetMandatoryNode(xmlTempRequestDoc, ".//MORTGAGESUBQUOTE")
    xmlNode.appendChild xmlMortSubQuoteNo
    
    ' Get the mortgage sub quote data to retrieve the amount requested and LTV
    strResponse = objMortSubQuoteBO.GetData(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, xmlResponseElem, True
   
    Set xmlAmtRequested = xmlGetMandatoryNode(xmlResponseDoc, ".//AMOUNTREQUESTED")
    Set xmlLTV = xmlGetMandatoryNode(xmlResponseDoc, ".//LTV")
    
'Stage5:

    ' Create request to ValidateUserMandateLevel
    
    Set xmlTempRequest = xmlGetMandatoryNode(xmlTempRequestDoc, ".//REQUEST")
    xmlTempRequest.removeChild xmlTempRequest.firstChild
    xmlTempRequest.appendChild xmlAmtRequested
    xmlTempRequest.appendChild xmlLTV
     
    ' Validate User Mandate Level
    strResponse = objOrganisationBO.ValidateUserMandateLevel(xmlTempRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, xmlResponseElem, True
    
    strValidUser = xmlGetMandatoryNodeText(xmlResponseDoc, ".//VALIDUSER")
   
'Stage6:

    Set xmlElement = xmlResponseDoc.createElement("USER")
    xmlElement.setAttribute "VALIDUSER", strValidUser
    xmlResponseElem.appendChild xmlElement
    
    Set IApplicationQuoteBO_ValidateUserMandateLevel = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
validateUserMandateLevelExit:

    '   clean-up section, release object references
    Set xmlRequestNode = Nothing
    Set xmlApplication = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlMortSubQuoteNo = Nothing
    Set xmlAmtRequested = Nothing
    Set xmlLTV = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDoc = Nothing
    Set objOrganisationBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Set objApplicationBO = Nothing
    Set objQuotationBO = Nothing
    Set objMortSubQuoteBO = Nothing
    Set objOrganisationBO = Nothing
    
    Exit Function

ValidateUserMandateLevelVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_ValidateUserMandateLevel = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    ' go to clean-up section
    Resume validateUserMandateLevelExit
    
End Function

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set m_objErrAssist = New ErrAssist
    Set m_objXMLAssist = New XMLAssist
    Set m_objContext = GetObjectContext()
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objErrAssist = Nothing
    Set m_objXMLAssist = Nothing
    Set m_objContext = Nothing
End Sub

Public Function CalculateMaximumBorrowing(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------

On Error GoTo CalculateMaximumBorrowingVbErr

    Const strFunctionName As String = "CalculateMaximumBorrowing"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    'Dim xmlElement As IXMLDOMElement
    
    Dim objApplicationBO As ApplicationBO
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim xmlTemp As FreeThreadedDOMDocument40
    
    Set xmlTemp = New FreeThreadedDOMDocument40
    xmlTemp.validateOnParse = False
    xmlTemp.setProperty "NewParser", True
    Set xmlResponseElem = xmlTemp.createElement("MAXBORROWINGLIST")
    Set xmlDataNode = xmlTemp.appendChild(xmlResponseElem)
         
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO") '"omAQ"
    End If
                   
    Set xmlRequestNode = xmlIn.getElementsByTagName("CUSTOMERLIST").Item(0)
    
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERLIST tag not found"
    End If
                                      
    Dim strHighestEarners As String
    
    Dim objReturnNode As IXMLDOMNode
    
    strHighestEarners = objApplicationBO.GetHighestEarners(vstrXMLRequest)
                        
    Set objReturnNode = objICostModelAssistBO.GetBorrowingAndLenderName(strHighestEarners)
                    
    CalculateMaximumBorrowing = objReturnNode.xml
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateMaximumBorrowingExit:
    
    Set objApplicationBO = Nothing
    Set objICostModelAssistBO = Nothing
    Set xmlTemp = Nothing
    Set xmlOut = Nothing
    Set xmlIn = Nothing
    
    Exit Function

CalculateMaximumBorrowingVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
        
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    CalculateMaximumBorrowing = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalculateMaximumBorrowingExit

End Function

Public Function CalculateApproximateMonthlyCost(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'
' pass:         vstrXMLRequest  xml Request data stream
'
'   <REQUEST>
'       <APPLICATION>
'           <APPLICATIONNUMBER> </APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER> </APPLICATIONFACTFINDNUMBER>
'       </APPLICATION>
'   </REQUEST>
'
' return:       xml Response data stream
'               either: TYPE="SUCCESS"
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CalculateApproximateMonthlyCost"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DOMDocument based method
    
    Set xmlTempResponseNode = IApplicationQuoteBO_CalculateApproximateMonthlyCost(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    CalculateApproximateMonthlyCost = xmlResponseElem.Text

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CalculateApproximateMonthlyCost = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit

End Function

Public Function CalculateAffordability(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' NOTE: SOME more functionality is to be added to this method. Checked for backup sake
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'Input Format
'  <REQUEST>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>...</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>...</CUSTOMERVERSIONNUMBER>
'               <CUSTOMERROLETYPE>...</CUSTOMERROLETYPE>
'           </CUSTOMER>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>...</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>...</CUSTOMERVERSIONNUMBER>
'               <CUSTOMERROLETYPE>...</CUSTOMERROLETYPE>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'       <TYPEOFAPPLICATION>..</TYPEOFAPPLICATION>
'       <APPLICATIONNUMBER>...<APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>...<APPLICATIONFACTFINDNUMBER>
'       <MORTGAGESUBQUOTENUMBER>...</MORTGAGESUBQUOTENUMBER>
'       <LIFESUBQUOTENUMBER>...</LIFESUBQUOTENUMBER>
'       <QUOTATIONNUMBER>...</QUOTATIONNUMBER>
'  </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAffordabilityVbErr

    Const strFunctionName As String = "CalculateAffordability"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_CalculateAffordability(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalculateAffordability = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateAffordabilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CalculateAffordability = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalculateAffordabilityExit

End Function

Private Function IApplicationQuoteBO_CalculateAffordability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Calculates the Affordability and results are stored in the Quotation table.
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAffordabilityErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_CalculateAffordability"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIAppQuoteTxBO As IApplicationQuoteTxBO
    
    'Dim sQuotationType As String, strResponse As String
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIAppQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIAppQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIAppQuoteTxBO.CalculateAffordability(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_CalculateAffordability = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateAffordabilityExit:
    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIAppQuoteTxBO = Nothing

    Exit Function

CalculateAffordabilityErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CalculateAffordability = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalculateAffordabilityExit

End Function

Public Function CheckApplicantPPEligibility(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks if applicants are eligible for payment protection
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       xml Response data stream containing results of operation
'               either: TYPE="SUCCESS"
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CheckApplicantPPEligibilityVbErr
    
    Const strFunctionName As String = "CheckApplicantPPEligibility"

    Dim xmlMainApplicants As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlIn As FreeThreadedDOMDocument40
    
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    Dim objIApplicationDO As IApplicationDO
    
'    MS SYS0824 UK residency NOT a criterion for PP Eligibility
    
    'SG 28/05/02 SYS4767
    'Dim strApplicant1Uk As String, strApplicant2Uk As String
    
    Dim strApplicant1EmpStatus As String, strApplicant2EmpStatus As String
    Dim strApplicant1EmpProblem As String, strApplicant2EmpProblem As String
    Dim strCustomer1Number As String, strCustomer2Number As String
    Dim strCustomer1VersionNumber As String, strCustomer2VersionNumber As String
    Dim strCustomer1EmpSeqNo As String, strCustomer2EmpSeqNo As String
    
    Dim strXML As String
    
    Set objIApplicationDO = New ApplicationDO

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlMainApplicants = objIApplicationDO.FindMainApplicants(xmlIn.documentElement)
            
    '
    ' Trap Record not found errors
    '
    On Error Resume Next
    
    'Fetch Customer Details into local variables
    Call GetCustomerDetailsFromXml(xmlMainApplicants.ownerDocument, _
                             strCustomer1Number, strCustomer1VersionNumber, _
                             strCustomer2Number, strCustomer2VersionNumber)
            
    '
    ' Set the Emplyment Status and Employment Problem for Applicant 1
    '
    Call SetEmpDetailsOfApplicant(strCustomer1Number, strCustomer1VersionNumber, _
                                 strApplicant1EmpStatus, strApplicant1EmpProblem, _
                                 strCustomer1EmpSeqNo)
                                 
    '
    ' Record not found is ok
    '
    If m_objErrAssist.LastErr <> 0 Then
        If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
            On Error GoTo CheckApplicantPPEligibilityVbErr
            m_objErrAssist.ReRaise
        End If
    End If
    
    ' Set the Emplyment Status and Employment Problem for Applicant 2 (if existing)
    If Len(strCustomer2Number) > 0 And Len(strCustomer2VersionNumber) > 0 Then
       Call SetEmpDetailsOfApplicant(strCustomer2Number, strCustomer2VersionNumber, _
                                 strApplicant2EmpStatus, strApplicant2EmpProblem, _
                                 strCustomer2EmpSeqNo)
                                 
        '
        ' Record not found is ok
        '
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
                On Error GoTo CheckApplicantPPEligibilityVbErr
                m_objErrAssist.ReRaise
            End If
        End If
    
    End If
    
    '
    '
    '
    If (strApplicant1EmpStatus = "") Then
        strApplicant1EmpStatus = "N"
    End If

    '
    '
    '
    If (strApplicant2EmpStatus = "") Then
        strApplicant2EmpStatus = "N"
    End If

    '
    '    MS SYS0824 UK residency NOT a criterion for PP Eligibility
    '
    strXML = BuildXmlInForCheckEligibility("", strApplicant1EmpStatus, _
                                    strApplicant1EmpProblem, "", _
                                    strApplicant2EmpStatus, strApplicant2EmpProblem)
         
    
    If m_objContext Is Nothing Then
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
    Else
        Set objPPSubQuoteBO = m_objContext.CreateInstance(cstrPPSUBQUOTE_COMP & ".PaymentProtectionSubQuoteBO")
    End If
    
    strXML = objPPSubQuoteBO.CheckApplicantEligibility(strXML)
    
    Call m_objErrAssist.CheckXMLResponse(strXML, True)
    
    Set xmlDoc = m_objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
        
    Call ModifyXmlOutForCheckPPEligibility(xmlOut, xmlDoc, _
                                    strCustomer1Number, strCustomer1VersionNumber, _
                                    strCustomer2Number, strCustomer2VersionNumber)
                                    
    CheckApplicantPPEligibility = xmlOut.xml

    If Not m_objContext Is Nothing Then
       m_objContext.SetComplete
    End If
        
CheckApplicantPPEligibilityExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlIn = Nothing
    Set objPPSubQuoteBO = Nothing
    Set objIApplicationDO = Nothing

    Exit Function

CheckApplicantPPEligibilityVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CheckApplicantPPEligibility = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CheckApplicantPPEligibilityExit

End Function

Public Function CalculateRecommendedPPCover(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the recommended payment protection cover
'
' pass:         vstrXMLRequest  of the format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <APPLICATIONNUMBER> ... </APPLICATIONNUMBER>
'      <APPLICATIONFACTFINDNUMBER> ... </APPLICATIONFACTFINDNUMBER>
'       <MORTGAGESUBQUOTENUMBER> ... </MORTGAGESUBQUOTENUMBER>
'       <TYPEOFAPPLICATION> ... </TYPEOFAPPLICATION>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>....</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>..</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'           <CUSTOMER>
'               <CUSTOMERNUMBER> ... </CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>...</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </APPLICATIONQUOTE>
'</REQUEST>
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'               RecommendedCover : RecommendedLevel of payment protection cover
'------------------------------------------------------------------------------------------
On Error GoTo CalculateRecommendedPPCoverVbErr
    
    Const strFunctionName As String = "CalculateRecommendedPPCover"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    
    ' Variables to store the data from the input (vstrXmlRequest)
    Dim strApplicationNumber As String, strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String, strTypeofApplication As String
    Dim strResponse As String
    
    Dim dblNewRepayments As Double, dblExistingRepayments As Double
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    Set xmlDoc = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Get all the values from InXml and store them in local variables
    Call GetValFromInXmlForCalcRecomPPCover(xmlDoc, strApplicationNumber, _
                                             strApplicationFactFindNumber, _
                                            strMortgageSubQuoteNumber, _
                                            strTypeofApplication)
    
    ' Get existing loan details from Active Mortgage Sub-Quote
    dblNewRepayments = GetToalNetMonthlyCostFromMortSubQuote(strApplicationNumber, _
                                          strApplicationFactFindNumber, _
                                          strMortgageSubQuoteNumber)
                                          
    dblExistingRepayments = GetExitingRepayments_CalcRecomPPCover( _
                                xmlDoc.documentElement, strTypeofApplication)
       
    ' Calculate the Recommended Cover
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(xmlDoc.documentElement)
    Set xmlTableElem = xmlDoc.createElement("PAYMENTPROTECTIONSUBQUOTE")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = xmlDoc.createElement("NEWREPAYMENTS")
    xmlElement.Text = dblNewRepayments
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("EXISTINGREPAYMENTS")
    xmlElement.Text = dblExistingRepayments
    xmlTableElem.appendChild xmlElement
    
    If m_objContext Is Nothing Then
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
    Else
        Set objPPSubQuoteBO = m_objContext.CreateInstance(cstrPPSUBQUOTE_COMP & ".PaymentProtectionSubQuoteBO")
    End If
    
    strResponse = objPPSubQuoteBO.CalculateRecommendedCover(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
    'Append the Recommended Cover to Response
    xmlResponseElem.appendChild xmlDoc.getElementsByTagName("RECOMMENDEDCOVER").Item(0)
    
    CalculateRecommendedPPCover = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateRecommendedPPCoverExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objPPSubQuoteBO = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    
    Exit Function

CalculateRecommendedPPCoverVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateRecommendedPPCover = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalculateRecommendedPPCoverExit
    
End Function

Public Function CalculateAndSavePPCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks if applicants are eligible for payment protection. Also, it calculates the
'               monthly cost of cover and stores it to the database
'
' pass:         vstrXMLRequest  containing values of
'               ApplicationNumber, ApplicationFactFindNumber, PPSubQuoteNumber, Customer1Number,
'               Customer2Number, Customer1VersionNumber, Customer2VersionNumber, SplitCOver,
'               DistributedChannel, CoverType, MortgageCoverAmount, MortgageCoverForApplicant1
'               MortgageCoverForApplicant2, RecommendedMortgageCoverAmount,
'               SelectedPPProductXML
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAndSavePPCostsVbErr
    
    Const strFunctionName As String = "CalculateAndSavePPCosts"

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlDataElem As IXMLDOMElement
    Dim objPaymentProtectionBO As PaymentProtectionBO
' MS    08/07/00 SYS0824 UK residency NOT a criterion for PP Eligibility
' Change cover type to validation value for the calc
    Dim objComboDO As ComboDO
    
    If m_objContext Is Nothing Then
        Set objPaymentProtectionBO = New PaymentProtectionBO
        Set objComboDO = New ComboDO
    Else
        Set objPaymentProtectionBO = m_objContext.CreateInstance(gstrPP_COMPONENT & ".PaymentProtectionBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlDoc = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Local variables to store the values passed as XML string
    Dim strCustomer1Number As String, strCustomer2Number As String
    Dim strCustomer1VersionNumber As String, strCustomer2VersionNumber As String
    Dim strPPSubQuoteNumber As String, strDistributionChannel As String
    Dim strSplitCover As String, strCoverType As String
    Dim strMortgageCoverAmount As String, strMortgageCoverForApplicant1 As String
    Dim strMortgageCoverForApplicant2 As String, strRecomMortgageCoverAmount As String
    Dim strApplicationNumber As String, strApplicationFactFindNumber As String
    Dim strProductCode As String
    
    Dim strApplicant1Age As String, strApplicant2Age As String
    Dim strApplicant1Gender As String, strApplicant2Gender As String
    'Dim strCustomerVersionNumber As String, strApplicantAge As String
    'Dim strApplicantGender As String
    Dim strApplicant1Premium As String
    Dim strApplicant1IPT As String, strApplicant2Premium As String
    Dim strApplicant2IPT As String
    
    ' Get all the values from InXml and store them in local variables
    Call CalcAndSavePPCosts_GetValFromInXml _
                            (xmlDoc, _
                             strApplicationNumber, strApplicationFactFindNumber, _
                             strCustomer1Number, strCustomer2Number, _
                             strCustomer1VersionNumber, strCustomer2VersionNumber, _
                             strPPSubQuoteNumber, strDistributionChannel, _
                             strSplitCover, strCoverType, strMortgageCoverAmount, _
                             strMortgageCoverForApplicant1, strMortgageCoverForApplicant2, _
                             strRecomMortgageCoverAmount, strProductCode)
        
    ' Retreive Customer Details - Set Age and Gender for the Applicants
    If Len(strCustomer1Number) > 0 Then
        Call CalcAndSavePPCosts_GetCustDetails(strCustomer1Number, _
                                        strCustomer1VersionNumber, _
                                        strApplicant1Age, strApplicant1Gender)
    End If

    If Len(strCustomer2Number) > 0 Then
        Call CalcAndSavePPCosts_GetCustDetails(strCustomer2Number, _
                                        strCustomer2VersionNumber, _
                                        strApplicant2Age, strApplicant2Gender)
    End If
    
    ' Validate Sub-Quote Details and Calculate Costs
    
    ' Build the Input Xml for PaymentProtectionBO.CalculateCosts
    Set xmlElement = xmlDoc.getElementsByTagName("PAYMENTPROTECTION").Item(0)
    Set xmlDataElem = xmlDoc.createElement("APPLICANT1AGE")
    xmlElement.appendChild xmlDataElem
    xmlDataElem.Text = strApplicant1Age
    Set xmlDataElem = xmlDoc.createElement("APPLICANT1GENDER")
    xmlDataElem.Text = strApplicant1Gender
    xmlElement.appendChild xmlDataElem
    Set xmlDataElem = xmlDoc.createElement("APPLICANT2AGE")
    xmlDataElem.Text = strApplicant2Age
    xmlElement.appendChild xmlDataElem
    Set xmlDataElem = xmlDoc.createElement("APPLICANT2GENDER")
    xmlDataElem.Text = strApplicant2Gender
    xmlElement.appendChild xmlDataElem
    
' MS    08/07/00 SYS0824 UK residency NOT a criterion for PP Eligibility

    
    ' Change cover type to validation value for the calc
    Dim strCoverTypeValue As String
    Dim xmlCoverTypeElem As IXMLDOMElement
    'strTagName = "COVERTYPE"
    Set xmlCoverTypeElem = xmlDoc.getElementsByTagName("COVERTYPE").Item(0)
    
    If Not xmlCoverTypeElem Is Nothing Then
        strCoverTypeValue = xmlCoverTypeElem.Text
        xmlCoverTypeElem.Text = objComboDO.GetFirstComboValidation("PPCoverType", _
                                                                   strCoverTypeValue)
    End If
    
    
    
    Dim strXML As String



    strXML = objPaymentProtectionBO.CalculateCosts(xmlDoc.xml)
    Call m_objErrAssist.CheckXMLResponse(strXML, True)
    
    Set xmlDoc = m_objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Call CalcAndSavePPCosts_GetCostsFromXml(xmlDoc, _
                                 strApplicant1Premium, strApplicant1IPT, _
                                 strApplicant2Premium, strApplicant2IPT, _
                                 strMortgageCoverAmount, strMortgageCoverForApplicant1, _
                                 strMortgageCoverForApplicant2)

'   Save PPSubQuote
   Call CalcAndSavePPCosts_SavePPSubQuote(xmlOut, _
                                          strApplicationNumber, _
                                          strApplicationFactFindNumber, _
                                          strPPSubQuoteNumber, _
                                          strApplicant1Premium, _
                                          strApplicant1IPT, _
                                          strApplicant2Premium, _
                                          strApplicant2IPT, _
                                          strProductCode, _
                                          strSplitCover, _
                                          strMortgageCoverAmount, _
                                          strMortgageCoverForApplicant1, _
                                          strMortgageCoverForApplicant2, _
                                          strCoverType)
                                              
    CalculateAndSavePPCosts = xmlOut.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateAndSavePPCostsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlDoc = Nothing
    
    Exit Function

CalculateAndSavePPCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateAndSavePPCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateAndSavePPCostsExit

End Function

Public Function CalcAllowableInc(ByVal vstrXMLRequest As String) As String

On Error GoTo CalcAllowableIncVbErr

    Const strFunctionName As String = "CalcAllowableInc"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IApplicationQuoteBO_CalcAllowableInc(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalcAllowableInc = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalcAllowableIncExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalcAllowableIncVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalcAllowableInc = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalcAllowableIncExit

End Function

 Private Sub CalcAndSavePPCosts_GetCostsFromXml(ByVal xmlDoc As FreeThreadedDOMDocument40, _
                                                ByRef strApplicant1Premium As String, _
                                                ByRef strApplicant1IPT As String, _
                                                ByRef strApplicant2Premium As String, _
                                                ByRef strApplicant2IPT As String, _
                                                ByRef strMortgageCoverAmt As String, _
                                                ByRef strMortgageCoverForApplicant1 As String, _
                                                ByRef strMortgageCoverForApplicant2 As String)
                                                
' header ----------------------------------------------------------------------------------
' Description:  Fetches the required values from the Output of CalculateCosts and
'               saves them in local variables
'
' pass:         xmlDoc : Document contating the output of CalculateCosts
'               Other variables in which the values are to be stored
'------------------------------------------------------------------------------------------
  On Error GoTo CalcAndSavePPCosts_GetCostsFromXmlVbErr

    Const strFunctionName As String = "CalcAndSavePPCosts_GetCostsFromXml"
    
    strApplicant1Premium = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "APPLICANT1PREMIUM")
    strApplicant1IPT = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "APPLICANT1IPT")
    
    strApplicant2Premium = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "APPLICANT2PREMIUM")
    strApplicant2IPT = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "APPLICANT2IPT")
    ' DM    13/07/00 SYS0993 Wrong XML tags being used. Changed MORTGAGECOVERAMOUNT to MORTGAGEAMOUNT
    '                and COVERTYPEID to COVERTYPE and also PRODUCTCODE to PPPRODUCTCODE.
    strMortgageCoverAmt = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "MORTGAGEAMOUNT")
    strMortgageCoverForApplicant1 = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "MORTGAGECOVERFORAPPLICANT1")
    strMortgageCoverForApplicant2 = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "MORTGAGECOVERFORAPPLICANT2")
                            
    Exit Sub
CalcAndSavePPCosts_GetCostsFromXmlVbErr:

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
                            
End Sub

'EP1038/MAR1891 Not used
'Private Sub SetResidencyStatusOfApplicants(ByVal strCustomer1Number As String, _
'                                           ByVal strCustomer1VersionNumber As String, _
'                                           ByVal strCustomer2Number As String, _
'                                           ByVal strCustomer2VersionNumber As String, _
'                                           ByRef strApplicant1Uk As String, _
'                                           ByRef strApplicant2Uk As String)
'' header ----------------------------------------------------------------------------------
'' Description:  Sets the Residency Status for the Applicants
''
'' pass:         Customer Number, Customer Version Number of both the applicants
''               strApplicant1Uk      : Residency Status of Applicant1
''               strApplicant2Uk      : Residency Status of Applicant2
''------------------------------------------------------------------------------------------
'
'On Error GoTo SetResidencyStatusOfApplicantsVbErr
'
'    Const strFunctionName As String = "SetResidencyStatusOfApplicants"
'
'    Dim objADOAssist As ADOAssist
'    Dim objSQLAssist As SQLAssist
'    Dim strCondition  As String
'
'    If m_objContext Is Nothing Then
'        Set objADOAssist = New ADOAssist
'    Else
'        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
'    End If
'
'    Set objSQLAssist = New SQLAssist
'
'    'Get value of 'UKResidentIndicator' for applicant1
'    strCondition = "CUSTOMERNUMBER=" & objSQLAssist.FormatString(strCustomer1Number) _
'                   & " And CUSTOMERVERSIONNUMBER = " & strCustomer1VersionNumber
'    Call objADOAssist.GetValueFromTable("CUSTOMERVERSION", strCondition, _
'                                               "NORMALLYRESIDENTINUK", strApplicant1Uk)
'
'    'Get value of 'UKResidentIndicator' for applicant2 (if existing )
'    If Len(strCustomer2Number) > 0 And Len(strCustomer2VersionNumber) > 0 Then
'        strCondition = "CUSTOMERNUMBER=" & objSQLAssist.FormatString(strCustomer2Number) _
'                   & " And CUSTOMERVERSIONNUMBER = " & strCustomer2VersionNumber
'        Call objADOAssist.GetValueFromTable("CUSTOMERVERSION", strCondition, _
'                                               "NORMALLYRESIDENTINUK", strApplicant2Uk)
'    End If
'
'    Set objADOAssist = Nothing
'    Set objSQLAssist = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'    Exit Sub
'
'SetResidencyStatusOfApplicantsVbErr:
'
'    Set objADOAssist = Nothing
'    Set objSQLAssist = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
'
'    're-raise error for business object to interpret as appropriate
'    Err.Raise Err.Number, Err.Source, Err.Description
'
'End Sub

Public Function ValidateMortgageSubQuote(ByVal vstrXMLRequest As String) As String
' description:  validate
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>
'   <BASICQUOTATIONDETAILS>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'       <MORTGAGESUBQUOTENUMBER>7</MORTGAGESUBQUOTENUMBER>
'   </BASICQUOTATIONDETAILS>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS" <VALIDMORTGAGESUBQUOTE>1</VALIDMORTGAGESUBQUOTE>
'                               or: TYPE="SYSERR" and <ERROR> element
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ValidateMortgageSubQuote"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_ValidateMortgageSubQuote(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ValidateMortgageSubQuote = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidateMortgageSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit
End Function

Private Sub GetCustomerDetailsFromXml(ByVal xmlMainApplicantsDoc As FreeThreadedDOMDocument40, _
                                      ByRef strCustomer1Number As String, _
                                      ByRef strCustomer1VersionNumber As String, _
                                      ByRef strCustomer2Number As String, _
                                      ByRef strCustomer2VersionNumber As String)
' header ----------------------------------------------------------------------------------
' Description:  Fetches the customer details (customer number, customer version number) into
'               into the respective variables passed for both applicants
'
' pass:         xmlMainApplicantsDoc : xml Document containing the data of the applicants
'               strCustomer1Number       : Customer number of applicant 1
'               strCustomer1VersionNumber: Customer version number of applicant 1
'               strCustomer2Number       : Customer number of applicant 2
'               strCustomer2VersionNumber: Customer version number of applicant 2
'------------------------------------------------------------------------------------------

On Error GoTo GetCustomerDetailsFromXmlVbErr

    Const strFunctionName As String = "GetCustomerDetailsFromXml"

    Dim xmlNodeList As IXMLDOMNodeList

    Set xmlNodeList = xmlMainApplicantsDoc.getElementsByTagName("CUSTOMERNUMBER")
    
    strCustomer1Number = xmlNodeList.Item(0).firstChild.Text
    If xmlNodeList.length > 1 Then 'If one main applicant only
        strCustomer2Number = xmlNodeList.Item(1).firstChild.Text
    End If
    
    Set xmlNodeList = xmlMainApplicantsDoc.getElementsByTagName("CUSTOMERVERSIONNUMBER")
    strCustomer1VersionNumber = xmlNodeList.Item(0).firstChild.Text
    If xmlNodeList.length > 1 Then ' If one main applicant only
        strCustomer2VersionNumber = xmlNodeList.Item(1).firstChild.Text
    End If
    
    Set xmlNodeList = Nothing
    
    Exit Sub
GetCustomerDetailsFromXmlVbErr:
    
    Set xmlNodeList = Nothing

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub SetEmpDetailsOfApplicant(ByVal strCustomerNumber As String, _
                                     ByVal strCustomerVersionNumber As String, _
                                     ByRef strApplicantEmpStatus As String, _
                                     ByRef strApplicantEmpProblem As String, _
                                     ByRef strCustomerEmpSeqNo As String)

' header ----------------------------------------------------------------------------------
' Description:  Finds the Employment status, Employment Sequence Number and  respective
'               NoticeProblemIndicator for applicant identified
' pass:         Customer Number, Customer Version Number of both the applicants
'               strApplicantEmpStatus : Employmet Status of Applicant
'               strApplicantEmpProblem: NoticeProblemIndicator of Applicant
'               strApplicantEmpSeqNo  : Employment Sequence Number
'Return :       N/A
'Error      :
'------------------------------------------------------------------------------------------

On Error GoTo SetEmpDetailsOfApplicantsVbErr

    Const strFunctionName As String = "SetEmpDetailsOfApplicant"

    Dim objADOAssist As ADOAssist
    Dim objComboDO As ComboDO
    Dim objSQLAssist As SQLAssist
    
    Dim strCondition  As String
    Dim blnRecordFound As Boolean
    Dim blnIsEmployed As Boolean
    
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objComboDO = New ComboBO
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    Set objSQLAssist = New SQLAssist

    'Employment status of Applicant1
    strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                   & " AND CUSTOMERVERSIONNUMBER =" & strCustomerVersionNumber _
                   & " AND MAINSTATUS = '1'"
                   
    Call objADOAssist.GetValueFromTable("EMPLOYMENT", strCondition, "EMPLOYMENTSTATUS", _
                                         strApplicantEmpStatus, blnRecordFound)
    
    'If a record was not found, raise error
    If Not blnRecordFound Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, _
                                " In table EMPLOYMENT for Customer : " & strCustomerNumber _
                                & " And Customer Version : " & strCustomerVersionNumber
    End If

    blnIsEmployed = objComboDO.IsItemInValidation("EmploymentStatus", _
                                                strApplicantEmpStatus, "E")
    
    If blnIsEmployed Then
        
        'Get the Employment Sequence Number
        Call objADOAssist.GetValueFromTable("EMPLOYMENT", strCondition, _
                            "EMPLOYMENTSEQUENCENUMBER", strCustomerEmpSeqNo, blnRecordFound)
        
        ' Get the value of NoticeProblemIndicator and assign it to ApplicantEmpProblem
        strCondition = "CUSTOMERNUMBER = " & objSQLAssist.FormatString(strCustomerNumber) _
                   & " AND CUSTOMERVERSIONNUMBER =" & strCustomerVersionNumber _
                   & " AND EMPLOYMENTSEQUENCENUMBER = " & strCustomerEmpSeqNo
    
        Call objADOAssist.GetValueFromTable("EMPLOYEDDETAILS", strCondition, _
                                            "NOTICEPROBLEMINDICATOR", _
                                             strApplicantEmpProblem, blnRecordFound)
        
        'If a record was not found, raise error
        If Not blnRecordFound Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, _
                          " In table EMPLOYEDDETAILS for Customer : " & strCustomerNumber _
                          & " And Customer Version : " & strCustomerVersionNumber _
                          & " And EmploymentSequenceNumber : " & strCustomerEmpSeqNo
        End If
    End If
    
    '  Clear all the objects
    Set objADOAssist = Nothing
    Set objComboDO = Nothing
    Set objSQLAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub
SetEmpDetailsOfApplicantsVbErr:
    
    Set objADOAssist = Nothing
    Set objComboDO = Nothing
    Set objSQLAssist = Nothing
    
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
    
    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function BuildXmlInForCheckEligibility(ByVal strApplicant1Uk As String, _
                                               ByVal strApplicant1EmpStatus As String, _
                                               ByVal strApplicant1EmpProblem As String, _
                                               ByVal strApplicant2Uk As String, _
                                               ByVal strApplicant2EmpStatus As String, _
                                               ByVal strApplicant2EmpProblem As String) As String
' header ----------------------------------------------------------------------------------
' Description:  Builds the XML string to be passed to PPSubQuoteBO for checking status;
'               this is called from CheckApplicantPPEligibility
' pass:         Customer Number, Customer Version Number, Employmet Status and
'               and Employment Problem of both applicants
'Return :       XML string
'Error      :
'------------------------------------------------------------------------------------------
'    MS SYS0824 UK residency NOT a criterion for PP Eligibility
    Dim strXML As String
   
    strXML = "<PAYMENTPROTECTIONSUBQUOTE>"
'    strXML = strXML & "<APPLICANT1UK>" & strApplicant1Uk & "</APPLICANT1UK>"
    strXML = strXML & "<APPLICANT1EMPLOYMENTSTATUS>" & strApplicant1EmpStatus _
                    & "</APPLICANT1EMPLOYMENTSTATUS>"
    strXML = strXML & "<APPLICANT1EMPPROBLEM>" & strApplicant1EmpProblem _
                    & "</APPLICANT1EMPPROBLEM>"
    
'    strXML = strXML & "<APPLICANT2UK>" & strApplicant2Uk & "</APPLICANT2UK>"
    strXML = strXML & "<APPLICANT2EMPLOYMENTSTATUS>" & strApplicant2EmpStatus _
                    & "</APPLICANT2EMPLOYMENTSTATUS>"
    strXML = strXML & "<APPLICANT2EMPPROBLEM>" & strApplicant2EmpProblem _
                    & "</APPLICANT2EMPPROBLEM>"
    
    strXML = strXML & "</PAYMENTPROTECTIONSUBQUOTE>"
    
    BuildXmlInForCheckEligibility = strXML
End Function

Private Sub ModifyXmlOutForCheckPPEligibility(ByVal xmlOut As FreeThreadedDOMDocument40, _
                                              ByVal xmlDoc As FreeThreadedDOMDocument40, _
                                              ByVal strCustomer1Number As String, _
                                              ByVal strCustomer1VersionNumber As String, _
                                              ByVal strCustomer2Number As String, _
                                              ByVal strCustomer2VersionNumber As String)
' header ----------------------------------------------------------------------------------
' Description:  Modifies the XMLOut after finding the eligibility for each applicant;
'               Used in the mehtos CheckApplicantPPEligibility
' pass:         xmlOut - Document to be modified
'               xmlDoc - Document containing the response from other BO,  and to be added
'               xmlOut
'               Customer Number, Customer Version Number, of both applicants
'Return :       XML string
'Error      :
'------------------------------------------------------------------------------------------

On Error GoTo ModifyXmlOutForCheckPPEligibilityVbErr

    Const strFunctionName As String = "ModifyXmlOutForCheckPPEligibility"
    
    Dim strApplicantEligible As String
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDocResponseElem As IXMLDOMElement
    
    'Dim xmlElement As IXMLDOMElement
    Dim xmlCustomerElem As IXMLDOMElement
    Dim xmlCustomerVersionElem As IXMLDOMElement
    
    Dim intNodeIndex  As Integer
    
    Set xmlResponseElem = xmlOut.getElementsByTagName("RESPONSE").Item(0)
    Set xmlDocResponseElem = xmlDoc.getElementsByTagName("RESPONSE").Item(0)
    
    'Add children of Response Node from xmlDoc to xmlOut
    For intNodeIndex = 0 To xmlDocResponseElem.childNodes.length - 1
        xmlResponseElem.appendChild xmlDocResponseElem.childNodes(0)
    Next intNodeIndex
    
    ' Add nodes for customer1
    strApplicantEligible = m_objXMLAssist.GetTagValue(xmlOut.documentElement, _
                                                    "APPLICANT1ELIGIBLE")
    Set xmlCustomerElem = xmlOut.createElement("CUSTOMERNUMBER1")
    Set xmlCustomerVersionElem = xmlOut.createElement("CUSTOMERVERSIONNUMBER1")
    
    xmlResponseElem.appendChild xmlCustomerElem
    xmlResponseElem.appendChild xmlCustomerVersionElem
    
    'If applicant 1 is eligible assin values to the aove two elements
    If strApplicantEligible <> "0" Then
        xmlCustomerElem.Text = strCustomer1Number
        xmlCustomerVersionElem.Text = strCustomer1VersionNumber
    End If
    
    ' Add nodes for customer2
    strApplicantEligible = m_objXMLAssist.GetTagValue(xmlOut.documentElement, _
                                                    "APPLICANT2ELIGIBLE")
    Set xmlCustomerElem = xmlOut.createElement("CUSTOMERNUMBER2")
    Set xmlCustomerVersionElem = xmlOut.createElement("CUSTOMERVERSIONNUMBER2")
    
    xmlResponseElem.appendChild xmlCustomerElem
    xmlResponseElem.appendChild xmlCustomerVersionElem
    
    'If applicant 2 is eligible assin values to the aove two elements
    If strApplicantEligible <> "0" Then
        xmlCustomerElem.Text = strCustomer2Number
        xmlCustomerVersionElem.Text = strCustomer2VersionNumber
    End If
    
    Set xmlCustomerElem = Nothing
    Set xmlCustomerVersionElem = Nothing
    
    Exit Sub
ModifyXmlOutForCheckPPEligibilityVbErr:
    
    Set xmlCustomerElem = Nothing
    Set xmlCustomerVersionElem = Nothing
    Set xmlDocResponseElem = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
       
    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub CalcAndSavePPCosts_GetValFromInXml(ByVal xmlDoc As FreeThreadedDOMDocument40, _
                                               ByRef strApplicationNumber As String, _
                                               ByRef strApplicationFactFindNumber As String, _
                                               ByRef strCustomer1Number As String, _
                                               ByRef strCustomer2Number As String, _
                                               ByRef strCustomer1VersionNumber As String, _
                                               ByRef strCustomer2VersionNumber As String, _
                                               ByRef strPPSubQuoteNumber As String, _
                                               ByRef strDistributionChannel As String, _
                                               ByRef strSplitCover As String, _
                                               ByRef strCoverType As String, _
                                               ByRef strMortgageCoverAmount As String, _
                                               ByRef strMortgageCoverForApplicant1 As String, _
                                               ByRef strMortgageCoverForApplicant2 As String, _
                                               ByRef strRecomMortgageCoverAmount As String, _
                                               ByRef strProductCode As String)
' header ----------------------------------------------------------------------------------
' Description:  Fetches all the values from xmlIn and assign them to respective local variables ;
'               this is used in the method 'CalculateAndSavePPCosts'
'
' pass:         xmlIn : xml Document containing the Input data
'               ApplicationNumber, ApplicationFactFindNumber
'               strCustomer1Number,  strCustomer2Number
'               strCustomer1VersionNumber, strCustomer2VersionNumber
'               strPPSubQuoteNumber, strDistributionChannel, strSplitCover
'               strCoverType, strMortgageCoverAmount,
'               strMortgageCoverForApplicant1, strMortgageCoverForApplicant2
'               strRecomMortgageCoverAmount
' return:       N/A
' Error:       nil
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSavePPCosts_GetValFromInXmlVbErr

    Const strFunctionName As String = "CalcAndSavePPCosts_GetValFromInXml"
    
    strApplicationNumber = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "APPLICATIONFACTFINDNUMBER")
    
    strCustomer1Number = m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "CUSTOMERNUMBER1")
    strCustomer2Number = m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "CUSTOMERNUMBER2")

    strCustomer1VersionNumber = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "CUSTOMERVERSIONNUMBER1")
    strCustomer2VersionNumber = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "CUSTOMERVERSIONNUMBER2")
    
    strPPSubQuoteNumber = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "PPSUBQUOTENUMBER")
    strDistributionChannel = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "DISTRIBUTIONCHANNEL")
    ' DM    02/08/00 SYS0993 Wrong XML tag being used for SPLITAMOUNT, change from SPLITCOVER.
    strSplitCover = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "SPLITAMOUNT")
    strCoverType = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "COVERTYPE")
    ' DM    13/07/00 SYS0993 Wrong XML tags being used. Changed MORTGAGECOVERAMOUNT to MORTGAGEAMOUNT
    '                and COVERTYPEID to COVERTYPE and also PRODUCTCODE to PPPRODUCTCODE.
    strMortgageCoverAmount = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "MORTGAGEAMOUNT")
       
    strMortgageCoverForApplicant1 = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "MORTGAGECOVERFORAPPLICANT1")
       
    strMortgageCoverForApplicant2 = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "MORTGAGECOVERFORAPPLICANT2")
            
    strRecomMortgageCoverAmount = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "RECOMMENDEDMORTGAGECOVERAMOUNT")
    
    strProductCode = _
            m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "PPPRODUCTCODE")
    
    Exit Sub
CalcAndSavePPCosts_GetValFromInXmlVbErr:

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function GetToalNetMonthlyCostFromMortSubQuote( _
                                       ByVal strApplicationNumber As String, _
                                       ByVal strApplicationFactFindNumber As String, _
                                       ByVal strMortgageSubQuoteNumber As String) As Double
' header ----------------------------------------------------------------------------------
' Description:  Gets value of 'Total Net Monthly Cost' from the table 'MotgageSubQuote'
'               from the record that satisfy the condtion (passed in as parameters)
' pass:         strApplicationNumber,  strApplicationFactFindNumber
'               strMortgageSubQuoteNumber
'Errors:        oeRecordNotFound - if corresponding record is not found
'------------------------------------------------------------------------------------------
On Error GoTo GetToalNetMonthlyCostFromMortSubQuoteVbErr

    Const strFunctionName As String = "GetToalNetMonthlyCostFromMortSubQuote"

    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist
    
    Dim strCondition  As String
    Dim blnRecordFound As Boolean
    Dim dblNewRepayment As Double
        
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
   
    Set objSQLAssist = New SQLAssist
   
    'Build the condition and Get value from table 'MORGAGESUBQUOTE'
    strCondition = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber) _
                   & " And APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber _
                   & " And MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber

    Call objADOAssist.GetValueFromTable("MORTGAGESUBQUOTE", strCondition, _
                                        "TOTALNETMONTHLYCOST", dblNewRepayment, _
                                        blnRecordFound)
    
    If Not blnRecordFound Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, _
                                " in table MORTGAGESUBQUOTE  for " _
                                & "Application Number :" & strApplicationNumber _
                                & " And Application Fact Find Number : " & strApplicationFactFindNumber _
                                & " And Mortgage Subquote Number : " & strMortgageSubQuoteNumber
    End If
           
    'If the value fetched is Null, set it to zero
    If IsNull(dblNewRepayment) Then dblNewRepayment = 0
    
    GetToalNetMonthlyCostFromMortSubQuote = dblNewRepayment
    
    ' Clear references to all the objects
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function
    
GetToalNetMonthlyCostFromMortSubQuoteVbErr:
    
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetComplete
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetExitingRepayments_CalcRecomPPCover( _
                                ByVal vxmlRequest As IXMLDOMNode, _
                                ByVal strTypeofApplication As String) As Double
' header ----------------------------------------------------------------------------------
' Description:  This calculate the existing repayment of all the customers linked to an
'               application.
' Called From : CalculateRecommendedPPCover
'
' pass:         xmlDoc : DOM document containing CustomerList.
'------------------------------------------------------------------------------------------
On Error GoTo GetExistingRepaymentsVbErr

    Const strFunctionName As String = "GetExitingRepayments_CalcRecomPPCover"

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlListElem As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strResponse As String   ', strNewRequest As String
    Dim blnIsFurtherAdvance As Boolean
    
    Dim objComboDO As ComboDO
    Dim objCFBO As CustomerFinancialBO
    Dim objConvAssist As ConvertAssist
    
    Set objConvAssist = New ConvertAssist
    
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
        Set objCFBO = New CustomerFinancialBO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objCFBO = _
                m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    End If
        
    blnIsFurtherAdvance = objComboDO.IsFurtherAdvance(strTypeofApplication)
    If blnIsFurtherAdvance Then ' Calculate the existing repayments
        ' PSC 30/06/00 SYS1012 - Start
        strResponse = objCFBO.GetLoansNotRedeemedTotals(vxmlRequest.xml)
        Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
        
        Dim strPaymentsNotRedeemed As String
        strPaymentsNotRedeemed = m_objXMLAssist.GetTagValue(xmlDoc.documentElement, "PAYMENTSONLOANSNOTREDEEMED")
        
        GetExitingRepayments_CalcRecomPPCover = objConvAssist.CSafeDbl(strPaymentsNotRedeemed)
        ' PSC 30/06/00 SYS1012 - End
    Else
        GetExitingRepayments_CalcRecomPPCover = 0#
    End If
    
    Set xmlDoc = Nothing
    Set xmlListElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlNode = Nothing
      
    Set objComboDO = Nothing
    Set objCFBO = Nothing
    Set objConvAssist = Nothing
    
    Exit Function
GetExistingRepaymentsVbErr:
    Set xmlDoc = Nothing
    Set xmlListElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlNode = Nothing
      
    Set objComboDO = Nothing
    Set objCFBO = Nothing
    Set objConvAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetComplete
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub GetValFromInXmlForCalcRecomPPCover(ByVal xmlIn As FreeThreadedDOMDocument40, _
                                               ByRef strApplicationNumber As String, _
                                               ByRef strApplicationFactFindNumber As String, _
                                               ByRef strMortgageSubQuoteNumber As String, _
                                               ByRef strTypeofApplication As String)
' header ----------------------------------------------------------------------------------
' Description:  Fetches all the values from xmlIn and assign them to respective local variables ;
'               this is used in the method 'CalculateRecommendedPPCover'
'
' pass:         xmlIn : xml Document containing the Input data
'               ApplicationNumber,  ApplicationFactFindNumber
'               MortgageSubQuoteNumber, TypeOfApplication
'------------------------------------------------------------------------------------------
On Error GoTo GetValFromInXmlForCalcRecomPPCoverVbErr

    Const strFunctionName As String = "GetValFromInXmlForCalcRecomPPCover"
    
    strApplicationNumber = m_objXMLAssist.GetTagValue(xmlIn.documentElement, "APPLICATIONNUMBER")
    
    strApplicationFactFindNumber = m_objXMLAssist.GetTagValue(xmlIn.documentElement, _
                                                        "APPLICATIONFACTFINDNUMBER")
    
    strMortgageSubQuoteNumber = m_objXMLAssist.GetTagValue(xmlIn.documentElement, _
                                                         "MORTGAGESUBQUOTENUMBER")
    
    strTypeofApplication = m_objXMLAssist.GetTagValue(xmlIn.documentElement, _
                                                    "TYPEOFAPPLICATION")
    
    'If any of the above values are empty raise error
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 _
       Or Len(strMortgageSubQuoteNumber) = 0 Or Len(strTypeofApplication) = 0 Then
    
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
                            " All input values not found for CalculateRecommendedPPCover"
    End If
    
    Exit Sub
    
GetValFromInXmlForCalcRecomPPCoverVbErr:
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub CalcAndSavePPCosts_GetCustDetails(ByVal strCustomerNumber As String, _
                                        ByRef strCustomerVersionNumber As String, _
                                        ByRef strApplicantAge As String, _
                                        ByRef strApplicantGender As String)
' header ----------------------------------------------------------------------------------
' Description:  Fetches Customer Details based on Customer Key Data passed in and assigns
'               age and gender to local variables
' pass:         CustomerNumber, CustomerVersionNumber
'               Applicant's Age,  Applicant's Gender
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSavePPCosts_GetCustDetailsVbErr

    Const strFunctionName As String = "CalcAndSavePPCosts_GetCustDetails"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlData As IXMLDOMElement
    'Dim strXML As String
    
    Dim objICustomerDO As ICustomerDO
    'DM added the context check for creating this object
    If m_objContext Is Nothing Then
        Set objICustomerDO = New CustomerDO
    Else
        Set objICustomerDO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerDO")
    End If

    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.appendChild xmlRequestDoc.createElement("CUSTOMER")
    xmlRequestDoc.documentElement.appendChild _
        (xmlRequestDoc.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber
    xmlRequestDoc.documentElement.appendChild _
        (xmlRequestDoc.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber
    
    Set xmlData = objICustomerDO.GetPersonalDetails(xmlRequestDoc.documentElement)
    
    strApplicantAge = m_objXMLAssist.GetTagValue(xmlData, "AGE")
    strApplicantGender = m_objXMLAssist.GetTagValue(xmlData, "GENDER")
    
    ' Clear the references to the objects
    Set xmlRequestDoc = Nothing
    Set xmlData = Nothing
    Set objICustomerDO = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub
    
CalcAndSavePPCosts_GetCustDetailsVbErr:
    Set xmlRequestDoc = Nothing
    Set xmlData = Nothing
    Set objICustomerDO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

Public Sub CalcAndSavePPCosts_SavePPSubQuote(ByVal xmlOut As FreeThreadedDOMDocument40, _
                                             ByVal strApplicationNumber As String, _
                                             ByVal strApplicationFactFindNumber As String, _
                                             ByVal strPPSubQuoteNumber As String, _
                                             ByVal strApplicant1Premium As String, _
                                             ByVal strApplicant1IPT As String, _
                                             ByVal strApplicant2Premium As String, _
                                             ByVal strApplicant2IPT As String, _
                                             ByVal strProductCode As String, _
                                             ByVal strSplitCover As String, _
                                             ByVal strMortgageCoverAmount As String, _
                                             ByVal strMortgageCoverForApplicant1 As String, _
                                             ByVal strMortgageCoverForApplicant2 As String, _
                                             ByVal strCoverType As String)
' header ----------------------------------------------------------------------------------
' Description:  Builds the inputXML required for saving the Sub-Quote details and updates it.
'               Also modifies XMLOut
' pass:         xmlOut - document to which the results are to be attached
'               Other parameters required for updating SubQuote;
'               Application Number,ApplicationFactFindNumber, PPSubQuoteNumber,
'               Applicant1Premium etc.
'------------------------------------------------------------------------------------------
On Error GoTo CalcAndSavePPCosts_SavePPSubQuoteVbErr

    Const strFunctionName As String = "CalcAndSavePPCosts_SavePPSubQuote"
    
    Dim strXML As String
    
    Dim objConvAssist As ConvertAssist
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
        
    If m_objContext Is Nothing Then
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
    Else
        Set objPPSubQuoteBO = _
            m_objContext.CreateInstance(cstrPPSUBQUOTE_COMP & ".PaymentProtectionSubQuoteBO")
    End If
    
    'Build the XML string required for Saving Sub-Quote
    ' DM    13/07/00 SYS0993 Wrong XML tags being used. Changed MORTGAGECOVERAMOUNT to MORTGAGEAMOUNT
    '                and COVERTYPEID to COVERTYPE and also PRODUCTCODE to PPPRODUCTCODE.
    strXML = "<PAYMENTPROTECTIONSUBQUOTE>"
    strXML = strXML & "<APPLICATIONNUMBER>" & strApplicationNumber & "</APPLICATIONNUMBER>"
    strXML = strXML & "<APPLICATIONFACTFINDNUMBER>" & strApplicationFactFindNumber _
                    & "</APPLICATIONFACTFINDNUMBER>"
    strXML = strXML & "<PPSUBQUOTENUMBER>" & strPPSubQuoteNumber & "</PPSUBQUOTENUMBER>"
    strXML = strXML & "<APPLICANT1PREMIUM>" & strApplicant1Premium & "</APPLICANT1PREMIUM>"
    strXML = strXML & "<APPLICANT1IPT>" & strApplicant1IPT & "</APPLICANT1IPT>"
    strXML = strXML & "<APPLICANT2PREMIUM>" & strApplicant2Premium & "</APPLICANT2PREMIUM>"
    strXML = strXML & "<APPLICANT2IPT>" & strApplicant2IPT & "</APPLICANT2IPT>"
    strXML = strXML & "<PPPRODUCTCODE>" & strProductCode & "</PPPRODUCTCODE>"
    ' DM    02/08/00 SYS0993 Wrong XML tag being used for SPLITAMOUNT, change from SPLITCOVER.
    strXML = strXML & "<SPLITAMOUNT>" & strSplitCover & "</SPLITAMOUNT>"
    strXML = strXML & "<MORTGAGEAMOUNT>" & strMortgageCoverAmount _
                    & "</MORTGAGEAMOUNT>"
    strXML = strXML & "<MORTGAGECOVERFORAPPLICANT1>" & strMortgageCoverForApplicant1 _
                    & "</MORTGAGECOVERFORAPPLICANT1>"
    strXML = strXML & "<MORTGAGECOVERFORAPPLICANT2>" & strMortgageCoverForApplicant2 _
                    & "</MORTGAGECOVERFORAPPLICANT2>"
    strXML = strXML & "<COVERTYPE>" & strCoverType & "</COVERTYPE>"
    strXML = strXML & "</PAYMENTPROTECTIONSUBQUOTE>"
    
    ' Save the sub quote and check the response
    strXML = objPPSubQuoteBO.SaveDetails(strXML)
    m_objErrAssist.CheckXMLResponse strXML, True
        
    ' Add the new nodes (MonthlyCost, MortgageCoverAmount, MortgageCoverForApplicant1
    '  MortgageCoverForApplicant1 to xmlOut
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMElement
    
    Set xmlResponseNode = xmlOut.getElementsByTagName("RESPONSE").Item(0)
    
    Set xmlElement = xmlOut.createElement("MONTHLYCOST")
    xmlElement.Text = objConvAssist.CSafeDbl(strApplicant1Premium) + objConvAssist.CSafeDbl(strApplicant1IPT) _
                     + objConvAssist.CSafeDbl(strApplicant2Premium) + objConvAssist.CSafeDbl(strApplicant2IPT)
    xmlResponseNode.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("MORTGAGEAMOUNT")
    xmlElement.Text = strMortgageCoverAmount
    xmlResponseNode.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("MORTGAGECOVERFORAPPLICANT1")
    xmlElement.Text = strMortgageCoverForApplicant1
    xmlResponseNode.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("MORTGAGECOVERFORAPPLICANT2")
    xmlElement.Text = strMortgageCoverForApplicant2
    xmlResponseNode.appendChild xmlElement
    
    Set objConvAssist = Nothing
    Set objPPSubQuoteBO = Nothing
    
    Exit Sub
CalcAndSavePPCosts_SavePPSubQuoteVbErr:
    
    Set objConvAssist = Nothing
    Set objPPSubQuoteBO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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


Public Function GetValidatedQuotationDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetValidatedQuotationDetailsVbErr
' GD BMIDS00256 Various changes.
    Const strFunctionName As String = "GetValidatedQuotationDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    
    Dim objICostModelAssistDO As ICostModelAssistDO
    'SYS0977  DJP     Dim objCostModelAssistBO As ICostModelAssistBO
    'GD BMIDS00256
    Dim objCostModelAssistBO As ICostModelAssistBO
    
    Dim objQuotationBO As QuotationBO
    Dim objICustomerEmploymentDO As ICustomerEmploymentDO
    Dim objICustomerFinancialDO As ICustomerFinancialDO

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objQuotationBO = New QuotationBO
        Set objICustomerEmploymentDO = New CustomerEmploymentDO
        Set objICustomerFinancialDO = New CustomerFinancialDO
        ' sys0977 DJP SYS1505    Set objCostModelAssistBO = New CostModelAssistBO
        Set objCostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objICustomerEmploymentDO = m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentDO")
        Set objICustomerFinancialDO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialDO")
        ' sys0977 ' DJP SYS1505         Set objCostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objCostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    'Check the root node
    Set xmlRequestNode = xmlIn.getElementsByTagName("BASICQUOTATIONDETAILS").Item(0)
    
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BASICQUOTATIONDETAILS" & " tag not found"
    End If

    'Dim xmlRootNode As IXMLDOMNode
    Dim xmlReturnNode As IXMLDOMNode
    Dim lngError As Long
    Dim strActiveQuoteNumber As String
    Dim xmlLTV As FreeThreadedDOMDocument40
    Set xmlReturnNode = objICostModelAssistDO.GetDataForSubQuoteValidation(xmlRequestNode)
    'Data returned in xmlReturnNode check if any warnings
    'lngError = m_objErrAssist.CheckResponse(xmlReturnNode, xmlResponseNode, True)
    
    strActiveQuoteNumber = m_objXMLAssist.GetTagValue(xmlRequestNode, "ACTIVEQUOTENUMBER")
    If Len(strActiveQuoteNumber) <> 0 Then
        
        Dim xmlElement As IXMLDOMElement
        
        m_objXMLAssist.ChangeNodeName xmlRequestNode, "BASICQUOTATIONDETAILS", "QUICKQUOTEAPPLICANTDETAILS"
        'xmlResponseElem gets used in the next method
        Set xmlElement = xmlRequestNode.ownerDocument.createElement("ISSELFEMPLOYED")
        
        Dim TempDomDoc As FreeThreadedDOMDocument40
        Dim xmlTempNode As IXMLDOMNode
        
        Set TempDomDoc = New FreeThreadedDOMDocument40
        TempDomDoc.validateOnParse = False
        TempDomDoc.setProperty "NewParser", True
        xmlElement.Text = Abs(objICustomerEmploymentDO.IsMainSelfEmployed(xmlReturnNode, TempDomDoc.documentElement))
        
        Set xmlElement = xmlRequestNode.appendChild(xmlElement)
        
        m_objXMLAssist.ChangeNodeName xmlReturnNode, "EMPLOYMENT", "CUSTOMER"
        'NOW find loans to be ported !!!!
        
        ' APS 31/05/00 AQR SYS0793 - Handle record not found on FindLoansToBePorted
        Dim bPortedLoans As Boolean
        On Error Resume Next
        Set xmlTempNode = objICustomerFinancialDO.FindLoansToBePorted(xmlReturnNode)
        m_objErrAssist.SaveErr
        On Error GoTo GetValidatedQuotationDetailsVbErr
        
        ' Record not found is ok
        bPortedLoans = True
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
                m_objErrAssist.ReRaise
            Else
                bPortedLoans = False
            End If
        End If
        'and pass them on Integer value returned here true or false
        
        If (bPortedLoans = True) Then
            Set xmlElement = xmlRequestNode.appendChild(xmlTempNode)
        End If
        m_objXMLAssist.ChangeNodeName xmlRequestNode, "QUICKQUOTEAPPLICANTDETAILS", "QUOTATION"
        
        Dim xmlMSQElement As IXMLDOMElement
        Dim xmlTempMSQ As IXMLDOMElement
        Dim strGetDataResponse As String
        Dim xmlGetDataResponse As FreeThreadedDOMDocument40
        Dim xmlRequestClone As IXMLDOMElement
        
        Set xmlRequestClone = xmlRequestNode.cloneNode(True)
        m_objXMLAssist.ChangeNodeName xmlRequestClone, "ACTIVEQUOTENUMBER", "QUOTATIONNUMBER"
        ' Get the MORTGAGESUBQUOTENUMBER, if there is one.
        ' If there is one, then we need to recalculate the LTV (potentially)
        Set xmlMSQElement = xmlIn.childNodes.Item(0)
        strGetDataResponse = objQuotationBO.GetData(xmlRequestClone.xml)
        ' Check the response
        m_objErrAssist.CheckXMLResponse strGetDataResponse, True, xmlResponseNode
        Set xmlGetDataResponse = m_objXMLAssist.load(strGetDataResponse, TypeName(Me), strFunctionName)
        Set xmlTempMSQ = xmlGetDataResponse.selectSingleNode(".//MORTGAGESUBQUOTENUMBER")
        If Not (xmlTempMSQ Is Nothing) Then
            If Not (Len(xmlTempMSQ.Text) = 0) Then
            Set xmlTempNode = objCostModelAssistBO.GetLTV(xmlMSQElement, "COST MODELLING")
            ' Check the response
            lngError = m_objErrAssist.CheckResponse(xmlTempNode, xmlResponseNode, True)
            ' GD BMIDS00256 if AmountRequested is 0, then do not call CalcCostModelLTV.
            If CSafeDbl(xmlTempNode.selectSingleNode(".//AMOUNTREQUESTED").Text) <> 0 Then
                ' Create a temporary document
                Dim xmlDocCalcLTV As FreeThreadedDOMDocument40
                Dim xmlDocElement As IXMLDOMElement
                Set xmlDocCalcLTV = New FreeThreadedDOMDocument40
                xmlDocCalcLTV.validateOnParse = False
                xmlDocCalcLTV.setProperty "NewParser", True
        
                'copy over the request node no children
                Set xmlDocElement = m_objXMLAssist.GetRequestNodeEx(xmlIn.documentElement)
                Set xmlDocCalcLTV = xmlDocElement.ownerDocument
        
                ' copy over the LTV response node above including children.
                m_objXMLAssist.CopyNode "LTV", "LTV", xmlTempNode.firstChild, xmlDocCalcLTV.firstChild, True
                ' copy over the Customerlist from the employment list
                m_objXMLAssist.CopyNode "EMPLOYMENTLIST", "CUSTOMERLIST", xmlReturnNode, xmlDocCalcLTV.firstChild.firstChild, True
        
                ' Call the CalcCostModelLTV method on this object
                Dim strResponse As String
                
                Set xmlLTV = New FreeThreadedDOMDocument40
                xmlLTV.validateOnParse = False
                xmlLTV.setProperty "NewParser", True
        
                strResponse = CalcCostModelLTV(xmlDocCalcLTV.xml)
                ' Load the string into xml
                Set xmlLTV = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
                ' Check the response
                lngError = m_objErrAssist.CheckResponse(xmlLTV.documentElement, xmlResponseNode, True)

                
                Dim strXMLReturn As String
                Dim xmlDoc As FreeThreadedDOMDocument40

                Dim xmlNode As IXMLDOMNode
                ' DM Need to add the LTV xml to the xml being sent to the GetValidatedQuotationDetails method below.
        
                Set xmlNode = xmlLTV.selectSingleNode("//LTV")
        
                m_objXMLAssist.CopyNode "LTV", "LTV", xmlNode, xmlRequestNode, True
            End If
            End If
        End If

        strXMLReturn = objQuotationBO.GetValidatedQuotationDetails(xmlRequestNode.ownerDocument.xml)
        Set xmlDoc = m_objXMLAssist.load(strXMLReturn, TypeName(Me), strFunctionName)
        lngError = m_objErrAssist.CheckResponse(xmlDoc.documentElement, xmlResponseNode, True)


'sys0977        ' if the error wasn't Record Not Found then raise it
        If lngError <> 0 Then
            lngError = m_objErrAssist.GetOmigaErrorNumber(lngError)
            If lngError <> oeRecordNotFound Then
               m_objErrAssist.ThrowError gstrAPPLICATIONQUOTE, strFunctionName, lngError
            End If
        End If

        m_objXMLAssist.ChangeNodeName xmlRequestNode, "QUOTATION", "QUOTATIONDETAILS"
        'SA BMIDS00365 Need to return these numbers from this method too.{
        m_objXMLAssist.CopyNode "ACCEPTEDQUOTENUMBER", "ACCEPTEDQUOTENUMBER", xmlReturnNode.selectSingleNode(".//ACCEPTEDQUOTENUMBER"), xmlRequestNode, False
        m_objXMLAssist.CopyNode "RECOMMENDEDQUOTENUMBER", "RECOMMENDEDQUOTENUMBER", xmlReturnNode.selectSingleNode(".//RECOMMENDEDQUOTENUMBER"), xmlRequestNode, False
        'SA BMIDS00365 }
        
        Set xmlResponseNode = xmlResponseNode.appendChild(xmlRequestNode)
        m_objXMLAssist.AttachResponseData xmlResponseNode, xmlDoc.documentElement
               

' DC SYS0977 09/08/01
        Dim bValidMortgageSubQuote As Boolean
        Dim bValidBCSubQuote As Boolean
        Dim bValidPPSubQuote As Boolean
        bValidMortgageSubQuote = False
        bValidBCSubQuote = False
        bValidMortgageSubQuote = False
        Dim xmlTempFlag As IXMLDOMElement
        bValidMortgageSubQuote = False
        Set xmlTempFlag = xmlResponseNode.selectSingleNode(".//VALIDMORTGAGESUBQUOTE")
        If Not (xmlTempFlag Is Nothing) Then
            bValidMortgageSubQuote = (xmlTempFlag.Text = "1")
        End If
        'For a PP and BC then, being absent doesn't mean it's invalid.
        bValidBCSubQuote = True
        Set xmlTempFlag = xmlResponseNode.selectSingleNode(".//VALIDBCSUBQUOTE")
''GD     28/01/03 BM0287 START
'        If Not (xmlTempFlag Is Nothing) Then
'            bValidBCSubQuote = (xmlTempFlag.Text = "0")
'        End If
        If Not (xmlTempFlag Is Nothing) Then
            If xmlTempFlag.Text = "1" Then
                bValidBCSubQuote = True
            Else
                bValidBCSubQuote = False
            End If
        End If
''GD     28/01/03 BM0287 END

        bValidPPSubQuote = True
        Set xmlTempFlag = xmlResponseNode.selectSingleNode(".//VALIDPPSUBQUOTE")
''GD     28/01/03 BM0287 START
'        If Not (xmlTempFlag Is Nothing) Then
'            bValidPPSubQuote = (xmlTempFlag.Text = "0")
'        End If
        If Not (xmlTempFlag Is Nothing) Then
            If xmlTempFlag.Text = "1" Then
                bValidPPSubQuote = True
            Else
                bValidPPSubQuote = False
            End If
        End If
''GD     28/01/03 BM0287 END


        Set xmlTempFlag = Nothing
        
        ' These will be false iff. the xmlResponseNode returns a validXXsubquote flag is "0"
        
        If bValidMortgageSubQuote And bValidBCSubQuote And bValidPPSubQuote Then
            If Not (xmlLTV Is Nothing) Then
            If Len(m_objXMLAssist.GetNodeValue(xmlLTV, ".//LTV", False)) > 0 Then
                Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
                Set objMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
                Set xmlDocElement = m_objXMLAssist.GetRequestNodeEx(xmlIn.documentElement)
                Set xmlDoc = xmlDocElement.ownerDocument
                m_objXMLAssist.CopyNode "MORTGAGESUBQUOTE", "MORTGAGESUBQUOTE", xmlRequestNode, xmlDocElement, True
                m_objXMLAssist.CopyNode "APPLICATIONNUMBER", "APPLICATIONNUMBER", xmlRequestNode, xmlDoc.firstChild.firstChild, False
                m_objXMLAssist.CopyNode "APPLICATIONFACTFINDNUMBER", "APPLICATIONFACTFINDNUMBER", xmlRequestNode, xmlDoc.firstChild.firstChild, False
                m_objXMLAssist.CopyNode "LTV", "LTV", xmlNode, xmlDoc.firstChild.firstChild, True
                ' Update the MORTGAGESUBQUOTE table with the LTV.
                ' The LTV will only have changed if the Purchase Price/Estimated value has changed
                strResponse = objMortgageSubQuoteBO.Update(xmlDoc.xml)
                Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
        ' Check the response
                lngError = m_objErrAssist.CheckResponse(xmlDoc.firstChild, xmlResponseNode, False)
        ' if the error wasn't Record Not Found then raise it
                If lngError <> 0 Then
                    lngError = m_objErrAssist.GetOmigaErrorNumber(lngError)
                    If lngError <> oeRecordNotFound Then
                       m_objErrAssist.ThrowError gstrAPPLICATIONQUOTE, strFunctionName, lngError
                    End If
                End If

            End If
            End If
        End If

'GD BMIDS00256 END re-apply 5
    Else
     
        xmlResponseNode.appendChild xmlRequestNode
    
    End If
    
    GetValidatedQuotationDetails = xmlOut.xml

GetValidatedQuotationDetailsExit:

    Set objICostModelAssistDO = Nothing
    Set objQuotationBO = Nothing
    Set objMortgageSubQuoteBO = Nothing
    Set objCostModelAssistBO = Nothing
    Set objICustomerEmploymentDO = Nothing
    Set objICustomerFinancialDO = Nothing

    Exit Function

GetValidatedQuotationDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
 
    GetValidatedQuotationDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetValidatedQuotationDetailsExit

End Function

Public Function AcceptQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Sets the status of the current (ative) quotation to "accepted".
'
'
' pass:         vstrXMLRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "AcceptQuotation"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_AcceptQuotation(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    AcceptQuotation = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    AcceptQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationQuoteBO_AcceptQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    Const strFunctionName As String = "IApplicationQuoteBO_AcceptQuotation"
    Dim strNodeName As String
    Dim strAppNo As String
    Dim strFactFindNo As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objApplicationBO As ApplicationBO
    Dim objIQuotationBO As IQuotationBO
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    On Error GoTo AcceptQuotationErr

    strNodeName = "QUOTATION"
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIQuotationBO = New QuotationBO
        Set objApplicationBO = New ApplicationBO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    ' Check the quotation is not already recommended
    ' Need to get the application data by calling GetApplicationData
    ' The tag for this function is "QUOTATION", so need to rename it to "APPLICATION"
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "FACTFINDNUMBER is empty"
    End If
    
    ' Clone the whole thing, REQUEST included
    Set xmlRequestElem = vxmlRequest.cloneNode(True)
    
    ' Find "QUOTATION"
    'GD SYS2391 Change made to allow XML parser V3 to function correctly
    'ORIGINAL line:Set xmlAppDataElem = m_objXMLAssist.GetNode(xmlRequestElem, "/REQUEST/" & strNodeName, True)
    Set xmlAppDataElem = m_objXMLAssist.GetNode(xmlRequestElem, "./" & strNodeName, True)
    
    ' Rename it
    m_objXMLAssist.ChangeNodeName xmlAppDataElem, strNodeName, "APPLICATION"
    ' SR 23/12/2002 : BM0209 - Fetch data required for CostModelling only
    Set xmlNode = xmlOut.createElement("CALCSDATAONLY")
    xmlNode.Text = "1"
    xmlAppDataElem.appendChild xmlNode
    ' SR 23/12/2002 : BM0209 - End
    
    ' Call GetApplicationData
    Set xmlAppResponseElem = objIApplicationDO.GetApplicationData(xmlRequestElem)

    m_objXMLAssist.CheckNode xmlAppResponseElem, TypeName(Me), strFunctionName
    
    ' Get the ActiveQuoteNumber
    Dim sActiveQuoteNumber As String
    
    ' MS SYS0914 Accepted Quote NOT Recommended
'    Dim sRecommendedQuoteNumber As String
    Dim sAcceptedQuoteNumber As String
    
    sActiveQuoteNumber = m_objXMLAssist.GetTagValue(xmlAppResponseElem, "ACTIVEQUOTENUMBER")
    ' MS SYS0914 Accepted Quote NOT Recommended
'    sRecommendedQuoteNumber = m_objXMLAssist.GetTagValue(xmlAppResponseElem, "RECOMMENDEDQUOTENUMBER")
    sAcceptedQuoteNumber = m_objXMLAssist.GetTagValue(xmlAppResponseElem, "ACCEPTEDQUOTENUMBER")
    
    If Len(sActiveQuoteNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ACTIVEQUOTENUMBER is empty"
    End If
    
    ' MS SYS0914 Accepted Quote NOT Recommended
'    If Len(sRecommendedQuoteNumber) > 0 Then
    If Len(sAcceptedQuoteNumber) > 0 Then
    ' MS SYS0914 Accepted Quote NOT Recommended
'        If objConvertAssist.CSafeLng(sActiveQuoteNumber) = objConvertAssist.CSafeLng(sRecommendedQuoteNumber) Then
        If CSafeLng(sActiveQuoteNumber) = CSafeLng(sAcceptedQuoteNumber) Then
    ' MS SYS0914 Accepted Quote NOT Recommended raise correct error
'            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 200
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 203
        End If
    End If
    
    ' Now need to build up the XML to pass to IsQuotationAffordable
    'Dim xmlDOMQuotation As FreeThreadedDOMDocument40
    'Dim xmlQuotationElem As IXMLDOMElement
    'Dim xmlQuotationItem As IXMLDOMElement
    'Dim xmlQuotationRespElem As IXMLDOMElement
    '===========================================================================================================
    ' SR 20/12/2002 : BM0209 - Affordability check is not required for BMIDS
    
'    Set xmlRequestElem = xmlDOMQuotation.appendChild(xmlDOMQuotation.createElement("REQUEST"))
'    Set xmlQuotationElem = xmlRequestElem.appendChild(xmlDOMQuotation.createElement("SEARCH"))
'    Set xmlQuotationElem = xmlQuotationElem.appendChild(xmlDOMQuotation.createElement("QUOTATION"))
'
'    Set xmlQuotationItem = xmlQuotationElem.appendChild(xmlDOMQuotation.createElement("APPLICATIONNUMBER"))
'    xmlQuotationItem.Text = strAppNo
'
'    Set xmlQuotationItem = xmlQuotationElem.appendChild(xmlDOMQuotation.createElement("APPLICATIONFACTFINDNUMBER"))
'    xmlQuotationItem.Text = strFactFindNo
'
'    Set xmlQuotationItem = xmlQuotationElem.appendChild(xmlDOMQuotation.createElement("QUOTATIONNUMBER"))
'    xmlQuotationItem.Text = sActiveQuoteNumber
    
    ' Check Quotation is affordable
'    Set xmlQuotationRespElem = objIQuotationBO.IsQuotationAffordable(xmlDOMQuotation.documentElement)
'
'    ' Check the response to make sure it's valid
'    m_objErrAssist.CheckResponse xmlQuotationRespElem, xmlResponseElem, True
'
'    Dim bAffordable As Boolean
'
'    bAffordable = objConvertAssist.CSafeBool(m_objXmlAssist.GetTagValue(xmlQuotationRespElem, "ISQUOTATIONAFFORDABLE"))
'
'    If bAffordable = False Then
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 204
'    End If
    
    ' SR 20/12/2002 : BM0209 - End
    '===========================================================================================================
    
    ' Set the Quotation as Recommended
    ' Set the Quotation as Accepted !
    ' MS SYS0914 Accepted Quote NOT Recommended
'    Dim xmlRecommendedQuoteNode As IXMLDOMNode
    Dim xmlAcceptedQuoteNode As IXMLDOMNode
    
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlApplicationDOM As FreeThreadedDOMDocument40

'==============================================================================================================
'SR 23/12/2002 : BM0209 - This part is not required
'    ' Get the recommended quote node
'    ' Get the accepted quote node
'    Set xmlApplicationNode = m_objXmlAssist.GetNode(xmlAppResponseElem, "/APPLICATIONLATESTDETAILS/APPLICATION")
'
''    Set xmlRecommendedQuoteNode = m_objXMLAssist.GetNode(xmlApplicationNode, "./APPLICATIONFACTFIND/RECOMMENDEDQUOTENUMBER")
'        ' MS SYS0914 Accepted Quote NOT Recommended
'    Set xmlAcceptedQuoteNode = m_objXmlAssist.GetNode(xmlApplicationNode, "./APPLICATIONFACTFIND/ACCEPTEDQUOTENUMBER")
'
'    ' Is it valid?
''    m_objXMLAssist.CheckNode xmlRecommendedQuoteNode, TypeName(Me), strFunctionName
'        ' MS SYS0914 Accepted Quote NOT Recommended
'    m_objXmlAssist.CheckNode xmlAcceptedQuoteNode, TypeName(Me), strFunctionName
'
'   ' Update the Active Quote Number
''    xmlRecommendedQuoteNode.Text = sActiveQuoteNumber
'        ' MS SYS0914 Accepted Quote NOT Recommended
'    xmlAcceptedQuoteNode.Text = sActiveQuoteNumber
'SR 23/12/2002 : BM0209 - End
'==============================================================================================================
    Dim xmlUpdateNode As IXMLDOMNode
    
    ' Is the owner document valid?
    m_objXMLAssist.CheckNode vxmlRequest.ownerDocument, TypeName(Me), strFunctionName
    
    ' SR 23/12/2002 : BM0209 - Update the Accepted Quote Number only in the table ApplicationFactFind
    Set xmlUpdateNode = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    
    ' Is the request node valid?
    m_objXMLAssist.CheckNode xmlUpdateNode, TypeName(Me), strFunctionName
    
    Set xmlApplicationNode = xmlOut.createElement("APPLICATIONFACTFIND")
    xmlUpdateNode.appendChild xmlApplicationNode
    xmlApplicationNode.appendChild xmlAppDataElem.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlApplicationNode.appendChild xmlAppDataElem.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    
    Set xmlAcceptedQuoteNode = xmlOut.createElement("ACCEPTEDQUOTENUMBER")
    xmlAcceptedQuoteNode.Text = sActiveQuoteNumber
    xmlApplicationNode.appendChild xmlAcceptedQuoteNode
        
    ' Now do the update
    Dim strResp As String
    strResp = objApplicationBO.UpdateApplicationFactFind(xmlUpdateNode.xml)
    ' SR 23/12/2002 : BM0209 - End
    
    Set xmlApplicationDOM = New FreeThreadedDOMDocument40
    xmlApplicationDOM.validateOnParse = False
    xmlApplicationDOM.setProperty "NewParser", True
    xmlApplicationDOM.loadXML strResp
            
    m_objErrAssist.CheckResponse xmlApplicationDOM.documentElement, xmlResponseElem, True
        
    'MC 23/01/2001 SYS1855. Create Application Fee Types from Mortgage One Off Costs
    Set xmlTempResponse = CreateApplicationFeeTypes(strAppNo, strFactFindNo, sActiveQuoteNumber)
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    
    Set IApplicationQuoteBO_AcceptQuotation = xmlResponseElem
AcceptQuotationExit:
        
    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set objIApplicationDO = Nothing
    Set objApplicationBO = Nothing
    Set objIQuotationBO = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing

    Exit Function

AcceptQuotationErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_AcceptQuotation = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume AcceptQuotationExit
    
End Function

Private Function CreateApplicationFeeTypes(ByVal strApplicationNumber As String, _
                                            ByVal strApplicationFFNumber As String, _
                                            ByVal strAcceptedQuoteNumber As String) As IXMLDOMNode

'MC 23/01/2001 SYS1855. Create Application Fee Types from Mortgage One Off Costs

On Error GoTo CreateApplicationFeeTypesVbErr

Dim objPProcBO As PaymentProcessingBO
Dim objPProcDO As PaymentProcessingDO
Dim objQuotationDO As IQuotationDO
Dim objAppBO As ApplicationBO
Dim xmlRequestElement As IXMLDOMElement
Dim xmlTempRequestElement As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim xmlChildElement As IXMLDOMElement
Dim xmlRequest As IXMLDOMNode
Dim xmlResponse As IXMLDOMNode
Dim xmlResponseElem As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRespDoc As FreeThreadedDOMDocument40
' DRC BMIDS767 Rev - Start
Dim objIMortgageSubQuoteBO As MortgageSubQuoteBO
Dim xmlOneOffCostsList As IXMLDOMNodeList
Dim xmlOneOffCostsListNode As IXMLDOMNode
Dim xmlOneOffCostNode As IXMLDOMNode
'Dim xmlAddedToLoanNode As IXMLDOMNode
Dim strFeeType As String
Dim strPaymentEvent As String
Dim blnSingleRefundFee As Boolean
'Dim intCntFeePayments As Integer
Dim strFeeAmount As String
Dim intValue As Integer
'*=[MC]changed integer to double.
Dim intAmount As Double
' DRC BMIDS767 Rev - End

Dim xmlNode As IXMLDOMElement
Dim xmlFeeTypeList As IXMLDOMNodeList
Dim UndeletedFeetypeList As Collection
Dim xmlPayNode As IXMLDOMElement
Dim strResp As String
Dim strMSQNumber As String
Dim strEntry As String
Dim strErrMess As String
Dim lngErrNo As Long
'Dim intUnDeleteCount As Integer

' PSC 04/05/2005 MAR1497
Dim xmlFeePaymentList As IXMLDOMNodeList

'SR 15/07/2004 : BMIDS795 - Start
Dim intAppFeeTypeCount As Integer
Dim objADOAssist As ADOAssist
Dim strCondition As String
'SR 15/07/2004 : BMIDS795  - End

Const strFunctionName As String = "CreateApplicationFeeTypes"

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlRespDoc = New FreeThreadedDOMDocument40
    xmlRespDoc.validateOnParse = False
    xmlRespDoc.setProperty "NewParser", True
    Set xmlNode = xmlRespDoc.createElement("DUMMY")
    
    'Get the Mortgage Sub Quote Number for the Accepted Quote
    Set objQuotationDO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationDO")
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("QUOTATION")
    xmlRequest.appendChild xmlElement
    Set xmlChildElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlChildElement.Text = strApplicationNumber
    xmlElement.appendChild xmlChildElement
    Set xmlChildElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlChildElement.Text = strApplicationFFNumber
    xmlElement.appendChild xmlChildElement
    Set xmlChildElement = xmlDoc.createElement("ACTIVEQUOTENUMBER")
    xmlChildElement.Text = strAcceptedQuoteNumber
    xmlElement.appendChild xmlChildElement
    
    Set xmlResponse = objQuotationDO.GetQuotationDetails(xmlRequest)
    strMSQNumber = m_objXMLAssist.GetMandatoryElementText(xmlResponse, "MORTGAGESUBQUOTENUMBER")
    
    Set objPProcBO = m_objContext.CreateInstance(gstrPAYMENTPROCESSING & ".PaymentProcessingBO")
    Set objPProcDO = m_objContext.CreateInstance(gstrPAYMENTPROCESSING & ".PaymentProcessingDO")
    Set objAppBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Set xmlRequestElement = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("APPLICATIONFEETYPE")
    Set UndeletedFeetypeList = New Collection
    xmlElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    
    ' SR 15/07/2004: BMIDS795 - Do not create new Application Costs without deleting the existing ones
    If CSafeInt(strAcceptedQuoteNumber) = 1 Then
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        strCondition = " APPLICATIONNUMBER = '" & strApplicationNumber & "'"
        intAppFeeTypeCount = objADOAssist.GetNumberOfRecords("APPLICATIONFEETYPE", strCondition)
    End If
    ' SR 15/07/2004: BMIDS795 - End
    
    If CSafeInt(strAcceptedQuoteNumber) = 1 And intAppFeeTypeCount = 0 Then ' SR 15/07/2004: BMIDS795
        'Create Application Fee Type records
        xmlRequestElement.setAttribute "OPERATION", "CreateApplicationFeeTypes"
        xmlElement.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
        xmlElement.setAttribute "MORTGAGESUBQUOTENUMBER", strMSQNumber
        xmlRequestElement.appendChild xmlElement
        strResp = objPProcBO.omPayProcRequest(xmlRequestElement.xml)
        xmlRespDoc.loadXML strResp
        'AQR 3518 DRC BugFix added 'Resp'
        m_objErrAssist.CheckResponse xmlRespDoc.documentElement, xmlElement, True
    Else
    'DRC - BMIDS767 Rev start
        'Need to get a list of the OneOffCosts that aren't added to loan
        Set xmlRequest = xmlRequest.cloneNode(False)
        Set xmlElement = xmlDoc.createElement("ONEOFFCOST")
        xmlRequest.appendChild xmlElement
        Set xmlChildElement = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlChildElement.Text = strApplicationNumber
        xmlElement.appendChild xmlChildElement
        Set xmlChildElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlChildElement.Text = strApplicationFFNumber
        xmlElement.appendChild xmlChildElement
        Set xmlChildElement = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
        xmlChildElement.Text = strMSQNumber
        xmlElement.appendChild xmlChildElement
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        strResp = objIMortgageSubQuoteBO.GetOneOffCostsDetails(xmlRequest.xml)
        xmlDoc.loadXML strResp
        lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
        Set xmlOneOffCostsListNode = xmlDoc.selectSingleNode(".//MORTGAGEONEOFFCOSTLIST")
     'DRC - BMIDS767 Rev End
        'Delete existing Fee Types and replace
        xmlRequestElement.setAttribute "OPERATION", "FindFeeTypeList"
        Set xmlElement = xmlDoc.createElement("APPLICATIONFEETYPE")
        xmlElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber
        xmlRequestElement.appendChild xmlElement
        strResp = objPProcBO.omPayProcRequest(xmlRequestElement.xml)
        
        'Check response for errors. Ignore record not found
        xmlDoc.loadXML strResp
        lngErrNo = m_objErrAssist.GetErrorNumberFromResponse(xmlDoc.documentElement, True)
        
        'If no error...
        If lngErrNo = 0 Or lngErrNo = oeRecordNotFound Then
        'BMIDS767 DRC - Check for AdHoc FeeType & sort out reporting into memopad
          ' Start
            
            Set xmlFeeTypeList = xmlDoc.selectNodes(".//APPLICATIONFEETYPE")
            For Each xmlNode In xmlFeeTypeList
                ' PSC 04/05/2006 MAR1497 - Start
                blnSingleRefundFee = False
            
                Set xmlFeePaymentList = xmlNode.selectNodes(".//FEEPAYMENT")
                
                If xmlFeePaymentList.length > 0 Then
                    Set xmlPayNode = xmlFeePaymentList.Item(0)
                    
                    If xmlFeePaymentList.length = 1 Then
                        strPaymentEvent = xmlGetAttributeAsInteger(xmlFeePaymentList.Item(0), "PAYMENTEVENT", -1)
                        
                        If IsValidationTypeInValidationList("PaymentEvent", "RFV", strPaymentEvent) Then
                            blnSingleRefundFee = True
                        End If
                   End If
                End If
                ' PSC 04/05/2006 MAR1497 - End
            
                strFeeType = xmlNode.getAttribute("FEETYPE") 'SR 13/07/2004 : BMIDS767
                If ((Not xmlPayNode Is Nothing) And Not blnSingleRefundFee) _
                   Or ((xmlGetAttributeAsBoolean(xmlNode, "ADHOCIND") And _
                       Not IsValidationTypeInValidationList("OneOffCost", "APR", strFeeType))) Then  'SR 13/07/2004 : BMIDS767
                       
                   ' insert a memo into the memopad
                    'SR 14/07/2004 : Add memopad message only when payment exists
                    If Not xmlPayNode Is Nothing Then

                        strEntry = strEntry & IIf(Len(strEntry) > 0, ", ", "") & _
                                   comboAssistEx.GetComboText("OneOffCost", CInt(strFeeType))
                        
                        ' add application feetype node  to UNDeleted list
                        strErrMess = comboAssistEx.GetComboText("OneOffCost", CInt(strFeeType))
                        If (Len(xmlNode.getAttribute("AMOUNT")) > 0) Then
                            strErrMess = strErrMess & " Amount = " & xmlNode.getAttribute("AMOUNT")
                        End If
                    
                        UndeletedFeetypeList.Add (strErrMess)
                    End If
           'DRC BMIDS767 rev
               'Look for Matching One Off Cost Record
                    strFeeType = xmlNode.getAttribute("FEETYPE")
                    strFeeAmount = xmlNode.getAttribute("AMOUNT")
                    Set xmlOneOffCostNode = xmlOneOffCostsListNode.selectSingleNode(".//MORTGAGEONEOFFCOST[MORTGAGEONEOFFCOSTTYPE = '" & _
                                                 strFeeType & "' and AMOUNT = '" & strFeeAmount & "']")
                                                                   
                    If Not (xmlOneOffCostNode Is Nothing) Then
                        'Remove the Matching One Off Cost out of the list so we don't generate a new App Fee Type record later
                          xmlOneOffCostsListNode.removeChild xmlOneOffCostNode
                    End If

                'DRC BMIDS767 rev
                Else
                    If lngErrNo = 0 Then
                     'MAR377 - DRC - need to be able to delete the payment record and the feepayment record
                     '    in the case of a  single refund fee payment
                        If blnSingleRefundFee Then
                         'Delete existing PAYMENTRECORD  & FEEPAYMENT record
                            If (Not xmlNode.selectSingleNode(".//FEEPAYMENT") Is Nothing) Then
                                Set xmlTempRequestElement = xmlDoc.createElement("REQUEST")
                                xmlTempRequestElement.setAttribute "OPERATION", "DeleteFeeTypePayment"
                                xmlTempRequestElement.appendChild xmlNode.selectSingleNode(".//FEEPAYMENT")
                                xmlTempRequestElement.appendChild xmlNode.selectSingleNode(".//PAYMENTRECORD")
                                strResp = objPProcBO.omPayProcRequest(xmlTempRequestElement.xml)
                                xmlRespDoc.loadXML strResp
                                m_objErrAssist.CheckResponse xmlRespDoc.documentElement, xmlElement, True
                            End If
                        End If
                    'MAR377 - End
                    'Delete existing APPLICATIONFEETYPE record
                        Set xmlTempRequestElement = xmlDoc.createElement("REQUEST")
                        xmlTempRequestElement.setAttribute "OPERATION", "DeleteApplicationFeeTypes"
                        xmlTempRequestElement.appendChild xmlNode
                        strResp = objPProcBO.omPayProcRequest(xmlTempRequestElement.xml)
                        xmlRespDoc.loadXML strResp
                        m_objErrAssist.CheckResponse xmlRespDoc.documentElement, xmlElement, True

                    End If
                End If
            Next
            ' Now go thru the One Off Cost list, looking for any record that is Not to be added to loan
            ' Must also be > 0  and have a validation type <> XAC
            ' and generate a new  App Fee Type records from it
            Set xmlOneOffCostsList = xmlOneOffCostsListNode.selectNodes(".//MORTGAGEONEOFFCOST[ADDTOLOAN != '1']")
            For Each xmlOneOffCostNode In xmlOneOffCostsList
          'check the oneOffCost does not have a validationType of XAC
                intValue = xmlGetMandatoryNodeAsInteger(xmlOneOffCostNode, "MORTGAGEONEOFFCOSTTYPE")
                'intAmount = xmlGetMandatoryNodeAsInteger(xmlOneOffCostNode, "AMOUNT")
                '*=[MC]intAmount changed to Double type, variable notation hasn't been changed to minimise the impact on code
                intAmount = xmlGetMandatoryNodeAsDouble(xmlOneOffCostNode, "AMOUNT")
                
                If (IsValidationType("OneOffCost", intValue, "XAC") = False) And _
                   (intAmount > 0) Then
                    ' create a new blank application feetype record
                    Set xmlTempRequestElement = xmlDoc.createElement("REQUEST")
                    xmlTempRequestElement.setAttribute "OPERATION", "CreateApplicationFeeType"
                    Set xmlElement = xmlDoc.createElement("APPLICATIONFEETYPE")
                    xmlElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber
                    xmlElement.setAttribute "FEETYPE", m_objXMLAssist.GetNodeValue(xmlOneOffCostNode, ".//MORTGAGEONEOFFCOSTTYPE")
                    xmlElement.setAttribute "AMOUNT", m_objXMLAssist.GetNodeValue(xmlOneOffCostNode, ".//AMOUNT")
                    xmlElement.setAttribute "ADHOCIND", m_objXMLAssist.GetNodeValue(xmlOneOffCostNode, ".//ADHOCIND", , False)
                    'MAR28
                    xmlElement.setAttribute "REFUNDAMOUNT", m_objXMLAssist.GetNodeValue(xmlOneOffCostNode, ".//REFUNDAMOUNT", , False)
                    xmlTempRequestElement.appendChild xmlElement
                    strResp = objPProcBO.omPayProcRequest(xmlTempRequestElement.xml)
                    xmlRespDoc.loadXML strResp
                    m_objErrAssist.CheckResponse xmlRespDoc.documentElement, xmlElement, True
                End If
            Next
            
            'BMIDS977  Do not display an Interactive warning message
            
            'If UNDeleted list is not empty
            If UndeletedFeetypeList.Count > 0 Then
               
                'SR 15/07/2004 : BMIDS795 - Add memopad message (one for all the payment records)
                Set xmlTempRequestElement = xmlDoc.createElement("REQUEST")
                Set xmlElement = xmlDoc.createElement("MEMOPAD")
                xmlTempRequestElement.appendChild xmlElement
                Set xmlChildElement = xmlDoc.createElement("APPLICATIONNUMBER")
                xmlChildElement.Text = strApplicationNumber
                xmlElement.appendChild xmlChildElement
                Set xmlChildElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
                xmlChildElement.Text = strApplicationFFNumber
                xmlElement.appendChild xmlChildElement
                Set xmlChildElement = xmlDoc.createElement("ACTIONNAME")
                xmlChildElement.Text = "Accept"
                xmlElement.appendChild xmlChildElement
                Set xmlChildElement = xmlDoc.createElement("MEMOENTRY")
                xmlElement.appendChild xmlChildElement
                If UndeletedFeetypeList.Count > 1 Then
                    ' PSC 04/05/2006 MAR1497
                    strEntry = strEntry & " have payment/rebate/deduction/already refunded or not to be refunded records against them or have been returned as result of return of funds."
                    strEntry = "Fee Types: " & strEntry
                Else
                    ' PSC 04/05/2006 MAR1497
                    strEntry = strEntry & " has payment/rebate/deduction/already refunded or not to be refunded records against it or has been returned as the result of a return of funds."
                    strEntry = "Fee Type: " & strEntry
                End If
            
                xmlChildElement.Text = strEntry
                Set xmlChildElement = xmlDoc.createElement("SCREENDESCRIPTION")
                xmlElement.appendChild xmlChildElement
                xmlChildElement.Text = "Cost Modelling"
                Set xmlChildElement = xmlDoc.createElement("SCREENREFERENCE")
                xmlElement.appendChild xmlChildElement
                xmlChildElement.Text = "CM010"
                Set xmlChildElement = xmlDoc.createElement("ENTRYTYPE")
                xmlElement.appendChild xmlChildElement
                xmlChildElement.Text = comboAssistEx.GetFirstComboValueId("MemoPadEntryType", "RF")
                strResp = objAppBO.CreateMemoPad(xmlTempRequestElement.xml)
                xmlRespDoc.loadXML strResp
                m_objErrAssist.CheckResponse xmlRespDoc.documentElement, xmlElement, True
                'SR 15/07/2004 : BMIDS795 - End
            End If
        Else
            'Raise an error
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, lngErrNo
        End If
    End If
    
    Set CreateApplicationFeeTypes = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateApplicationFeeTypesExit:
    '   clean-up section, release object references
    Set objPProcBO = Nothing
    Set objPProcDO = Nothing
    Set objQuotationDO = Nothing
    Set xmlRequestElement = Nothing
    Set xmlElement = Nothing
    Set xmlChildElement = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlDoc = Nothing
    Set xmlRespDoc = Nothing
    Set xmlNode = Nothing
    Set xmlResponseElem = Nothing
    Set objADOAssist = Nothing
    Set UndeletedFeetypeList = Nothing
    Set objAppBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    ' PSC 04/05/2005 MAR1497
    Set xmlFeePaymentList = Nothing

    Exit Function
CreateApplicationFeeTypesVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    ' PSC 04/05/2006 MAR1497
    Set CreateApplicationFeeTypes = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
     
    ' go to clean-up section
    Resume CreateApplicationFeeTypesExit

End Function

Public Function GetLoanCompositionDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the loan composition details for a given mortgage subquote
'
' pass:         vstrXMLRequest  xml Request as follows:
'
'    <REQUEST>
'      <LOANCOMPOSITION>
'        <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
'        <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'      </LOANCOMPOSITION>
'   </REQUEST>"

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanCompositionDetailsVbErr

    Const cstrFunctionName As String = "GetLoanCompositionDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    'BMIDS622
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    'Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'BMIDS622 End
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
    
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IApplicationQuoteBO_GetLoanCompositionDetails(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    'm_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetLoanCompositionDetails = xmlResponseElem.xml
    
    'BMIDS622
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    'BMIDS622 End

GetLoanCompositionDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetLoanCompositionDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    GetLoanCompositionDetails = m_objErrAssist.CreateErrorResponse
    
    'BMIDS622
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    'BMIDS622 End
    
    ' go to clean-up section
    Resume GetLoanCompositionDetailsExit
End Function

Private Function IApplicationQuoteBO_GetLoanCompositionDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the loan composition details for a given mortgage subquote
'
'
' pass:         vstrXMLRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    On Error GoTo GetLoanCompositionDetailsErr
    
    Const cstrFunctionName As String = "IApplicationQuoteBO_GetLoanCompositionDetails"

    Dim strNodeName As String
    Dim strAppNo As String
    Dim strFactFindNo As String
    Dim strMortgageSubQuote As String
    Dim strLocation As String
    Dim sTypeOfValutation As String
    'Dim sSharedOwnerShipPercentage As String
    Dim strPurchasePrice As String
    Dim strApplicationType As String
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlDomAppData As FreeThreadedDOMDocument40
    Dim xmlAppDataItem As IXMLDOMElement
    Dim xmlAppFactFindNode As IXMLDOMNode
    Dim xmlNewPropertyResp As IXMLDOMElement
    Dim xmlNewPropertyRoot As IXMLDOMElement
    Dim xmlNewPropertyElem As IXMLDOMElement
    Dim xmlNewPropertyItem As IXMLDOMElement
    Dim xmlDOMNewProperty As FreeThreadedDOMDocument40
    'Dim xmlSharedOwnerShipPercentage As IXMLDOMNode
    Dim xmlSubQuoteResp As IXMLDOMElement
    Dim xmlDOMSubQuoteReq As FreeThreadedDOMDocument40
    Dim xmlSubQuoteRoot As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlSubQuoteItem As IXMLDOMElement
    'Dim xmlAppNode  As IXMLDOMNode
    
    Dim objIApplicationDO As IApplicationDO
    'Dim objApplicationBO As ApplicationBO
    Dim objINewPropertyDO As INewPropertyDO
    'Dim objIQuotationBO As IQuotationBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    
    Dim xmlValuationTypeNode As IXMLDOMNode
    Dim xmlLocationNode As IXMLDOMNode
    
    '*=[MC]BMIDS763 - CC075 - APPLICATIONDATE
    Dim strApplicationDate As String
    'END SECTION
    
    strNodeName = "LOANCOMPOSITION"
    
    ' Create default repsonse block

    Set xmlOut = New FreeThreadedDOMDocument40        'BMIDS622
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objINewPropertyDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyDO")
    End If

'Stage1:

    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), cstrFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuote = m_objXMLAssist.GetTagValue(xmlRequestElem, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strMortgageSubQuote) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMissingPrimaryTag, "MORTGAGESUBQUOTENUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMissingPrimaryTag, "FACTFINDNUMBER is empty"
    End If

'Stage2:
    
    'Get Application Data
    
    ' Build the request
    Set xmlAppDataElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    
    m_objXMLAssist.CheckNode xmlAppDataElem, TypeName(Me), cstrFunctionName
       
    Set xmlDomAppData = New FreeThreadedDOMDocument40
    xmlDomAppData.validateOnParse = False
    xmlDomAppData.setProperty "NewParser", True
    Set xmlDomAppData.documentElement = xmlAppDataElem
    Set xmlAppDataElem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATION"))
    
    m_objXMLAssist.CheckNode xmlAppDataElem, TypeName(Me), cstrFunctionName
    
    ' Add the application number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONNUMBER"))
    m_objXMLAssist.CheckNode xmlAppDataItem, TypeName(Me), cstrFunctionName
    xmlAppDataItem.Text = strAppNo
    
    ' Add the application fact find number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlAppDataItem.Text = strFactFindNo
    
    ' SR 19/12/2002 : BM0209 - get data required for CostModelling only
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("CALCSDATAONLY"))
    xmlAppDataItem.Text = "1"
    ' SR 19/12/2002 : BM0209 - End
    
    ' Get the ApplicationData
    Set xmlAppResponseElem = objIApplicationDO.GetApplicationData(xmlDomAppData.documentElement)
    
    Set xmlAppFactFindNode = m_objXMLAssist.GetNode(xmlAppResponseElem, "//APPLICATIONFACTFIND")
    strPurchasePrice = m_objXMLAssist.GetTagValue(xmlAppFactFindNode, "PURCHASEPRICEORESTIMATEDVALUE")
    strApplicationType = m_objXMLAssist.GetTagValue(xmlAppFactFindNode, "TYPEOFAPPLICATION")
    xmlResponseElem.appendChild xmlAppFactFindNode
    
    '*=[MC]BMIDS763 - APPEND APPLICATION DATE ELEMENT TO THE REQUEST
    Call vxmlRequest.appendChild(m_objXMLAssist.GetNode(xmlAppFactFindNode, "APPLICATIONDATE").cloneNode(True))
    'SECTION END
    
    ' SR 19/12/2002 : BM0209 - Start
'    Set xmlAppNode = m_objXmlAssist.GetNode(xmlAppResponseElem, ".//APPLICATION")
'    xmlResponseElem.appendChild xmlAppNode
    ' SR 19/12/2002 : BM0209 - End
'Stage2:

    ' Get New Property details
    Set xmlNewPropertyRoot = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDOMNewProperty = xmlNewPropertyRoot.ownerDocument
    
    Set xmlNewPropertyElem = xmlNewPropertyRoot.appendChild(xmlDOMNewProperty.createElement("NEWPROPERTY"))
    
    Set xmlNewPropertyItem = xmlNewPropertyElem.appendChild(xmlDOMNewProperty.createElement("APPLICATIONNUMBER"))
    xmlNewPropertyItem.Text = strAppNo
    
    Set xmlNewPropertyItem = xmlNewPropertyElem.appendChild(xmlDOMNewProperty.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlNewPropertyItem.Text = strFactFindNo
    
    Set xmlNewPropertyResp = objINewPropertyDO.GetValuationTypeAndLocation(xmlNewPropertyRoot)
    
    strLocation = m_objXMLAssist.GetTagValue(xmlNewPropertyResp, "PROPERTYLOCATION")
    sTypeOfValutation = m_objXMLAssist.GetTagValue(xmlNewPropertyResp, "VALUATIONTYPE")
    
    'Set xmlSharedOwnerShipPercentage = m_objXMLAssist.GetNode(xmlNewPropertyResp, "SHAREDOWNERSHIPPERCENTAGE", False)
    'Set xmlSharedOwnerShipPercentage = m_objXMLAssist.GetNode(xmlNewPropertyResp, "SHAREDPERCENTAGE", False)
        
    'If xmlSharedOwnerShipPercentage Is Nothing Then
            
        'Dim xmlTempElement As IXMLDOMElement
        'Set xmlTempElement = xmlOut.createElement("SHAREDOWNERSHIPPERCENTAGE")
        'Set xmlTempElement = xmlOut.createElement("SHAREDPERCENTAGE")
        'xmlTempElement.Text = "100"
        'xmlTempElement.Text = xmlSharedOwnerShipPercentage.Text
        'xmlNewPropertyResp.appendChild xmlTempElement
        'xmlNewPropertyResp.appendChild xmlSharedOwnerShipPercentage
    
    'Else
    '    Dim xmlTempElement As IXMLDOMElement
    '    Set xmlTempElement = xmlOut.createElement("SHAREDPERCENTAGE")
    '    If Len(xmlSharedOwnerShipPercentage.Text) = 0 Then
    '        xmlSharedOwnerShipPercentage.Text = "100"
    '    End If
    'End If
    
    If Len(sTypeOfValutation) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, "TYPEOFVALUATION is empty"
    End If
    
    If Len(strLocation) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, "PROPERTYLOCATION is empty"
    End If
    
    '***************************************************************************
    'DPF (29/10/2002 - Append Valuation Type & Property Location to the Response
    '***************************************************************************
    'Set xmlAppFactFindNode = m_objXmlAssist.GetNode(xmlResponseElem, "APPLICATIONFACTFIND")
        
    Set xmlValuationTypeNode = m_objXMLAssist.GetNode(xmlNewPropertyResp, "VALUATIONTYPE")
    Set xmlLocationNode = m_objXMLAssist.GetNode(xmlNewPropertyResp, "PROPERTYLOCATION")
            
    xmlAppFactFindNode.appendChild xmlValuationTypeNode
    xmlAppFactFindNode.appendChild xmlLocationNode
    
    'END OF CHANGE
    '***************************************************************************
        
'Stage3:

    ' Get Loan Composition Details
    
    Set xmlSubQuoteRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlDOMSubQuoteReq = xmlSubQuoteRoot.ownerDocument
    
    Set xmlSubQuoteElem = xmlSubQuoteRoot.appendChild(xmlDOMSubQuoteReq.createElement("MORTGAGESUBQUOTE"))
    
    ' Now add the application number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONNUMBER"))
    xmlSubQuoteItem.Text = strAppNo
        
    '... and application fact find number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlSubQuoteItem.Text = strFactFindNo
    
    '*=[MC] BMIDS763 APPLICATIONDATE
    strApplicationDate = m_objXMLAssist.GetTagValue(xmlAppFactFindNode, "APPLICATIONDATE", True, True)
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONDATE"))
    xmlSubQuoteItem.Text = strApplicationDate
    '*=SECTION END - MC
    
    ' Mortgage sub quote number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("MORTGAGESUBQUOTENUMBER"))
    xmlSubQuoteItem.Text = strMortgageSubQuote
        
    ' Location
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("LOCATION"))
    xmlSubQuoteItem.Text = strLocation
    
    ' Type of Valuation
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("TYPEOFVALUATION"))
    xmlSubQuoteItem.Text = sTypeOfValutation
    
    'MAR1061 don't send purchaseprice, it is picked up from MSQ
    ' Purchase Price
    'Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("PURCHASEPRICE"))
    'xmlSubQuoteItem.Text = strPurchasePrice
    
    ' Type of Application
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("TYPEOFAPPLICATION"))
    xmlSubQuoteItem.Text = strApplicationType
    
    Set xmlSubQuoteResp = objIMortgageSubQuoteBO.GetLoanCompositionDetails(xmlSubQuoteRoot)
    m_objErrAssist.CheckResponse xmlSubQuoteResp, xmlResponseElem, True
    
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlSubQuoteResp
            
    Set IApplicationQuoteBO_GetLoanCompositionDetails = xmlResponseElem

GetLoanCompositionDetailsExit:

    '   clean-up section, release object references
    'BMIDS622 set objects to nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlDomAppData = Nothing
    Set xmlAppDataItem = Nothing
    Set xmlAppFactFindNode = Nothing
    Set xmlNewPropertyResp = Nothing
    Set xmlNewPropertyRoot = Nothing
    Set xmlNewPropertyElem = Nothing
    Set xmlNewPropertyItem = Nothing
    Set xmlDOMNewProperty = Nothing
    Set xmlSubQuoteResp = Nothing
    Set xmlDOMSubQuoteReq = Nothing
    Set xmlSubQuoteRoot = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlSubQuoteItem = Nothing
    Set objIApplicationDO = Nothing
    Set objINewPropertyDO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set xmlValuationTypeNode = Nothing
    Set xmlLocationNode = Nothing
    'BMIDS622 End

    Exit Function

GetLoanCompositionDetailsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetLoanCompositionDetails = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetLoanCompositionDetailsExit
    
End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>"
'   <LOANCOMPONENTDETAILS>
'      <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'        <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'        <MORTGAGESUBQUOTENUMBER>1</MORTGAGESUBQUOTENUMBER>
'        <AMOUNTREQUESTED>100000</AMOUNTREQUESTED>"
'      </LOANCOMPONENTDETAILS>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "GetDefaultsForNewLoanComponent"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_GetDefaultsForNewLoanComponent(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    GetDefaultsForNewLoanComponent = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    GetDefaultsForNewLoanComponent = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationQuoteBO_GetDefaultsForNewLoanComponent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vxmlRequest     xml Request data stream
'
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "GetDefaultsForNewLoanComponent"
    'Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseRoot As IXMLDOMElement
    'Dim xmlResponseElem As IXMLDOMElement
    'Dim xmlRequestElem As IXMLDOMElement
    Dim xmlGetDefaultsResp As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO

    On Error GoTo GetDefaultsForNewLoanComponentErr
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    ' Call the cost model assist BO
    Set xmlGetDefaultsResp = objICostModelAssistBO.GetDefaultsForNewLoanComponent(vxmlRequest)
    
    m_objErrAssist.CheckResponse xmlGetDefaultsResp, xmlResponseRoot, True
    
    m_objXMLAssist.AttachResponseData xmlResponseRoot, xmlGetDefaultsResp
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationQuoteBO_GetDefaultsForNewLoanComponent = xmlResponseRoot

GetDefaultsForNewLoanComponentExit:

    '   clean-up section, release object references

    Exit Function

GetDefaultsForNewLoanComponentErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetDefaultsForNewLoanComponent = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetDefaultsForNewLoanComponentExit
    
End Function

Public Function CreateFirstMortgageLifeSubQuotes(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates the first Mortgage Life Sub quote.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <LIFESUBQUOTES>
'       <QUOTATIONTYPE>CostModelling</QUOTATIONTYPE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   </LIFESUBQUOTES>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateFirstMortgageLifeSubQuotes"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_CreateFirstMortgageLifeSubQuotes(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateFirstMortgageLifeSubQuotes = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CreateFirstMortgageLifeSubQuotes = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Public Function UpdateMortgageSubquote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Updates Mortgage Subquote.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGESUBQUOTE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   </MORTGAGESUBQUOTE>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "UpdateMortgageSubquote"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response blocks
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_UpdateMortgageSubquote(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    UpdateMortgageSubquote = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateMortgageSubquote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit
End Function

' BMIDS977 Public Function UpdateQuotationForAddedCosts removed
'Public Function UpdateQuotationForAddedCosts(ByVal vstrXMLRequest As String) As String
'' header ----------------------------------------------------------------------------------
'' description:  Remodels the mortgage when fees are added to the loan
'' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
'' return:                       xml Response data stream containing results of operation
''                               either: TYPE="SUCCESS"
''                               or: TYPE="SYSERR" and <ERROR> element
''------------------------------------------------------------------------------------------
'On Error GoTo UpdateQuotationForAddedCostsVbErr
'    Const strFunctionName As String = "UpdateQuotationForAddedCosts"
'
'    Dim xmlIn As FreeThreadedDOMDocument40
'    Dim objIScriptInterface As IScriptInterface
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlTempResponseNode As IXMLDOMNode
'
'    ' Create default response block
'    Set xmlOut = New FreeThreadedDOMDocument40
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
'
'    ' Run validation checks
'    Set xmlTempResponseNode = IApplicationQuoteBO_Validate(xmlIn.documentElement, bomUpdateQuotationForAddedCosts)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'
'    ' Run scripting interface
'    Set objIScriptInterface = New ScriptInterface
'    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
'                                                           TypeName(Me), _
'                                                            strFunctionName, _
'                                                            xmlIn.documentElement)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'
'    ' Delegate to FreeThreadedDOMDocument40 based method
'    Set xmlTempResponseNode = IApplicationQuoteBO_UpdateQuotationForAddedCosts(xmlIn.documentElement)
'    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'    UpdateQuotationForAddedCosts = xmlResponseElem.xml
'    m_objContext.SetComplete
'
'UpdateQuotationForAddedCostsExit:
'
'    ' Clean-up section, release object references
'    Set xmlIn = Nothing
'    Set objIScriptInterface = Nothing
'    Set xmlOut = Nothing
'    Set xmlResponseElem = Nothing
'    Set xmlTempResponseNode = Nothing
'    Exit Function
'
'UpdateQuotationForAddedCostsVbErr:
'    If m_objErrAssist.IsWarning = True Then
'        m_objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'    m_objErrAssist.AddToErrSource strFunctionName
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'    UpdateQuotationForAddedCosts = m_objErrAssist.CreateErrorResponse
'    m_objContext.SetAbort
'
'    ' Go to clean-up section
'    Resume UpdateQuotationForAddedCostsExit
'End Function

Private Function IApplicationQuoteBO_CreateFirstMortgageLifeSubQuotes(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the first Mortgage Life Sub quote.
'
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <LIFESUBQUOTES>
'       <QUOTATIONTYPE>2</QUOTATIONTYPE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   </LIFESUBQUOTES>
' </REQUEST>

' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IApplicationQuoteBO_CreateFirstMortgageLifeSubQuotes"
    Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim objAppManagerBO As ApplicationManagerBO
    Dim objQuotationBO As QuotationBO
    Dim objCustFinancialDO As ICustomerFinancialDO
    'Dim xmlAppDataElem As IXMLDOMNode
    'Dim xmlAppResponseElem As IXMLDOMNode
    
    On Error GoTo CreateFirstMortgageLifeSubQuotesErr

    strNodeName = "LIFESUBQUOTES"

    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objAppManagerBO = New ApplicationManagerBO
        Set objQuotationBO = New QuotationBO
        Set objCustFinancialDO = New CustomerFinancialDO
    Else
        Set objAppManagerBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objCustFinancialDO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialDO")
    End If
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    Dim strAppNo As String
    Dim strFactFindNo As String
    Dim strQuotationNo As String
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    strQuotationNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "QUOTATIONNUMBER")

    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "FACTFINDNUMBER is empty"
    End If
    
    Dim strCustomers As String
    Dim xmlDOMCustomerResp As FreeThreadedDOMDocument40
    Dim xmlCustomerReqElem As IXMLDOMElement
    Dim xmlCustomerReqItem As IXMLDOMElement
    
    m_objXMLAssist.CheckNode vxmlRequest.ownerDocument, TypeName(Me), strFunctionName
    
    Set xmlCustomerReqElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    
    Set xmlDOMCustomerResp = New FreeThreadedDOMDocument40
    xmlDOMCustomerResp.validateOnParse = False
    xmlDOMCustomerResp.setProperty "NewParser", True
    Set xmlCustomerReqItem = xmlCustomerReqElem.appendChild(xmlDOMCustomerResp.createElement("APPLICATIONNUMBER"))
    xmlCustomerReqItem.Text = strAppNo
    
    Set xmlCustomerReqItem = xmlCustomerReqElem.appendChild(xmlDOMCustomerResp.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlCustomerReqItem.Text = strFactFindNo
    
    ' Find the customers for this application
    strCustomers = objAppManagerBO.FindCustomersForApplication(xmlCustomerReqElem.xml)
    Set xmlDOMCustomerResp = m_objXMLAssist.load(strCustomers, TypeName(Me), strFunctionName)

    m_objErrAssist.CheckResponse xmlDOMCustomerResp.documentElement, xmlResponseElem

    ' Get a list of loans which are to be ported
    ' Need some XML that contains Application number, Application FactFind number, and the
    ' Customer Version XML
    Dim xmlPortedLoansRoot As IXMLDOMElement
    Dim xmlPortedLoansElem As IXMLDOMElement
    Dim xmlPortedLoansResp As IXMLDOMElement
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomerNode As IXMLDOMNode
    
    Set xmlPortedLoansRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)

    Set xmlCustomerList = xmlDOMCustomerResp.getElementsByTagName("CUSTOMER")
    m_objXMLAssist.CheckNode xmlCustomerList, TypeName(Me), strFunctionName
    
    If xmlCustomerList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "List of customers for application is empty"
     
    End If
    
    ' Create a customer list node
    Set xmlPortedLoansElem = xmlPortedLoansRoot.appendChild(xmlDOMCustomerResp.createElement("CUSTOMERLIST"))
    
    For Each xmlCustomerNode In xmlCustomerList
        xmlPortedLoansElem.appendChild xmlCustomerNode
    Next
    
    On Error Resume Next
    Dim bPortedLoans As Boolean
    Set xmlPortedLoansResp = objCustFinancialDO.FindLoansToBePorted(xmlPortedLoansRoot)
    bPortedLoans = True
    
    m_objErrAssist.SaveErr
    On Error GoTo CreateFirstMortgageLifeSubQuotesErr
    
    ' Record not found is ok
    If m_objErrAssist.LastErr <> 0 Then
        If m_objErrAssist.LastErr <> omiga4RecordNotFound Then
            m_objErrAssist.ReRaise
        Else
            bPortedLoans = False
        End If
    End If

    ' Now Create Mortgage and Life SuqQuotes
    ' Need to convert the quotation type passed in first
    Dim sQuotationType As String
    Dim xmlDOMLifeSubQuote As FreeThreadedDOMDocument40
    Dim xmlLifeSubQuoteRoot As IXMLDOMElement
    Dim xmlLifeSubQuoteElem As IXMLDOMElement
    Dim xmlLifeSubQuoteItem As IXMLDOMElement
    
    sQuotationType = m_objXMLAssist.GetTagValue(vxmlRequest, "QUOTATIONTYPE")
    
    ' Now create the input XML
    ' Copy the request block
    Set xmlLifeSubQuoteRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlDOMLifeSubQuote = New FreeThreadedDOMDocument40
    xmlDOMLifeSubQuote.validateOnParse = False
    xmlDOMLifeSubQuote.setProperty "NewParser", True
    Set xmlLifeSubQuoteElem = xmlLifeSubQuoteRoot.appendChild(xmlDOMLifeSubQuote.createElement("QUOTATION"))
    
    ' Now quotation type
    Set xmlLifeSubQuoteItem = xmlLifeSubQuoteElem.appendChild(xmlDOMLifeSubQuote.createElement("QUOTATIONTYPE"))
    xmlLifeSubQuoteItem.Text = sQuotationType

    ' Application number
    Set xmlLifeSubQuoteItem = xmlLifeSubQuoteElem.appendChild(xmlDOMLifeSubQuote.createElement("APPLICATIONNUMBER"))
    xmlLifeSubQuoteItem.Text = strAppNo
    
    ' Application fact find number
    Set xmlLifeSubQuoteItem = xmlLifeSubQuoteElem.appendChild(xmlDOMLifeSubQuote.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlLifeSubQuoteItem.Text = strFactFindNo
    
    ' Quotation number
    Set xmlLifeSubQuoteItem = xmlLifeSubQuoteElem.appendChild(xmlDOMLifeSubQuote.createElement("QUOTATIONNUMBER"))
    xmlLifeSubQuoteItem.Text = strQuotationNo
   
    ' Now the loans to be ported

    If bPortedLoans = True Then
        xmlLifeSubQuoteElem.appendChild xmlPortedLoansResp
    End If
    
    'Dim xmlCreateSubQuoteElem As IXMLDOMElement
    Dim xmlDOMSubQuoteResp As FreeThreadedDOMDocument40
    Dim strSubQuoteResp As String
    
    strSubQuoteResp = objQuotationBO.CreateFirstMortgageLifeSubQuotes(xmlLifeSubQuoteRoot.xml)
    Set xmlDOMSubQuoteResp = m_objXMLAssist.load(strSubQuoteResp, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDOMSubQuoteResp.documentElement, xmlResponseElem, True
    
    ' The created mortgage subquite number and lifes ub quite number should be returned
    'Dim strMortgageSubQuoteNumber As String
    'Dim strLifeSubQuoteNumber As String
    
    xmlResponseElem.appendChild m_objXMLAssist.GetNode(xmlDOMSubQuoteResp.documentElement, "/RESPONSE/MORTGAGESUBQUOTE/MORTGAGESUBQUOTENUMBER")
    'xmlResponseElem.appendChild m_objXMLAssist.GetNode(xmlDOMSubQuoteResp.documentElement, "/RESPONSE/MORTGAGESUBQUOTE/LIFESUBQUOTENUMBER")
    
    ' End of function
    Set IApplicationQuoteBO_CreateFirstMortgageLifeSubQuotes = xmlResponseElem

CreateFirstMortgageLifeSubQuotesExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set objAppManagerBO = Nothing
    Set objQuotationBO = Nothing
    Set objCustFinancialDO = Nothing

    Exit Function

CreateFirstMortgageLifeSubQuotesErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CreateFirstMortgageLifeSubQuotes = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume CreateFirstMortgageLifeSubQuotesExit
    
End Function

Private Function IApplicationQuoteBO_FindMortgageProducts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IApplicationQuoteBO_FindMortgageProducts"

    
On Error GoTo FindMortgageProductsVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".CostModelAssistBO")
    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationQuoteBO_FindMortgageProducts = objICostModelAssistBO.FindMortgageProducts(vxmlRequest)

FindMortgageProductsExit:

    '   clean-up section, release object references

    Exit Function

FindMortgageProductsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_FindMortgageProducts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume FindMortgageProductsExit

End Function

Public Function FindMortgageProducts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGEPRODUCTREQUEST>
'        <SEARCHCONTEXT>TEST01</SEARCHCONTEXT>
'        <CHANNELID>1</CHANNELID>
'        <MORTGAGESUBQUOTENUMBER>1</MORTGAGESUBQUOTENUMBER>
'        <APPLICATIONNUMBER></APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER></MORTGAGESUBQUOTENUMBER>
'        <PURPOSEOFLOAN></PURPOSEOFLOAN>
'        <TERMINYEARS></TERMINYEARS>
'        <TERMINMONTHS></TERMINMONTHS>
'        <AMOUNTREQUESTED></AMOUNTREQUESTED>
'   </MORTGAGEPRODUCTREQUEST>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "FindMortgageProducts"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DOMDocument based method
    
    Set xmlTempResponseNode = IApplicationQuoteBO_FindMortgageProducts(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    FindMortgageProducts = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    FindMortgageProducts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    ' go to clean-up section
    Resume CreateExit
End Function

Public Function CalcCostModelLTV(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <LTV>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'       <AMOUNTREQUESTED>70000</AMOUNTREQUESTED>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>1234</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>1234</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LTV>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CalcCostModelLTV"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DOMDocument based method
    
    Set xmlTempResponseNode = IApplicationQuoteBO_CalcCostModelLTV(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    CalcCostModelLTV = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CalcCostModelLTV = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationQuoteBO_CalcCostModelLTV(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCalcLTVNode As IXMLDOMNode
    
Const strFunctionName As String = "IApplicationQuoteBO_CalcCostModelLTV"


On Error GoTo CalcCostModelLTVVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Delegate to the assist object
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PSC 30/06/00 SYS1012 - Start
    Set xmlCalcLTVNode = objICostModelAssistBO.CalcCostModelLTV(vxmlRequest, "COST MODELLING")
    ' PSC 30/06/00 SYS1012 - End
    m_objErrAssist.CheckResponse xmlCalcLTVNode, xmlResponseRoot, True
    m_objXMLAssist.AttachResponseData xmlResponseRoot, xmlCalcLTVNode
    
    Set IApplicationQuoteBO_CalcCostModelLTV = xmlResponseRoot
    ' All done
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalcCostModelLTVExit:

    '   clean-up section, release object references

    Exit Function

CalcCostModelLTVVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CalcCostModelLTV = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume CalcCostModelLTVExit

End Function

Public Function CalculateMortgageCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates and saves the mortgage costs
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' <REQUEST>
' <MORTGAGECOSTS>
'   <CONTEXT>Cost Modelling</CONTEXT>
'   <QUOTATIONNUMBER>1</QUOTATIONNUMBER>
'   <LIFESUBQUOTENUMBER>1</LIFESUBQUOTENUMBER>
'   <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
'   <CUSTOMERLIST>
'       <CUSTOMER>
'           <CUSTOMERNUMBER>1333</CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'       </CUSTOMER>
'       <CUSTOMER>
'           <CUSTOMERNUMBER>1325</CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'       </CUSTOMER>
'   </CUSTOMERLIST>
'   <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'   "<APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
' </MORTGAGECOSTS>
' </REQUEST>

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CalculateMortgageCosts"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_CalculateMortgageCosts(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    CalculateMortgageCosts = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CalculateMortgageCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationQuoteBO_CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
        
    
    On Error GoTo CalculateMortgageCostsErr
    Const strFunctionName As String = "IApplicationQuoteBO_CalculateMortgageCosts"

    Dim strNodeName As String
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlMortgageCostsResp As IXMLDOMNode
    Dim xmlLoanCompReq As IXMLDOMNode
    Dim xmlLoanCompReqElem As IXMLDOMNode
    Dim xmlLoanCompItem As IXMLDOMNode
    Dim xmlLoanCompResp As IXMLDOMNode
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    Dim xmlMortgageSubQuoteNumber As IXMLDOMNode
    'Dim xmlLifeSubQuoteNumber As IXMLDOMNode
    
    Dim objIApplicationDO As IApplicationDO
    Dim objApplicationBO As ApplicationBO
    Dim objIQuotationBO As IQuotationBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Dim blnCalcFeesOnly As Boolean ' PSC 10/11/2006 EP2_41

    strNodeName = "MORTGAGECOSTS"
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".CostModelAssistBO")
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteTxBO")
    End If
    
'Stage1:
    ' Get loan composition details - specific to App Quote

    Set xmlLoanCompReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDOMDoc = xmlLoanCompReq.ownerDocument
    
    Set xmlAppNo = m_objXMLAssist.GetNode(xmlRequestElem, "APPLICATIONNUMBER")
    Set xmlAppFactFindNo = m_objXMLAssist.GetNode(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    Set xmlMortgageSubQuoteNumber = m_objXMLAssist.GetNode(xmlRequestElem, "MORTGAGESUBQUOTENUMBER")
    'Set xmlLifeSubQuoteNumber = m_objXMLAssist.GetNode(xmlRequestElem, "LIFESUBQUOTENUMBER")
    Set xmlLoanCompReqElem = xmlLoanCompReq.appendChild(xmlDOMDoc.createElement("LOANCOMPOSITION"))
    
    xmlLoanCompReqElem.appendChild xmlAppNo.cloneNode(True)
    xmlLoanCompReqElem.appendChild xmlAppFactFindNo.cloneNode(True)
    xmlLoanCompReqElem.appendChild xmlMortgageSubQuoteNumber.cloneNode(True)
    
    Set xmlLoanCompResp = IApplicationQuoteBO_GetLoanCompositionDetails(xmlLoanCompReq)
    m_objErrAssist.CheckResponse xmlLoanCompResp, xmlResponseElem, True
    
    '*=[MC]BMIDS763 - APPEND APPLICATION DATE
    Call vxmlRequest.appendChild(m_objXMLAssist.GetNode(xmlLoanCompReq, "APPLICATIONDATE").cloneNode(True))
    'END SECTION
    
'Stage2:
    
    ' CalculateMortgageCosts
    
    'EP1038/MAR1893 Change xmlLifesubQuoteNumber to Nothing
    Set xmlMortgageCostsResp = objIApplicationQuoteTxBO.CalculateMortgageCosts(vxmlRequest, _
                                                                               xmlLoanCompResp, _
                                                                               xmlAppNo, _
                                                                               xmlAppFactFindNo, _
                                                                               Nothing, _
                                                                               xmlMortgageSubQuoteNumber)
    m_objErrAssist.CheckResponse xmlMortgageCostsResp, xmlResponseElem, True
    
    ' PSC 10/11/2006 EP2_41 - Start
    blnCalcFeesOnly = xmlGetNodeAsBoolean(vxmlRequest, "MORTGAGECOSTS/CALCULATEFEESONLY")
    
    If blnCalcFeesOnly Then
        xmlResponseElem.appendChild xmlGetMandatoryNode(xmlMortgageCostsResp, "ONEOFFCOSTLIST")
        xmlResponseElem.appendChild xmlGetMandatoryNode(xmlMortgageCostsResp, "MORTGAGELENDER")
    End If
    ' PSC 10/11/2006 EP2_41 - End

    Set IApplicationQuoteBO_CalculateMortgageCosts = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateMortgageCostsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlMortgageCostsResp = Nothing
    Set xmlLoanCompReq = Nothing
    Set xmlLoanCompReqElem = Nothing
    Set xmlLoanCompItem = Nothing
    Set xmlLoanCompResp = Nothing
    Set xmlDOMDoc = Nothing
    Set xmlAppNo = Nothing
    Set xmlAppFactFindNo = Nothing
    Set xmlMortgageSubQuoteNumber = Nothing
    'Set xmlLifeSubQuoteNumber = Nothing
    
    Set objIApplicationDO = Nothing
    Set objApplicationBO = Nothing
    Set objIQuotationBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set objICostModelAssistBO = Nothing
    Set objIApplicationQuoteTxBO = Nothing
    
    Exit Function

CalculateMortgageCostsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CalculateMortgageCosts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalculateMortgageCostsExit
End Function

Private Function IApplicationQuoteBO_CalculateAndSaveLifeCoverCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
'EP1038/MAR1893 No longer used
'    Dim xmlResponseRoot As IXMLDOMElement
'    Dim xmlRequestElem As IXMLDOMElement
'    Dim xmlOut As FreeThreadedDOMDocument40
'    Const strFunctionName As String = "IApplicationQuoteBO_CalculateAndSaveLifeCoverCosts"
'    Dim objICostModelAssistBO As ICostModelAssistBO
'
'On Error GoTo CalculateAndSaveLifeCoverCostsVbErr
'
'    Set xmlOut = New FreeThreadedDOMDocument40
'    xmlOut.validateOnParse = False
'    xmlOut.setProperty "NewParser", True
'
'    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
'    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
'
'    If m_objContext Is Nothing Then
'        Set objICostModelAssistBO = New CostModelAssistBO
'    Else
'        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".CostModelAssistBO")
'    End If
'
'   ' Set IApplicationQuoteBO_CalculateAndSaveLifeCoverCosts = objICostModelAssistBO.CalculateAndSaveLifeCoverCosts(vxmlRequest)
'
'    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "LIFECOVERCOSTS")
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Build the Life Cover calc input interface
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim xmlLifeCoverRespNode As IXMLDOMNode
'    'Dim xmlLifeCoverReqNode As IXMLDOMNode
'
'    Set xmlLifeCoverRespNode = IApplicationQuoteBO_BuildLifeCoverCalcInputInterface(vxmlRequest)
'    m_objErrAssist.CheckResponse xmlLifeCoverRespNode, xmlResponseRoot, True
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Calculate the premiums, then save the life cover details
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim xmlLifeCoverReq As IXMLDOMNode
'    Dim xmlLifeCoverResp As IXMLDOMNode
'    Dim xmlLifeCoverElem As IXMLDOMNode
'    Dim xmlAppNo As IXMLDOMNode
'    Dim xmlAppFactFindNo As IXMLDOMNode
'    Dim xmlSubQuote As IXMLDOMNode
'
'    Set xmlLifeCoverReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
'
'    Set xmlLifeCoverElem = m_objXMLAssist.GetNode(xmlLifeCoverRespNode, "LIFECOVER").cloneNode(True)
'    Set xmlAppNo = m_objXMLAssist.GetNode(xmlRequestElem, "APPLICATIONNUMBER")
'    Set xmlAppFactFindNo = m_objXMLAssist.GetNode(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
'    Set xmlSubQuote = m_objXMLAssist.GetNode(xmlRequestElem, "LIFESUBQUOTENUMBER")
'
'    xmlLifeCoverElem.appendChild xmlAppNo.cloneNode(True)
'    xmlLifeCoverElem.appendChild xmlAppFactFindNo.cloneNode(True)
'    xmlLifeCoverElem.appendChild xmlSubQuote.cloneNode(True)
'
'
'    ' Need to modify the customer XML
'    Dim xmlCustomerList As IXMLDOMNodeList
'    Dim xmlCustomer As IXMLDOMNode
'    Dim xmlItem As IXMLDOMNode
'    Dim intCustomerNumber As Integer
'
'    Set xmlCustomerList = xmlRequestElem.getElementsByTagName("CUSTOMER")
'    intCustomerNumber = 1
'
'    For Each xmlCustomer In xmlCustomerList
'        Set xmlItem = m_objXMLAssist.GetNode(xmlCustomer, "CUSTOMERNUMBER")
'        m_objXMLAssist.ChangeNodeName xmlItem, "CUSTOMERNUMBER", "CUSTOMERNUMBER" & intCustomerNumber
'
'        Set xmlItem = m_objXMLAssist.GetNode(xmlCustomer, "CUSTOMERVERSIONNUMBER")
'        m_objXMLAssist.ChangeNodeName xmlItem, "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER" & intCustomerNumber
'        intCustomerNumber = intCustomerNumber + 1
'    Next
'    xmlLifeCoverElem.appendChild m_objXMLAssist.GetNode(xmlRequestElem, "CUSTOMERLIST")
'    xmlLifeCoverReq.appendChild xmlLifeCoverElem.cloneNode(True)
'
'    Set xmlLifeCoverResp = objICostModelAssistBO.CalcAndSaveLifeCover(xmlLifeCoverReq)
'
'    ' DM 27-06-00 Added return response checking to this method
'    m_objErrAssist.CheckResponse xmlLifeCoverResp, xmlResponseRoot, True
'
'    ' Add the response to our repsonse, removing "RESPONSE"
'    m_objXMLAssist.AttachResponseData xmlResponseRoot, xmlLifeCoverResp
'
'    Set IApplicationQuoteBO_CalculateAndSaveLifeCoverCosts = xmlResponseRoot
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'CalculateAndSaveLifeCoverCostsExit:
'
'    '   clean-up section, release object references
'
'    Exit Function
'
'CalculateAndSaveLifeCoverCostsVbErr:
'
'    If m_objErrAssist.IsWarning = True Then
'        m_objErrAssist.AddWarning xmlResponseRoot
'        Resume Next
'    End If
'
'    m_objErrAssist.AddToErrSource strFunctionName
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    Set IApplicationQuoteBO_CalculateAndSaveLifeCoverCosts = m_objErrAssist.CreateErrorResponseEx
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetAbort
'    End If
'
'    ' Append
'    ' go to clean-up section
'    Resume CalculateAndSaveLifeCoverCostsExit

End Function
'EP1038/MAR1893 End

Public Function CalculateAndSaveLifeCoverCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates and saves the life cover costs
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' <REQUEST>
'   <LIFECOVERCOSTS>
'      <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'      <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'      <LIFESUBQUOTENUMBER>1</LIFESUBQUOTENUMBER>
'      <CUSTOMERLIST>
'         <CUSTOMER>
'            <CUSTOMERNUMBER>1333</CUSTOMERNUMBER>
'            <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'         </CUSTOMER>
'         <CUSTOMER>
'            <CUSTOMERNUMBER>1325</CUSTOMERNUMBER>
'            <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'         </CUSTOMER>
'      </CUSTOMERLIST>
'      <LOANCOMPONENTLIST>
'         <LOANCOMPONENT>
'         .......
'         <LOANCOMPONENT>
'      <LOANCOMPONENTLIST>
'   <LIFECOVERCOSTS>
' <REQUEST>
On Error GoTo CalculateAndSaveLifeCoverCostsVbErr

    Const strFunctionName As String = "CalculateAndSaveLifeCoverCosts"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DOMDocument based method
    
    'Set xmlTempResponseNode = IApplicationQuoteBO_CalculateAndSaveLifeCoverCosts(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXMLAssist.WriteXMLToFile xmlResponseElem.xml
    CalculateAndSaveLifeCoverCosts = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateAndSaveLifeCoverCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalculateAndSaveLifeCoverCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateAndSaveLifeCoverCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CalculateAndSaveLifeCoverCostsExit
End Function

Public Function BuildLifeCoverCalcInputInterface(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculate the Term and Endowment Needs and Terms excluding non-ported loan
'               components
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>"
'   <LIFECOVER>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'       <LOANCOMPONENTLIST>
'           <LOANCOMPONENT>
'           </LOANCOMPONENT>
'       </LOANCOMPONENTLIST>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER></CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LIFECOVER>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "BuildLifeCoverCalcInputInterface"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_BuildLifeCoverCalcInputInterface(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    BuildLifeCoverCalcInputInterface = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    BuildLifeCoverCalcInputInterface = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationQuoteBO_BuildLifeCoverCalcInputInterface(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' description:  Calculate the Term and Endowment Needs and Terms excluding non-ported loan
'               components
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>"
'   <LIFECOVERCOSTS>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'       <LOANCOMPONENTLIST>
'           <LOANCOMPONENT>
'           </LOANCOMPONENT>
'       </LOANCOMPONENTLIST>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER></CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LIFECOVERCOSTS>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
    Const strFunctionName As String = "IApplicationQuoteBO_BuildLifeCoverCalcInputInterface"
    'Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseItem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlGetDefaultsResp As IXMLDOMNode
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objICustomerDO As ICustomerDO
    Dim objIApplicationDO As IApplicationDO
    Dim objIComboDO As ComboDO
    
    On Error GoTo BuildLifeCoverCalcInputInterfaceErr
        
    'strNodeName = "LIFECOVERCOSTS"
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    Set xmlResponseElem = xmlResponseRoot.appendChild(xmlOut.createElement("LIFECOVER"))
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "LIFECOVERCOSTS")
    
    Set xmlAppNo = m_objXMLAssist.GetNode(xmlRequestElem, "APPLICATIONNUMBER")
    Set xmlAppFactFindNo = m_objXMLAssist.GetNode(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objICustomerDO = New CustomerDO
        Set objIApplicationDO = New ApplicationDO
        Set objIComboDO = New ComboDO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objICustomerDO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerDO")
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' Call the cost model assist BO
    Set xmlGetDefaultsResp = objICostModelAssistBO.BuildLifeCoverCalcInputInterface(vxmlRequest)
    m_objErrAssist.CheckResponse xmlGetDefaultsResp, xmlResponseRoot, True
    
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlGetDefaultsResp
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get Customer Details - Build up a list
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlCustomerResp As IXMLDOMNode
    'Dim xmlCustomerReq As IXMLDOMElement
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomerRespList As IXMLDOMElement
    'Dim xmlCustomerRespNode As IXMLDOMNode
    Dim xmlCustomerDoc As FreeThreadedDOMDocument40
    
    Set xmlCustomerDoc = New FreeThreadedDOMDocument40
    xmlCustomerDoc.validateOnParse = False
    xmlCustomerDoc.setProperty "NewParser", True
    Set xmlCustomerList = xmlRequestElem.getElementsByTagName("CUSTOMER")
    
    Set xmlCustomerRespList = xmlCustomerDoc.createElement("CUSTOMERLIST")
    
    For Each xmlCustomerNode In xmlCustomerList
        Set xmlCustomerResp = objICustomerDO.GetPersonalDetails(xmlCustomerNode)
        xmlCustomerRespList.appendChild xmlCustomerResp
    Next
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Determine the relationship of customer 1 and customer 2, if more than
    ' one customer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Dim intNoCustomers As Integer
    Dim strLifeCoverRelationship As String
    
    Set xmlCustomerList = xmlCustomerRespList.getElementsByTagName("CUSTOMER")
    
    ' More than one customer?
    If xmlCustomerList.length > 1 Then
        Dim xmlCustomerRelationshipReq As IXMLDOMNode
        Dim xmlCustomerRelationshipElem As IXMLDOMElement
        Dim xmlCustomerRelationshipResp As IXMLDOMNode
        Dim strRelationship As String
        
        Set xmlCustomerRelationshipReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlDOMDoc = xmlCustomerRelationshipReq.ownerDocument
        
        Set xmlCustomerRelationshipElem = xmlCustomerRelationshipReq.appendChild(xmlDOMDoc.createElement("CUSTOMERRELATIONSHIP"))
        xmlCustomerRelationshipElem.appendChild xmlAppNo.cloneNode(True)
        xmlCustomerRelationshipElem.appendChild xmlAppFactFindNo.cloneNode(True)
        
        ' APS 31/05/00 - SYS0796 Handle record not found and assume a single relationship
        On Error Resume Next
        Set xmlCustomerRelationshipResp = objIApplicationDO.GetCustomerRelationship(xmlCustomerRelationshipReq)
        m_objErrAssist.SaveErr
        On Error GoTo BuildLifeCoverCalcInputInterfaceErr
        
        If (m_objErrAssist.LastErr <> 0) Then
            If (m_objErrAssist.LastErr <> omiga4RecordNotFound) Then
                m_objErrAssist.ReRaise
            End If
        Else
            strRelationship = m_objXMLAssist.GetNodeValue(xmlCustomerRelationshipResp, "RELATIONSHIPTOOWNER")
            strRelationship = objIComboDO.GetFirstComboValidation("ApplicantRelationship", strRelationship)
        End If
        
        If strRelationship = cstrHusband Or strRelationship = cstrWife Or strRelationship = cstrUnmarried Then
            strLifeCoverRelationship = "Joint"
        Else
            strLifeCoverRelationship = "Single"
        End If
    Else
        strLifeCoverRelationship = "Single"
    End If
    
    Set xmlResponseItem = xmlResponseElem.appendChild(xmlOut.createElement("RELATIONSHIP"))
    xmlResponseItem.Text = strLifeCoverRelationship
    
    ' Call CostModelAssist to build the xml from the customer version XML
    Dim xmlLifeCoverElem As IXMLDOMElement
    Set xmlLifeCoverElem = objICostModelAssistBO.BuildLifeCoverInterfaceXML(xmlCustomerRespList)
    m_objErrAssist.CheckResponse xmlLifeCoverElem, xmlResponseRoot, True

    ' All done
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlLifeCoverElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
        
    
    Set IApplicationQuoteBO_BuildLifeCoverCalcInputInterface = xmlResponseRoot

BuildLifeCoverCalcInputInterfaceExit:

    '   clean-up section, release object references
    Set objIApplicationDO = Nothing

    Exit Function

BuildLifeCoverCalcInputInterfaceErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_BuildLifeCoverCalcInputInterface = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume BuildLifeCoverCalcInputInterfaceExit
    
End Function

Private Function IApplicationQuoteBO_FindStoredQuoteDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindStoredQuoteDetailsVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_FindStoredQuoteDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
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
                
    Set xmlTempResponseNode = objICostModelAssistBO.FindStoredQuoteDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
    Set IApplicationQuoteBO_FindStoredQuoteDetails = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindStoredQuoteDetailsExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindStoredQuoteDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_FindStoredQuoteDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit

End Function

Public Function FindStoredQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindStoredQuoteDetailsVbErr

    Const strFunctionName As String = "FindStoredQuoteDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DOMDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationQuoteBO_FindStoredQuoteDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    FindStoredQuoteDetails = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindStoredQuoteDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindStoredQuoteDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindStoredQuoteDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit

End Function

Public Function ProcessAddedOneOffCosts(ByVal vstrXMLRequest As String) As String
'--------------------------------------------------------------------------------------
' description:  Process one off costs
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>"
'      <ADDEDONEOFFCOSTS>
'        <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
'        <CUSTOMERLIST>
'            <CUSTOMER>
'                <CUSTOMERNUMBER>1333</CUSTOMERNUMBER>
'                <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'            </CUSTOMER>
'            <CUSTOMER>
'                <CUSTOMERNUMBER>1325</CUSTOMERNUMBER>
'                <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'            </CUSTOMER>
'        </CUSTOMERLIST>
'        <ONEOFFCOSTLIST>
'            <ONEOFFCOST>
'                <COMBOVALIDATIONTYPE>MIG</COMBOVALIDATIONTYPE>
'                <AMOUNT>1000</AMOUNT>
'                <ADDEDTOLOAN>1</ADDEDTOLOAN>
'            </ONEOFFCOST>
'            <ONEOFFCOST>
'                <COMBOVALIDATIONTYPE>TID</COMBOVALIDATIONTYPE>
'                <AMOUNT>500</AMOUNT>
'            </ONEOFFCOST>
'    </ADDEDONEOFFCOSTS>
' </REQUEST>
' return:       xml Response data stream
'               either: TYPE="SUCCESS"
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "ProcessAddedOneOffCosts"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_ProcessAddedOneOffCosts(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ProcessAddedOneOffCosts = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    ProcessAddedOneOffCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationQuoteBO_ProcessAddedOneOffCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' description:  Process one off costs
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>
'      <ADDEDONEOFFCOSTS>
'        <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
'        <CUSTOMERLIST>
'            <CUSTOMER>
'                <CUSTOMERNUMBER>1333</CUSTOMERNUMBER>
'                <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'            </CUSTOMER>
'            <CUSTOMER>
'                <CUSTOMERNUMBER>1325</CUSTOMERNUMBER>
'                <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'            </CUSTOMER>
'        </CUSTOMERLIST>
'        <ONEOFFCOSTLIST>
'            <ONEOFFCOST>
'                <COMBOVALIDATIONTYPE>MIG</COMBOVALIDATIONTYPE>
'                <AMOUNT>1000</AMOUNT>
'                <ADDEDTOLOAN>1</ADDEDTOLOAN>
'            </ONEOFFCOST>
'            <ONEOFFCOST>
'                <COMBOVALIDATIONTYPE>TID</COMBOVALIDATIONTYPE>
'                <AMOUNT>500</AMOUNT>
'            </ONEOFFCOST>
'    </ADDEDONEOFFCOSTS>
' </REQUEST>
' return:       xml Response data stream
'               either: TYPE="SUCCESS"
'               or: TYPE="SYSERR" and <ERROR> element
'-------------------------------------------------------------------------------------
On Error GoTo ProcessAddedOneOffCostsErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_ProcessAddedOneOffCosts"

    
    Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    
    Dim xmlAccountList As IXMLDOMElement
    Dim xmlAccountDom As FreeThreadedDOMDocument40
    Dim xmlLoanList As IXMLDOMNode
    Dim xmlLoansNotRedeemedNode As IXMLDOMNode
    
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objICostModelAssistDO As ICostModelAssistDO
    Dim objICustomerFinancialBO As ICustomerFinancialBO
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    
    ' Validate the node passed in
    strNodeName = "ADDEDONEOFFCOSTS"
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objICustomerFinancialBO = New CustomerFinancialBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objICostModelAssistDO = m_objContext.CreateInstance(App.Title & ".CostModelAssistDO")
        Set objICustomerFinancialBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get New Property Details
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strLocation As String, strValuationType As String, strSharedPercentage  As String
    Set xmlTempResponseNode = objICostModelAssistDO.GetValuationTypeAndLocation(vxmlRequest)
    
    strLocation = m_objXMLAssist.GetTagValue(xmlTempResponseNode, "PROPERTYLOCATION")
    strValuationType = m_objXMLAssist.GetTagValue(xmlTempResponseNode, "VALUATIONTYPE")
    strSharedPercentage = m_objXMLAssist.GetTagValue(xmlTempResponseNode, "SHAREDPERCENTAGE")
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get Loans not redeemed
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PSC 30/06/00 SYS1012 - Start
    Dim strLoansNotRedeemed As String
    Set xmlLoansNotRedeemedNode = objICostModelAssistBO.GetLoansNotRedeemed(vxmlRequest)
    m_objErrAssist.CheckResponse xmlLoansNotRedeemedNode, xmlResponseRoot
    strLoansNotRedeemed = m_objXMLAssist.GetNodeValue(xmlLoansNotRedeemedNode, _
                                                      ".//LOANSNOTREDEEMED", _
                                                      True, False)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Process the one off costs
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set xmlTempResponseNode = _
            objICostModelAssistBO.ProcessAddedOneOffCosts(vxmlRequest, _
                                                          "COSTMODELLING", _
                                                          strLocation, _
                                                          strValuationType, _
                                                          strLoansNotRedeemed, _
                                                          strSharedPercentage)
    ' PSC 30/06/00 SYS1012 - End

    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseRoot, True
    m_objXMLAssist.AttachResponseData xmlResponseRoot, xmlTempResponseNode
    
    ' All done
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationQuoteBO_ProcessAddedOneOffCosts = xmlResponseRoot

ProcessAddedOneOffCostsExit:
    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlDOMDoc = Nothing
    Set xmlResponseRoot = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNewRequest = Nothing
    
    Set xmlAccountList = Nothing
    Set xmlAccountDom = Nothing
    Set xmlLoanList = Nothing
    
    Set objICostModelAssistBO = Nothing
    Set objICostModelAssistDO = Nothing
    Set objICustomerFinancialBO = Nothing
    Set xmlLoansNotRedeemedNode = Nothing
    
    Exit Function
    
ProcessAddedOneOffCostsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_ProcessAddedOneOffCosts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ProcessAddedOneOffCostsExit
    
End Function

Public Function RecostMortgageComponentAndLifeCover( _
                        ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'     <APPLICATIONQUOTE>
'        <APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER>
'        <LIFESUBQUOTENUMBER>
'        <CONTEXT>
'        <INTERESTONLYAMOUNT>
'        <CAPITALANDINTERESTAMOUNT>
'        <APPLICATIONDATE>
'        <ONEOFFCOSTLIST>
'            ...
'        </ONEOFFCOSTLIST>
'        <CUSTOMERLIST>
'            <CUSTOMER>
'               <CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>
'            </CUSTOMER>
'            <CUSTOMER>
'               <CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>
'            </CUSTOMER>
'        </CUSTOMERLIST>
'        <LOANCOMPONENT>
'            <LOANCOMPONENTSEQUENCENUMBER>
'        </LOANCOMPONENT>
'        <LOANCOMPONENTLIST>
'            ...
'        </LOANCOMPONENTLIST>
'    </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo RecostMortgageComponentAndLifeCoverVbErr

    Const strFunctionName As String = "RecostMortgageComponentAndLifeCover"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = _
            IApplicationQuoteBO_RecostMortgageComponentAndLifeCover(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RecostMortgageComponentAndLifeCover = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RecostMortgageComponentAndLifeCoverExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

RecostMortgageComponentAndLifeCoverVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    RecostMortgageComponentAndLifeCover = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit
    
End Function

Public Function RecommendQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a new quotation record as active quotation
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <CONTEXT>..</CONTEXT>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'   </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo RecommendQuotationVbErr

    Const strFunctionName As String = "RecommendQuotation"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_RecommendQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RecommendQuotation = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RecommendQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    RecommendQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecommendQuotationExit
    
End Function

Public Function CreateNewQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a new quotation record as active quotation
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <ACTIVEQUOTENUMBER>..</ACTIVEQUOTENUMBER>
'   </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewQuotationVbErr

    Const strFunctionName As String = "CreateNewQuotation"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_CreateNewQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateNewQuotation = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateNewQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit
    
End Function

Private Function IApplicationQuoteBO_CreateNewQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateNewQuotationVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_CreateNewQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem

    If m_objContext Is Nothing Then
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.CreateNewQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_CreateNewQuotation = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewQuotationExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_CreateNewQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit

End Function

Public Function ReinstateQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: This reinstates a stored quotation but, after checking the quotation is still
'              valid
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <QUOTATIONNUMBER>..</QUOTATIONNUMBER>
'   </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ReinstateQuotationVbErr

    Const strFunctionName As String = "ReinstateQuotation"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_ReinstateQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ReinstateQuotation = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ReinstateQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
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

    ReinstateQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateQuotationExit
    
End Function

Public Function ResetMortgageSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: This method resets an un-calculated mortgageSubQuote.
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <MORTGAGESUBQUOTENUMBER>..</MORTGAGESUBQUOTENUMBER>
'       <AMOUNTREQUESTED>..</AMOUNTREQUESTED>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>..</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>..</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       <CUSTOMERLIST>
'   </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "ResetMortgageSubQuote"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_ResetMortgageSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ResetMortgageSubQuote = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ResetMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit
    
End Function

Public Function CalculateAndSaveBCSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: This method resets an un-calculated mortgageSubQuote.
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>..</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>..</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       <CUSTOMERLIST>
'       <BUILDINGANDCONTENTS>
'           .................
'       </BUILDINGANDCONTENTS>
'   </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAndSaveBCSubQuoteVbErr

    Const strFunctionName As String = "CalculateAndSaveBCSubQuote"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_CalculateAndSaveBCSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalculateAndSaveBCSubQuote = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateAndSaveBCSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    CalculateAndSaveBCSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit
    
End Function

Public Function GetHighestEarners(ByVal vstrXMLRequest As String) As String
' header ------------------------------------------------------------------------------
' description:  This determines which applicants are to be used in 'Maximum Borriwing'
'               calculation
'SR - 25/04/00
'** FIX ME  --- This method is copied from ApplicationBO.GetHighestEarners, and some changes are made.
'            ApplicationBO.GetHighestEarners is called from 'AppQuoteBO.CalculateMaxBorrowings'.
'           Change that reference to AppQuoteBO.GetHighestEarners and remove the one in ApplicationBO
'
'Request Format :
' <REQUEST>
'   <CUSTOMERLIST>
'       <CUSTOMER>
'           <CUSTOMERNUMBER>...</CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER>...</CUSTOMERVERSIONNUMBER>
'           <CUSTOMERROLETYPE>...</CUSTOMERROLETYPE>
'        </CUSTOMER>
'    </CUSTOMERLIST>
' </REQUEST>
'--------------------------------------------------------------------------------------
On Error GoTo GetHighestEarnersVbErr

    Const strFunctionName As String = "GetHighestEarners"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DOMDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationQuoteBO_GetHighestEarners(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetHighestEarners = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetHighestEarnersExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetHighestEarnersVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetHighestEarners = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetHighestEarnersExit

End Function

Public Function GetAcceptedQuoteData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get mortgagesubquote data for the accepted quote
' pass:         <REQUEST>
'                <APPLICATION>
'                   <APPLICATIONNUMBER> </APPLICATIONNUMBER>
'                   <APPLICATIONFACTFINDNUMBER> </APPLICATIONFACTFINDNUMBER>
'                </APPLICATION>
'               </REQUEST>
'
' return:
'<RESPONSE TYPE="SUCCESS">
'</RESPONSE>
'
'history JLD 18/12/01 Created Method
'------------------------------------------------------------------------------------------
On Error GoTo GetAcceptedQuoteDataVbErr

    Const strFunctionName As String = "GetAcceptedQuoteData"
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DOMDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationQuoteBO_GetAcceptedQuoteData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetAcceptedQuoteData = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetAcceptedQuoteDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetAcceptedQuoteDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetAcceptedQuoteData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetAcceptedQuoteDataExit

End Function
Public Function CreateAdHocFeeAndRemodel(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create the fee in the correct table (one off costs and/or application fee type)
'               Then call remodel quotation.
'
On Error GoTo CreateAdHocFeeAndRemodelVbErr

    Const strFunctionName As String = "CreateAdHocFeeAndRemodel"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DOMDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationQuoteBO_CreateAdHocFeeAndRemodel(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateAdHocFeeAndRemodel = xmlResponseElem.xml
    
    m_objContext.SetComplete

CreateAdHocFeeAndRemodelExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
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

    CreateAdHocFeeAndRemodel = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume CreateAdHocFeeAndRemodelExit

End Function


Public Function GetApplicationSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get summary data for Appliaction
' pass:         <REQUEST>
'                <APPLICATIONDETAILS>
'                   <APPLICATIONNUMBER> </APPLICATIONNUMBER>
'                   <APPLICATIONFACTFINDNUMBER> </APPLICATIONFACTFINDNUMBER>
'                   <CUSTOMERNUMBER> </CUSTOMERNUMBER>
'                   <CUSTOMERVERSIONNUMBER> </CUSTOMERVERSIONNUMBER>
'                   </APPLICATIONDETAILS>
'               </REQUEST>
'
' return:
'<RESPONSE TYPE="SUCCESS">
'   <CUSTOMERNAME>
'       <TITLE> </TITLE>
'       <FIRSTFORENAME> </FIRSTFORENAME>
'       <SURNAME> </SURNAME>
'   </CUSTOMERNAME>
'   <PROPERTYADDRESS>
'       <FLATNUMBER> </FLATNUMBER>
'       <BUILDINGORHOUSENAME> </BUILDINGORHOUSENAME>
'       <BUILDINGORHOUSENUMBER> </BUILDINGORHOUSENUMBER>
'       <STREET> </STREET>
'       <DISTRICT> </DISTRICT>
'       <TOWN> </TOWN>
'       <COUNTY> </COUNTY>
'       <POSTCODE> </POSTCODE>
'    </PROPERTYADDRESS>
'    <CUSTOMERCONTACTDETAILS>
'       <TELEPHONENUMBER> </TELEPHONENUMBER>
'       <USAGE> </USAGE>
'    </CUSTOMERCONTACTDETAILS>
'    <TYPEOFAPPLICATIONDETAILS>
'        <TYPEOFAPPLICATION> </TYPEOFAPPLICATION>
'    </TYPEOFAPPLICATIONDETAILS>
'    <BUSINESSTYPE>
'        <DIRECTINDIRECTBUSINESS> </DIRECTINDIRECTBUSINESS>
'    </BUSINESSTYPE>
'    <BUSINESSSOURCE>
'        <NAME> </NAME>
'        <UNITNAME> </UNITNAME>
'    </BUSINESSSOURCE>
'    <BUSINESSSOURCECONTACTDETAILS>
'        <TELEPHONENUMBER> </TELEPHONENUMBER>
'        <TELEPHONEEXTENSIONNUMBER> </TELEPHONEEXTENSIONNUMBER>
'    </BUSINESSSOURCECONTACTDETAILS>
'    <PRODUCTANDLOANLIST>
'       <PRODUCTANDLOANDETAILS>
'           <LOANAMOUNT> </LOANAMOUNT>
'           <TERMINYEARS> </TERMINYEARS>
'           <TERMINMONTHS> </TERMINMONTHS>
'           <REPAYMENTMETHOD> </REPAYMENTMETHOD>
'           <GROSSMONTHLYCOST> </GROSSMONTHLYCOST>
'           <MORTGAGEPRODUCTDESCRIPTION>  </MORTGAGEPRODUCTDESCRIPTION>
'           <INTERESTRATE> </INTERESTRATE>
'           <MONTHLYREPAYMENT> </MONTHLYREPAYMENT>
'           <PRODUCTNAME> </PRODUCTNAME>
'           <WITHDRAWNDATE> </WITHDRAWNDATE>
'           <RATECHARGED> </RATECHARGED>
'       </PRODUCTANDLOANDETAILS>
'    </PRODUCTANDLOANLIST>
'    <LEGALREPS>
'        <COMPANYNAME> </COMPANYNAME>
'        <DXID> </DXID>
'        <DXLOCATION> </DXLOCATION>
'        <CONTACTFORENAME> </CONTACTFORENAME>
'        <CONTACTSURNAME> </CONTACTSURNAME>
'        <TELEPHONENUMBER> </TELEPHONENUMBER>
'    </LEGALREPS>
'    <HOMEINSQUOTE>
'        <TOTALBCMONTHLYCOST> </TOTALBCMONTHLYCOST>
'    </HOMEINSQUOTE>
'    <PAYMENTPROTECTIONQUOTE>
'        <TOTALPPMONTHLYCOST> </TOTALPPMONTHLYCOST>
'    </PAYMENTPROTECTIONQUOTE>
'    <APPLICATIONSTAGE>
'        <STAGENAME> </STAGENAME>
'    </APPLICATIONSTAGE>
'    <APPLICANTDETAILSLIST>
'       <APPLICANTDETAILS>
'           <TITLE> </TITLE>
'           <FIRSTFORENAME> </FIRSTFORENAME>
'           <SECONDFORENAME> </SECONDFORENAME>
'           <SURNAME> </SURNAME>
'        </APPLICANTDETAILS>
'    <APPLICANTDETAILSLIST>
'</RESPONSE>
'
'history BG 15/01/01 Created Method
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationSummaryVbErr

    Const strFunctionName As String = "GetApplicationSummary"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DOMDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationQuoteBO_GetApplicationSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetApplicationSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetApplicationSummaryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetApplicationSummaryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetApplicationSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetApplicationSummaryExit

End Function

Public Function RemodelQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Remodels the repayment costs and APR to ensure that the latest rates
'               have been applied prior to issuing the offer
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo RemodelQuotationVbErr

    Const strFunctionName As String = "RemodelQuotation"
    
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IApplicationQuoteBO_Validate(xmlIn.documentElement, bomRemodelQuotation)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_RemodelQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RemodelQuotation = xmlResponseElem.xml
    
    m_objContext.SetComplete

RemodelQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
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

    RemodelQuotation = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RemodelQuotationExit
    
End Function

Private Function IApplicationQuoteBO_RemodelQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Remodels the repayment costs and APR to ensure that the latest rates
'               have been applied prior to issuing the offer
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo RemodelQuotationVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_RemodelQuotation"

    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.RemodelQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_RemodelQuotation = xmlResponseElem

    m_objContext.SetComplete
    
RemodelQuotationExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
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

    Set IApplicationQuoteBO_RemodelQuotation = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RemodelQuotationExit

End Function

Private Function IApplicationQuoteBO_Validate(ByVal vxmlRequest As IXMLDOMElement, _
                                    ByVal veboMethod As IApplicationQuoteBOMethod) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IApplicationQuoteBO_Validate"

        
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    m_objContext.SetComplete
    
    Set IApplicationQuoteBO_Validate = xmlResponseElem
    
ValidateExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    
    Exit Function

ValidateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Public Function RecostQuotationForRateChange(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Remodels the repayment costs and APR to ensure that the latest rates
'               have been applied prior to issuing the offer
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'Request for GetOneOffCosts:
'                   <REQUEST>
'                       <ONEOFFCOSTS>
'                           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                           <MORTGAGESUBQUOTENUMBER></MORTGAGESUBQUOTENUMBER
'                       </ONEOFFCOSTS>
'                   </REQUEST>
'Request for GetLoanCompositionDetails:
'                   <REQUEST>
'                       <LOANCOMPOSITION>
'                           <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                           <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                           <MORTGAGESUBQUOTENUMBER></MORTGAGESUBQUOTENUMBER
'                       </LOANCOMPOSITION>
'                   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo RecostQuotationForRateChangeVbErr

    Const strFunctionName As String = "RecostQuotationForRateChange"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IApplicationQuoteBO_RecostQuotationForRateChange(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RecostQuotationForRateChange = xmlResponseElem.xml
    
    m_objContext.SetComplete

RecostQuotationForRateChangeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

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

    RecostQuotationForRateChange = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RecostQuotationForRateChangeExit
    
End Function

Public Function ValidateManualAdjustmentPercent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks the user and product parameters to validate the adjustment.
' pass:         vstrXmlRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateManualAdjustmentVbErr

    Const strFunctionName As String = "ValidateManualAdjustmentPercent"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_ValidateManualAdjustmentPercent(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ValidateManualAdjustmentPercent = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateManualAdjustmentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateManualAdjustmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidateManualAdjustmentPercent = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateManualAdjustmentExit
    
End Function

Public Function ValidateUserMandateLevel(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo ValidateUserMandateLevelVbErr

    Const strFunctionName As String = "ValidateUserMandateLevel"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IApplicationQuoteBO_ValidateUserMandateLevel(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ValidateUserMandateLevel = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

validateUserMandateLevelExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateUserMandateLevelVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidateUserMandateLevel = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume validateUserMandateLevelExit
    
End Function

Public Function GetAcceptedOrActiveQuoteData(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Get mortgagesubquote data for the accepted quote
' pass:         <REQUEST>
'                <APPLICATION>
'                   <APPLICATIONNUMBER> </APPLICATIONNUMBER>
'                   <APPLICATIONFACTFINDNUMBER> </APPLICATIONFACTFINDNUMBER>
'                </APPLICATION>
'               </REQUEST>
'


'
'history GD 08/07/02 Created Method for BMIDS00165
'------------------------------------------------------------------------------------------
On Error GoTo GetAcceptedOrActiveQuoteDataVbErr

    Const strFunctionName As String = "GetAcceptedOrActiveQuoteData"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IApplicationQuoteBO_GetAcceptedOrActiveQuoteData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetAcceptedOrActiveQuoteData = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetAcceptedOrActiveQuoteDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetAcceptedOrActiveQuoteDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetAcceptedOrActiveQuoteData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetAcceptedOrActiveQuoteDataExit

End Function

Private Function IApplicationQuoteBO_GetAcceptedOrActiveQuoteData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Gets the mortgagesubquote data for the accepted or active quote.
'
'
' pass:         vstrXMLRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "IApplicationQuoteBO_GetAcceptedOrActiveQuoteData"
    On Error GoTo GetAcceptedOrActiveQuoteDataErr
        
    Dim strNodeName As String
    Dim strAppNo As String
    Dim strFactFindNo As String
    Dim strAcceptedQuoteNumber As String
    Dim strActiveQuoteNumber As String
    Dim strQuoteNumber As String
    Dim strMSQuoteNumber As String
    
    Dim xmlAppFactFindNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlDomAppData As FreeThreadedDOMDocument40
    Dim xmlAppDataItem As IXMLDOMElement
    Dim xmlSubQuoteResp As IXMLDOMElement
    Dim xmlDOMSubQuoteReq As FreeThreadedDOMDocument40
    Dim xmlSubQuoteRoot As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlSubQuoteItem As IXMLDOMElement
    Dim xmlMSQNode As IXMLDOMNode
    Dim xmlMSubQuoteResp As IXMLDOMElement
    Dim xmlDOMMSubQuoteReq As FreeThreadedDOMDocument40
    Dim xmlMSubQuoteRoot As IXMLDOMElement
    Dim xmlMSubQuoteElem As IXMLDOMElement
    Dim xmlMSubQuoteItem As IXMLDOMElement
    
    Dim objIApplicationDO As IApplicationDO
    Dim objIQuotationBO As IQuotationBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    
    Dim blnThrowError As Boolean    'BM0206 MDC 19/12/2002
    
    strNodeName = "APPLICATION"
    
'Stage1:

    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objIQuotationBO = New QuotationBO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objIQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    End If
    
    ' Validate the node passed in
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXMLAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    strAppNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXMLAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "FACTFINDNUMBER is empty"
    End If
        
'Stage2:

    ' Need to build the request to pass to GetApplicationData
    Set xmlAppDataElem = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
    
    m_objXMLAssist.CheckNode xmlAppDataElem, TypeName(Me), strFunctionName
    
    Set xmlDomAppData = New FreeThreadedDOMDocument40
    xmlDomAppData.validateOnParse = False
    xmlDomAppData.setProperty "NewParser", True
    Set xmlDomAppData.documentElement = xmlAppDataElem
    Set xmlAppDataElem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONFACTFIND"))
    
    m_objXMLAssist.CheckNode xmlAppDataElem, TypeName(Me), strFunctionName
    
    ' Add the application number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONNUMBER"))
    m_objXMLAssist.CheckNode xmlAppDataItem, TypeName(Me), strFunctionName
    xmlAppDataItem.Text = strAppNo
    
    ' Add the application fact find number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlAppDataItem.Text = strFactFindNo
    
    ' Get the ApplicationData
    Set xmlAppResponseElem = objIApplicationDO.GetApplicationFactFindData(xmlDomAppData.documentElement)
    
    Set xmlAppFactFindNode = m_objXMLAssist.GetNode(xmlAppResponseElem, "//APPLICATIONFACTFIND")
    strAcceptedQuoteNumber = m_objXMLAssist.GetTagValue(xmlAppFactFindNode, "ACCEPTEDQUOTENUMBER")
    strActiveQuoteNumber = m_objXMLAssist.GetTagValue(xmlAppFactFindNode, "ACTIVEQUOTENUMBER")
    
'Stage3:

    'BM0206 MDC 19/12/2002 - Add option to throw error if quote not found
    blnThrowError = CSafeBool(m_objXMLAssist.GetAttributeValue(vxmlRequest, "REQUEST", "THROWERROR"))
    If blnThrowError And Len(strAcceptedQuoteNumber) = 0 And Len(strActiveQuoteNumber) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Accepted, and Active Quote not found"
    End If
    'BM0206 MDC 19/12/2002 - End
    
    'BM0145 - MV - 06/12/2002
    If Len(strAcceptedQuoteNumber) > 0 Or Len(strActiveQuoteNumber) > 0 Then
    
        'If (Len(strAcceptedQuoteNumber) = 0 And Len(strActiveQuoteNumber) = 0) Then
        '    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Accepted, and Active Quote not found"
        'End If
        
        If Len(strAcceptedQuoteNumber) <> 0 Then
            strQuoteNumber = strAcceptedQuoteNumber
        Else
            strQuoteNumber = strActiveQuoteNumber
        End If
    
'Stage4:

        ' Get quotation Details
        Set xmlSubQuoteRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
        Set xmlDOMSubQuoteReq = xmlSubQuoteRoot.ownerDocument
        
        Set xmlSubQuoteElem = xmlSubQuoteRoot.appendChild(xmlDOMSubQuoteReq.createElement("QUOTATION"))
        
        'Application number
        Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONNUMBER"))
        xmlSubQuoteItem.Text = strAppNo
            
        'Application fact find number
        Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONFACTFINDNUMBER"))
        xmlSubQuoteItem.Text = strFactFindNo
            
        'Quote number
        Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("QUOTATIONNUMBER"))
        xmlSubQuoteItem.Text = strQuoteNumber
            
        Set xmlSubQuoteResp = objIQuotationBO.GetData(xmlSubQuoteRoot)
        m_objErrAssist.CheckResponse xmlSubQuoteResp, xmlResponseElem, True
        
        'MAR7 GHun
        m_objXMLAssist.AttachResponseData xmlResponseElem, xmlSubQuoteResp
        'MAR7 End
        
        'Get the mortgagesubquote number of the accepted quote
        Set xmlMSQNode = m_objXMLAssist.GetNode(xmlSubQuoteResp, "//QUOTATION")
        strMSQuoteNumber = m_objXMLAssist.GetTagValue(xmlMSQNode, "MORTGAGESUBQUOTENUMBER")
        'BMIDS00654 MDC 04/11/2002
        If Len(Trim$(strMSQuoteNumber)) = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 239, "Accepted, and Active Quote not found"
        End If
        'BMIDS00654 MDC 04/11/2002 - End
    
'Stage5:

        'Get the Mortgagesubquote data
        Set xmlMSubQuoteRoot = m_objXMLAssist.GetRequestNode(vxmlRequest.ownerDocument)
        Set xmlDOMMSubQuoteReq = xmlMSubQuoteRoot.ownerDocument
        
        Set xmlMSubQuoteElem = xmlMSubQuoteRoot.appendChild(xmlDOMSubQuoteReq.createElement("MORTGAGESUBQUOTE"))
        
        ' Now add the application number
        Set xmlMSubQuoteItem = xmlMSubQuoteElem.appendChild(xmlDOMMSubQuoteReq.createElement("APPLICATIONNUMBER"))
        xmlMSubQuoteItem.Text = strAppNo
            
        '... and application fact find number
        Set xmlMSubQuoteItem = xmlMSubQuoteElem.appendChild(xmlDOMMSubQuoteReq.createElement("APPLICATIONFACTFINDNUMBER"))
        xmlMSubQuoteItem.Text = strFactFindNo
            
        ' mortgagesubquote number
        Set xmlMSubQuoteItem = xmlMSubQuoteElem.appendChild(xmlDOMMSubQuoteReq.createElement("MORTGAGESUBQUOTENUMBER"))
        xmlMSubQuoteItem.Text = strMSQuoteNumber
            
        Set xmlMSubQuoteResp = objIMortgageSubQuoteBO.GetMortgageSubQuoteData(xmlMSubQuoteRoot)
        m_objErrAssist.CheckResponse xmlMSubQuoteResp, xmlResponseElem, True
        
        ' Now build the XML to be returned from this method.
        ' Need to return Mortgage sub quote
        m_objXMLAssist.AttachResponseData xmlResponseElem, xmlMSubQuoteResp
    
    End If
    
    Set IApplicationQuoteBO_GetAcceptedOrActiveQuoteData = xmlResponseElem

GetAcceptedOrActiveQuoteDataExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set objIApplicationDO = Nothing
    Set objIQuotationBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    Set xmlDomAppData = Nothing
    Set xmlAppDataItem = Nothing
    Set xmlSubQuoteResp = Nothing
    Set xmlDOMSubQuoteReq = Nothing
    Set xmlSubQuoteRoot = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlSubQuoteItem = Nothing
    Set xmlMSQNode = Nothing
    Set xmlMSubQuoteResp = Nothing
    Set xmlDOMMSubQuoteReq = Nothing
    Set xmlMSubQuoteRoot = Nothing
    Set xmlMSubQuoteElem = Nothing
    Set xmlMSubQuoteItem = Nothing

    Exit Function

GetAcceptedOrActiveQuoteDataErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetAcceptedOrActiveQuoteData = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetAcceptedOrActiveQuoteDataExit
    
End Function

Public Function RefreshMortgageAccountData(ByVal vstrXMLRequest As String) As String

On Error GoTo RefreshMortgageAccountDataVbErr

    Const strFunctionName As String = "RefreshMortgageAccountData"
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                              
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_RefreshMortgageAccountData(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RefreshMortgageAccountData = xmlTempResponseNode.xml
    
    m_objContext.SetComplete

RefreshMortgageAccountDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

RefreshMortgageAccountDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    RefreshMortgageAccountData = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RefreshMortgageAccountDataExit
    
End Function

Private Function IApplicationQuoteBO_RefreshMortgageAccountData( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RefreshMortgageAccountDataVbErr

    Const strFunctionName As String = "IApplicationQuoteBO_RefreshMortgageAccountData"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.RefreshMortgageAccountData(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IApplicationQuoteBO_RefreshMortgageAccountData = xmlTempResponseNode
    
RefreshMortgageAccountDataExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIApplicationQuoteTxBO = Nothing
    
    Exit Function

RefreshMortgageAccountDataVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_RefreshMortgageAccountData = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RefreshMortgageAccountDataExit
    
End Function

Public Function GetAndSavePortedStepAndPeriodFromMortgageAccount(ByVal vstrXMLRequest As String) As String

On Error GoTo GetAndSavePortedStepAndPeriodFromMortgageAccountVbErr

    Const strFunctionName As String = "GetAndSavePortedStepAndPeriodFromMortgageAccount"
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                              
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_GetAndSavePortedStepAndPeriodFromMortgageAccount(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    GetAndSavePortedStepAndPeriodFromMortgageAccount = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetAndSavePortedStepAndPeriodFromMortgageAccountExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    GetAndSavePortedStepAndPeriodFromMortgageAccount = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume GetAndSavePortedStepAndPeriodFromMortgageAccountExit
    
End Function

Private Function IApplicationQuoteBO_GetAndSavePortedStepAndPeriodFromMortgageAccount( _
                                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode


On Error GoTo GetAndSavePortedStepAndPeriodFromMortgageAccountVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_GetAndSavePortedStepAndPeriodFromMortgageAccount"


    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIApplicationQuoteTxBO As IApplicationQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIApplicationQuoteTxBO = New ApplicationQuoteTxBO
    Else
        Set objIApplicationQuoteTxBO = m_objContext.CreateInstance(App.Title & ".ApplicationQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIApplicationQuoteTxBO.GetAndSavePortedStepAndPeriodFromMortgageAccount(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_GetAndSavePortedStepAndPeriodFromMortgageAccount = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetAndSavePortedStepAndPeriodFromMortgageAccountExit:

    Set objIApplicationQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
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

    Set IApplicationQuoteBO_GetAndSavePortedStepAndPeriodFromMortgageAccount = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetAndSavePortedStepAndPeriodFromMortgageAccountExit
End Function

' PSC 05/12/2006 EP2_249 - Start
Public Function GetIntroducerProcFeeDetails(ByVal vstrXMLRequest As String) As String

On Error GoTo GetIntroducerProcFeeDetailsVbErr

    Const strFunctionName As String = "GetIntroducerProcFeeDetails"
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                              
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = IApplicationQuoteBO_GetIntroducerProcFeeDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetIntroducerProcFeeDetails = xmlResponseElem.xml
    
GetIntroducerProcFeeDetailsExit:

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

GetIntroducerProcFeeDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetIntroducerProcFeeDetails = m_objErrAssist.CreateErrorResponse

    ' go to clean-up section
    Resume GetIntroducerProcFeeDetailsExit
    
End Function

Private Function IApplicationQuoteBO_GetIntroducerProcFeeDetails( _
                                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode


On Error GoTo GetIntroducerProcFeeDetailsVbErr
    
    Const strFunctionName As String = "IApplicationQuoteBO_GetIntroducerProcFeeDetails"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICMAssist As ICostModelAssistBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICMAssist = New CostModelAssistBO
    Else
        Set objICMAssist = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    Set xmlTempResponseNode = objICMAssist.GetIntroducerProcFeeDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IApplicationQuoteBO_GetIntroducerProcFeeDetails = xmlTempResponseNode

    
GetIntroducerProcFeeDetailsExit:

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set objICMAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

GetIntroducerProcFeeDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationQuoteBO_GetIntroducerProcFeeDetails = _
                                        m_objErrAssist.CreateErrorResponseEx()
    
    ' go to clean-up section
    Resume GetIntroducerProcFeeDetailsExit
End Function
' PSC 05/12/2006 EP2_249 - End


