Attribute VB_Name = "KFIDataHelper"
'********************************************************************************
'** Module:         KFIDataHelper
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    This module contains functions/procedures that are common to
'**                 all/most KFI document types.
'********************************************************************************
'********************************************************************************
'BBG Specific History
'
'Prog   Date        Description
'AW     18/01/2005  E2EM00001894 - Added IsButToLet function.Amended GetFixedRateTerm().
'MV     22/02/2005  BBG1932 - Amended GetFeeDetailDescription()
'TK     17/03/2005  BBG2037 - Amended BuildWhatFeesYouMustPaySection()
'********************************************************************************
'Mars Specific History
'
'Prog   Date        Description
'BC     14/09/2005  MAR54 and MAR88 - Project MARS
'TW     09/11/2005  MAR442  - A number of changes to allow omCK.CreateKFI to work
'                             Note that these changes prevent errors where certain data is missing
'                             See the specific methods/functions for the changes
'BC     16/11/2005  MAR589 - KFI Changes
'BC     06/12/2005  MAR774 - Incorrect Fee Description
'BC     15/12/2005  MAR893 - Fee Description for Fixed Rate Products
'BC     10/01/2066  MAR907 - Product Switches
'BC     02/03/2006  MAR1347  Use first Validation Type for RepaymentMethod, rather than the first
'BC     03/03/2006  MAR1347  Create either RATEPERIOD or RATEENDDATE in AddRatePeriodAttribute
'BC     06/03/2006  MAR1367  Set CHARGEAMOUNT to LoanAmount from LoanComponent table not TotalLoanAmount
'PE     14/03/2006  MAR1061 The estimated purchase price is now set using the mortgage sub quote value, not application fact find
'BC     01/05/2006  MAR1685  Inhibit TERMINMONTHS attribute if vobjCommon.TermInMonths = zero
'********************************************************************************
'Epsom Specific History
'
'Prog   Date        Description
'SAB    21/04/2006  EP454   Created new function AddIntermediaryContactDetails
'SAB    03/05/2006  EP479   Created new Function IsSection3MemoPad
'SAB    03/05/2006  EP479   Amended AddProductNameAttribute to use the PRODUCTTEXTDETAILS column rather than PRODUCTNAME
'SAB    03/05/2006  EP479   Amended BuildWhatFeesYouMustPaySection to allow for 3 separate Fee Descriptions when calling GetFeeDetailDescription
'SAB    03/05/2006  EP479   Amended BuildWhatFeesYouMustPaySection to check whether the solictor is PANEL
'SAB    03/05/2006  EP479   Amended GetFeeDetailDescription to allow for 3 separate Fee Descriptions fields (needad as 3 bulletpoints required on Offer Doc)
'SAB    03/05/2006  EP479   Amended GetFeeDetailDescription for 2 new Fee Types
'IK     10/05/2006  EP531   add MAR1721 amendments
'PB     12/05/2006  EP529   Merged MAR1716 - Calculation of MaxCharge in wrong place
'PB     12/05/2006  EP529   Merged MAR1720 - If FinalPayment > 0 reduce number of monthly payments by 1
'PB     12/05/2006  EP529   Merged MAR1720 - Ensure that PART from multicomponent loan is numbererd corr3ectly
'AW     15/05/2006  EP520   CC16 New fee description (Third party valuation)
'PB     16/05/2006  EP529   Merged MAR1731 - Fulfilment - Offer Document - oncorrect validity period shown
'PM     18/05/2006  EP584   Rewritten AddIntermediaryContactDetails
'PM     18/05/2006  EP584   Amended AddLegalRepContactDetails to add solicitor name
'PM     18/05/2006  EP584   Amended AddMortgageTermAttributes. We require TermInMonths without additional <MONTHS> sub-element
'PM     19/05/2006  EP584   Amended BuildUsingIntermediarySection so that IntermediaryName is the broker
'PM     19/05/2006  EP584   Amended BuildInsuranceSection
'AW     22/05/2006  EP590   Amendments for Section 8 (Fees yo will pay) and Section 13 (Intermediary fees)
'PB     24/05/2006  EP603   MAR1777 Changed 'DateDiff("m"' to 'MonthDiff(' and added CalcExpectedCompletionDate
'PB     24/05/2006  EP603   MAR1736 Show Base Rate Set description in text describing changes in rates (UAT3101)
'                           NOTE: the following was not added because it does not apply to this code:
'                           MAR1736 Estimated Legal Fees were being added to wrong node (UAT3102)
'DRC    25/05/2006  EP610   Epsom specific changes as noted
'PM     25/05/2006  EP610   Amended AddAddressAttributes
'PB     25/05/2006  EP603   MAR1788 Set APR attribute from mdblAPR (MortgageSubQuote) rather than LoanComponent
'PM     26/05/2006  EP628   Amended BuildMortgageNoMoreSection - added attribute MAXCHARGE & REDADMINAMOUNT for
'                           section 10
'AW     01/06/06    EP520   Do not create BROKERREFUND element for zero refund
'PB     05/06/2006  EP651   MAR1590 KFI Part and Part printing issue
'PM     05/06/2006  EP652   Amended AddApplicantNameAttribute to allow for title other
'PM     06/06/2006  EP610   Amended GetFeeDetailDescription to change wording
'PB     07/06/2006  EP696   MAR1826 No 'NOFEES' element appears when there are costs but all are zero
'PB/PM  12/06/2006  EP697   Offer Document Changes
'PB     12/06/2006  EP730   MAR1831 Change description of 'Repayment' payment method
'PM     19/06/2006  EP697   Amended AddUntilDateOrForPeriodElements to add PREFPERIOD along with TERMYEARS and TERMMONTHS
'PE     20/06/2006  EP773   Modify AddAddressAttributes and BuildContactNames to return results on a single line.
'PB     20/06/2006  EP700   Wrong address appearing on offer document - should be same as valuation address
'DRC    21/06/2006  EP801   Change to GetFeeDetailDescription for Redemption Admin Fee
'PM     22/06/2006  EP826   Change BuildWhatFeesYouMustPaySection
'PB     22/06/2006  EP821   ACC185 - Wording for valuation fee in the offer is incorrect. Check lender if adding fees is ok.
'PB     22/06/2006  EP827   ACC193 - Summary: Broker fee wording incorrect when refundable and non refundable.
'                           The rebate is pulling through to the wrong section of the offer.
'                           Other wording also corrected.
'PB     22/06/2006  EP841   ACC208 - Arrangement Fee for the fixed rate wording is incorrect.
'PM     26/06/2006  EP845   Amended IsFirstTimeBuyer to test for 'F' not 'S'
'PB     26/06/2006  EP773   Get address etc from NEWPROPERTY if no valuation details
'PM     27/06/2006  EP832   Updated BuildAdditionalFeaturesSection to comment out overpayments processing
'PE     27/06/2006  EP842   Update decription for SRF cost type. (Telegraphic Transfer Fee - Retention/Stage Advance)
'PB     27/06/2006  EP827   Modifications to broker fee in section 8
'PB     27/06/2006  EP841   Modifications to arrangement fee in section 8
'PB     (as above)          Minor change to cope with negative payments
'PB     29/06/2006  EP845   Ensure that BTL mortgages are not shown as first-time buyer on offer doc
'PB     29/06/2006  EP846   Ensure that LTB mortgages are not shown as first-time buyer on offer doc
'PE     29/06/2006  EP895   Set the intermediary name to broker, packager or "your broker"
'PB     29/06/2006  EP700   Correct faulty XPATH causing valuation address to be missed
'PE     29/06/2006  EP796   Removed repayment fee from charge amount
'PB     30/06/2006  EP700   Test for postcode within valuation address; if blank, use the new property address
'PB     13/07/2006  EP821   Make sure "Third Party Valuation Fee" shows as just "Valuation Fee" in offer doc
'DRC    13/07/2006  EP966   Don't put in Intermediary address if there's no intermediary
'PB     17/07/2006  EP821   Extend previous fix to check for "0.00" broker fee as well as ""
'DRC    17/08/2006  EP892   Concatenate Broker & Packager Name for section13
'PB     18/08/2006  EP1082  Show ContactTitleOther for legal rep on offer doc
'PB     22/08/2006  EP1082  (As above) Ensure other title is shown correctly for all third parties in offer doc
'PE     25/08/2006  EP892   Change BuildUsingIntermediarySection to retrieve intermediary and packager names as discrete values.
'PE     01/09/2006  EP1100  CC72 - Phase 1 Offer - Section 8 - Retype Valuation Fee - Change text
'PE     04/09/2006  EP1113  CC118 - Valuation moved from ‘fees payable to db mortgages’ section to ‘other fees’ section.
'PE     05/09/2006  EP1114  Nominal completion date incorrect in the offer document.
'PE     05/09/2006  EP1120  DBM132 - Amendment to: KFI Offer document - section 4
'PE     07/09/2006  EP1100  Valuation type now retrieved from valuerinstruction instead of newproperty.
'PE     25/09/2006  EP1166  Modified Set2DP to correctly format numbers.
'AW     04/10/2006  EP1185  Amended ReType validation match to "RT"
'PSC    09/11/2006  EP2_41  Amend IsFirstTimeBuyer to only check loancomponent if there is an active quote
'INR    30/11/2006  EP2_139 Offer and KFI changes.
'PB     01/12/2006  EP2_139 GetLegalFeeText().
'INR    30/11/2006  EP2_422 Section13 changes.
'PB     03/01/2007  EP2_648 KFI error when trying to view - needed else clause for empty MORTGAGEINCENTIVE element
'GHun   26/01/2006  EP2_974 Merged EP1212, EP1226 and EP1238
'INR    30/01/2007  EP2_718 changed Legal Fees from  template coded default
'INR    07/02/2007  EP2_583  KFI changes
'INR    20/02/2007  EP2_704 THIRDPARTIES should be INTERMEDIARYNAMES
'INR    28/02/2007  EP2_1449 add the MORTGAGETYPE to REQUIREDINSURANCE
'INR    01/03/2007  EP2_1742 should be xmlSingleComponent
'INR    02/03/2007  EP2_1382 add the OTHERFEES element for AdditionalBrokerFees
'PB     05/03/2007  EP2_1629 Section numbers in the text were incorrect for some cases.
'PB     07/03/2007  EP2_1861 Added REFUNDOFVALUATION
'PB     09/03/2007  EP2_1627 Utilised gblnSection2 flag
'PB     15/03/2007  EP2_1931 Valuation fee should only appear in 'Other fees' if it is for a third party valuation
'INR    19/03/2007  EP2_1883 Check for PuposeOfLoan
'PB     20/03/2007  EP2_1861 Correct valuation fee text
'PB     20/03/2007  EP2_1931 No redemption fees for ABO
'SR     23/03/2007  EP2_1938 modified 'AddRepaymentMethodAttribute' and 'GetInterestRates'
'PB     23/03/2007  EP2_1937 Corrected section numbering
'INR    23/03/2007  EP2_1980 Sort section8 Fee Detail Descriptions
'INR    23/03/2007  EP2_1753 DBM186 Mods
'PB     26/03/2007  EP2_1937 Re-arranged section 10 as some text should only appear once
'PB     27/03/2007  EP2_1932 Added free-valuation details to section 10
'INR    27/03/2007  EP2_2057 Need  Contact Telephone details
'INR    28/03/2007  EP2_1860 Get the Broker name
'PB     30/03/2007  EP2_2192 Bug in BuildMortgageNoMoreSection causiung error with portable cases
'INR    01/04/2007  EP2_1351 Use description from combo for TPV
'DS    03/04/2007   EP2_1931 Put null check before adding attributes to xmlPortable and vxmlNode
'INR    04/04/2007  EP2_2203 Won't always be viewing/printing the active quote
'INR    05/04/2007  EP2_1915 Changes to displaying Legal costs
'INR    06/04/2007  EP2_1931 Rework section 10
'SR     11/04/2007  EP2_2270 modified 'AddIntermediaryContactDetails'
'SR     12/04/2007  EP2_2341
'GHun   13/04/2007  EP2_1930 Changed GetFeeDetailDescription
'GHun   13/04/2007  EP2_1930 Further change to GetFeeDetailDescription
'INR    15/04/2007  EP2_2395 AddInterestRateTypeAttributeSect6
'INR    15/04/2007  EP2_2395  Section10 change
'INR    18/04/2007  EP2_2448 Only add dblRedAdminAmount if single component
'GHun   18/04/2007  EP2_2477 Changed GetFeeDetailDescription and BuildWhatFeesYouMustPaySection
'INR    20/04/2007  EP2_2478 New check for additional borrowing and/or TOE and modify cashbacks
'INR    20/04/2007  EP2_2534 Wording needs to take into account refund and rebate for TPV/TPVA
'********************************************************************************
Option Explicit

Private Const mcstrModuleName As String = "KFIDataHelper"

Private Type InterestRateDataType
    intCumulativePeriods As Integer
    strIntRateType As String
End Type


'********************************************************************************
'** Function:       AddReferenceAttribute
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Adds the REFERENCE attribute to the node.
'** Parameters:     vxmlData - the source data XML.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddReferenceAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddReferenceAttribute"
    Dim strValue As String
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler
    
    '*-get the value for and add the REFERENCE attribute
    strValue = xmlGetAttributeText(vxmlData, "APPLICATIONNUMBER")
    
    Set xmlItem = vxmlData.selectSingleNode(gcstrQUOTATION_PATH)
    If Not xmlItem Is Nothing Then
        strValue = strValue & "(" & xmlGetAttributeText(xmlItem, "QUOTATIONNUMBER") & ")" 'SR 27/09/2004 : CORE82
    End If
    
    Call xmlSetAttributeValue(vxmlNode, "REFERENCE", strValue)
    Set xmlItem = Nothing

Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddApplicantNameAttribute
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Adds the APPLICANTNAME attribute to the node.
'** Parameters:     vxmlData - the source data XML.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddApplicantNameAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddApplicantNameAttribute"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode, xmlCustomerVersion As IXMLDOMNode 'SR 30/09/2004 : BBG1509
    Dim strName As String
'PM 05/06/2006 EPSOM EP652   Start
    Dim strTitle As String
'PM 05/06/2006 EPSOM EP652   End
    Dim strValue As String, strCustomerRoleType As String 'SR 30/09/2004 : BBG1509
    Dim strValidationToSearchFor As String

    On Error GoTo ErrHandler
    
    '====================================================================================
    'SR 30/09/2004 : BBG1509 - get the value(s) for and add the APPLICANTNAME attribute
    '                if any of the applicant is 'COMPANY' then display company only
    '                in other cases, do not add Guarantors to the list
    '=====================================================================================
    'Check whether a company is mentioned in customer list
    Set xmlList = vxmlData.selectNodes(gcstrCUSTOMERROLE_PATH & "[contains(@CUSTOMERROLETYPE,'C,')]")
    strValidationToSearchFor = IIf(xmlList.length > 0, "C,", "A,")
    
    Set xmlList = vxmlData.selectNodes(gcstrCUSTOMERROLE_PATH)
    For Each xmlItem In xmlList
        strCustomerRoleType = xmlGetAttributeText(xmlItem, "CUSTOMERROLETYPE")
        If InStr(UCase$(strCustomerRoleType), strValidationToSearchFor) > 0 Then
            Set xmlCustomerVersion = xmlItem.selectSingleNode("./CUSTOMERVERSION")
'SR 13/12/2004 : BBG1858 -
'            strName = BuildName(xmlGetAttributeText(xmlCustomerVersion, "TITLE_TEXT"), _
'                xmlGetAttributeText(xmlCustomerVersion, "FIRSTFORENAME"), _
'                xmlGetAttributeText(xmlCustomerVersion, "SURNAME"))

'PM 05/06/2006 EPSOM EP652   Start
            strTitle = xmlGetAttributeText(xmlCustomerVersion, "TITLEOTHER")
            If strTitle = "" Then
                strTitle = xmlGetAttributeText(xmlCustomerVersion, "TITLE_TEXT")
            End If
            strName = BuildName(strTitle, _
                                GetAllForeNames(xmlCustomerVersion), _
                                xmlGetAttributeText(xmlCustomerVersion, "SURNAME"))
'            strName = BuildName(xmlGetAttributeText(xmlCustomerVersion, "TITLE_TEXT"), _
'                                GetAllForeNames(xmlCustomerVersion), _
'                                xmlGetAttributeText(xmlCustomerVersion, "SURNAME"))
'SR 13/12/2004 : BBG1858 - End
'PM 05/06/2006 EPSOM EP652   End
            If Len(strValue) > 0 Then
                strValue = strValue & " and "
            End If
            strValue = strValue & strName
        End If
    Next xmlItem
    'SR 30/09/2004 : BBG1509 - End
    Call xmlSetAttributeValue(vxmlNode, "APPLICANTNAME", strValue)
    Set xmlList = Nothing
    Set xmlItem = Nothing

Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'SR 13/12/2004 : BBG1858 - Get all the forenames as single string
Public Function GetAllForeNames(ByVal vxmlCustomerVersion As IXMLDOMNode) As String

    Const cstrFunctionName As String = "GetAllForeNames"
    Dim strAllForeNames As String, strTemp As String
    
    On Error GoTo ErrHandler
    strTemp = xmlGetAttributeText(vxmlCustomerVersion, "FIRSTFORENAME")
    strAllForeNames = strTemp
    
    strTemp = xmlGetAttributeText(vxmlCustomerVersion, "SECONDFORENAME")
    If Len(strTemp) > 0 Then
        strAllForeNames = strAllForeNames & " " & strTemp
    End If
    
    strTemp = xmlGetAttributeText(vxmlCustomerVersion, "OTHERFORENAMES")
    If Len(strTemp) > 0 Then
        strAllForeNames = strAllForeNames & " " & strTemp
    End If
       
    GetAllForeNames = strAllForeNames
    
    Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function
'SR 13/12/2004 : BBG1858 - End

'********************************************************************************
'** Function:       AddApplicantAddressAttributes
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Adds the APPLICANTADDRESSLINE1, APPLICANTADDRESSLINE2,
'**                 APPLICANTADDRESSLINE3, APPLICANTPOSTCODE attributes to the
'**                 node.
'** Parameters:     vxmlData - the source data XML.
'**                 vxmlNode - the node to add the attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddApplicantAddressAttributes(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddApplicantAddressAttributes"
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlAddress As IXMLDOMNode
    Dim xmlPropertyLocation As IXMLDOMNode 'PB EP529 / MAR1731
    Dim xmlList As IXMLDOMNodeList
      
    Dim strPropertyLocation As String 'PB EP529 / MAR1731
        
    On Error GoTo ErrHandler
    
    '*-get the value(s) for and add the address attributes
    'SR 15/12/2004 : BBG1813 - Add Address of company, if one exists in the XML
    Set xmlList = vxmlData.selectNodes(gcstrCUSTOMERROLE_PATH & "[contains(@CUSTOMERROLETYPE,'C,')]")
    If xmlList.length > 0 Then
        Set xmlCustomer = xmlList.Item(0)
    Else
        Set xmlCustomer = vxmlData.selectSingleNode(gcstrCUSTOMERROLE_PATH & "[@CUSTOMERORDER=1]")
    End If
    'SR 15/12/2004 : BBG1813 - End
    '*-get the address with type 2 first, if one
    Set xmlAddress = xmlCustomer.selectSingleNode("CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE=2]/ADDRESS")
    If xmlAddress Is Nothing Then
        '*-there isn't so try to get the address with type 1
        Set xmlAddress = xmlCustomer.selectSingleNode("CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE=1]/ADDRESS")
    End If
    
    If Not xmlAddress Is Nothing Then
 'PM 25/05/2006 EPSOM EP610   Start
        Call AddAddressAttributes(xmlAddress, vxmlNode, "APPLICANTADDRESS")
'        Call AddAddressAttributes(xmlAddress, vxmlNode, "APPLICANTADDRESSLINE1", "APPLICANTADDRESSLINE2", "APPLICANTADDRESSLINE3", "APPLICANTPOSTCODE")
'PM 25/05/2006 EPSOM EP610   End
    End If
    
    ' PB 16/05/2006 EP529 / MAR1731
    Set xmlPropertyLocation = vxmlData.selectSingleNode(gcstrPROPERTYLOCATION_PATH)
    strPropertyLocation = xmlGetAttributeText(xmlPropertyLocation, "PROPERTYLOCATION_TEXT")
    
    Call xmlSetAttributeValue(vxmlNode, "PROPERTYLOCATION", strPropertyLocation)
    ' EP529 / MAR1731 End
    
    Set xmlCustomer = Nothing
    Set xmlAddress = Nothing
    Set xmlList = Nothing
    Set xmlPropertyLocation = Nothing 'PB EP529 / MAR1731

Exit Sub
ErrHandler:
    Set xmlCustomer = Nothing
    Set xmlAddress = Nothing
    Set xmlList = Nothing
    Set xmlPropertyLocation = Nothing 'PB EP529 / MAR1731
    
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'PM 25/05/2006 EPSOM EP610   Start
'********************************************************************************
'** Function:       AddAddressAttributes
'** Created by:     Andy Maggs
'** Date:           24/05/2004
'** Description:    Adds the address attributes from the node to the specified
'**                 node.
'** Parameters:     vxmlAddress - the node containing the address data.
'**                 vxmlNode - the address node.
'**                 vstrAddressAttribName - the first address line attribute name.
'**                 vblnFlat - return the address on a single line.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddAddressAttributes(ByVal vxmlAddress As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, _
        Optional ByVal vstrAddressAttribName As String = "ADDRESS", _
        Optional ByVal vblnFlat As Boolean = False)
    Const cstrFunctionName As String = "AddAddressAttributes"
    
    Dim strAddress2 As String
    Dim strAddress3 As String
    Dim strPostcode As String
    Dim strAddress As String
    
    On Error GoTo ErrHandler
        
    If vxmlAddress Is Nothing Then
        '*-there is nothing we can do, so quit
        Exit Sub
    End If
    
    strAddress = BuildAddress1(xmlGetAttributeText(vxmlAddress, "FLATNUMBER"), _
            xmlGetAttributeText(vxmlAddress, "BUILDINGORHOUSENAME"), _
            xmlGetAttributeText(vxmlAddress, "BUILDINGORHOUSENUMBER"), _
            xmlGetAttributeText(vxmlAddress, "STREET"))
    
    strAddress2 = BuildAddress2(xmlGetAttributeText(vxmlAddress, "DISTRICT"), _
            xmlGetAttributeText(vxmlAddress, "TOWN"))
    
    strAddress3 = BuildAddress3(xmlGetAttributeText(vxmlAddress, "COUNTY"), _
            xmlGetAttributeText(vxmlAddress, "COUNTRY_TEXT"))
    
    strPostcode = xmlGetAttributeText(vxmlAddress, "POSTCODE")
    
    If Not vblnFlat Then
        If Not strAddress2 = "" Then strAddress = IIf(strAddress = "", strAddress2, strAddress + "\par " + strAddress2)
        If Not strAddress3 = "" Then strAddress = IIf(strAddress = "", strAddress3, strAddress + "\par " + strAddress3)
        If Not strPostcode = "" Then strAddress = IIf(strAddress = "", strPostcode, strAddress + "\par " + strPostcode)
    Else
        If Not strAddress2 = "" Then strAddress = IIf(strAddress = "", strAddress2, strAddress + ", " + strAddress2)
        If Not strAddress3 = "" Then strAddress = IIf(strAddress = "", strAddress3, strAddress + ", " + strAddress3)
        If Not strPostcode = "" Then strAddress = IIf(strAddress = "", strPostcode, strAddress + ", " + strPostcode)
    End If
    
    Call xmlSetAttributeValue(vxmlNode, vstrAddressAttribName, strAddress)
    
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'PM 25/05/2006 EPSOM EP610   End
'EP2_2057
Public Sub AddIntermedAddressAttributes(ByVal vxmlAddress As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, _
        Optional ByVal vstrAddressAttribName As String = "ADDRESS", _
        Optional ByVal vblnTelephone As Boolean = False)
    Const cstrFunctionName As String = "AddIntermedAddressAttributes"
    
    Dim strAddress As String
    Dim strAddress2 As String
    Dim strAddress3 As String
    Dim strAddress4 As String
    Dim strAddress5 As String
    Dim strAddress6 As String
    Dim strPostcode As String
    Dim strAreaCode As String
    Dim strTelephone As String
    
    On Error GoTo ErrHandler
        
    If vxmlAddress Is Nothing Then
        '*-there is nothing we can do, so quit
        Exit Sub
    End If
    
    strAddress = xmlGetAttributeText(vxmlAddress, "ADDRESSLINE1")
    strAddress2 = xmlGetAttributeText(vxmlAddress, "ADDRESSLINE2")
    strAddress3 = xmlGetAttributeText(vxmlAddress, "ADDRESSLINE3")
    strAddress4 = xmlGetAttributeText(vxmlAddress, "ADDRESSLINE4")
    strAddress5 = xmlGetAttributeText(vxmlAddress, "ADDRESSLINE5")
    strAddress6 = xmlGetAttributeText(vxmlAddress, "ADDRESSLINE6")
    strPostcode = xmlGetAttributeText(vxmlAddress, "POSTCODE")
    strAreaCode = xmlGetAttributeText(vxmlAddress, "TELEPHONEAREACODE")
    strTelephone = xmlGetAttributeText(vxmlAddress, "TELEPHONENUMBER")
   
    If Not strAddress2 = "" Then strAddress = IIf(strAddress = "", strAddress2, strAddress + "\par " + strAddress2)
    If Not strAddress3 = "" Then strAddress = IIf(strAddress = "", strAddress3, strAddress + "\par " + strAddress3)
    If Not strAddress4 = "" Then strAddress = IIf(strAddress = "", strAddress4, strAddress + "\par " + strAddress4)
    If Not strAddress5 = "" Then strAddress = IIf(strAddress = "", strAddress5, strAddress + "\par " + strAddress5)
    If Not strAddress6 = "" Then strAddress = IIf(strAddress = "", strAddress6, strAddress + "\par " + strAddress6)
    If Not strPostcode = "" Then strAddress = IIf(strAddress = "", strPostcode, strAddress + "\par " + strPostcode)
    If vblnTelephone Then
        If Not strAreaCode = "" Then strAddress = IIf(strAddress = "", strAreaCode, strAddress + "\par " + strAreaCode)
        If Not strTelephone = "" Then strAddress = IIf(strAddress = "", strTelephone, strAddress + " " + strTelephone)
    End If
    
    Call xmlSetAttributeValue(vxmlNode, vstrAddressAttribName, strAddress)
    
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'********************************************************************************
'** Function:       AddValidToDateAttribute
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Adds the VALIDTODATE attribute to the node.
'** Parameters:     vxmlData - the source data XML.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddValidToDateAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddValidToDateAttribute"
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
    Dim lngDays As Long

    On Error GoTo ErrHandler
    
    '*-get the number of days a quote is valid for
    Set xmlItem = vxmlData.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "KFIDaysValidFor" & Chr$(34) & "]")
    lngDays = xmlGetAttributeAsInteger(xmlItem, "AMOUNT")
    '*-calculate and set the VALIDTODATE attribute
    strValue = Format$(DateAdd("d", lngDays, Now), "dd/mm/yyyy")
    Call xmlSetAttributeValue(vxmlNode, "VALIDTODATE", strValue)
    Set xmlItem = Nothing
    
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'SR 19/10/2001 : BBG1657
Public Sub AddOfferValidToDateAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddOfferValidToDateAttribute"
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
    Dim lngDays As Long

    On Error GoTo ErrHandler
    
    '*-get the number of days a quote is valid for
    Set xmlItem = vxmlData.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "OfferExpirationDate" & Chr$(34) & "]")
    lngDays = xmlGetAttributeAsInteger(xmlItem, "AMOUNT")
    '*-calculate and set the VALIDTODATE attribute
    strValue = Format$(DateAdd("d", lngDays, Now), "dd/mm/yyyy")
    Call xmlSetAttributeValue(vxmlNode, "VALIDTODATE", strValue)
    Set xmlItem = Nothing
    
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'SR 19/10/2001 : BBG1657 - End


'********************************************************************************
'** Function:       AddCompletedByDateAttribute
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Adds the COMPLETEDBYDATE attribute to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddCompletedByDateAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddCompletedByDateAttribute"
    Dim strValue As String
        
    On Error GoTo ErrHandler
        
    '*-get and set the COMPLETEDBYDATE attribute
    strValue = Format$(vobjCommon.ExpectedCompletionDate, "dd/mm/yyyy")
    Call xmlSetAttributeValue(vxmlNode, "COMPLETEDBYDATE", strValue)
    
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddTodayAttribute
'** Created by:     Andy Maggs
'** Date:           30/03/2004
'** Description:    Adds the specified 'today' attribute to the node.
'** Parameters:     vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddTodayAttribute(ByVal vxmlNode As IXMLDOMNode, _
        Optional ByVal vstrAttribName As String = "TODAY")
    Const cstrFunctionName As String = "AddTodayAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler
    
    '*-get the value for and add the attribute
    strValue = Format$(Now, "dd/mm/yyyy")
    Call xmlSetAttributeValue(vxmlNode, vstrAttribName, strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetPurposeOfLoanAttribute
'** Created by:     Andy Maggs
'** Date:           26/03/2004
'** Description:    Gets the value of the purpose of the loan attribute.
'** Parameters:     vxmlData - the data to get the attribute value from.
'** Returns:        The value if found, or 0.
'** Errors:         None Expected
'********************************************************************************
Public Function GetPurposeOfLoanAttribute(ByVal vxmlData As IXMLDOMNode) As Long
    Const cstrFunctionName As String = "GetPurposeOfLoanAttribute"
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler
    
    Set xmlItem = vxmlData.selectSingleNode(gcstrMORTGAGELOAN_PATH)
    If Not xmlItem Is Nothing Then
        GetPurposeOfLoanAttribute = xmlGetAttributeAsLong(xmlItem, "PURPOSEOFLOAN")
    End If
    Set xmlItem = Nothing

Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddFeesAndPremiumsElement
'** Created by:     Andy Maggs
'** Date:           30/03/2004
'** Description:    Adds the appropriate FEESANDPREMIUMS element together with
'**                 it's attributes.
'** Parameters:     vxmlOut - the XML document containing the formatted KFI data.
'**                 vxmlData - the XML data.
'**                 vxmlNode - the node to create the elements on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddFeesAndPremiumsElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddFeesAndPremiumsElement"
    Dim intNum As Integer
    Dim xmlFeeNode As IXMLDOMNode
    Dim xmlPurposeNode As IXMLDOMNode
    Dim aintFeePremNum(1 To 4, 1 To 4) As Integer

    On Error GoTo ErrHandler
    
    '*-in order to be able to determine which of the 16 possible elements to
    '*-create we need to extract the relevant data
    '                                           FEES
    '                       -------------------------------------
    '                       None    All Pay     Part    All Added
    'INSURANCE  None        1       2           3       4
    'PREMIUMS   All Pay     5       6           7       8
    '           Part        9       10          11      12
    '           All Added   13      14          15      16
    aintFeePremNum(1, 1) = 1
    aintFeePremNum(1, 2) = 5
    aintFeePremNum(1, 3) = 9
    aintFeePremNum(1, 4) = 13
    aintFeePremNum(2, 1) = 2
    aintFeePremNum(2, 2) = 6
    aintFeePremNum(2, 3) = 10
    aintFeePremNum(2, 4) = 14
    aintFeePremNum(3, 1) = 3
    aintFeePremNum(3, 2) = 7
    aintFeePremNum(3, 3) = 11
    aintFeePremNum(3, 4) = 15
    aintFeePremNum(4, 1) = 4
    aintFeePremNum(4, 2) = 8
    aintFeePremNum(4, 3) = 12
    aintFeePremNum(4, 4) = 16
    
    intNum = aintFeePremNum(vobjCommon.FeesType, vobjCommon.InsurancePremiumType)
    Set xmlFeeNode = vobjCommon.CreateNewElement("FEESANDPREMIUMS" & CStr(intNum), vxmlNode)
        
    '*-now add one of the mandatory PURPOSEPURCHASE, PURPOSEREMORTGAGE or
    '*-PURPOSEFURTHERADVANCE elements
    'CORE82  Don't want to do this
'    Set xmlPurposeNode = vobjCommon.AddLoanPurposeElement(xmlFeeNode)
'    If Not xmlPurposeNode Is Nothing Then
'        '*-add the mandatory TOTALLOANAMOUNT attribute required for 1-16
'        Call xmlSetAttributeValue(xmlPurposeNode, "TOTALLOANAMOUNT", CStr(vobjCommon.LoanAmount))
        
        'CORE82 Don't add AMOUNTOFFEES for 16
        '*-add the AMOUNTOFFEES attribute required for 3,7,11,15,16
        If intNum = 3 Or intNum = 7 Or intNum = 11 Or intNum = 15 Then
            Call xmlSetAttributeValue(xmlFeeNode, "AMOUNTOFFEES", CStr(vobjCommon.FeesAddedToLoanAmount))
        End If
        
        'CORE82 Add TOTALFEES for 16
        '*-add the TOTALFEES attribute required for 4,8,12,16
        If intNum = 4 Or intNum = 8 Or intNum = 12 Or intNum = 16 Then
            Call xmlSetAttributeValue(xmlFeeNode, "TOTALFEES", CStr(vobjCommon.TotalFees))
        End If
        
        '*-add the INSURANCEPREMIUM attribute required for 9,10,11,12
        If intNum = 9 Or intNum = 10 Or intNum = 11 Or intNum = 12 Then
            Call xmlSetAttributeValue(xmlFeeNode, "INSURANCEPREMIUM", CStr(vobjCommon.PremiumsAddedToLoanAmount))
        End If
            
        '*-add the TOTALPREMIUMS attribute required for 13,14,15,16
        If intNum = 13 Or intNum = 14 Or intNum = 15 Or intNum = 16 Then
            Call xmlSetAttributeValue(xmlFeeNode, "TOTALPREMIUMS", CStr(vobjCommon.TotalInsurancePremiums))
        End If
        
        '*-add the SECTIONNUMBER attribute
        AddFeesAndPremiumsSectionNumber xmlFeeNode, intNum, vobjCommon
        
'    End If
    Set xmlFeeNode = Nothing
    Set xmlPurposeNode = Nothing

Exit Sub
ErrHandler:
    Set xmlFeeNode = Nothing
    Set xmlPurposeNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddFeesAndPremiumsSectionNumber
'** Created by:     Paul Buck
'** Date:           05/03/2007
'** Description:    Adds the SECTIONNUMBER attribute(s), depending on the FEESANDPREMIUMS number
'                   and whether or not there is a section 2.
'** Parameters:     xmlNode    - the XML node affected.
'                   intNum     - the specific FEESANDPREMIUMS number used.
'                   vobjCommon - instance of class CommonDataHelper
'** Returns:        passes back complete node by reference
'** Errors:         None Expected
'********************************************************************************
Sub AddFeesAndPremiumsSectionNumber(ByVal xmlNode As IXMLDOMNode, ByVal intNum As Integer, vobjCommon As CommonDataHelper)
    '
    Const cstrFunctionName As String = "AddFeesAndPremiumsSectionNumber"
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    '
    Select Case intNum
        Case 1
            ' none yet for FEESANDPREMIUMS1
        Case 2, 3, 4, 7
            ' FEESANDPREMIUMS2 etc.
            If gblnSection2 Then
                strValue = "8"
            Else
                strValue = "7"
            End If
        Case 5, 9, 10, 13
            ' FEESANDPREMIUMS5 etc.
            If gblnSection2 Then
                strValue = "9"
            Else
                strValue = "8"
            End If
        Case 6, 11, 12, 16
            ' FEESANDPREMIUMS6 etc.
            If gblnSection2 Then
                strValue = "8 and 9"
            Else
                strValue = "7 and 8"
            End If
        Case 8, 15
            ' FEESANDPREMIUMS8 etc.
            If gblnSection2 Then
                strValue1 = "8"
                strValue2 = "9"
            Else
                strValue1 = "7"
                strValue2 = "8"
            End If
        Case 14
            ' FEESANDPREMIUMS14
            If gblnSection2 Then
                strValue1 = "9"
                strValue2 = "8"
            Else
                strValue1 = "8"
                strValue2 = "7"
            End If
    End Select
    '
    ' Check whether there is 1 or 2 values to add
    If strValue <> "" Then
        xmlSetAttributeValue xmlNode, "SECTIONNUMBER", strValue
    Else
        xmlSetAttributeValue xmlNode, "SECTIONNUMBER1", strValue1
        xmlSetAttributeValue xmlNode, "SECTIONNUMBER2", strValue2
    End If
    '
End Sub
'********************************************************************************
'** Function:       IsCatStandardMortgage
'** Created by:     Andy Maggs
'** Date:           31/03/2004
'** Description:    Gets whether the mortgage is CAT standard or not.
'** Parameters:     vxmlData - the XML data to search.
'** Returns:        True if it is, else False.
'** Errors:         None Expected
'********************************************************************************
Public Function IsCatStandardMortgage(ByVal vxmlData As IXMLDOMNode) As Boolean
    Const cstrFunctionName As String = "IsCatStandardMortgage"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim blnIsCatStandard As Boolean

    On Error GoTo ErrHandler
    
    Set xmlList = vxmlData.selectNodes(gcstrMORTGAGEPRODUCT_PATH)
    '*-if any mortgage product in the quote is CAT standard, then so is thw quote
    For Each xmlItem In xmlList
        blnIsCatStandard = xmlGetAttributeAsBoolean(xmlItem, "CATIND")
        If blnIsCatStandard Then Exit For
    Next xmlItem
    
    IsCatStandardMortgage = blnIsCatStandard
    Set xmlList = Nothing
    Set xmlItem = Nothing

Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddLenderContactAttributes
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Adds the APPLICANTNAME attribute to the node.
'** Parameters:     vxmlData - the source data XML.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddLenderContactAttributes(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddLenderContactAttributes"
    Dim xmlDirectoryItem As IXMLDOMNode
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
    Dim strContactTitleText As String ' PB 22/08/2006 EP1082

    On Error GoTo ErrHandler
    
    '*-get the value for the CONTACTCOMPANY attribute
    Set xmlDirectoryItem = vxmlData.selectSingleNode(gcstrNANDADIRECTORY_PATH)
    If Not xmlDirectoryItem Is Nothing Then
        strValue = xmlGetAttributeText(xmlDirectoryItem, "COMPANYNAME")
        Call xmlSetAttributeValue(vxmlNode, "CONTACTCOMPANY", strValue)
        
        '*-get the value(s) for and add the CONTACTNAME attribute
        Set xmlItem = xmlDirectoryItem.selectSingleNode("CONTACTDETAILS")
        If Not xmlItem Is Nothing Then
            'PB 22/08/2006 EP1082 Begin
            strContactTitleText = xmlGetAttributeText(xmlItem, "CONTACTTITLE_TEXT")
            If UCase(strContactTitleText) = "OTHER" Then
                strValue = BuildName(xmlGetAttributeText(xmlItem, "CONTACTTITLEOTHER"), _
                        xmlGetAttributeText(xmlItem, "CONTACTFORENAME"), _
                        xmlGetAttributeText(xmlItem, "CONTACTSURNAME"))
            Else
                strValue = BuildName(strContactTitleText, _
                        xmlGetAttributeText(xmlItem, "CONTACTFORENAME"), _
                        xmlGetAttributeText(xmlItem, "CONTACTSURNAME"))
            End If
            'PB EP1082 End
            Call xmlSetAttributeValue(vxmlNode, "CONTACTNAME", strValue)
        End If
            
        'CORE82 Rework of Telephone and Fax Contact Numbers
        '*-get the value(s) for and add the CONTACTTELEPHONENUMBER attribute
        Set xmlItem = xmlDirectoryItem.selectSingleNode("CONTACTTELEPHONEDETAILS[@USAGE_VALIDID ='W']") 'BC MAR1503
        If Not xmlItem Is Nothing Then
            '*-get the value for and add the CONTACTTELEPHONENUMBER attribute
            strValue = xmlGetAttributeText(xmlItem, "EXTENSIONNUMBER")
            If Len(strValue) > 0 Then
                strValue = " Extn: " & strValue
            End If
            strValue = xmlGetAttributeText(xmlItem, "TELENUMBER") & strValue
            strValue = xmlGetAttributeText(xmlItem, "AREACODE") & " " & strValue
            Call xmlSetAttributeValue(vxmlNode, "CONTACTTELEPHONENUMBER", strValue)
        End If
            
        '*-get the value(s) for and add the CONTACTFAXNUMBER attribute
        Set xmlItem = xmlDirectoryItem.selectSingleNode("CONTACTTELEPHONEDETAILS[@USAGE_VALIDID='F']") 'BC MAR1503
        If Not xmlItem Is Nothing Then
            '*-get the value for and add the CONTACTFAXNUMBER attribute
            strValue = xmlGetAttributeText(xmlItem, "TELENUMBER")
            strValue = xmlGetAttributeText(xmlItem, "AREACODE") & " " & strValue
            Call xmlSetAttributeValue(vxmlNode, "CONTACTFAXNUMBER", strValue)
        End If
        
        '*-get the value(s) for and add the address attributes
        Set xmlItem = xmlDirectoryItem.selectSingleNode("ADDRESS")
        If Not xmlItem Is Nothing Then
            'INR CORE82 Added "CONTACTADDRESSLINE1", "CONTACTADDRESSLINE2", "CONTACTADDRESSLINE3"
'PM 25/05/2006 EPSOM EP610   Start
'            Call AddAddressAttributes(xmlItem, vxmlNode, "CONTACTADDRESSLINE1", "CONTACTADDRESSLINE2", "CONTACTADDRESSLINE3", "CONTACTPOSTCODE")
            Call AddAddressAttributes(xmlItem, vxmlNode, "CONTACTADDRESS")
'PM 25/05/2006 EPSOM EP610 End
        End If
    End If
    Set xmlDirectoryItem = Nothing
    Set xmlItem = Nothing
    
Exit Sub
ErrHandler:
    Set xmlDirectoryItem = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Public Sub AddIntermediaryContactAttributes(ByVal vobjCommon As CommonDataHelper, _
                                            ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddIntermediaryContactAttributes"
    Dim strContactName As String
    Dim strCompanyName As String
    Dim strAddrLine1 As String
    Dim strAddrLine2 As String
    Dim strTemp As String
    Dim strAddrLine3 As String
    Dim strPostcode As String
    Dim strFaxNumber As String
    Dim strTelNo As String
    Dim xmlInterAddress As IXMLDOMNode
    Dim xmlContactDetails As IXMLDOMNode, xmlContactDetailsList As IXMLDOMNodeList
    Dim xmlIntermediaryCompany As IXMLDOMNode 'SR 27/10/2004 : BBG171

    On Error GoTo ErrHandler
    
    '*- Get all the values to be attached
    'SR 21/09/2004 : CORE82
'    strCompanyName = GetIntermediaryCompanyName(vobjCommon)
'    strContactName = GetIntermediaryContactName(vobjCommon)
    strCompanyName = vobjCommon.IntermediaryOrganisationName
    strContactName = vobjCommon.IntermediaryContactName
    'SR 21/09/2004 : CORE82 - End
    
    'SR 27/10/2004 : BBG1717
    Set xmlIntermediaryCompany = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONINTERMEDIARY_PATH & "[contains(@INTERMEDIARYTYPE,'ECROSD')]" & "/INTERMEDIARY")
    If xmlIntermediaryCompany Is Nothing Then Exit Sub
    
    Set xmlInterAddress = xmlIntermediaryCompany.selectSingleNode(".//ADDRESS")
    'Set xmlInterAddress = vobjCommon.Data.selectSingleNode(gcstrINTERMEDIARYADDRESS_PATH)
    'SR 27/10/2004 : BBG1717 - End
    
    If Not xmlInterAddress Is Nothing Then
        strTemp = xmlGetAttributeText(xmlInterAddress, "FLATNUMBER")
        If strTemp <> "" Then
            strAddrLine1 = strTemp & " "
        End If
        
        strTemp = xmlGetAttributeText(xmlInterAddress, "BUILDINGORHOUSENAME")
        If strTemp <> "" Then
            strAddrLine1 = strAddrLine1 & strTemp & " "
        End If
        
        strTemp = xmlGetAttributeText(xmlInterAddress, "BUILDINGORHOUSENUMBER")
        If strTemp <> "" Then
            strAddrLine2 = strAddrLine2 & strTemp & " "
        End If
        
        strTemp = xmlGetAttributeText(xmlInterAddress, "STREET")
        strAddrLine2 = strAddrLine2 & strTemp & " "
        
        strTemp = xmlGetAttributeText(xmlInterAddress, "TOWN")
        strAddrLine3 = strTemp
        
        strTemp = xmlGetAttributeText(xmlInterAddress, "COUNTY") 'SR 11/09/2004 : CORE82
        strAddrLine3 = strAddrLine3 & " " & strTemp
        
        strPostcode = xmlGetAttributeText(xmlInterAddress, "POSTCODE")
    End If
        'SR 11/09/2004 : CORE82
'        Set xmlContactDetailsList = vobjCommon.Data.selectNodes(gcstrINTERMEDIARYCONTACTDETAILS_PATH) 'SR 27/10/2004 : BBG1717
    Set xmlContactDetailsList = _
        xmlIntermediaryCompany.selectNodes(".//INTERMEDIARYORGANISATION/INTERMEDIARYCONTACT/CONTACTTELEPHONEDETAILS") 'SR 27/10/2004 : BBG1717
    For Each xmlContactDetails In xmlContactDetailsList
        If xmlGetAttributeText(xmlContactDetails, "USAGE_VALIDID") = "W" Then  'BC MAR1503
            strTelNo = "{" & _
                       xmlGetAttributeText(xmlContactDetails, "COUNTRYCODE") & " " & _
                       xmlGetAttributeText(xmlContactDetails, "AREACODE") & " " & _
                       xmlGetAttributeText(xmlContactDetails, "TELENUMBER") & "}"
        Else
            strFaxNumber = "{" & _
                           xmlGetAttributeText(xmlContactDetails, "COUNTRYCODE") & " " & _
                           xmlGetAttributeText(xmlContactDetails, "AREACODE") & " " & _
                           xmlGetAttributeText(xmlContactDetails, "TELENUMBER") & "}"
        End If
    Next xmlContactDetails 'SR 11/09/2004 : CORE82 - End
    
    '*-Add the address data as attributes
    xmlSetAttributeValue vxmlNode, "CONTACTNAME", strContactName
    xmlSetAttributeValue vxmlNode, "CONTACTCOMPANY", strCompanyName
    xmlSetAttributeValue vxmlNode, "CONTACTADDRESSLINE1", strAddrLine1
    xmlSetAttributeValue vxmlNode, "CONTACTADDRESSLINE2", strAddrLine2
    xmlSetAttributeValue vxmlNode, "CONTACTADDRESSLINE3", strAddrLine3
    xmlSetAttributeValue vxmlNode, "CONTACTPOSTCODE", strPostcode
    xmlSetAttributeValue vxmlNode, "CONTACTTELEPHONENUMBER", strTelNo
    xmlSetAttributeValue vxmlNode, "CONTACTFAXNUMBER", strFaxNumber
   
    Set xmlInterAddress = Nothing
    Set xmlContactDetails = Nothing
    Set xmlContactDetailsList = Nothing
    Set xmlIntermediaryCompany = Nothing
Exit Sub
ErrHandler:
    Set xmlInterAddress = Nothing
    Set xmlContactDetails = Nothing
    Set xmlContactDetailsList = Nothing
    Set xmlIntermediaryCompany = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Private Function GetIntermediaryContactName(ByVal vobjCommon As CommonDataHelper) As String
    Const cstrFunctionName As String = "GetIntermediaryContactName"
    Dim xmlIntermedInd As IXMLDOMNode
    Dim strContact As String
    
    On Error GoTo ErrHandler
    
    Set xmlIntermedInd = vobjCommon.Data.selectSingleNode(gcstrINTERMEDIARYINDIVIDUAL_PATH)
    If Not xmlIntermedInd Is Nothing Then
        strContact = BuildName(xmlGetAttributeText(xmlIntermedInd, "TITLE_TEXT"), _
                xmlGetAttributeText(xmlIntermedInd, "FORENAME"), _
                xmlGetAttributeText(xmlIntermedInd, "SURNAME"))
    Else
        strContact = ""
    End If
    
    GetIntermediaryContactName = strContact
    Set xmlIntermedInd = Nothing
    
Exit Function
ErrHandler:
    Set xmlIntermedInd = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

Private Function GetIntermediaryCompanyName(ByVal vobjCommon As CommonDataHelper) As String
    Const cstrFunctionName As String = "GetIntermediaryCompanyName"
    Dim xmlIntemediaryOrganisation As IXMLDOMNode
    Dim xmlIntemediaryOrganisationList As IXMLDOMNodeList, xmlNode As IXMLDOMNode  'SR 11/09/2004 : CORE82
    Dim strValidationType As String 'SR 11/09/2004 : CORE82
    
    On Error GoTo ErrHandler
    
    'SR 11/09/2004 : CORE82
    Set xmlIntemediaryOrganisationList = vobjCommon.Data.selectNodes(gcstrAPPLICATIONINTERMEDIARY_PATH)
    If xmlIntemediaryOrganisationList.length > 1 Then
        For Each xmlNode In xmlIntemediaryOrganisationList
            strValidationType = xmlGetAttributeText(xmlNode, "INTERMEDIARYTYPE")
            If (InStr((UCase$(strValidationType)), "ECROSD,") > 0) Then
                Set xmlIntemediaryOrganisation = xmlNode.selectSingleNode(".//INTERMEDIARY/INTERMEDIARYORGANISATION")
                Exit For
            End If
        Next xmlNode
    Else
        Set xmlIntemediaryOrganisation = vobjCommon.Data.selectSingleNode(gcstrINTERMEDIARYORGANISATION_PATH)
    End If
    'SR 11/09/2004 : CORE82
    
    If Not xmlIntemediaryOrganisation Is Nothing Then
        GetIntermediaryCompanyName = xmlGetAttributeText(xmlIntemediaryOrganisation, "NAME")
    Else
        GetIntermediaryCompanyName = ""
    End If
    Set xmlIntemediaryOrganisation = Nothing
    Set xmlIntemediaryOrganisationList = Nothing
    Set xmlNode = Nothing

Exit Function
ErrHandler:
    Set xmlIntemediaryOrganisation = Nothing
    Set xmlIntemediaryOrganisationList = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddRatePeriodElements
'** Created by:     Andy Maggs
'** Date:           05/04/2004
'** Description:    Adds a RATEPERIOD (and children) elements for each
'**                 InterestRateType record in the loan component.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmLoanComponent - the input data representing a single
'**                 LOANCOMPONENT
'**                 vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddRatePeriodElements(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode, ByVal vxmlNode As IXMLDOMNode, _
        ByVal vblnFirstPeriodOnly As Boolean, ByVal vblnNotFirstPeriod As Boolean, _
        Optional ByRef rblnHasMoreRateTypes As Boolean)
    Const cstrFunctionName As String = "AddRatePeriodElements"
    Dim eType As MortgageInterestRateType
    Dim xmlRatePeriodNode As IXMLDOMNode
    Dim xmlPeriodNode As IXMLDOMNode
    Dim xmlRateNode As IXMLDOMNode
    Dim dblRatePayable As Double
    Dim dblDiscountRate As Double
    Dim dblBaseRate As Double
    Dim dblCappedRate As Double
    Dim dblCollaredRate As Double
    Dim strBaseRateName As String
    Dim strValue As String
    Dim xmlList As IXMLDOMNodeList
    Dim xmlRate As IXMLDOMNode
    Dim blnIsFullTerm As Boolean
    Dim blnIsRemainingTerm As Boolean
    Dim blnIsFinalRate As Boolean
    Dim blnHasEndDate As Boolean
    Dim dtEndDate As Date
    Dim intPeriod As Integer
    Dim xmlNewNode As IXMLDOMNode
    Dim dblFinalRatePayable As Double
    Dim strFinalRateName As String
    Dim intFrom As Integer
    Dim intTo As Integer
    Dim intIndex As Integer
    Dim strPart As String
    'SR 28/09/2004 : CORE82
    Dim dblBaseRateAdjustment As Double
    Dim xmlLC As IXMLDOMNode
    Dim xmlNotTrackerBelow As IXMLDOMNode
    Dim blnTrackerBelowProduct As Boolean
    'SR 28/09/2004 : CORE82 - End

    On Error GoTo ErrHandler

    '*-get all the interest rates for this loan component
    Set xmlList = vobjCommon.GetLoanComponentInterestRates(vxmlLoanComponent)
    Set xmlLC = vobjCommon.SingleLoanComponent 'SR 28/09/2004 : CORE82
        
    If vblnFirstPeriodOnly Then
        intFrom = 0
        intTo = 0
        rblnHasMoreRateTypes = (xmlList.length > 1)
    ElseIf vblnNotFirstPeriod Then
        intFrom = 1
        intTo = xmlList.length - 1
        strPart = xmlGetAttributeText(vxmlLoanComponent, "LOANCOMPONENTSEQUENCENUMBER")
    Else
        intFrom = 0
        intTo = xmlList.length - 1
    End If
    
    '*-add the rate periods
    For intIndex = intFrom To intTo
        Set xmlRate = xmlList.Item(intIndex)
               
        '*-get the rate type for this item
        eType = GetInterestRateType(xmlRate)
        '*-get the actual rates for this item
        Call GetInterestRates(vobjCommon, vxmlLoanComponent, xmlRate, eType, _
                dblRatePayable, dblDiscountRate, dblBaseRate, dblCappedRate, _
                dblCollaredRate, strBaseRateName, dblBaseRateAdjustment) 'SR 28/09/2004 : CORE82
        
        'CORE82 Moved from below. Need blnIsFinalRate
        '*-get the term for this rate period
        Call GetInterestRateTermData(vobjCommon, xmlList, xmlRate, _
                blnIsFullTerm, blnIsRemainingTerm, blnIsFinalRate, blnHasEndDate, _
                dtEndDate, intPeriod, dblFinalRatePayable, strFinalRateName)
                
        'CORE82 Create RATESTEP if Finalrate is FALSE and this is a SINGLECOMPONENT
        'i.e. if vblnFirstPeriodOnly = False and vblnNotFirstPeriod = False
        
        'MAR88 - BC 05 Oct
        If ((blnIsFinalRate = False) And (vblnFirstPeriodOnly = False) And (vblnNotFirstPeriod = False) And (blnTrackerBelowProduct = False)) Then
           'If ((blnIsFinalRate = False) And (vblnFirstPeriodOnly = False) And (vblnNotFirstPeriod = False) And (blnTrackerBelowProduct = False)) Then
            '*-for each rate, add the RATEPERIOD element
            Set xmlRatePeriodNode = vobjCommon.CreateNewElement("RATEPERIOD", vxmlNode)
            If vblnNotFirstPeriod Then
                '*-add the PART attribute
                Call xmlSetAttributeValue(xmlRatePeriodNode, "PART", strPart)
            End If
            '*-add the RATESTEP element
            Set xmlPeriodNode = vobjCommon.CreateNewElement("RATESTEP", xmlRatePeriodNode)
        Else
            'CORE82 Only Create RATEPERIOD if Finalrate is TRUE
            '*-for each rate, add the RATEPERIOD element
            Set xmlPeriodNode = vobjCommon.CreateNewElement("RATEPERIOD", vxmlNode)
            If vblnNotFirstPeriod Then
                '*-add the PART attribute
                Call xmlSetAttributeValue(xmlPeriodNode, "PART", strPart)
            End If

        End If
        
        If eType = mrtStandardVariableRate Then    ' Or eType = mrtDiscountedRate Then 'SR 28/09/2004 : CORE82
            '*-add the VARIABLE element
            Set xmlRateNode = vobjCommon.CreateNewElement("VARIABLE", xmlPeriodNode)
            '*-add the INITIALRATE attribute
            If eType = mrtStandardVariableRate Then
                strValue = CStr(dblRatePayable)
            Else
                strValue = CStr(dblBaseRate)
            End If
            ' PB 24/05/2006 EP603/MAR1777
            Call xmlSetAttributeValue(xmlRateNode, "NAMEOFSVR", strFinalRateName)
            ' EP603/MAR1777 End
            Call xmlSetAttributeValue(xmlRateNode, "INITIALRATE", strValue)
        End If
        
        If eType = mrtFixedRate Then
            '*-or add the FIXED element
            Set xmlRateNode = vobjCommon.CreateNewElement("FIXED", xmlPeriodNode)
            '*-add the INITIALRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "INITIALRATE", CStr(dblRatePayable))
        End If
        
        If eType = mrtCappedRate Then
            '*-or add the CAPPED element
            Set xmlRateNode = vobjCommon.CreateNewElement("CAPPED", xmlPeriodNode)
            '*-add the CEILINGRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "CEILINGRATE", CStr(dblCappedRate))
            '*-add the INITIALRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "INITIALRATE", CStr(dblBaseRate))
        End If
        
        If eType = mrtCollaredRate Then
            '*-or add the COLLARED element
            Set xmlRateNode = vobjCommon.CreateNewElement("COLLARED", xmlPeriodNode)
            '*-add the FLOORRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "FLOORRATE", CStr(dblCollaredRate))
            '*-add the INITIALRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "INITIALRATE", CStr(dblBaseRate))
        End If
        
        If eType = mrtCappedAndCollaredRate Then
            '*-or add the CAPPEDANDCOLLARED element
            Set xmlRateNode = vobjCommon.CreateNewElement("CAPPEDANDCOLLARED", xmlPeriodNode)
            '*-add the CEILINGRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "CEILINGRATE", CStr(dblCappedRate))
            '*-add the FLOORRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "FLOORRATE", CStr(dblCollaredRate))
            '*-add the INITIALRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "INITIALRATE", CStr(dblBaseRate))
        End If
        
        If eType = mrtTrackerAbove Then
            '*-or add the TRACKERABOVE element
            Set xmlRateNode = vobjCommon.CreateNewElement("TRACKERABOVE", xmlPeriodNode)
            '*-add the VARIABLERATE attribute
            'INR CORE82 VARIABLERATE not INITIALRATE
            Call xmlSetAttributeValue(xmlRateNode, "VARIABLERATE", CStr(dblBaseRate))
            '*-add the NAMEOFTRACKEDRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "NAMEOFTRACKEDRATE", strBaseRateName)
            '*-add the PERCENTABOVE attribute
            Call xmlSetAttributeValue(xmlRateNode, "PERCENTABOVE", CStr(Abs(dblDiscountRate)))
        
        End If

'       MAR44 BC 12 Sep 05
'        If eType = mrtDiscountedRate Then
'            'SR 28/09/2004 : CORE82
'            Set xmlRateNode = vobjCommon.CreateNewElement("TRACKERBELOW", xmlPeriodNode)
'            xmlSetAttributeValue xmlRateNode, "DISCOUNTRATE", set2DP(CStr(dblDiscountRate))
'            xmlSetAttributeValue xmlRateNode, "BASERATE", set2DP(CStr(dblBaseRate - dblBaseRateAdjustment))
'            xmlSetAttributeValue xmlRateNode, "LOADEDRATE", set2DP(CStr(dblBaseRate))
'            xmlSetAttributeValue xmlRateNode, "NAMEOFTRACKEDRATE", strBaseRateName
'            xmlSetAttributeValue xmlRateNode, "PERCENTABOVE", set2DP(CStr(dblBaseRateAdjustment))
'            xmlSetAttributeValue xmlRateNode, "RESOLVEDRATE", set2DP(CStr(xmlGetAttributeAsDouble(xmlLC, "RESOLVEDRATE")))
'            Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
'                intPeriod, xmlRateNode, , , , True)
'            'E2EM00002949 Just call this once, AFTERPERIOD is now dealt with in
'            'AddUntilDateOrForPeriodElements Otherwise we get an extra UNTILDATE,
'            'which we don't want if we have an enddate.
''            Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
''                intPeriod, xmlRateNode, , , "AFTERPERIOD")
'            If intIndex = 0 Then
'                blnTrackerBelowProduct = True
'            End If
'            'SR 28/09/2004 : CORE82 - End
'        End If
        
        If eType = mrtDiscountedRate Or eType = mrtTrackerAbove Then
            '*-add the ADJUSTEDRATE element
            Set xmlRateNode = vobjCommon.CreateNewElement("ADJUSTEDRATE", xmlPeriodNode)
            '*-add the INITIALRATE attribute
            Call xmlSetAttributeValue(xmlRateNode, "INITIALRATE", set2DP(CStr(dblRatePayable))) 'SR 28/09/2004 : CORE82
        End If
        
        If Not blnTrackerBelowProduct Then  'SR 28/09/2004 : CORE82
            Set xmlNotTrackerBelow = vobjCommon.CreateNewElement("NOTTRACKERBELOW", xmlPeriodNode)
            
            '*-add the UNTILDATE or FORPERIOD elements
            Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
                    intPeriod, xmlNotTrackerBelow)
            
            If blnIsRemainingTerm Then
                '*-add the REMAININGTERM element
                Call vobjCommon.CreateNewElement("REMAININGTERM", xmlNotTrackerBelow)
            End If
            
            If blnIsFullTerm Then
                '*-add the FULLTERM element
                Call vobjCommon.CreateNewElement("FULLTERM", xmlNotTrackerBelow)
            End If
            
            If blnIsFinalRate And (Not blnTrackerBelowProduct) Then 'SR 28/09/2004 : CORE82
                '*-add the FINALRATE element
                Set xmlNewNode = vobjCommon.CreateNewElement("FINALRATE", xmlPeriodNode) 'SR 28/09/2004 : CORE82
                '*-add the VARIABLERATE attribute
                Call xmlSetAttributeValue(xmlNewNode, "VARIABLERATE", CStr(dblFinalRatePayable))
                '*-add the NAMEOFSVR attribute
                Call xmlSetAttributeValue(xmlNewNode, "NAMEOFSVR", strFinalRateName)
            End If
            
            If vblnNotFirstPeriod Then
                '*-this applies to MULTICOMPONENT section only and then only for 2nd and
                '*-subsequent interest rate periods
                '*-get the previous rate period
                Set xmlRate = xmlList.Item(intIndex - 1)
                '*-get the rate type for this item
                eType = GetInterestRateType(xmlRate)
                '*-get the term data for this previous rate
                Call GetInterestRateTermData(vobjCommon, xmlList, xmlRate, _
                        blnIsFullTerm, blnIsRemainingTerm, blnIsFinalRate, _
                        blnHasEndDate, dtEndDate, intPeriod, dblFinalRatePayable, _
                        strFinalRateName)
                
                '*-add the UNTILDATE or FORPERIOD elements
                Call AddUntilDateOrForPeriodElements(vobjCommon, blnHasEndDate, dtEndDate, _
                        intPeriod, xmlNotTrackerBelow, "ONDATE", "ONDATE", "AFTERPERIOD")
            
            End If
        End If 'SR 28/09/2004 : CORE82 - End
    Next intIndex
    
    Set xmlRatePeriodNode = Nothing
    Set xmlPeriodNode = Nothing
    Set xmlRateNode = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlNewNode = Nothing
    Set xmlLC = Nothing
    Set xmlNotTrackerBelow = Nothing
Exit Sub
ErrHandler:
    Set xmlRatePeriodNode = Nothing
    Set xmlPeriodNode = Nothing
    Set xmlRateNode = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlNewNode = Nothing
    Set xmlLC = Nothing
    Set xmlNotTrackerBelow = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddUntilDateOrForPeriodElements
'** Created by:     Andy Maggs
'** Date:           13/07/2004
'** Description:    Adds an UNTILDATE or FORPERIOD element to the supplied element.
'** Parameters:     vblnHasEndDate - whether a valid end date is supplied or not.
'**                 vdtEndDate - the end date.
'**                 vintPeriod - the period in months.
'**                 vxmlNode - the node to add the elements to.
'**                 vstrUntilDateElemName - the name of the UNTILDATE element.
'**                 vstrEndDateAttribName - the name of the ENDDATE attribute.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddUntilDateOrForPeriodElements(ByVal vobjCommon As CommonDataHelper, _
        ByVal vblnHasEndDate As Boolean, ByVal vdtEndDate As Date, _
        ByVal vintPeriod As Integer, ByVal vxmlNode As IXMLDOMNode, _
        Optional ByVal vstrUntilDateElemName As String = "UNTILDATE", _
        Optional ByVal vstrEndDateAttribName As String = "ENDDATE", _
        Optional ByVal vstrForPeriodElemName As String = "FORPERIOD", _
        Optional ByVal blnAfterAttribute As Boolean = False)
        
    Const cstrFunctionName As String = "AddUntilDateOrForPeriodElements"
    
    Dim xmlNewNode As IXMLDOMNode
    Dim intYears As Integer
    Dim intMths As Integer
    Dim strPrefPeriod As String

    On Error GoTo ErrHandler

    If vblnHasEndDate Then
        '*-add the UNTILDATE element
        Set xmlNewNode = vobjCommon.CreateNewElement(vstrUntilDateElemName, vxmlNode)
        '*-add the ENDDATE attribute
        Call xmlSetAttributeValue(xmlNewNode, vstrEndDateAttribName, _
                Format$(vdtEndDate, "dd/mm/yyyy"))
                
        '*E2EM00002949 - add the AFTERDATE element
        '(ONLY CURRENTLY USED BY SECTION4 TRACKERBELOW,)
        If blnAfterAttribute Then
            Set xmlNewNode = vobjCommon.CreateNewElement("AFTERDATE", vxmlNode)
            '*-add the ENDDATE attribute
            Call xmlSetAttributeValue(xmlNewNode, vstrEndDateAttribName, _
                    Format$(vdtEndDate, "dd/mm/yyyy"))
        End If

    ElseIf vintPeriod <> -1 Then
        '*-add the FORPERIOD element
        Set xmlNewNode = vobjCommon.CreateNewElement(vstrForPeriodElemName, vxmlNode)
        'PM 19/06/2006 EP697- START
        'if no months are applicable then months should not appear.
        'If no whole years is applicable then year should not appear
        intYears = vintPeriod / 12
        intMths = vintPeriod Mod 12
        If intYears <> 0 Then
            strPrefPeriod = CStr(intYears) & IIf(intYears > 1, " Years", " Year")
        End If
        If intMths <> 0 Then
            If strPrefPeriod <> "" Then
                strPrefPeriod = strPrefPeriod & " and "
            End If
            strPrefPeriod = strPrefPeriod & CStr(intMths) & IIf(intMths > 1, " Months", "Month")
        End If
        
        '*-add the TERMYEARS attribute
        Call xmlSetAttributeValue(xmlNewNode, "TERMYEARS", CStr(intYears))
'        Call xmlSetAttributeValue(xmlNewNode, "TERMYEARS", CStr(vintPeriod \ 12))
        '*-add the TERMMONTHS attribute
        Call xmlSetAttributeValue(xmlNewNode, "TERMMONTHS", CStr(intMths))
'        Call xmlSetAttributeValue(xmlNewNode, "TERMMONTHS", CStr(vintPeriod Mod 12))
        Call xmlSetAttributeValue(xmlNewNode, "PREFPERIOD", strPrefPeriod)
        
        '*E2EM00002949 - add the AFTERPERIOD element
        '(ONLY CURRENTLY USED BY SECTION4 TRACKERBELOW,)
        If blnAfterAttribute Then
            Set xmlNewNode = vobjCommon.CreateNewElement("AFTERPERIOD", vxmlNode)
            '*-add the TERMYEARS attribute
            Call xmlSetAttributeValue(xmlNewNode, "TERMYEARS", CStr(intYears))
 '           Call xmlSetAttributeValue(xmlNewNode, "TERMYEARS", CStr(vintPeriod \ 12))
            '*-add the TERMMONTHS attribute
            Call xmlSetAttributeValue(xmlNewNode, "TERMMONTHS", CStr(intMths))
'            Call xmlSetAttributeValue(xmlNewNode, "TERMMONTHS", CStr(vintPeriod Mod 12))
            Call xmlSetAttributeValue(xmlNewNode, "PREFPERIOD", strPrefPeriod)
        End If
        'PM 19/06/2006 EP697- END
    End If
    Set xmlNewNode = Nothing
Exit Sub
ErrHandler:
    Set xmlNewNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetInterestRates
'** Created by:     Andy Maggs
'** Date:           16/04/2004
'** Description:    Gets the various interest rates depending on the type of the
'**                 interestratetype element.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlLoanComponent - the loan component.
'**                 vxmlInterestRate - the particular interest rate element.
'**                 veType - the type of interest rate element.
'**                 rdblRatePayable - the actual rate payable
'**                 rdblDiscountRate - the amount of discount
'**                 rdblBaseRate - the base rate before any discount
'**                 rdblCappedRate - the capped rate
'**                 rdblCollaredRate - the collared rate
'**                 rstrBaseRateName - the name of the base rate
'** Returns:        Values via the ByRef parameters.
'** Errors:         None Expected
'********************************************************************************
Public Sub GetInterestRates(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode, ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal veType As MortgageInterestRateType, _
        ByRef rdblRatePayable As Double, _
        Optional ByRef rdblDiscountRate As Double = 0, _
        Optional ByRef rdblBaseRate As Double = 0, _
        Optional ByRef rdblCappedRate As Double = 0, _
        Optional ByRef rdblCollaredRate As Double = 0, _
        Optional ByRef rstrBaseRateName As String = "", _
        Optional ByRef rdblBaseRateAdjustment As Double = 0) 'SR 28/09/2004 : CORE82
    Const cstrFunctionName As String = "GetInterestRates"
    Dim intSequence As Integer
    Dim dblRate As Double

    On Error GoTo ErrHandler
    
    intSequence = xmlGetAttributeAsInteger(vxmlInterestRate, "INTERESTRATETYPESEQUENCENUMBER")
    dblRate = xmlGetAttributeAsDouble(vxmlInterestRate, "RATE")
    
    If veType = mrtStandardVariableRate Then
        If intSequence = 1 Then
            rdblRatePayable = xmlGetAttributeAsDouble(vxmlLoanComponent, "RESOLVEDRATE")
        Else
            rdblRatePayable = GetBaseInterestRate(vobjCommon, vxmlInterestRate, _
                    rstrBaseRateName)
        End If
    Else
        If veType = mrtFixedRate Then
            'SR EP2_1938 : For Epsom Fixed rates are not store in table INTERESTRATETYPE but in BASERATE and BASERATEBAND
            rdblRatePayable = dblRate + GetBaseInterestRate(vobjCommon, vxmlInterestRate, rstrBaseRateName) 'SR EP2_1938
        Else
            If veType = mrtDiscountedRate Or veType = mrtTrackerAbove Then
                rdblBaseRate = GetBaseInterestRate(vobjCommon, vxmlInterestRate, _
                        rstrBaseRateName, rdblBaseRateAdjustment)  'SR 28/09/2004 : CORE82
                rdblDiscountRate = dblRate
                rdblRatePayable = rdblBaseRate - rdblDiscountRate
            Else
                If veType = mrtCappedRate Or veType = mrtCollaredRate _
                        Or veType = mrtCappedAndCollaredRate Then
                    rdblBaseRate = GetBaseInterestRate(vobjCommon, vxmlInterestRate, _
                            rstrBaseRateName)
                    rdblCappedRate = xmlGetAttributeAsDouble(vxmlInterestRate, "CEILINGRATE")
                    rdblCollaredRate = xmlGetAttributeAsDouble(vxmlInterestRate, "FLOOREDRATE")
                    rdblDiscountRate = dblRate
                    rdblRatePayable = rdblBaseRate - rdblDiscountRate
                End If
            End If
        End If
    End If

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetBaseInterestRate
'** Created by:     Andy Maggs
'** Date:           16/04/2004
'** Description:    Gets the appropriate base interest rate for the specified
'**                 interest rate type.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlInterestRate - the interest rate record.
'**                 rstrBaseRateName - the name of the returned base rate.
'** Returns:        The appropriate base rate.
'** Errors:         None Expected
'********************************************************************************
Public Function GetBaseInterestRate(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlInterestRate As IXMLDOMNode, ByRef rstrBaseRateName As String, _
        Optional ByRef rdblBaseRateAdjustment As Double, _
        Optional ByRef rdblCurrentBaseInterestRate As Double, _
        Optional ByRef rintCurrentRateType As Integer _
        ) As Double
    'SR 28/09/2004 : CORE82
    'PM 08/06/2006 : EP697 Added last 2 optional output params above [added PB 12/06/2006]
    Const cstrFunctionName As String = "GetBaseInterestRate"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim dtDate As Date
    Dim dtLatestDate As Date
    Dim dblCurrentBaseRate As Double
    Dim xmlBand As IXMLDOMNode
    Dim dblRateDifference As Double
    Dim dblActualRate As Double
    Dim xmlBaseRateSet As IXMLDOMNode
    'PB 12/06/2006 EP697 Begin
    Dim intCurrentRateType As Integer
    Dim dblCurrentBaseInterestRate As Double
    'EP697 End
    
    On Error GoTo ErrHandler

    Set xmlList = vxmlInterestRate.selectNodes(".//BASERATE")
    For Each xmlItem In xmlList
        dtDate = xmlGetAttributeAsDate(xmlItem, "BASERATESTARTDATE")
        'If dtDate <= Now And dtDate > dtLatestDate Then 'SR 28/09/2004 : CORE82
        If dtDate > dtLatestDate Then 'SR 28/09/2004 : CORE82
            dtLatestDate = dtDate
            dblCurrentBaseRate = xmlGetAttributeAsDouble(xmlItem, "BASEINTERESTRATE")
'PM 08/06/2006 EP697 Start [PB 12/06/2006]
            intCurrentRateType = xmlGetAttributeAsDouble(xmlItem, "RATETYPE")
            dblCurrentBaseInterestRate = xmlGetAttributeAsDouble(xmlItem, "BASEINTERESTRATE")
'PM 08/06/2006 EP697 End
            'BBG1567 changed to use BASERATESETDESCRIPTION
'            rstrBaseRateName = xmlGetAttributeText(xmlItem, "RATEDESCRIPTION")
        End If
    Next xmlItem
    
    'BBG1567 changed from BASERATE/RATEDESCRIPTION to use BASERATESETDESCRIPTION
    Set xmlBaseRateSet = vxmlInterestRate.selectSingleNode("//BASERATESETDATA")
    rstrBaseRateName = xmlGetAttributeText(xmlBaseRateSet, "BASERATESETDESCRIPTION")
    
    Set xmlBand = SelectBaseRateBand(vxmlInterestRate, vobjCommon.LoanAmount, _
            vobjCommon.LTV)
'TW 09/11/2005 MAR442
    On Error Resume Next
    dblRateDifference = xmlGetAttributeAsDouble(xmlBand, "RATEDIFFERENCE") 'SR 28/09/2004 : CORE82
    On Error GoTo ErrHandler
'TW 09/11/2005 MAR442 End

    dblActualRate = dblCurrentBaseRate + dblRateDifference
    
    rdblBaseRateAdjustment = dblRateDifference     'SR 28/09/2004 : CORE82
'PM 08/06/2006 EP697 Start [PB 12/06/2006]
    rintCurrentRateType = intCurrentRateType
    rdblCurrentBaseInterestRate = dblCurrentBaseInterestRate
'PM 08/06/2006 EP697 End

    GetBaseInterestRate = dblActualRate
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlBand = Nothing
    Set xmlBaseRateSet = Nothing
Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlBand = Nothing
    Set xmlBaseRateSet = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       SelectBaseRateBand
'** Created by:     Andy Maggs
'** Date:           29/03/2004
'** Description:    Selects the base rate band having the smallest maximum loan
'**                 amount greater than the supplied loan amount and also has the
'**                 smallest maximum LTV value greater than the supplied LTV
'**                 amount.
'** Parameters:     vxmlData - the INTERESTRATETYPE node containing the
'**                 BASERATEBAND nodes to search.
'**                 vlngLoanAmount - the loan amount to filter by.
'**                 vdblLTV - the LTV amount to filter by.
'** Returns:        The appropriate base rate band node.
'** Errors:         None Expected
'********************************************************************************
Private Function SelectBaseRateBand(ByVal vxmlData As IXMLDOMNode, _
        ByVal vlngLoanAmount As Long, ByVal vdblLTV As Double) As IXMLDOMNode
    Const cstrFunctionName As String = "SelectBaseRateBand"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim lngMaxLoanAmount As Long
    Dim dblMaxLTV As Double
    Dim lngSmallestMaxLoanAmount As Long
    Dim dblSmallestMaxLTV As Double
    Dim xmlSelection As IXMLDOMNode
    Dim intItem As Integer

    On Error GoTo ErrHandler
    
    Set xmlList = vxmlData.selectNodes("BASERATESETDATA/BASERATEBAND")
    '*-search the list to find the band that has the smallest maximum loan amount
    '*-that is also bigger than or equal to the supplied loan amount AND that has
    '*-the smallest maximum LTV that is also bigger than or equal to the supplied
    '*-LTV value
    intItem = 0
    For Each xmlItem In xmlList
        lngMaxLoanAmount = xmlGetAttributeAsLong(xmlItem, "MAXIMUMLOANAMOUNT")
        dblMaxLTV = xmlGetAttributeAsDouble(xmlItem, "MAXIMUMLTV")
        
        If lngMaxLoanAmount >= vlngLoanAmount And dblMaxLTV >= vdblLTV Then
            intItem = intItem + 1
            If intItem = 1 Then
                '*-set the initial values of the smallest
                lngSmallestMaxLoanAmount = lngMaxLoanAmount
                dblSmallestMaxLTV = dblMaxLTV
            End If
            '*-this is a valid rate band, but are the values the smallest
            If lngMaxLoanAmount <= lngSmallestMaxLoanAmount Then
                lngSmallestMaxLoanAmount = lngMaxLoanAmount
                If dblMaxLTV <= dblSmallestMaxLTV Then
                    dblSmallestMaxLTV = dblMaxLTV
                    '*-this is the closest match so far
                    Set xmlSelection = xmlItem
                End If
            End If
        End If
    Next xmlItem
    
    Set SelectBaseRateBand = xmlSelection
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlSelection = Nothing
Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlSelection = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

Public Sub GetBasicInterestRateTermData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlRateList As IXMLDOMNodeList, _
        ByVal vxmlInterestRate As IXMLDOMNode, _
        ByRef rblnHasEndDate As Boolean, _
        ByRef rdtEndDate As Date, _
        ByRef rintPeriod As Integer)
    Dim blnIsFullTerm As Boolean
    Dim blnIsRemainingTerm As Boolean
    Dim blnIsFinalRate As Boolean
    Dim dblFinalRatePayable As Double
    Dim strFinalRateName As String
    
    '*-pass the call through to the method actually doing the work
    Call GetInterestRateTermData(vobjCommon, vxmlRateList, vxmlInterestRate, _
            blnIsFullTerm, blnIsRemainingTerm, blnIsFinalRate, rblnHasEndDate, _
            rdtEndDate, rintPeriod, dblFinalRatePayable, strFinalRateName)
            
End Sub

'********************************************************************************
'** Function:       GetInterestRateTermData
'** Created by:     Andy Maggs
'** Date:           21/04/2004
'** Description:    Gets the term related data for a single InterestRateType
'**                 record.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlInterestRate - the interestratetype record.
'**                 rblnIsFullTerm - whether this rate is for the full term.
'**                 rblnIsRemainingTerm - whether this rate is for the remaining
'**                 term.
'**                 rblnIsFinalRate - whether this is the final rate in the
'**                 sequence.
'**                 rblnHasEndDate - whether this rate has an end date or period.
'**                 rdtEndDate - the end date.
'**                 rintPeriod - the period in months.
'**                 rdblFinalRatePayable - the final rate payable if this is the
'**                 final interest rate record.
'**                 rstrFinalRateName - the name of this rate if it is the final
'**                 interest rate record.
'** Returns:        Values via the ByRef parameters.
'** Errors:         None Expected
'********************************************************************************
Private Sub GetInterestRateTermData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlRateList As IXMLDOMNodeList, _
        ByVal vxmlInterestRate As IXMLDOMNode, _
        ByRef rblnIsFullTerm As Boolean, _
        ByRef rblnIsRemainingTerm As Boolean, _
        ByRef rblnIsFinalRate As Boolean, _
        ByRef rblnHasEndDate As Boolean, _
        ByRef rdtEndDate As Date, _
        ByRef rintPeriod As Integer, _
        ByRef rdblFinalRatePayable As Double, _
        ByRef rstrFinalRateName As String)
    Const cstrFunctionName As String = "GetInterestRateTermData"
    Dim strRateType As String
    Dim dtDate As Date

    On Error GoTo ErrHandler
    
    '*-initialise ByRef params
    rblnIsFullTerm = False
    rblnIsRemainingTerm = False
    rblnIsFinalRate = False
    rblnHasEndDate = False
    rdtEndDate = dtDate
    rdblFinalRatePayable = 0
    rstrFinalRateName = vbNullString
    
    rintPeriod = xmlGetAttributeAsInteger(vxmlInterestRate, "INTERESTRATEPERIOD")
    '*-is this the last interest rate type record?
    If rintPeriod = -1 Then
        If vxmlRateList.length = 1 Then
            rblnIsFullTerm = True
        Else
            '*-check to see if it is a base rate record
            strRateType = xmlGetAttributeText(vxmlInterestRate, "RATETYPE")
            If strRateType <> "B" Then
                rblnIsRemainingTerm = True
            Else
                rblnIsFinalRate = True
                '*-get the name of the final rate
                rdblFinalRatePayable = GetBaseInterestRate(vobjCommon, vxmlInterestRate, rstrFinalRateName)
            End If
        End If
    Else
        '*-it isn't, so we need to work out the period of this record
        '*-if there is an end date use it otherwise use the period
        If xmlAttributeValueExists(vxmlInterestRate, "INTERESTRATEENDDATE") Then
            rblnHasEndDate = True
            'TK 07/11/2004 BBG1782 Start
            Dim strRateEndDate As String
            strRateEndDate = xmlGetAttributeText(vxmlInterestRate, "INTERESTRATEENDDATE")
            
            If InStr(1, strRateEndDate, "T") Then
                strRateEndDate = Left$(strRateEndDate, InStr(1, strRateEndDate, "T") - 1)

                If IsDate(strRateEndDate) Then
                    rdtEndDate = CDate(strRateEndDate)
                End If
            Else
                rdtEndDate = xmlGetAttributeAsDate(vxmlInterestRate, "INTERESTRATEENDDATE")
            End If
            'TK 07/11/2004 BBG1782 End
        End If
    End If
    
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddPurchPriceEstValueAttribute
'** Created by:     Andy Maggs
'** Date:           08/04/2004
'** Description:    Adds the PURCHASEPRICE or ESTIMATEDVALUE attribute.
'** Parameters:     vxmlData - the XML source data being the omRB response data.
'**                 vxmlNode - the node to add the attribute to.
'**                 vblnPurposePurchase - true if we are adding the purchase
'**                 price or valuation estimate for a purchase, else false if we
'**                 are adding the value for a remortgage or further advance.
'**                 rblnIsValuation - true if the value used for the attribute is
'**                 a valuation value else false.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddPurchPriceEstValueAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnPurposePurchase As Boolean, _
        ByRef rblnIsValuation As Boolean, Optional ByVal vstrAttribName As String = "")
    Const cstrFunctionName As String = "AddPurchPriceEstValueAttribute"
    Dim lngValue As Long

    On Error GoTo ErrHandler
    
    Call PurchasePriceOrEstimatedValue(vxmlData, vblnPurposePurchase, rblnIsValuation, _
            lngValue)
    
    If Len(vstrAttribName) = 0 Then
        If vblnPurposePurchase Then
            vstrAttribName = "PURCHASEPRICE"
        Else
            vstrAttribName = "ESTIMATEDVALUE"
        End If
    End If
    
    '*-set the value for the attribute
    Call xmlSetAttributeValue(vxmlNode, vstrAttribName, CStr(lngValue))

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'E2EM00002842 New Sub
'********************************************************************************
'** Function:       AddPurchPriceValueAttribute
'** Created by:     Ian Ross
'** Date:           08/04/2004
'** Description:    Adds the PURCHASEPRICE attribute.
'** Parameters:     vxmlData - the XML source data being the omRB response data.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddPurchPriceValueAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddPurchPriceValueAttribute"
    
    Dim xmlItem As IXMLDOMNode
    Dim ppOrEstimate As Long
    Dim discountAmount As Long
    Dim typeOfApp As String
    
    On Error GoTo ErrHandler
    
    'EP2_583
    Set xmlItem = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        typeOfApp = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
    End If
    Set xmlItem = vxmlData.selectSingleNode(gcstrPROPERTYLOCATION_PATH)
    If Not xmlItem Is Nothing Then
        discountAmount = xmlGetAttributeAsLong(xmlItem, "DISCOUNTAMOUNT")
    End If
       
    
    Set xmlItem = vxmlData.selectSingleNode(gcstrMORTGAGESUBQUOTE_PATH)
    If Not xmlItem Is Nothing Then
        
        ppOrEstimate = xmlGetAttributeAsLong(xmlItem, "PURCHASEPRICEORESTIMATEDVALUE")
        'EP2_583/EP2_1038 If RTB, take off the Discount Amount
        If CheckForValidationType(typeOfApp, "RTB") Then
            ppOrEstimate = ppOrEstimate - discountAmount
        End If
    End If
    
    '*-set the value for the attribute
    Call xmlSetAttributeValue(vxmlNode, "PURCHASEPRICE", ppOrEstimate)
    
    Set xmlItem = Nothing

Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'********************************************************************************
'** Function:       PurchasePriceOrEstimatedValue
'** Created by:     Andy Maggs
'** Date:           18/05/2004
'** Description:    Gets the purchase price or the estimated value from a
'**                 valuation report if applicable.
'** Parameters:     vxmlData - the full KFI input data.
'**                 vblnPurposePurchase - whether the mortgage is for a new
'**                 purchase or not.
'**                 rblnIsEstimated - True if a valuation is returned.
'**                 rlngValue - the valuation or purchase price as appropriate.
'** Returns:        rblnEstimated, rlngValue.
'** Errors:         None Expected
'********************************************************************************
Public Sub PurchasePriceOrEstimatedValue(ByVal vxmlData As IXMLDOMNode, _
        ByVal vblnPurposePurchase As Boolean, ByRef rblnIsEstimated As Boolean, _
        ByRef rlngValue As Long)
    Const cstrFunctionName As String = "PurchasePriceOrEstimatedValue"
    Dim xmlItems As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlHighestVal As IXMLDOMNode
    Dim xmlValRepVal As IXMLDOMNode
    Dim lngPPOrEstimate As Long
    Dim lngValue As Long
    
    On Error GoTo ErrHandler

    '*-are there any valuation instructions?
    Set xmlItems = vxmlData.selectNodes(gcstrVALUATIONINSTRUCTIONS_PATH)
    Set xmlHighestVal = GetElementWithMaxLngAttribValue(xmlItems, "INSTRUCTIONSEQUENCENO")
    If Not xmlHighestVal Is Nothing Then
        'BBG1738/E2EM00002840 Start searching from the selected node
        Set xmlValRepVal = xmlHighestVal.selectSingleNode(".//VALNREPVALUATION")
    End If
    
    ' MAR1061 - Peter Edney - 14/03/2006
    ' Set xmlItem = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    Set xmlItem = vxmlData.selectSingleNode(gcstrMORTGAGESUBQUOTE_PATH)
    If Not xmlItem Is Nothing Then
        lngPPOrEstimate = xmlGetAttributeAsLong(xmlItem, "PURCHASEPRICEORESTIMATEDVALUE")
    End If
    
    If Not xmlValRepVal Is Nothing Then
        lngValue = xmlGetAttributeAsLong(xmlValRepVal, "POSTWORKSVALUATION")
        If lngValue > 0 Then
            rlngValue = lngValue
            rblnIsEstimated = True
        Else
            lngValue = xmlGetAttributeAsLong(xmlValRepVal, "PRESENTVALUATION")
            If vblnPurposePurchase Then
                If (lngValue > 0) And (lngValue < lngPPOrEstimate) Then
                    rlngValue = lngValue
                    rblnIsEstimated = True
                Else
                    rlngValue = lngPPOrEstimate
                    rblnIsEstimated = False
                End If
            Else
                If lngValue > 0 Then
                    rlngValue = lngValue
                    rblnIsEstimated = True
                Else
                    rlngValue = lngPPOrEstimate
                    rblnIsEstimated = False
                End If
            End If
        End If
    Else
        rlngValue = lngPPOrEstimate
        rblnIsEstimated = False
    End If
    Set xmlItems = Nothing
    Set xmlItem = Nothing
    Set xmlHighestVal = Nothing
    Set xmlValRepVal = Nothing
Exit Sub
ErrHandler:
    Set xmlItems = Nothing
    Set xmlItem = Nothing
    Set xmlHighestVal = Nothing
    Set xmlValRepVal = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddRecommendsNotRecommendsElement
'** Created by:     Andy Maggs
'** Date:           14/04/2004
'** Description:    Adds the RECOMMENDS or the NOTRECOMMENDS element according
'**                 to the value of the LEVELOFADVICE field.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddRecommendsNotRecommendsElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddRecommendsNotRecommendsElement"
    Dim xmlItem As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim strValue As String

    On Error GoTo ErrHandler
    
    Set xmlItem = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        strValue = xmlGetAttributeText(xmlItem, "LEVELOFADVICE")
        If InStr(UCase$(strValue), "NADV,") > 0 Then 'SR 25/10/2004 : BBG1684
            Set xmlNewNode = vobjCommon.CreateNewElement("NOTRECOMMENDS", vxmlNode)
        ElseIf InStr(UCase$(strValue), "ADV,") > 0 Then  'SR 25/10/2004 : BBG1684
            Set xmlNewNode = vobjCommon.CreateNewElement("RECOMMENDS", vxmlNode)
        End If
    End If
    
    Set xmlItem = Nothing
    Set xmlNewNode = Nothing
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlNewNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       IsFirstTimeBuyer
'** Created by:     Andy Maggs
'** Date:           14/04/2004
'** Description:    Gets whether the quote is for a first time buyer from the
'**                 TYPEOFBUYER attribute.
'** Parameters:     vxmlData - the input data.
'** Returns:        True if it is a first time buyer else False.
'** Errors:         None Expected
'********************************************************************************
Public Function IsFirstTimeBuyer(ByVal vxmlData As IXMLDOMNode) As Boolean
    Const cstrFunctionName As String = "IsFirstTimeBuyer"
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
    
    'PB 28/06/2006 EP845/EP846 Begin
    Dim xmlFactFind As IXMLDOMNode
    Dim strNatureOfLoan As String
    Dim intActiveQuote As Integer
    Dim xmlActiveQuote As IXMLDOMNode
    Dim xmlLoanComponent As IXMLDOMNode
    Dim intSubQuote As Integer
    Dim xmlSubQuote As IXMLDOMNode
    'PB 28/06/2006 EP845/EP846 End
    
    On Error GoTo ErrHandler
    
    Set xmlItem = vxmlData.selectSingleNode(gcstrAPPLICATION_PATH)
    '
    strValue = xmlGetAttributeText(xmlItem, "TYPEOFBUYER")
    
    '*-return whether this is a first time buyer or not
    
    IsFirstTimeBuyer = False
    If (InStr((UCase$(strValue)), "F") > 0) Then
        IsFirstTimeBuyer = True
    End If
    '
    'PB 29/06/2006 EP845/EP846 Begin
    ' Now check each loan component within the active quote to make sure they aren't an LTB or BTL type.
    ' If they are, this cannot be a first time mortgage.
    Set xmlFactFind = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    intActiveQuote = xmlGetAttributeAsInteger(xmlFactFind, "ACTIVEQUOTENUMBER")
    
    'PSC 09/11/2006 EP2_41 - Start
    If intActiveQuote > 0 Then
        Set xmlActiveQuote = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/QUOTATION[@QUOTATIONNUMBER=" & intActiveQuote & "]")
        'EP2_2203 Won't always be viewing/printing the active quote
        If xmlActiveQuote Is Nothing Then
            Set xmlActiveQuote = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/QUOTATION")
        End If
        If Not xmlActiveQuote Is Nothing Then
            intSubQuote = xmlGetAttributeAsInteger(xmlActiveQuote, "MORTGAGESUBQUOTENUMBER")
            Set xmlSubQuote = xmlActiveQuote.selectSingleNode("MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=" & intSubQuote & "]")
            For Each xmlLoanComponent In xmlSubQuote.selectNodes("LOANCOMPONENT")
                'EP2_1883
                If Not xmlLoanComponent.Attributes.getNamedItem("PURPOSEOFLOAN") Is Nothing Then
                    If CheckForValidationType(xmlLoanComponent.Attributes.getNamedItem("PURPOSEOFLOAN").Text, "BTL") Or _
                       CheckForValidationType(xmlLoanComponent.Attributes.getNamedItem("PURPOSEOFLOAN").Text, "LTB") Then
                        ' This is either Buy To Let or Let To Buy, so cannot be First Time Buyer
                        IsFirstTimeBuyer = False
                        Exit For
                    End If
                End If
            Next xmlLoanComponent
        End If
    End If
    'PSC 09/11/2006 EP2_41 - End
    
    Set xmlLoanComponent = Nothing
    Set xmlSubQuote = Nothing
    Set xmlActiveQuote = Nothing
    Set xmlFactFind = Nothing
    'PB 29/06/2006 EP845/EP846 End
    '
    Set xmlItem = Nothing
   Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function


'********************************************************************************
'** Function:       IsSection3MemoPad
'** Created by:     Steve Badman
'** Date:           26/04/2006
'** Description:    Checks whether there are any MemoPad entries and if they relate to the Section3 type
'** Parameters:     vxmlData - the XML data to search.
'** Returns:        True if it is, else False.
'** Errors:         None Expected
'********************************************************************************
Public Function IsSection3MemoPad(ByVal vxmlData As IXMLDOMNode, ByRef strMemoPadText As String) As Boolean
    Const cstrFunctionName As String = "IsSection3MemoPad"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim blnIsSection3MemoPad As Boolean
    Dim strMemoPadEntryType As String

    On Error GoTo ErrHandler
    
    Set xmlList = vxmlData.selectNodes(gcstrMEMOPAD_PATH)
    blnIsSection3MemoPad = False
    For Each xmlItem In xmlList
        strMemoPadEntryType = xmlGetAttributeText(xmlItem, "ENTRYTYPE_VALIDID")
        '*-if any MemoPad record for Section 3 then it needs to be included
        If strMemoPadEntryType = "S3" Then   'Validation type for Section 3 = "S3"
            strMemoPadText = xmlGetAttributeText(xmlItem, "MEMOENTRY")
            blnIsSection3MemoPad = True
        End If
    Next xmlItem
    
    IsSection3MemoPad = blnIsSection3MemoPad
    Set xmlList = Nothing
    Set xmlItem = Nothing

Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddAPRAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the APR attribute to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddAPRAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddAPRAttribute"
    Dim xmlItems As IXMLDOMNodeList
    Dim dblValue As Double
    Dim strAPRAsString As String

    On Error GoTo ErrHandler
    
    '*-get the value for and add the APR attribute
    Set xmlItems = vobjCommon.LoanComponents
    ' PB 25/05/2006 EP603/MAR1788
    'dblValue = GetMaxDblAttribValue(xmlItems, "APR")
    dblValue = vobjCommon.APR
    ' EP603/MAR1788 End
    'BBG1588 Turn this into a string or the merge will format a numeric to 2 decimal places
    'SR 09/12/2004 : E2EM00003166
    'strAPRAsString = CStr(Round(dblValue, 1)) + "%"
    strAPRAsString = Format(dblValue, "#0.0") + "%"
    'SR 09/12/2004 : E2EM00003166 - End
    'BBG1567 Need to Round APR to 1 Decimal Place
    Call xmlSetAttributeValue(vxmlNode, "APR", strAPRAsString)

    Set xmlItems = Nothing

Exit Sub
ErrHandler:
    Set xmlItems = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetMortgageInterestRateType
'** Created by:     Andy Maggs
'** Date:           06/04/2004
'** Description:    Gets the appropriate mortgage interest rate type by analysing
'**                 the collection of interest types on the loan component.
'** Parameters:     vxmlLoanComponent - the loan component XML to check.
'** Returns:        The mortgage interest rate type.
'** Errors:         None Expected
'********************************************************************************
Public Function GetMortgageInterestRateType(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode) As MortgageInterestRateType
    Const cstrFunctionName As String = "GetMortgageInterestRateType"
    Dim xmlItem As IXMLDOMNode
    Dim eResult As MortgageInterestRateType

    On Error GoTo ErrHandler

    '*-select the first item in the sequence
    Set xmlItem = vobjCommon.GetLoanComponentFirstInterestRate(vxmlLoanComponent)
    If Not xmlItem Is Nothing Then
        eResult = GetInterestRateType(xmlItem)
    End If
    
    '*-return the result
    GetMortgageInterestRateType = eResult

    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddPaybackPerPoundAttribute
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the PAYBACKPERPOUND attribute to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddPaybackPerPoundAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddPaybackPerPoundAttribute"
    Dim dblValue As Double
    Dim strValue As String

    On Error GoTo ErrHandler
    
    If vobjCommon.LoanAmount > 0 Then
        ' PB 25/05/2006 EP603/MAR1788
        'dblValue = CDbl(vobjCommon.LoanAmountRepayable) / CDbl(vobjCommon.LoanAmount)
        dblValue = vobjCommon.AmountPerPound
        ' EP603/MAR1788 End
    End If
    strValue = Format$(dblValue, "0.00")
    Call xmlSetAttributeValue(vxmlNode, "PAYBACKPERPOUND", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddPaymentAttribute
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the PAYMENT attribute to the node.
'** Parameters:     vxmlLoanComponent - the loan component to set the payment for.
'**                 vxmlNode - the node to add the attribute to.
'**                 vstrAttributeName - the name of the attribute to add.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddPaymentAttribute(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vstrAttributeName As String)
    Const cstrFunctionName As String = "AddPaymentAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler
    
    strValue = xmlGetAttributeText(vxmlLoanComponent, "GROSSMONTHLYCOST")
    Call xmlSetAttributeValue(vxmlNode, vstrAttributeName, strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddRateAttribute
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the specified attribute to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlComponent - the loan component element.
'**                 vxmlItem - the interest rate type element.
'**                 vxmlNode - the node to add the attribute to.
'**                 vstrAttribName - the name to use for the attribute.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddRateAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlComponent As IXMLDOMNode, ByVal vxmlItem As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vstrAttribName As String)
    Const cstrFunctionName As String = "AddRateAttribute"
    Dim eType As MortgageInterestRateType
    Dim dblRate As Double

    On Error GoTo ErrHandler
    
'TW 09/11/2005 MAR442
    If Not vxmlItem Is Nothing Then
'TW 09/11/2005 MAR442 End
        eType = GetInterestRateType(vxmlItem)
        Call GetInterestRates(vobjCommon, vxmlComponent, vxmlItem, eType, dblRate)
'TW 09/11/2005 MAR442
    End If
'TW 09/11/2005 MAR442 End
    Call xmlSetAttributeValue(vxmlNode, vstrAttribName, CStr(dblRate))
 
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddNumberOfPaymentsAttribute
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the NUMBEROFPAYMENTS attribute to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        The actual number of payments value.
'** Errors:         None Expected
'********************************************************************************
Public Function AddNumberOfPaymentsAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlRate As IXMLDOMNode, ByVal vxmlNode As IXMLDOMNode) As Integer
    Const cstrFunctionName As String = "AddNumberOfPaymentsAttribute"
    Dim intPeriod As Integer
    
    On Error GoTo ErrHandler
    
    '*-the logic for adding the NUMBEROFPAYMENTS attribute is the same as for
    '*-adding the period attribute, so reuse it.
    intPeriod = AddPeriodAttribute(vxmlRate, "INTERESTRATEPERIOD", _
            "INTERESTRATEENDDATE", vobjCommon.ExpectedCompletionDate, _
            "NUMBEROFPAYMENTS", vxmlNode)
    
    '*-return the actual period for information purposes
    AddNumberOfPaymentsAttribute = intPeriod

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddMortgageCommencementDateAttribute
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the MORTGAGECOMMENCEMENTDATE attribute to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddMortgageCommencementDateAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddMortgageCommencementDateAttribute"
    Dim dtValue As Date
    Dim strValue As String

    On Error GoTo ErrHandler
    
    '*-select the application fact find node
    dtValue = vobjCommon.ExpectedCompletionDate
    strValue = Format$(dtValue, "dd/mm/yyyy")
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGECOMMENCEMENTDATE", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'MAR88 BC - Pass in boolean bCalcDatePlus2Years to calculate Completion Date PLus 2 Years
Public Sub AddCalculationDateAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vbCalcDatePlus2Years As Boolean)
    Const cstrFunctionName As String = "AddCalculationDateAttribute"
    Dim dtValue As Date
    Dim strValue As String
    Dim blnDateFound As Boolean
    Dim xmlNode As IXMLDOMNode

    On Error GoTo ErrHandler
    
    '*-select the application fact find node
    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/REPORTONTITLE")
    If Not xmlNode Is Nothing Then
        If xmlAttributeValueExists(xmlNode, "COMPLETIONDATE") Then
            dtValue = xmlGetAttributeAsDate(xmlNode, "COMPLETIONDATE")
            blnDateFound = True
        End If
    End If
    
    If Not blnDateFound Then
        Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
        If Not xmlNode Is Nothing Then
            If xmlAttributeValueExists(xmlNode, "EXPECTEDCOMPLETIONDATE") Then
                dtValue = xmlGetAttributeAsDate(xmlNode, "EXPECTEDCOMPLETIONDATE")
                blnDateFound = True
            End If
        End If
    End If
    
    If Not blnDateFound Then
        Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrQUOTATION_PATH)
        If Not xmlNode Is Nothing Then
            If xmlAttributeValueExists(xmlNode, "DATEANDTIMEGENERATED") Then
                dtValue = xmlGetAttributeAsDate(xmlNode, "DATEANDTIMEGENERATED")
                dtValue = DateAdd("m", 1, dtValue)
                blnDateFound = True
            End If
        End If
    End If
    
    'MAR88 vbCalcDatePlus2Years included to add 2 years to expected completion date
    
    If vbCalcDatePlus2Years Then
        strValue = Format$(DateAdd("m", 24, dtValue), "dd/mm/yyyy")
    Else
        strValue = Format$(dtValue, "dd/mm/yyyy")
    End If
    
    Call xmlSetAttributeValue(vxmlNode, "CALCULATIONDATE", strValue)

    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'********************************************************************************
'** Function:       GetInterestRateType
'** Created by:     Andy Maggs
'** Date:           16/04/2004
'** Description:    Gets the type of interest rate.
'** Parameters:     vxmlInterestRate - the element containing the data for an
'**                 interestratetype element
'** Returns:        The interest rate type.
'** Errors:         None Expected
'********************************************************************************
Public Function GetInterestRateType(ByVal vxmlInterestRate As IXMLDOMNode) As MortgageInterestRateType
    Const cstrFunctionName As String = "GetInterestRateType"
    Dim strValue As String
    Dim eResult As MortgageInterestRateType
    Dim intRate As Integer
    Dim blnCollared As Boolean
    Dim blnCapped As Boolean

    On Error GoTo ErrHandler

    strValue = xmlGetAttributeText(vxmlInterestRate, "RATETYPE")
    Select Case strValue
        Case "B"
            eResult = mrtStandardVariableRate
            
        Case "D"
            intRate = xmlGetAttributeAsInteger(vxmlInterestRate, "RATE")
            If intRate < 0 Then
                '*-a tracker rate has a negative rate!
                eResult = mrtTrackerAbove
            Else
                eResult = mrtDiscountedRate
            End If
            
        Case "F"
            eResult = mrtFixedRate
            
        Case Else
            '*-this is either capped, collared or both
            '*-if the floored rate is null then this is just a capped rate
            '*-otherwise, if the ceiling rate is null then this is just a
            '*-collared rate, otherwise it is both
            blnCollared = xmlAttributeValueExists(vxmlInterestRate, "FLOOREDRATE")
            blnCapped = xmlAttributeValueExists(vxmlInterestRate, "CEILINGRATE")
            If blnCollared And blnCapped Then
                eResult = mrtCappedAndCollaredRate
            ElseIf blnCollared Then
                eResult = mrtCollaredRate
            Else
                eResult = mrtCappedRate
            End If
    End Select
    
    '*-return the result
    GetInterestRateType = eResult

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetFixedRateTerm
'** Created by:     Andy Maggs
'** Date:           16/04/2004
'** Description:    Gets the term in years and months for which the mortgage has
'**                 a fixed interest rate.
'** Parameters:     vobjCommon - the common data helper.
'**                 rintYears - the years part of the fixed term.
'**                 rintMonths - the months part of the fixed term.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub GetFixedRateTerm(ByVal vobjCommon As CommonDataHelper, _
        ByRef rintYears As Integer, ByRef rintMonths As Integer)
    Const cstrFunctionName As String = "GetFixedRateTerm"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim eType As MortgageInterestRateType
    Dim xmlIntRateList As IXMLDOMNodeList
    Dim xmlIntRate As IXMLDOMNode
    Dim intPeriod As Integer
    Dim intTotal As Integer
    
    'AW     18/01/2005  E2EM00001894
    Dim dtRateEndDate As Date, dtExpectedCompletion As Date
    Dim blnExpectedCompDateFound As Boolean
    Dim xmlNode As IXMLDOMNode
    'AW     18/01/2005  E2EM00001894 - end
    
    On Error GoTo ErrHandler

    '*-initialise byref parameters
    rintYears = 0
    rintMonths = 0

    'AW     18/01/2005  E2EM00001894
    blnExpectedCompDateFound = False
    
    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlNode Is Nothing Then
        If xmlAttributeValueExists(xmlNode, "EXPECTEDCOMPLETIONDATE") Then
            dtExpectedCompletion = xmlGetAttributeAsDate(xmlNode, "EXPECTEDCOMPLETIONDATE")
            blnExpectedCompDateFound = True
        End If
    End If
       
    Set xmlList = vobjCommon.LoanComponents
    
    For Each xmlItem In xmlList
        
        eType = GetMortgageInterestRateType(vobjCommon, xmlItem)
        
        If eType = mrtFixedRate Then
        
            Set xmlIntRateList = xmlItem.selectNodes("//INTERESTRATETYPE")
                
            For Each xmlIntRate In xmlIntRateList
                
                If (xmlGetAttributeText(xmlIntRate, "RATETYPE") = "F") Then
                
                    If xmlAttributeValueExists(xmlIntRate, "INTERESTRATEENDDATE") Then
                    
                        dtRateEndDate = xmlGetAttributeAsDate(xmlIntRate, "INTERESTRATEENDDATE")
                        
                        If (blnExpectedCompDateFound = True) Then
                            
                            ' PB 24/05/2006 EP603/MAR1777
                            'intTotal = intTotal + DateDiff("M", dtExpectedCompletion, dtRateEndDate)
                            intTotal = intTotal + MonthDiff(dtExpectedCompletion, dtRateEndDate)    'MAR1777 GHun
                            ' EP603/MAR1777 End
                            
                        Else
                            
                            ' PB 24/05/2006 EP603/MAR1777
                            'intTotal = intTotal + DateDiff("M", Now, dtRateEndDate)
                            intTotal = intTotal + MonthDiff(Now, dtRateEndDate) 'MAR1777 GHun
                            ' EP603/MAR1777 End
                            
                        End If
                    Else
                        intPeriod = xmlGetAttributeAsInteger(xmlIntRate, "INTERESTRATEPERIOD")
                        If intPeriod > 0 Then
                            intTotal = intTotal + intPeriod
                        End If
                        
                    End If
                    
                End If
                
            Next xmlIntRate
            
         End If
         
    Next xmlItem
    'AW     18/01/2005  E2EM00001894 - end
    rintYears = intTotal \ 12
    rintMonths = intTotal Mod 12

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlIntRateList = Nothing
    Set xmlIntRate = Nothing
    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlIntRateList = Nothing
    Set xmlIntRate = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddVariableFixedElement
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the VARIABLE or FIXED element to the node as appropriate.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddVariableFixedElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlItem As IXMLDOMNode, ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddVariableFixedElement"
    Dim eType As MortgageInterestRateType

    On Error GoTo ErrHandler
    
    eType = GetInterestRateType(vxmlItem)
    If eType = mrtFixedRate Then
        Call vobjCommon.CreateNewElement("FIXED", vxmlNode)
    Else
        Call vobjCommon.CreateNewElement("VARIABLE", vxmlNode)
    End If

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildName
'** Created by:     Andy Maggs
'** Date:           02/04/2004
'** Description:    Builds a name using the supplied values.
'** Parameters:     vstrTitle - the title.
'**                 vstrForenames - the forenames.
'**                 vstrSurname - the surname.
'** Returns:        The full name.
'** Errors:         None Expected
'********************************************************************************
Public Function BuildName(ByVal vstrTitle As String, ByVal vstrForenames As String, _
        ByVal vstrSurname As String) As String
    Const cstrFunctionName As String = "BuildName"
    Dim strName As String

    On Error GoTo ErrHandler
    
    strName = "%1% %2% %3%"
    strName = ReplaceTemplateValue(strName, "%1%", " ", vstrTitle)
    strName = ReplaceTemplateValue(strName, "%2%", " ", vstrForenames)
    strName = ReplaceTemplateValue(strName, "%3%", "", vstrSurname)
    
    BuildName = strName

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildAddress1
'** Created by:     Andy Maggs
'** Date:           02/04/2004
'** Description:    Builds the first address line using the supplied values.
'** Parameters:     vstrFlatNumber - the flat number.
'**                 vstrBuildingName - the name of the building.
'**                 vstrBuildingNumber - the number of the building.
'**                 vstrStreet - the street name.
'** Returns:        The first address line.
'** Errors:         None Expected
'********************************************************************************
Private Function BuildAddress1(ByVal vstrFlatNumber As String, _
        ByVal vstrBuildingName As String, ByVal vstrBuildingNumber As String, _
        ByVal vstrStreet As String) As String
    Const cstrFunctionName As String = "BuildAddress1"
    Dim strAddress As String

    On Error GoTo ErrHandler
    
    strAddress = "%1%, %2%, %3% %4%"
    If Len(vstrFlatNumber) > 0 Then
        vstrFlatNumber = "Flat " & vstrFlatNumber
    End If
    strAddress = ReplaceTemplateValue(strAddress, "%1%", ", ", vstrFlatNumber)
    strAddress = ReplaceTemplateValue(strAddress, "%2%", ", ", vstrBuildingName)
    strAddress = ReplaceTemplateValue(strAddress, "%3%", " ", vstrBuildingNumber)
    strAddress = ReplaceTemplateValue(strAddress, "%4%", "", vstrStreet)
    
    BuildAddress1 = strAddress

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildAddress2
'** Created by:     Andy Maggs
'** Date:           02/04/2004
'** Description:    Builds the second address line using the supplied values.
'** Parameters:     vstrDistrict - the district.
'**                 vstrTown - the town.
'** Returns:        The second address line.
'** Errors:         None Expected
'********************************************************************************
Private Function BuildAddress2(ByVal vstrDistrict As String, _
        ByVal vstrTown As String) As String
    Const cstrFunctionName As String = "BuildAddress2"
    Dim strAddress As String

    On Error GoTo ErrHandler
    
    strAddress = "%1%, %2%"
    strAddress = ReplaceTemplateValue(strAddress, "%1%", ", ", vstrDistrict)
    strAddress = ReplaceTemplateValue(strAddress, "%2%", "", vstrTown)
    
    BuildAddress2 = strAddress

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildAddress3
'** Created by:     Andy Maggs
'** Date:           02/04/2004
'** Description:    Builds the third address line using the supplied values.
'** Parameters:     vstrCounty - the county.
'**                 vstrCountry - the country.
'** Returns:        The second address line.
'** Errors:         None Expected
'********************************************************************************
Private Function BuildAddress3(ByVal vstrCounty As String, _
        ByVal vstrCountry As String) As String
    Const cstrFunctionName As String = "BuildAddress3"
    Dim strAddress As String

    On Error GoTo ErrHandler
    ' SR 22/09/2004 : CORE82 - Do not display country in the address
    'strAddress = "%1%, %2%"
    strAddress = "%1%"
    ' SR 22/09/2004 : CORE82 - End
    
    strAddress = ReplaceTemplateValue(strAddress, "%1%", "", vstrCounty)
    'strAddress = ReplaceTemplateValue(strAddress, "%2%", "", vstrCountry) ' SR 22/09/2004 : CORE82
    
    BuildAddress3 = strAddress

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       ReplaceTemplateValue
'** Created by:     Andy Maggs
'** Date:           29/03/2004
'** Description:    Replaces the value specified by vstrFind in vstrExpression
'**                 with the value in vstrValue, however, if vstrValue is an
'**                 empty string then the method has the effect of removing the
'**                 template item altogether.
'** Parameters:     vstrExpression - the expression containing the value to
'**                 Replace.
'**                 vstrFind - the value to replace.
'**                 vstrValue - the value to replace vstrFind with.
'** Returns:        The modified vstrExpression.
'** Errors:         None Expected
'********************************************************************************
Private Function ReplaceTemplateValue(ByVal vstrExpression As String, _
        ByVal vstrFind As String, ByVal vstrSeparator As String, _
        ByVal vstrValue As String) As String
    Const cstrFunctionName As String = "ReplaceTemplateValue"

    On Error GoTo ErrHandler
    
    If Len(vstrValue) = 0 Then
        '*-remove the template item altogether, including the separator
        vstrFind = vstrFind & vstrSeparator
    End If
    ReplaceTemplateValue = Replace(vstrExpression, vstrFind, vstrValue)

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddProductNameAttribute
'** Created by:     Andy Maggs
'** Date:           05/04/2004
'** Description:    Adds the product name attribute to the node.
'** Parameters:     vxmlData - the input data representing a single LOANCOMPONENT
'**                 node.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddProductNameAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, _
        ByVal vobjCommon As CommonDataHelper)
    Const cstrFunctionName As String = "AddProductNameAttribute"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlMPLang As IXMLDOMNode
    Dim strValue As String
    Dim dtStartDate As Date
    Dim dtStoredDate As Date
    Dim xmlSpecialGroup As IXMLDOMNode 'EP1120
    Dim strSpecialGroup As String 'EP1120
    Dim strProductCode As String 'EP1120
    Dim strProductName As String 'EP1238
        
    On Error GoTo ErrHandler
    
    'MAR54 read MORTGAGEPRODUCTLANGUAGE from a list, checking for latest date
    '*-get the latest product language node
    Set xmlList = vxmlData.selectNodes("MORTGAGEPRODUCT/MORTGAGEPRODUCTLANGUAGE")
   'Set xmlMPLang = vxmlData.selectSingleNode("MORTGAGEPRODUCT/MORTGAGEPRODUCTLANGUAGE")
    For Each xmlMPLang In xmlList
        dtStartDate = xmlGetAttributeAsDate(xmlMPLang, "STARTDATE")
        If dtStartDate > dtStoredDate Then
            'SAB 26/04/2006 - EPSOM EP454 : START
            'Use the PRODUCTTEXTDETAILS column rather than PRODUCTNAME
            'strValue = xmlGetAttributeText(xmlMPLang, "PRODUCTNAME")
            strValue = xmlGetAttributeText(xmlMPLang, "PRODUCTTEXTDETAILS")
            strProductName = strValue 'EP1238
            
            'Peter Edney - 05/09/2006 - EP1120
            Set xmlSpecialGroup = vxmlData.selectSingleNode("MORTGAGEPRODUCT/SPECIALGROUP")
            If Not (xmlSpecialGroup Is Nothing) Then
                'LH 13/10/2006 - EP1226
                strSpecialGroup = xmlGetAttributeText(xmlSpecialGroup, "GROUPTYPESEQUENCENUMBER_TEXT", "")
            Else
                strSpecialGroup = vobjCommon.GetSpecialSchemeGroupType
            End If
            strProductCode = xmlGetAttributeText(xmlMPLang, "MORTGAGEPRODUCTCODE")
            
            If strSpecialGroup <> "" Then
                strValue = strValue & ", " & strSpecialGroup
            End If
            
            If strProductCode <> "" Then
                strValue = strValue & " (Product Code " & strProductCode & ")"
            End If
            
            'SAB 26/04/2006 - EPSOM EP454 : END
            Call xmlSetAttributeValue(vxmlNode, "PRODUCTNAME", strValue)
            Call xmlSetAttributeValue(vxmlNode, "PRODUCTNAMESHORT", strProductName) 'EP1238
        End If
        dtStoredDate = dtStartDate
    Next xmlMPLang
    
    Set xmlMPLang = Nothing
Exit Sub
ErrHandler:
    Set xmlMPLang = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddInterestRateTypeAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the RATETYPE attribute to the node.
'** Parameters:     vxmlData - the input data.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddInterestRateTypeAttribute(ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddInterestRateTypeAttribute"
    Dim eType As MortgageInterestRateType
    Dim strValue As String

    On Error GoTo ErrHandler
'TW 09/11/2005 MAR442
    If vxmlInterestRate Is Nothing Then
        strValue = ""
    Else
'TW 09/11/2005 MAR442 End
    eType = GetInterestRateType(vxmlInterestRate)
    Select Case eType
        Case mrtStandardVariableRate
            strValue = "Variable"
        Case mrtFixedRate
            strValue = "Fixed"
        Case mrtDiscountedRate
            strValue = "Discounted"
        Case Else
            strValue = "Capped/Collared"

    End Select
'TW 09/11/2005 MAR442
    End If
'TW 09/11/2005 MAR442 End
    
    Call xmlSetAttributeValue(vxmlNode, "RATETYPE", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
Public Sub AddInterestRateTypeAttributeSect6(ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddInterestRateTypeAttribute"
    Dim eType As MortgageInterestRateType
    Dim strValue As String

    On Error GoTo ErrHandler
    If vxmlInterestRate Is Nothing Then
        strValue = ""
    Else
    eType = GetInterestRateType(vxmlInterestRate)
    Select Case eType
        Case mrtFixedRate
            strValue = "Fixed"
        Case Else
            strValue = "Variable"

    End Select
    End If
    
    Call xmlSetAttributeValue(vxmlNode, "RATETYPE", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'********************************************************************************
'** Function:       AddRepaymentMethodAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the REPAYMENTMETHOD attribute to the node.
'** Parameters:     vxmlLoanComponent - the loan component.
'**                 vxmlNode - the node to add the attribute to.
'**                 vstrAttribName - the optional attribute name.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddRepaymentMethodAttribute(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode, _
        Optional ByVal vstrAttribName As String = "REPAYMENTMETHOD")
    Const cstrFunctionName As String = "AddRepaymentMethodAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler
    
    strValue = xmlGetAttributeText(vxmlLoanComponent, "REPAYMENTMETHOD")

    If (InStr(1, strValue, "1", vbTextCompare)) Then
        strValue = "Interest Only"
    ElseIf (InStr(1, strValue, "2", vbTextCompare)) Then
        'PB 12/06/2006 EP730/MAR1831 Begin
        'strValue = "Capital and Interest"
         strValue = "Repayment" 'SR EP2_1938 - Use 'Repayment' instead of 'Repayment (Capital And Interest)
        'PB EP730/MAR1831 End
    ElseIf (InStr(1, strValue, "3", vbTextCompare)) Then
        strValue = "Part and Part"
    End If
        
    Call xmlSetAttributeValue(vxmlNode, vstrAttribName, strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddTermInMonthsAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the TERMINMONTHS attribute to the node.
'** Parameters:     vxmlLoanComponent - the loan component.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddTermInMonthsAttribute(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddTermInMonthsAttribute"
    Dim intYears As Integer
    Dim intMonths As Integer

    On Error GoTo ErrHandler
    
    intYears = xmlGetAttributeAsInteger(vxmlLoanComponent, "TERMINYEARS")
    intMonths = xmlGetAttributeAsInteger(vxmlLoanComponent, "TERMINMONTHS")
    intMonths = (12 * intYears) + intMonths
    Call xmlSetAttributeValue(vxmlNode, "TERMINMONTHS", CStr(intMonths))

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddLoanComponentAmountAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the LOANCOMPONENTAMOUNT attribute to the node.
'** Parameters:     vxmlLoanComponent - the loan component.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddLoanComponentAmountAttribute(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddLoanComponentAmountAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler
    'SR EP2_2341 : user TotalLoanComponentAmount instead of LOANAMOUNT
    strValue = xmlGetAttributeText(vxmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT")
    Call xmlSetAttributeValue(vxmlNode, "LOANCOMPONENTAMOUNT", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddComponentPartAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the PART attribute to the node.
'** Parameters:     vxmlComponent - the loan or balanceschedule component.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddComponentPartAttribute(ByVal vxmlComponent As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddComponentPartAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler
    
    strValue = xmlGetAttributeText(vxmlComponent, "LOANCOMPONENTSEQUENCENUMBER")
    Call xmlSetAttributeValue(vxmlNode, "PART", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddIntRatePaymentAttribute
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Adds the PAYMENT attribute to the node.
'** Parameters:     vxmlLoanComponent - the loan component to set the payment for.
'**                 vxmlNode - the node to add the attribute to.
'**                 vstrAttributeName - the name of the attribute to add.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddIntRatePaymentAttribute(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vxmlRate As IXMLDOMNode, ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddIntRatePaymentAttribute"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim intRateSeqNum As Integer
    Dim intItemSeqNum As Integer
    Dim strValue As String

    On Error GoTo ErrHandler
    
    '*-get the sequence number of this interest rate
    intRateSeqNum = xmlGetAttributeAsInteger(vxmlRate, "INTERESTRATETYPESEQUENCENUMBER")
    
    '*-get the list of loan component payment schedule records
    Set xmlList = vxmlLoanComponent.selectNodes("//LOANCOMPONENTPAYMENTSCHEDULE")
    For Each xmlItem In xmlList
        intItemSeqNum = xmlGetAttributeAsInteger(xmlItem, "INTERESTRATETYPESEQUENCENUMBER")
        If intItemSeqNum = intRateSeqNum Then
            strValue = xmlGetAttributeText(xmlItem, "MONTHLYCOST")
            Exit For
        End If
    Next xmlItem
    Call xmlSetAttributeValue(vxmlNode, "PAYMENT", strValue)

    Set xmlList = Nothing
    Set xmlItem = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddNewAmountAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the NEWAMOUNT attribute to the node.
'** Parameters:     vxmlBalance - the balanceschedule component.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddNewAmountAttribute(ByVal vxmlBalance As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddNewAmountAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler

'MAR54 Retrieve MonthlyCost from LoanComponentPaymentSchedule
'MAR54 rather than Balance from LoanComponentBalanceSchedule
'   strValue = xmlGetAttributeText(vxmlBalance, "BALANCE")
    strValue = xmlGetAttributeText(vxmlBalance, "MONTHLYCOST")
    Call xmlSetAttributeValue(vxmlNode, "NEWAMOUNT", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddOfferTypeAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the OFFERTYPE attribute to the node.
'** Parameters:     vxmlInterestRate - the interest rate element.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddOfferTypeAttribute(ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddOfferTypeAttribute"
    Dim strValue As String

    On Error GoTo ErrHandler
    
'TW 09/11/2005 MAR442
    If vxmlInterestRate Is Nothing Then
        strValue = ""
    Else
'TW 09/11/2005 MAR442 End
    
    strValue = xmlGetAttributeText(vxmlInterestRate, "RATETYPE")
    Select Case strValue
        Case "F"
            strValue = "Fixed Rate"
        Case "D"
            strValue = "Discount Rate"
'MAR54
        Case "B"
            strValue = "Variable Rate"
        Case "C"
            strValue = "Capped/Collared Rate"
        Case Else
            strValue = ""
    End Select
'TW 09/11/2005 MAR442
    End If
'TW 09/11/2005 MAR442 End
    Call xmlSetAttributeValue(vxmlNode, "OFFERTYPE", strValue)

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'SR EP2_2159
Public Function GetProductRateType(ByVal xmlLoanComponent As IXMLDOMNode) As String
    
    Const cstrFunctionName As String = "GetProductRateType"
    
    Dim xmlInterestRateType As IXMLDOMNode
    Dim strValue As String
    
    Set xmlInterestRateType = xmlLoanComponent.selectSingleNode( _
                    ".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]")
    If xmlInterestRateType Is Nothing Then
        strValue = ""
    Else
        strValue = xmlGetAttributeText(xmlInterestRateType, "RATETYPE")
        Select Case strValue
            Case "F"
                strValue = "Fixed Rate"
            Case "D"
                strValue = "Discount Rate"
            Case "B"
                strValue = "Variable Rate"
            Case "C"
                strValue = "Capped/Collared Rate"
            Case Else
                strValue = ""
        End Select
    End If
    
    GetProductRateType = strValue
Exit Function

ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function
'SR EP2_2159 - End


'********************************************************************************
'** Function:       AddRatePeriodAttribute
'** Created by:     Andy Maggs
'** Date:           15/04/2004
'** Description:    Adds the RATEPERIOD attribute to the node.
'** Parameters:     vxmlInterestRate - the interest rate element.
'**                 vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddRatePeriodAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlInterestRate As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddRatePeriodAttribute"
    
    'MAR54
    Dim xmlNewNode As IXMLDOMNode
    Dim strValue As String

    On Error GoTo ErrHandler
    
'TW 09/11/2005 MAR442
    If Not vxmlInterestRate Is Nothing Then
'TW 09/11/2005 MAR442 End
        If xmlAttributeValueExists(vxmlInterestRate, "INTERESTRATEPERIOD") Then
            strValue = xmlGetAttributeText(vxmlInterestRate, "INTERESTRATEPERIOD")
            Set xmlNewNode = vobjCommon.CreateNewElement("RATEPERIOD", vxmlNode)
            Call xmlSetAttributeValue(xmlNewNode, "RATEPERIOD", strValue)
        ElseIf xmlAttributeValueExists(vxmlInterestRate, "INTERESTRATEENDDATE") Then
            strValue = xmlGetAttributeText(vxmlInterestRate, "INTERESTRATEENDDATE")
            Set xmlNewNode = vobjCommon.CreateNewElement("RATEENDDATE", vxmlNode)
            Call xmlSetAttributeValue(xmlNewNode, "PERIODENDDATE", strValue)
        End If
'TW 09/11/2005 MAR442
    End If
'TW 09/11/2005 MAR442 End
    'MAR54 Add RATEPERIOD Element


Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddSecurityAddress
'** Created by:     Andy Maggs
'** Date:           24/05/2004
'** Description:    Adds the address of the property securing the mortgage.
'** Parameters:     vxmlData - the full input data.
'**                 vxmNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddSecurityAddress(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddSecurityAddress"
    Dim xmlAddress As IXMLDOMNode
    Dim strFlatNumber As String
    Dim strAddress As String

    On Error GoTo ErrHandler
   'SR 16/09/2004 : CORE82
    'PB 20/06/2006 EP700 Start
    'Set xmlAddress = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS")
    'PB 29/06/2006 EP700
    'Set xmlAddress = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/VALNPROPERTYDETAILS[last()]")
    Set xmlAddress = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/VALNREPPROPERTYDETAILS[last()]")
    'PB 29/06/2006 EP700 End
    'PB EP700 End
    
    'PB 30/06/2006 EP700
    ' Test for postcode; if none, don't use this address as it is probably blank
    If xmlGetAttributeText(xmlAddress, "POSTCODE") = "" Then
        Set xmlAddress = Nothing
    End If
    'PB EP700 End
    
    'PB 26/06/2006 EP773 Begin
    If xmlAddress Is Nothing Then
        ' Get 'New property' address instead
        Set xmlAddress = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS")
    End If
    'PB EP773 End
    
    If Not xmlAddress Is Nothing Then
        '*-extract the address data
        strFlatNumber = xmlGetAttributeText(xmlAddress, "FLATNUMBER")
        
        '*-and write it into a single line
        strAddress = "%1%, %2%, %3% %4%, %5%, %6%, %7%, %8%"
        If Len(strFlatNumber) > 0 Then
            strFlatNumber = "Flat " & strFlatNumber
        End If
        strAddress = ReplaceTemplateValue(strAddress, "%1%", ", ", strFlatNumber)
        strAddress = ReplaceTemplateValue(strAddress, "%2%", ", ", _
                xmlGetAttributeText(xmlAddress, "BUILDINGORHOUSENAME"))
        strAddress = ReplaceTemplateValue(strAddress, "%3%", " ", _
                xmlGetAttributeText(xmlAddress, "BUILDINGORHOUSENUMBER"))
        strAddress = ReplaceTemplateValue(strAddress, "%4%", ", ", _
                xmlGetAttributeText(xmlAddress, "STREET"))
        strAddress = ReplaceTemplateValue(strAddress, "%5%", ", ", _
                xmlGetAttributeText(xmlAddress, "DISTRICT"))
        strAddress = ReplaceTemplateValue(strAddress, "%6%", ", ", _
                xmlGetAttributeText(xmlAddress, "TOWN"))
        strAddress = ReplaceTemplateValue(strAddress, "%7%", ", ", _
                xmlGetAttributeText(xmlAddress, "COUNTY"))
        strAddress = ReplaceTemplateValue(strAddress, "%8%", "", _
                xmlGetAttributeText(xmlAddress, "POSTCODE"))
                
        ' ik_20041222_E2EM00002921 trim output (can be left with trailing ", ")
        strAddress = Trim(strAddress)
        If Right(strAddress, 1) = "," Then
            strAddress = Left(strAddress, Len(strAddress) - 1)
        End If
        ' ik_20041222_E2EM00002921_ends
        
        Call xmlSetAttributeValue(vxmlNode, "OFFERSECURITYADDRESS", strAddress) 'SR 24/09/2004 : CORE82
    End If
    
    Set xmlAddress = Nothing

Exit Sub
ErrHandler:
    Set xmlAddress = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'********************************************************************************
'** Function:       AddIntermediaryContactDetails
'** Created by:     Pat Morse
'** Date:           19/05/2006
'** Description:    Adds Broker details to INTERMEDIARY element:-
'**                     INTERMEDIARYNAME
'**                     INTERMEDIARYADDRESS
'**                 Adds the following to the TEMPLATE element
'**                     BROKERNAME
'**                     PACKAGERNAME
'**
'**                 EPSOM EP584 - rewritten from previous version
'** Parameters:     vxmlData - the full input data.
'**                 vxmNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
'EP2_139 Offer and KFI changes
'EP2_2057 rewrite
Public Sub AddIntermediaryContactDetails( _
    ByVal vobjCommon As CommonDataHelper _
    , ByVal vxmlData As IXMLDOMNode _
    , ByVal vxmlNode As IXMLDOMNode _
    , ByVal blnHorizontal As Boolean _
    , Optional ByVal blnTelephone As Boolean = False)
    
    Const cstrFunctionName As String = "AddIntermediaryContactDetails"
    
    Dim xmlIntermediaryInfo     As IXMLDOMNode
    Dim xmlIntroducerAddress  As IXMLDOMNode
    Dim strIntroducerName        As String
    Dim strCompanyName        As String
    Dim contactTelephone        As String
    Dim xmlARFirm  As IXMLDOMNode
    Dim xmlPrincipalFirm  As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Set xmlIntroducerAddress = vobjCommon.Data.selectSingleNode(gcstrINTRODUCER_PATH)
    Set xmlARFirm = vobjCommon.Data.selectSingleNode(gcstrAPPINTRODUCER_PATH & "/ARFIRM")
    'SR EP2_2270 - Get the principal Firm which is not a packager to get the address
    Set xmlPrincipalFirm = vobjCommon.Data.selectSingleNode(gcstrAPPINTRODUCER_PATH & "/PRINCIPALFIRM[not(@PACKAGERIND) or not(@PACKAGERIND=0)]")
    If Not xmlIntroducerAddress Is Nothing Then
        'set the broker name
        strIntroducerName = vobjCommon.IntroducerName
        strCompanyName = vobjCommon.BrokerName
        If Len(strIntroducerName) > 0 Then
            '*-create the main element
            Set xmlIntermediaryInfo = vobjCommon.CreateNewElement("INTERMEDIARY", vxmlNode)
            
            Call xmlSetAttributeValue(xmlIntermediaryInfo, "INTERMEDIARYCOMBINEDNAME", strIntroducerName)
            Call xmlSetAttributeValue(xmlIntermediaryInfo, "INTERMEDIARYCOMPANYNAME", strCompanyName)
            If Not xmlARFirm Is Nothing Then
                Call AddIntermedAddressAttributes(xmlARFirm, xmlIntermediaryInfo, "INTERMEDIARYADDRESS", blnTelephone)
            ElseIf Not xmlPrincipalFirm Is Nothing Then
                Call AddIntermedAddressAttributes(xmlPrincipalFirm, xmlIntermediaryInfo, "INTERMEDIARYADDRESS", blnTelephone)
            End If
        End If
    End If
    Set xmlIntroducerAddress = Nothing
    Set xmlIntermediaryInfo = Nothing
    Set xmlARFirm = Nothing
    Set xmlPrincipalFirm = Nothing

Exit Sub

ErrHandler:
    Set xmlIntroducerAddress = Nothing
    Set xmlIntermediaryInfo = Nothing
    Set xmlARFirm = Nothing
    Set xmlPrincipalFirm = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'PM 19/05/2006 : EPSOM - EP584 : End

'********************************************************************************
'** Function:       AddLegalRepContactDetails
'** Created by:     Steve Badman
'** Date:           21/04/2006
'** Description:    Adds the address of the solicitor
'** Parameters:     vobjCommon - the common data helper
'**                 vxmlData - the full input data
'**                 vxmNode - the node to add the attribute to
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddLegalRepContactDetails(ByVal vobjCommon As CommonDataHelper, _
    ByVal vxmlData As IXMLDOMNode, _
    ByVal vxmlNode As IXMLDOMNode)
    
    Const cstrFunctionName As String = "AddLegalRepContactDetails"
    Dim xmlLegalInfo As IXMLDOMNode
    Dim xmlLegal As IXMLDOMNode
    Dim xmlCompany As IXMLDOMNode
    Dim xmlLegalAddress As IXMLDOMNode
    Dim xmlLegalContact As IXMLDOMNode
    Dim strTemp                 As String
    Dim strCompanyName          As String
    Dim strLegalName            As String
    Dim strLegalCombinedName    As String
    
    On Error GoTo ErrHandler
   
    
    Set xmlLegal = vobjCommon.Data.selectSingleNode(gcstrLEGALREP_PATH)
    If xmlLegal Is Nothing Then Exit Sub
    
    '*-create the main element
    Set xmlLegalInfo = vobjCommon.CreateNewElement("SEPARATELEGALREPRESENTATIVE", vxmlNode)
    
    Set xmlCompany = xmlLegal.selectSingleNode("//LEGALREP/THIRDPARTY")
    If xmlCompany Is Nothing Then
        Set xmlCompany = xmlLegal.selectSingleNode("//LEGALREP/NAMEANDADDRESSDIRECTORY")
    End If
    
    If Not xmlCompany Is Nothing Then
        'EPSOM EP610 - 25/05/2006 Pat Morse - START
        
        '*-extract the Company Name data
        strCompanyName = ""
        strCompanyName = xmlGetAttributeText(xmlCompany, "COMPANYNAME")
        
        '*-get a handle on the contact details node
        Set xmlLegalContact = xmlCompany.selectSingleNode("CONTACTDETAILS")
        
        '*-build Contact Name and Combined Name
        '*- Contact Name = contact name if present else company name
        '*- Combined Name = combined contact name with a new line and company name if both are present
        '*-                 else contact name or company name
        Call BuildContactNames(vobjCommon, xmlLegalContact, strCompanyName, strLegalName, strLegalCombinedName, True)
                
        If Not strLegalName = "" Then _
            Call xmlSetAttributeValue(xmlLegalInfo, "SOLICITORNAME", strLegalName)
            
        If Not strLegalCombinedName = "" Then _
            Call xmlSetAttributeValue(xmlLegalInfo, "SOLICITORCOMBINEDNAME", strLegalCombinedName)
            
        
        '*-extract the Address data
        Set xmlLegalAddress = xmlCompany.selectSingleNode("ADDRESS")
        If Not xmlLegalAddress Is Nothing Then
'PM 25/05/2006 EPSOM EP610
            Call AddAddressAttributes(xmlLegalAddress, xmlLegalInfo, "SOLICITORADDRESS", True)
        End If
    End If
    
    Set xmlLegalInfo = Nothing
    Set xmlLegal = Nothing
    Set xmlCompany = Nothing
    Set xmlLegalAddress = Nothing
    Set xmlLegalContact = Nothing

Exit Sub

ErrHandler:
    Set xmlLegalInfo = Nothing
    Set xmlLegal = Nothing
    Set xmlCompany = Nothing
    Set xmlLegalAddress = Nothing
    Set xmlLegalContact = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'PM 26/05/2006 EPSOM EP610   Start
'********************************************************************************
'** Function:       BuildContactNames
'** Created by:     Pat Morse
'** Date:           26/05/2006
'** Description:    Adds the address attributes from the node to the specified
'**                 node.
'** Parameters:     vobjCommon As CommonDataHelper,
'**                 vxmlContact - the node containing the contact data
'**                 istrCompanyName - the input company name
'**                 ostrContactName - the output contact name or company name
'**                 ostrCombinedName - the output contact name and new line with company name if present
'**                 vblnFlat - Return the names on one line.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildContactNames(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlContactDetails As IXMLDOMNode, _
        ByVal istrCompanyName As String, _
        ByRef ostrContactName As String, _
        ByRef ostrCombinedName As String, _
        Optional ByVal vblnFlat As Boolean = False)
        
    Const cstrFunctionName As String = "BuildContactNames"
        
        
    On Error GoTo ErrHandler
        
        
    ostrContactName = ""
    ostrCombinedName = ""
        
    If Not vxmlContactDetails Is Nothing Then
        'PB 18/08/2006 EP1082 Begin
        If UCase(xmlGetAttributeText(vxmlContactDetails, "CONTACTTITLE_TEXT")) = "OTHER" Then
            ostrContactName = vobjCommon.StandardContactName( _
                    xmlGetAttributeText(vxmlContactDetails, "CONTACTTITLEOTHER"), _
                    xmlGetAttributeText(vxmlContactDetails, "CONTACTFORENAME"), _
                    xmlGetAttributeText(vxmlContactDetails, "CONTACTSURNAME"))
        Else
            ostrContactName = vobjCommon.StandardContactName( _
                    xmlGetAttributeText(vxmlContactDetails, "CONTACTTITLE_TEXT"), _
                    xmlGetAttributeText(vxmlContactDetails, "CONTACTFORENAME"), _
                    xmlGetAttributeText(vxmlContactDetails, "CONTACTSURNAME"))
        End If
        'PB EP1082 End
    End If
                
    ostrCombinedName = ostrContactName
    If ostrCombinedName = "" Then
        ostrCombinedName = istrCompanyName
    Else
        If Not ostrContactName = "" Then
            If Not vblnFlat Then
                ostrCombinedName = ostrCombinedName + "\par " + istrCompanyName
            Else
                ostrCombinedName = ostrCombinedName + ", " + istrCompanyName
            End If
        End If
    End If
    
    
Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'PM 26/05/2006 EPSOM EP610   End


'********************************************************************************
'** Function:       AddMortgageTermAttributes
'** Created by:     Andy Maggs
'** Date:           25/05/2004
'** Description:    Adds the TERMINYEARS and TERMINMONTHS attributes.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attributes to.
'**                 vstrYearsAttribName - the name of the term in years attribute.
'**                 vstrMonthsAttribName - the name of the term in months attribute.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddMortgageTermAttributes(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, _
        Optional ByVal vstrYearsAttribName As String = "TERMINYEARS", _
        Optional ByVal vstrMonthsAttribName As String = "TERMINMONTHS")
    
    Dim xmlNewNode As IXMLDOMNode 'BC MAR1685 01/05/2006
    
    Const cstrFunctionName As String = "AddMortgageTermAttributes"

    On Error GoTo ErrHandler
    
    Call xmlSetAttributeValue(vxmlNode, vstrYearsAttribName, vobjCommon.TermInYears)
    Call xmlSetAttributeValue(vxmlNode, vstrMonthsAttribName, vobjCommon.TermInMonths)
    
    'Pat Morse EP584 18/06/2006 - START. We require TermInMonths without additional <MONTHS> sub-element
    'BC MAR1685 01/05/2006
'    If Not (vobjCommon.TermInMonths = 0) Then
'        Set xmlNewNode = vobjCommon.CreateNewElement("MONTHS", vxmlNode)
'        Call xmlSetAttributeValue(xmlNewNode, vstrMonthsAttribName, vobjCommon.TermInMonths)
'    End If
    'Pat Morse EP584 18/06/2006 - END

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddRestrictionsElements
'** Created by:     Andy Maggs
'** Date:           23/04/2004
'** Description:    Adds the RESTRICTIONS and child elements to the node.
'** Parameters:     vxmlList - the list of loan components.
'**                 vintLowestSeqNum - the lowest sequence number used for loan
'**                 component records, this is required to rebase the PART numbers
'**                 so that they always start at 1.
'**                 vxmlNode - the node to add the element(s) to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddRestrictionsElements(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlList As IXMLDOMNodeList, ByVal vintLowestSeqNum As Integer, _
        ByVal vxmlNode As IXMLDOMNode, ByVal dblMaxPercent As Double)
    Const cstrFunctionName As String = "AddRestrictionsElements"
    Dim xmlRestrictions As IXMLDOMNode
    Dim xmlItem As IXMLDOMNode
    Dim xmlPart As IXMLDOMNode
    Dim intPart As Integer
    Dim xmlNode As IXMLDOMNode
    Dim xmlProdRest As IXMLDOMNode
    Dim strValue As String
    Dim xmlAddRestrict As IXMLDOMNode
    Dim blnAddFinancialElement As Boolean
    'CORE82
    Dim xmlLTV As IXMLDOMNode
    Dim xmlPurpose As IXMLDOMNode
    'BBG1545
    Dim strSpecialGroup As String

    On Error GoTo ErrHandler
    
'SR 25/10/2004 : BBG1684 - BBG do not need these restrictions at the moment.
    '*-add the RESTRICTIONS element
'    Set xmlRestrictions = vobjCommon.CreateNewElement("RESTRICTIONS", vxmlNode)
'
'    'CORE82 Moved LTV and PURPOSEPURCHASE from AddMortgageComponentDetails
'    '*-add the LTV element
''    Set xmlLTV = vobjCommon.CreateNewElement("LTV", xmlRestrictions)
''    '*-add the MAXPERCENT attribute
''    Call xmlSetAttributeValue(xmlLTV, "MAXPERCENT", CStr(dblMaxPercent))
'
''    If vobjCommon.LoanPurposeText = "PURPOSEPURCHASE" Then
''        '*-add the PURPOSEPURCHASE element
''        Set xmlPurpose = vobjCommon.AddLoanPurposeElement(xmlLTV)
''    End If
'
'    '*-add the PRODUCTRESTRICTIONS element
'    Set xmlProdRest = vobjCommon.CreateNewElement("PRODUCTRESTRICTIONS", xmlRestrictions)
'    '*-add the PART element for each loan component
'    For Each xmlItem In vxmlList
'        Set xmlPart = vobjCommon.CreateNewElement("PART", xmlProdRest)
'        '*-add the RESTRICTION attribute
'        strValue = GetRestrictionText(xmlItem, "MORTGAGEPRODUCTTYPEOFBUYER", _
'                "TYPEOFBUYER_TEXT")
'        Call xmlSetAttributeValue(xmlPart, "RESTRICTION", strValue)
'        '*-add the RESTRICTIONTYPE attribute
'        Call xmlSetAttributeValue(xmlPart, "RESTRICTIONTYPE", "Type of borrower")
'        '*-add the PART attribute
'        'SR 16/09/2004 : CORE82 - Add PART element only when the loan has multiple components
'        If vobjCommon.LoanComponents.length > 1 Then
'            intPart = xmlGetAttributeAsInteger(xmlItem, "LOANCOMPONENTSEQUENCENUMBER") - vintLowestSeqNum
'            Call xmlSetAttributeValue(xmlPart, "PART", CStr(intPart))
'        End If 'SR 16/09/2004 : CORE82 - End
'
'        '*-first add the employment additional restrictions
'        '*-add the ADDITIONALRESTRICTIONTYPE element
'        Set xmlAddRestrict = vobjCommon.CreateNewElement("ADDITIONALRESTRICTIONTYPE", xmlPart)
'        '*-add the RESTRICTION attribute
'        strValue = GetRestrictionText(xmlItem, "MORTGAGEPRODUCTEMPLOYMENT", _
'                "EMPLOYMENTSTATUS_TEXT")
'        Call xmlSetAttributeValue(xmlAddRestrict, "RESTRICTION", strValue)
'        '*-add the RESTRICTIONTYPE attribute
'        Call xmlSetAttributeValue(xmlAddRestrict, "RESTRICTIONTYPE", "Type of employment")
'
'        '*-finally add the property location restrictions
'        '*-add the ADDITIONALRESTRICTIONTYPE element
'        Set xmlAddRestrict = vobjCommon.CreateNewElement("ADDITIONALRESTRICTIONTYPE", xmlPart)
'        '*-add the RESTRICTION attribute
'        strValue = GetRestrictionText(xmlItem, "MORTGAGEPRODUCTPROPLOCATION", _
'                "MORTGAGEPRODUCTLOCATIONTYPE_TEXT")
'        Call xmlSetAttributeValue(xmlAddRestrict, "RESTRICTION", strValue)
'        '*-add the RESTRICTIONTYPE attribute
'        Call xmlSetAttributeValue(xmlAddRestrict, "RESTRICTIONTYPE", "Property location(s)")
'
'' BBG1545 Use SpecialGroup, & search for adverse credit
'' IMPAIREDCREDIT not linked to the product search.
'        '*-get the mortgage product
''        Set xmlNode = xmlItem.selectSingleNode("//MORTGAGEPRODUCT")
''        If Not xmlNode Is Nothing Then
''            If xmlGetAttributeAsBoolean(xmlNode, "IMPAIREDCREDIT") Then
''                blnAddFinancialElement = True
''            End If
''        End If
'
'    Next xmlItem
'SR 25/10/2004 : BBG1684 - End ===================================================
    
' BBG1545 *-get the APPLICATIONFACTFIND Node
    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlNode Is Nothing Then
        'SR 27/10/2004 : BBG1724
        strSpecialGroup = xmlGetAttributeText(xmlNode, "SPECIALGROUP")
        Set xmlRestrictions = vobjCommon.CreateNewElement("RESTRICTIONS", vxmlNode)
        If InStr(1, strSpecialGroup, "HPC", vbTextCompare) Then
            Call vobjCommon.CreateNewElement("PRODUCT100", xmlRestrictions)
        End If
        
        If InStr(1, strSpecialGroup, "MAX130", vbTextCompare) Then
            Call vobjCommon.CreateNewElement("PRODUCTMAX130", xmlRestrictions)
        End If
        
        If InStr(1, strSpecialGroup, "AC", vbTextCompare) Then
            Call vobjCommon.CreateNewElement("FINANCIALSTATUS", xmlRestrictions)
        End If
    End If
    'SR 27/10/2004 : BBG1724 - End
    Set xmlRestrictions = Nothing
    Set xmlItem = Nothing
    Set xmlPart = Nothing
    Set xmlNode = Nothing
    Set xmlProdRest = Nothing
    Set xmlAddRestrict = Nothing
    Set xmlLTV = Nothing
    Set xmlPurpose = Nothing
Exit Sub
ErrHandler:
    Set xmlRestrictions = Nothing
    Set xmlItem = Nothing
    Set xmlPart = Nothing
    Set xmlNode = Nothing
    Set xmlProdRest = Nothing
    Set xmlAddRestrict = Nothing
    Set xmlLTV = Nothing
    Set xmlPurpose = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetRestrictionText
'** Created by:     Andy Maggs
'** Date:           23/04/2004
'** Description:    Builds and returns a comma delimited list of restrictions.
'** Parameters:     vxmlLoanComponent - the loan component data.
'**                 vstrElementName - the name of the element containing the
'**                 restriction information.
'**                 vstrAttribName - the name of the attribute containing the
'**                 restriction text.
'** Returns:        The comma delimited list of restrictions.
'** Errors:         None Expected
'********************************************************************************
Private Function GetRestrictionText(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vstrElementName As String, ByVal vstrAttribName As String) As String
    Const cstrFunctionName As String = "GetRestrictionText"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String

    On Error GoTo ErrHandler
    
    Set xmlList = vxmlLoanComponent.selectNodes("//" & vstrElementName)
    For Each xmlItem In xmlList
        If Len(strValue) > 0 Then
            strValue = strValue & ", "
        End If
        strValue = strValue & xmlGetAttributeText(xmlItem, vstrAttribName)
    Next xmlItem
    
    GetRestrictionText = strValue
    Set xmlList = Nothing
    Set xmlItem = Nothing

Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddTiedInsuranceElements
'** Created by:     Andy Maggs
'** Date:           23/04/2004
'** Description:    Adds the INSURANCE elements for any compulsory insurance.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the elements to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddTiedInsuranceElements(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddTiedInsuranceElements"
    Dim blnHasBCElement As Boolean
    Dim blnGetBCElement As Boolean
    Dim blnHasPPElement As Boolean
    Dim blnGetPPElement As Boolean
    Dim strBCCover As String
    Dim dblBCCost As Double
    Dim strPPCover As String
    Dim dblPPCost As Double
    Dim xmlInsurance As IXMLDOMNode

    On Error GoTo ErrHandler
    
    blnGetBCElement = Not xmlGetAttributeAsBoolean(vxmlNode, "HASBCINSURANCE")
    blnGetPPElement = Not xmlGetAttributeAsBoolean(vxmlNode, "HASPPINSURANCE")
    Call GetTiedInsuranceData(vobjCommon, blnGetBCElement, blnGetPPElement, _
            blnHasBCElement, strBCCover, dblBCCost, blnHasPPElement, _
            strPPCover, dblPPCost)
    
    If blnHasBCElement Then
        '*-add an INSURANCE element for the compulsory Buildings and Contents
        '*-insurance
        Set xmlInsurance = vobjCommon.CreateNewElement("INSURANCE", vxmlNode)
        '*-add the INSURANCESOURCE attribute
        Call xmlSetAttributeValue(xmlInsurance, "INSURANCESOURCE", vobjCommon.Provider)
        '*-add the type of insurance
        Call xmlSetAttributeValue(xmlInsurance, "TYPE", strBCCover)
        Call xmlSetAttributeValue(vxmlNode, "HASBCINSURANCE", "TRUE")
    End If
    
    If blnHasPPElement Then
        '*-add an INSURANCE element for the compulsory Payment Protection
        '*-insurance
        Set xmlInsurance = vobjCommon.CreateNewElement("INSURANCE", vxmlNode)
        '*-add the INSURANCESOURCE attribute
        Call xmlSetAttributeValue(xmlInsurance, "INSURANCESOURCE", vobjCommon.Provider)
        '*-add the type of insurance
        Call xmlSetAttributeValue(xmlInsurance, "TYPE", strPPCover)
        Call xmlSetAttributeValue(vxmlNode, "HASPPINSURANCE", "TRUE")
    End If
    
    Set xmlInsurance = Nothing

Exit Sub
ErrHandler:
    Set xmlInsurance = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetTiedInsuranceData
'** Created by:     Andy Maggs
'** Date:           23/04/2004
'** Description:    Adds the INSURANCE elements for any compulsory insurance.
'** Parameters:     vxmlLoanComponent - the loan component.
'**                 vxmlNode - the node to add the elements to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub GetTiedInsuranceData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vblnGetBC As Boolean, ByVal vblnGetPP As Boolean, _
        ByRef rblnHasBC As Boolean, ByRef rstrBCCoverType As String, _
        ByRef rdblBCMonthlyCost As Double, ByRef rblnHasPP As Boolean, _
        ByRef rstrPPCoverType As String, ByRef rdblPPMonthlyCost As Double)
    Const cstrFunctionName As String = "GetTiedInsuranceData"
    Dim xmlParams As IXMLDOMNode
    Dim strCover As String
    Dim xmlComponents As IXMLDOMNodeList
    Dim xmlComponent As IXMLDOMNode
    Dim xmlQuote As IXMLDOMNode
    Dim xmlDetails As IXMLDOMNode

    On Error GoTo ErrHandler

    Set xmlComponents = vobjCommon.LoanComponents
    For Each xmlComponent In xmlComponents
        Set xmlParams = xmlComponent.selectSingleNode("//MORTGAGEPRODUCTPARAMETERS")
        If Not xmlParams Is Nothing Then
            '*-only get the buildings and contents quote data once
            If vblnGetBC Then
                rblnHasBC = xmlGetAttributeAsBoolean(xmlParams, "COMPULSORYBC")
                vblnGetBC = False
                If rblnHasBC Then
                    '*-get the quote
                    Set xmlQuote = vobjCommon.Data.selectSingleNode("//BUILDINGSANDCONTENTSSUBQUOTE")
                    rdblBCMonthlyCost = xmlGetAttributeAsDouble(xmlQuote, "TOTALBCMONTHLYCOST")
                    '*-get the type of insurance
                    Set xmlDetails = vobjCommon.Data.selectSingleNode("//BUILDINGSANDCONTENTSDETAILS")
                    strCover = xmlGetAttributeText(xmlDetails, "COVERTYPE")
                    Select Case strCover
                        Case "BC"
                            rstrBCCoverType = "buildings & contents"
                        Case "B"
                            rstrBCCoverType = "buildings cover"
                        Case Else
                            rstrBCCoverType = "contents cover"
                    End Select
                End If
            End If
            
            '*-only get the payment protection quote data once
            If vblnGetPP Then
                rblnHasPP = xmlGetAttributeAsBoolean(xmlParams, "COMPULSORYPP")
                vblnGetPP = False
                If rblnHasPP Then
                    Set xmlQuote = vobjCommon.Data.selectSingleNode("//PAYMENTPROTECTIONSUBQUOTE")
                    rdblPPMonthlyCost = xmlGetAttributeAsDouble(xmlQuote, "TOTALPPMONTHLYCOST")
                    '*-get the type of insurance
                    rstrPPCoverType = "mortgage payment protection"
                End If
            End If
        End If
    Next xmlComponent
    
    Set xmlParams = Nothing
    Set xmlComponents = Nothing
    Set xmlComponent = Nothing
    Set xmlQuote = Nothing
    Set xmlDetails = Nothing
Exit Sub
ErrHandler:
    Set xmlParams = Nothing
    Set xmlComponents = Nothing
    Set xmlComponent = Nothing
    Set xmlQuote = Nothing
    Set xmlDetails = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddFeesPremiumsElement
'** Created by:     Andy Maggs
'** Date:           20/04/2004
'** Description:    Adds the appropriate element and attributes for NOFEESORPREMIUMS
'**                 or INCLUDEFEES (other types are not supported by Omiga)
'** Parameters:     vxmlNode - the node to add the elements to.
'** Returns:        The fees and/or premiums node.
'** Errors:         None Expected
'********************************************************************************
Public Function AddFeesPremiumsElement(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName As String = "AddFeesPremiumsElement"
    Dim xmlNewNode As IXMLDOMNode

    On Error GoTo ErrHandler
    
    If vobjCommon.FeesAddedToLoanAmount = 0 Then
        '*-add the NOFEESORPREMIUMS element
        Set xmlNewNode = vobjCommon.CreateNewElement("NOFEESORPREMIUMS", vxmlNode)
        '*-add the mandatory TOTALLOANAMOUNTPLUSFEES attribute
        Call xmlSetAttributeValue(xmlNewNode, "TOTALLOANAMOUNTPLUSFEES", _
                CStr(vobjCommon.LoanAmount))
    
    ElseIf vobjCommon.FeesAddedToLoanAmount > 0 Then
        '*-add the INCLUDEFEES element
        Set xmlNewNode = vobjCommon.CreateNewElement("INCLUDEFEES", vxmlNode)
        '*-add the mandatory TOTALLOANAMOUNTPLUSFEES attribute
        Call xmlSetAttributeValue(xmlNewNode, "TOTALLOANAMOUNTPLUSFEES", _
                CStr(vobjCommon.LoanAmount + vobjCommon.FeesAddedToLoanAmount))
    
    End If
    
    Set AddFeesPremiumsElement = xmlNewNode
    
    Set xmlNewNode = Nothing
Exit Function
ErrHandler:
    Set xmlNewNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddAutonumberingAttributes
'** Created by:     Andy Maggs
'** Date:           01/06/2004
'** Description:    Adds the attributes required for the autonumbering section.
'** Parameters:     vxmlNode - the autonumbering section.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddAutonumberingAttributes(ByVal vxmlNode As IXMLDOMNode, _
                                      Optional ByVal vstrName As String = "SECTION")
    Const cstrFunctionName As String = "AddAutonumberingAttributes"

    On Error GoTo ErrHandler

    '*-add the NAME attribute
    Call xmlSetAttributeValue(vxmlNode, "NAME", vstrName)
    '*-add the START attribute
    Call xmlSetAttributeValue(vxmlNode, "START", "1")
    '*-add the STEP attribute
    Call xmlSetAttributeValue(vxmlNode, "STEP", "1")

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetPropertyTenureType
'** Created by:     Srini Rao
'** Date:           02/06/2004
'** Description:    Gets the TenureType (Validation) from the NewProperty record
'** Parameters:     vobjCommon
'** Returns:        Tenure Type of the Property
'** Errors:         None Expected
'********************************************************************************
Public Function GetPropertyTenureType(ByVal vobjCommon As CommonDataHelper) As String
    Const cstrFunctionName As String = "GetPropertyTenureType"
    Dim xmlNode As IXMLDOMNode
    Dim strTenureType As String
    
    On Error GoTo ErrHandler

    strTenureType = ""
    
    Set xmlNode = vobjCommon.Data.selectSingleNode("//NEWPROPERTY")
    If Not xmlNode Is Nothing Then
        strTenureType = xmlGetAttributeText(xmlNode, "TENURETYPE")
    End If
    
    GetPropertyTenureType = strTenureType
    
    Set xmlNode = Nothing
Exit Function
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       IsPropertyFreeHold
'** Created by:     Srini Rao
'** Date:           02/06/2004
'** Description:    Find whether the property is FreeHold (TenureType)
'** Parameters:     vobjCommon
'** Returns:        Tenure Type of the Property
'** Errors:         None Expected
'********************************************************************************
Public Function IsPropertyFreeHold(ByVal vobjCommon As CommonDataHelper) As Boolean
    Const cstrFunctionName As String = "IsPropertyFreeHold"
    Dim strTenureType As String

    On Error GoTo ErrHandler
    
    strTenureType = GetPropertyTenureType(vobjCommon)
    IsPropertyFreeHold = (InStr(strTenureType, "F,") > 0)

Exit Function
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Function

Public Function GetYoungestApplicantsAge(ByVal vobjCommon As CommonDataHelper) As Integer
    Const cstrFunctionName As String = "GetYoungestApplicantsAge"
    Dim xmlcustomeRolerList As IXMLDOMNodeList
    Dim xmlCustomerRole As IXMLDOMNode
    Dim xmlCustomerVersion As IXMLDOMNode
    Dim strRoleType As String
    Dim intMinAge As Integer
    Dim intAge As Integer
    
    On Error GoTo ErrHandler

    intMinAge = 32767
    Set xmlcustomeRolerList = vobjCommon.Data.selectNodes(gcstrCUSTOMERROLE_PATH)
    For Each xmlCustomerRole In xmlcustomeRolerList
        strRoleType = xmlGetAttributeText(xmlCustomerRole, "CUSTOMERROLETYPE")
        If InStr(strRoleType, "A,") > 0 Then
            Set xmlCustomerVersion = xmlCustomerRole.selectSingleNode(".//CUSTOMERVERSION")
            intAge = xmlGetAttributeAsInteger(xmlCustomerVersion, "AGE")
            If intAge < intMinAge Then
                intMinAge = intAge
            End If
        End If
    Next xmlCustomerRole
        
    GetYoungestApplicantsAge = intMinAge

    Set xmlcustomeRolerList = Nothing
    Set xmlCustomerRole = Nothing
    Set xmlCustomerVersion = Nothing
Exit Function
ErrHandler:
    Set xmlcustomeRolerList = Nothing
    Set xmlCustomerRole = Nothing
    Set xmlCustomerVersion = Nothing
     errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddOtherMortgagesData
'** Created by:     Srini Rao
'** Date:           02/06/2004
'** Description:    Find whether any other mortgages exist on this property.
'**                 if Yes, add the details of the loan
'** Parameters:     vobjCommon - Data helper
'                   vxmlNode - node to attach the data to
'** Returns:
'** Errors:         None Expected
'********************************************************************************
Public Sub AddOtherMortgagesData(ByVal vobjCommon As CommonDataHelper, _
                                 ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddOtherMortgagesData"
    Dim xmlMortgageAccount As IXMLDOMNode
    Dim xmlMortgageLoan As IXMLDOMNode
    Dim xmlAccount As IXMLDOMNode
    Dim xmlAccountRelationshipList As IXMLDOMNodeList
    Dim xmlAccountRelationship As IXMLDOMNode
    Dim xmlNPAddress As IXMLDOMNode
    Dim xmlTPDIRNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim strNPAddressGuid As String
    Dim strSecurityAddress As String
    Dim blnMAFound As Boolean
    Dim strRedemptionStatus As String
    Dim strLenderName As String
    Dim dblOSBalance As Double
    
    On Error GoTo ErrHandler
        
    Set xmlNPAddress = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/NEWPROPERTY/NEWPROPERTYADDRESS")
    If Not xmlNPAddress Is Nothing Then
        strNPAddressGuid = xmlGetAttributeText(xmlNPAddress, "ADDRESSGUID")
    Else
        Exit Sub
    End If
        
    Set xmlAccountRelationshipList = vobjCommon.Data.selectNodes(gcstrACCOUNTRELATIONSHIP_PATH)
    For Each xmlAccountRelationship In xmlAccountRelationshipList
        Set xmlAccount = xmlAccountRelationship.selectSingleNode(".//ACCOUNT")
        Set xmlMortgageAccount = xmlAccountRelationship.selectSingleNode(".//ACCOUNT/MORTGAGEACCOUNT")
        If Not xmlMortgageAccount Is Nothing Then
            strSecurityAddress = xmlGetAttributeText(xmlMortgageAccount, "ADDRESSGUID")
        End If
        ' There can only be one mortgage on the property. So, if the NewPropertyAddress matches
        ' with the SecurityAddress (MortgageAccount.AddressGUID)
        If (strSecurityAddress = strNPAddressGuid) And strSecurityAddress <> "" Then
            blnMAFound = True
            Exit For
        End If
    Next xmlAccountRelationship
    
    If Not blnMAFound Then
         Set xmlNode = vobjCommon.CreateNewElement("NOOTHERMORTGAGE", vxmlNode)
    Else
        Set xmlMortgageLoan = xmlMortgageAccount.selectSingleNode("./MORTGAGELOAN")
        If xmlMortgageLoan Is Nothing Then
            Set xmlNode = vobjCommon.CreateNewElement("NOOTHERMORTGAGE", vxmlNode)
        Else
            Set xmlNode = vobjCommon.CreateNewElement("OTHERMORTGAGE", vxmlNode)
            
            If xmlAttributeValueExists(xmlAccount, "THIRDPARTYGUID") Then
                Set xmlTPDIRNode = xmlAccount.selectSingleNode(".//THIRDPARTY")
            Else
                Set xmlTPDIRNode = xmlAccount.selectSingleNode(".//NAMEANDADDRESSDIRECTORY")
            End If
            
            strLenderName = xmlGetAttributeText(xmlTPDIRNode, "COMPANYNAME")
            dblOSBalance = xmlGetAttributeAsDouble(xmlMortgageLoan, "OUTSTANDINGBALANCE")
            
            xmlSetAttributeValue xmlNode, "OTHERLENDER", strLenderName
            xmlSetAttributeValue xmlNode, "EXISTINGMORTGAGEAMOUNT", dblOSBalance
                        
            strRedemptionStatus = xmlGetAttributeText(xmlMortgageLoan, "REDEMPTIONSTATUS")
            If InStr(strRedemptionStatus, "R,") > 0 Then
                 Call vobjCommon.CreateNewElement("NOTBEREPAID", xmlNode)
            End If
        End If
    End If
    
    Set xmlMortgageAccount = Nothing
    Set xmlMortgageLoan = Nothing
    Set xmlAccount = Nothing
    Set xmlAccountRelationshipList = Nothing
    Set xmlAccountRelationship = Nothing
    Set xmlNPAddress = Nothing
    Set xmlTPDIRNode = Nothing
    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlMortgageAccount = Nothing
    Set xmlMortgageLoan = Nothing
    Set xmlAccount = Nothing
    Set xmlAccountRelationshipList = Nothing
    Set xmlAccountRelationship = Nothing
    Set xmlNPAddress = Nothing
    Set xmlTPDIRNode = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Public Function GetNumberOfApplicants(ByVal vobjCommon As CommonDataHelper) As Integer
    Const cstrFunctionName As String = "GetNumberOfApplicants"
    Dim xmlCustomerRoleList As IXMLDOMNodeList

    On Error GoTo ErrHandler

    ' Find the customers linked to this application whole role (CustomerRole.CustomerRoleType)
    ' is Applicant
    Set xmlCustomerRoleList = vobjCommon.Data.selectNodes(gcstrCUSTOMERROLE_PATH & "[@CUSTOMERROLETYPE='A,']") 'SR 11/09/2004 : CORE82
    
    GetNumberOfApplicants = xmlCustomerRoleList.length
    
    Set xmlCustomerRoleList = Nothing
Exit Function
ErrHandler:
    Set xmlCustomerRoleList = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       AddMortgageTypeAttribute
'** Created by:     Andy Maggs
'** Date:           27/05/2004
'** Description:    Adds the MORTGAGETYPE attribute to the specified node.
'** Parameters:     vxmlNode - the node to add the attribute to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddMortgageTypeAttribute(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddMortgageTypeAttribute"

    On Error GoTo ErrHandler

    If Not vxmlNode Is Nothing Then
        Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
    End If

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddPersonalCircumstancesElement
'** Created by:     Andy Maggs
'** Date:           08/06/2004
'** Description:
'** Parameters:
'** Returns:
'** Errors:         None Expected
'********************************************************************************
'Public Sub AddPersonalCircumstancesElement(ByVal vobjCommon As CommonDataHelper, _
'        ByVal vxmlNode As IXMLDOMNode)
'    Const cstrFunctionName As String = "AddPersonalCircumstancesElement"
'    Dim xmlAppFF As IXMLDOMNode
'    Dim xmlNewNode As IXMLDOMNode
'
'    On Error GoTo ErrHandler
'
'    Set xmlAppFF = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
'    If xmlGetAttributeAsInteger(xmlAppFF, "ACCEPTEDQUOTENUMBER") <> _
'            xmlGetAttributeAsInteger(xmlAppFF, "RECOMMENDEDQUOTENUMBER") Then
'        '*-add the PERSONALCIRCUMSTANCES element
'        Set xmlNewNode = vobjCommon.CreateNewElement("PERSONALCIRCUMSTANCES", vxmlNode)
'        '*-add the DISAGREED element
'        Call vobjCommon.CreateNewElement("DISAGREED", xmlNewNode)
'    End If
'
'Exit Sub
'ErrHandler:
'    '*-check the error and throw it on
'    errCheckError cstrFunctionName, mcstrModuleName
'End Sub

'********************************************************************************
'** Function:       BuildWhatFeesYouMustPaySection
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the section that is
'**                 virtually identical for the different documents but which has
'**                 different section numbers in the different documents.
'** Parameters:     vxmlNode - the section element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildWhatFeesYouMustPaySection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildWhatFeesYouMustPaySection"

    Dim xmlList As IXMLDOMNodeList
    Dim xmlCost As IXMLDOMNode
    Dim xmlPayable As IXMLDOMNode
    Dim xmlDetail As IXMLDOMNode
    Dim dblOutstanding As Double
    Dim strDetailDescription As String
    Dim xmlLender As IXMLDOMNode
    Dim strCostType As String
    Dim xmlNotAdded As IXMLDOMNode
    Dim lngAmountCouldAddToLoan As Long
    Dim xmlMIGCost As IXMLDOMNode
    Dim xmlHLC As IXMLDOMNode
    Dim lngAmount As Long
    Dim blnPurposePurchase As Boolean
    Dim blnIsEstimated As Boolean
    Dim xmlNoFees As IXMLDOMNode
    Dim xmlLegal As IXMLDOMNode
    Dim lngMIGAmount As Long
    Dim xmlBroker As IXMLDOMNode
    Dim lngBrokerAmount As Long
    Dim xmlTPVFee As IXMLDOMNode, xmlNode As IXMLDOMNode 'SR 11/10/2004 : BBG1545
    Dim strTemp As String 'SR 17/10/2004 : BBG1643
    Dim xmlValFee As IXMLDOMNode, dblValFee As Double, dblTPVFee As Double  'SR 17/10/2004 : BBG1643
     
    'SAB 27/04/2006 - EPSOM EP454 - START
    Dim strDetailDescription2 As String
    Dim strDetailDescription3 As String
    Dim xmlDetail2 As IXMLDOMNode
    Dim xmlDetail3 As IXMLDOMNode
    Dim xmlPanelLegal As IXMLDOMNode
    'SAB 27/04/2006 - EPSOM EP454 - END
    
    'PB 07/06/2006 EP696/MAR1826
    Dim intNoPayableCostRecords As Integer
    'EP696/MAR1826 End
     
    'PB 04/12/2006 EP2_139
    Dim strLegalFeeText As String
    'EP2_718
    Dim legalFee As Double
    Dim xmlHasFee As IXMLDOMNode
    'EP2_1382
    Dim strBrokerFeeDescr As String
    'EP2_1753
    Dim xmlDetail4 As IXMLDOMNode
    Dim strDetailDescription4 As String
    strDetailDescription4 = ""
    Dim legalFeeAmount As Long
    legalFeeAmount = 0
    
    On Error GoTo ErrHandler
    
    'PB 07/06/2006 EP696/MAR1826
    intNoPayableCostRecords = 0
    'EP696/MAR1826 End
    
    '*-add the provider attribute
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", vobjCommon.Provider)
    
    ' EP2_2477 GHun Copy the TPVA fee before it gets removed
    ' EP2_2477 IK 19/04/2007
    If Not vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'TPVA')]") Is Nothing Then
        Set xmlTPVFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'TPVA')]").cloneNode(False)
    End If
        
    '*-get the list of fees paid to the provider
    Set xmlList = vobjCommon.GetFeesPaidToProvider()
    
    '*-get the lender details
    Set xmlLender = vobjCommon.Data.selectSingleNode(gcstrMORTGAGELENDER_PATH)
    
    'PB 07/06/2006 EP696/MAR1826 End
    dblOutstanding = TotalDblAttributeValues(xmlList, "AMOUNT")
    For Each xmlCost In xmlList
        lngAmount = xmlGetAttributeAsLong(xmlCost, "AMOUNT")
        'EP2_1753 moved this to distinguish costtype for legal
        strCostType = xmlGetAttributeText(xmlCost, "MORTGAGEONEOFFCOSTTYPE")
        If lngAmount > 0 Then
            
            'EP2_2477 GHun Also exclude the TPVA fee as it gets added to TPV
            'PB 20/03/2007 EP2_1931
            If Not ((vobjCommon.IsAdditionalBorrowing And CheckForValidationType(strCostType, "END")) Or _
                CheckForValidationType(strCostType, "TPVA")) Then
                
                'EP2_2477 GHun Add the TPVA fee to the TPV fee
                ' EP2_2477 IK 19/04/2007
                If Not xmlTPVFee Is Nothing Then
                    If CheckForValidationType(strCostType, "TPV") Then
                        lngAmount = lngAmount + xmlGetAttributeAsLong(xmlTPVFee, "AMOUNT")
                    End If
                End If
                'EP2_2477 End
                
                'PB 07/06/2006 EP696/MAR1826
                intNoPayableCostRecords = intNoPayableCostRecords + 1
                'PB EP696/MAR1826 End
                
                '*-add the FEESPAYABLE element
                'Peter Edney - EP1113 - 04/09/2006
                'PB 15/03/2007 EP2_1931 Begin
                'If (InStr(strCostType, "VAL") Or InStr(strCostType, "TPV") _
                '    Or InStr(strCostType, "REI")) Then
                'EP2_1753 New validation type to indicate whether fee should be displayed
                'as being from the lender, or in the other section
                If (InStr(strCostType, "NLEN")) Then
                'PB EP2_1931 End
                    Set xmlPayable = vobjCommon.CreateNewElement("INTERMEDIARYFEESPAYABLE", vxmlNode)
                Else
                    'should have a validation type of "LEN"
                    Set xmlPayable = vobjCommon.CreateNewElement("FEESPAYABLE", vxmlNode)
                End If
                '*-add the mandatory FEEAMOUNT attribute
                Call xmlSetAttributeValue(xmlPayable, "FEEAMOUNT", CStr(lngAmount))
                '*-add the mandatory DESCRIPTION attribute
                'AW 12/10/06  EP1212 - End
                
                'BC MAR893 Begin
                If InStr(strCostType, "ARR") > 0 And _
                   vobjCommon.blnFixedRateMortgage = True Then
                    'PB 22/06/2006 EP841 Begin
                    'Call xmlSetAttributeValue(xmlPayable, "DESCRIPTION", "Fixed Rate Fee")
                    Call xmlSetAttributeValue(xmlPayable, "DESCRIPTION", "Arrangement Fee")
                    'PB EP841 End
                ElseIf InStr(strCostType, "DEE") > 0 Then  'SR EP2_2159 - modify description for Redemptino Admin Fee
                    strTemp = xmlGetAttributeText(xmlCost, "MORTGAGEONEOFFCOSTTYPE_TEXT")
                    'strTemp = strTemp & ", currently"  ' SR EP2_2392 : Do not add 'currently' as it is coming the Fee Description (first line)
                    xmlSetAttributeValue xmlPayable, "DESCRIPTION", strTemp
                Else
                    Call xmlCopyAttributeValue(xmlCost, xmlPayable, "MORTGAGEONEOFFCOSTTYPE_TEXT", "DESCRIPTION")
                End If
                'BC MAR893 End
                
                'PB 13/07/2006 EP821 Begin
                'EP2_1753 could have TPV or VAL, but not both
                'EP2_1351 Use description from combo for TPV
'                    If CheckForValidationType(strCostType, "TPV")
                 If CheckForValidationType(strCostType, "VAL") Then
                    ' Third party valuation or valuation - make sure it comes out as just "Valuation Fee"
                    xmlSetAttributeValue xmlPayable, "DESCRIPTION", "Valuation Fee"
                End If
                'PB EP821 End
                
                '*-add the DETAIL element
                 Set xmlDetail = vobjCommon.CreateNewElement("DETAIL", xmlPayable)
                '*-add the DESCRIPTION attribute
                
                'SAB 27/04/2006 - EPSOM EP454 - START
                strDetailDescription2 = ""
                strDetailDescription3 = ""
                'EP2_1753 extra line of description
                strDetailDescription4 = ""
                strDetailDescription = GetFeeDetailDescription(xmlCost, vobjCommon, strDetailDescription2, strDetailDescription3, strDetailDescription4, vblnIsOfferDocument)
                Call xmlSetAttributeValue(xmlDetail, "DESCRIPTION", strDetailDescription)
                
                '*-add the DETAIL2 element
                If Not strDetailDescription2 = "" Then
                    Set xmlDetail2 = vobjCommon.CreateNewElement("DETAIL2", xmlPayable)
                    Call xmlSetAttributeValue(xmlDetail2, "DESCRIPTION2", strDetailDescription2)
                End If
                
                '*-add the DETAIL3 element
                If Not strDetailDescription3 = "" Then
                    Set xmlDetail3 = vobjCommon.CreateNewElement("DETAIL3", xmlPayable)
                    Call xmlSetAttributeValue(xmlDetail3, "DESCRIPTION3", strDetailDescription3)
                End If
                
                '*-add the DETAIL4 element
                If Not strDetailDescription4 = "" Then
                    Set xmlDetail4 = vobjCommon.CreateNewElement("DETAIL4", xmlPayable)
                    Call xmlSetAttributeValue(xmlDetail4, "DESCRIPTION4", strDetailDescription4)
                End If
                'SAB 27/04/2006 - EPSOM EP454 - END
                            
                If (InStr(strCostType, "ARR") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWARRGEMTFEEADDED")) _
                        Or (InStr(strCostType, "POR") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWPORTINGFEEADDED")) _
                        Or (InStr(strCostType, "VAL") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWVALNFEEADDED")) _
                        Or (InStr(strCostType, "REI") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWREINSPTFEEADDED")) _
                        Or (InStr(strCostType, "ADM") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWADMINFEEADDED")) _
                        Or (InStr(strCostType, "TTF") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWTTFEEADDED")) _
                        Or (InStr(strCostType, "MIG") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWMIGFEEADDED")) _
                        Or (InStr(strCostType, "PSF") > 0 And xmlGetAttributeAsBoolean(xmlLender, "ALLOWPRECOMPSWITCHFEEADDED")) Then 'TK 17/03/2005 BBG2037
                    If Not xmlGetAttributeAsBoolean(xmlCost, "ADDTOLOAN") Then
                        '*-add the NOTADDEDTOLOAN element
                        Set xmlNotAdded = vobjCommon.CreateNewElement("NOTADDEDTOLOAN", xmlPayable)
                        '*-and add the amount to the can be added to loan list
                        lngAmountCouldAddToLoan = xmlGetAttributeAsLong(xmlCost, "AMOUNT")
                        '*-add the INCREASEDLOANAMOUNT attribute
                        Call xmlSetAttributeValue(xmlNotAdded, "INCREASEDLOANAMOUNT", CStr(vobjCommon.TotalLoanAmount + lngAmountCouldAddToLoan))
                        If vblnIsOfferDocument Then
                            '*-add the MORTGAGETYPE attribute
                            Call xmlSetAttributeValue(xmlNotAdded, "MORTGAGETYPE", _
                                    vobjCommon.MortgageTypeText)
                            '*-add the FEENAMES attribute
                            Call xmlCopyAttributeValue(xmlCost, xmlNotAdded, "MORTGAGEONEOFFCOSTTYPE_TEXT", "FEENAMES")
                        End If
                    End If
                End If
            End If
        Else
            'EP2_1753
            If CheckForValidationType(strCostType, "LEG") Then
                legalFeeAmount = xmlGetAttributeAsLong(xmlCost, "LEGALAMOUNT", "")
            End If

        End If
    Next xmlCost
        
    
    'PB 07/06/2006 EP696/MAR1826 Begin
    'End If
    If xmlList.length = 0 Or intNoPayableCostRecords = 0 Then
        '*-add the NOFEES element
        Set xmlNoFees = vobjCommon.CreateNewElement("NOFEES", vxmlNode)
        If vblnIsOfferDocument Then
            '*-add the provider attribute
            Call xmlSetAttributeValue(xmlNoFees, "PROVIDER", vobjCommon.Provider)
        End If
    End If  '
    'PB EP696/MAR1826 End
    
    '*-get the one off cost representing the mortgage idemnity premium if applicable
    Set xmlMIGCost = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'MIG')]")
    If Not xmlMIGCost Is Nothing Then
        'INR CORE82 Only want to display HIGHERLENDINGCHARGE if amount > 0
        lngMIGAmount = xmlGetAttributeAsLong(xmlMIGCost, "AMOUNT")

        If lngMIGAmount > 0 Then
            '*-add the HIGHERLENDINGCHARGE element
            Set xmlHLC = vobjCommon.CreateNewElement("HIGHERLENDINGCHARGE", vxmlNode)
            '*-add the mandatory FEEAMOUNT attribute
            Call xmlCopyAttributeValue(xmlMIGCost, xmlHLC, "AMOUNT", "FEEAMOUNT")
            '*-add the mandatory LTV attribute
            Call xmlSetAttributeValue(xmlHLC, "LTV", vobjCommon.LTV)
            
            blnPurposePurchase = (vobjCommon.LoanPurposeText = "PURPOSEPURCHASE")
            Call PurchasePriceOrEstimatedValue(vobjCommon.Data, blnPurposePurchase, _
                    blnIsEstimated, lngAmount)
            If blnIsEstimated Then
                '*-add the HLESTIMATEDVALUE element
                Call vobjCommon.CreateNewElement("HLESTIMATEDVALUE", xmlHLC)
            Else
                '*-add the HLPURCHASEPRICE element
                Call vobjCommon.CreateNewElement("HLPURCHASEPRICE", xmlHLC)
            End If
            
            '*-add the DETAIL element
            Set xmlDetail = vobjCommon.CreateNewElement("DETAIL", xmlHLC)
            '*-add the DESCRIPTION attribute
            'strDetailDescription = GetFeeDetailDescription(xmlMIGCost, vobjCommon)
            strDetailDescription = GetFeeDetailDescription(xmlCost, vobjCommon, strDetailDescription2, strDetailDescription3, strDetailDescription4, vblnIsOfferDocument)
            Call xmlSetAttributeValue(xmlDetail, "DESCRIPTION", strDetailDescription)
        End If
    End If
    
    'CORE82 add the OTHERFEES element for BrokerFees Only
    Set xmlBroker = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'BRK')]")
    If Not xmlBroker Is Nothing Then
    
        lngBrokerAmount = xmlGetAttributeAsLong(xmlBroker, "AMOUNT")
        If lngBrokerAmount > 0 Then
            'AW  22/05/2006  EP590
            'EP2_1382 Use INTERMEDIARYFEESPAYABLE to improve displayed text
            Set xmlCost = vobjCommon.CreateNewElement("INTERMEDIARYFEESPAYABLE", vxmlNode)
    '        Call xmlCopyAttributeValue(xmlBroker, xmlCost, "AMOUNT", "FEEAMOUNT")
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngBrokerAmount)
                
            '*-add the DESCRIPTION attribute
'            strDetailDescription = GetFeeDetailDescription(xmlBroker, vobjCommon)
            'strDetailDescription = GetFeeDetailDescription(xmlCost, vobjCommon, strDetailDescription2, strDetailDescription3)
            strDetailDescription = GetFeeDetailDescription(xmlBroker, vobjCommon, strDetailDescription2, strDetailDescription3, strDetailDescription4, vblnIsOfferDocument)
            strBrokerFeeDescr = xmlGetAttributeText(xmlBroker, "MORTGAGEONEOFFCOSTTYPE_TEXT")
            Call xmlSetAttributeValue(xmlCost, "DESCRIPTION", strBrokerFeeDescr)
            
            '*-add the DETAIL element
            If Len(strDetailDescription) > 0 Then
                Set xmlDetail2 = vobjCommon.CreateNewElement("DETAIL", xmlCost)
                Call xmlSetAttributeValue(xmlDetail2, "DESCRIPTION", strDetailDescription)
            End If
            
            '*-add the DETAIL2 element
            If Len(strDetailDescription2) > 0 Then
                Set xmlDetail2 = vobjCommon.CreateNewElement("DETAIL2", xmlCost)
                Call xmlSetAttributeValue(xmlDetail2, "DESCRIPTION2", strDetailDescription2)
            End If
            
            '*-add the DETAIL3 element
            If Len(strDetailDescription3) > 0 Then
                Set xmlDetail3 = vobjCommon.CreateNewElement("DETAIL3", xmlCost)
                Call xmlSetAttributeValue(xmlDetail3, "DESCRIPTION3", strDetailDescription3)
            End If

'            If Len(strBrokerFeeDescr) > 0 Then
'                 strDetailDescription = strBrokerFeeDescr & ". " & strDetailDescription
'            End If
'            Call xmlSetAttributeValue(xmlCost, "DESCRIPTION", strDetailDescription)
'            'SR 11/10/2004 : BBG1545
'            Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
'            strDetailDescription = xmlGetAttributeText(xmlNode, "BROKERFEEDETAILS")
'
'            If Len(strDetailDescription) > 0 Then
'                xmlSetAttributeValue xmlCost, "BROKERFEEDETAILS", strDetailDescription 'SR 17/10/2004 : BBG1643
'            End If
            'SR 11/10/2004 : BBG1545 - End
            
        End If
    End If
    
    'EP2_1382 add the OTHERFEES element for AdditionalBrokerFees
    Set xmlBroker = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'ABRK')]")
    If Not xmlBroker Is Nothing Then
    
        lngBrokerAmount = xmlGetAttributeAsLong(xmlBroker, "AMOUNT")
        If lngBrokerAmount > 0 Then
            'EP2_1382 Use INTERMEDIARYFEESPAYABLE to improve displayed text
            Set xmlCost = vobjCommon.CreateNewElement("INTERMEDIARYFEESPAYABLE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngBrokerAmount)
                
            '*-add the DESCRIPTION attribute
            strDetailDescription = GetFeeDetailDescription(xmlBroker, vobjCommon, strDetailDescription2, strDetailDescription3, strDetailDescription4, vblnIsOfferDocument)
            strBrokerFeeDescr = xmlGetAttributeText(xmlBroker, "MORTGAGEONEOFFCOSTTYPE_TEXT")
            Call xmlSetAttributeValue(xmlCost, "DESCRIPTION", strBrokerFeeDescr)
            
            '*-add the DETAIL element
            If Len(strDetailDescription) > 0 Then
                Set xmlDetail2 = vobjCommon.CreateNewElement("DETAIL", xmlCost)
                Call xmlSetAttributeValue(xmlDetail2, "DESCRIPTION", strDetailDescription)
            End If
            
            '*-add the DETAIL2 element
            If Len(strDetailDescription2) > 0 Then
                Set xmlDetail2 = vobjCommon.CreateNewElement("DETAIL2", xmlCost)
                Call xmlSetAttributeValue(xmlDetail2, "DESCRIPTION2", strDetailDescription2)
            End If
            
            '*-add the DETAIL3 element
            If Len(strDetailDescription3) > 0 Then
                Set xmlDetail3 = vobjCommon.CreateNewElement("DETAIL3", xmlCost)
                Call xmlSetAttributeValue(xmlDetail3, "DESCRIPTION3", strDetailDescription3)
            End If
            
        End If
    End If
    
    'SR 11/10/2004 : BBG1545
    '*- add OTHERVALUATION FEE  element
   'SR 17/10/2004 : BBG1643
    Set xmlTPVFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'TPV')]")
    Set xmlValFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'VAL')]")
    If Not xmlValFee Is Nothing Then
        dblValFee = xmlGetAttributeAsDouble(xmlValFee, "AMOUNT")
    End If
    
    If Not xmlTPVFee Is Nothing Then
        dblTPVFee = xmlGetAttributeAsDouble(xmlTPVFee, "AMOUNT")
    End If
    
    If dblTPVFee > 0 Or dblValFee > 0 Then
        Set xmlCost = vobjCommon.CreateNewElement("OTHERVALUATIONFEE", vxmlNode)
        If dblTPVFee > 0 Then
            xmlCopyAttributeValue xmlTPVFee, xmlCost, "AMOUNT", "FEEAMOUNT"
            xmlSetAttributeValue xmlCost, "DESCRIPTION", _
                    (xmlGetAttributeText(xmlTPVFee, "MORTGAGEONEOFFCOSTTYPE_TEXT"))
        Else
            If vobjCommon.IsApplicationPackaged Then
                xmlCopyAttributeValue xmlValFee, xmlCost, "AMOUNT", "FEEAMOUNT"
                xmlSetAttributeValue xmlCost, "DESCRIPTION", _
                    (xmlGetAttributeText(xmlValFee, "MORTGAGEONEOFFCOSTTYPE_TEXT"))
                
                vobjCommon.CreateNewElement "STANDARDVALUATIONFEE", xmlCost
            End If
        End If
    End If
    'SR 17/10/2004 : BBG1643 - End

    'EP2_718 Easier to do it all here
    'Get the correct legal fee node name and add to SECTION8
'    strLegalFeeText = GetLegalFeeText(vobjCommon)
'    vobjCommon.CreateNewElement strLegalFeeText, vxmlNode
    
    'EP2_718Get the correct legal fee node name and add to SECTION8
    If vobjCommon.IsAdditionalBorrowingTOE Or vobjCommon.IsProductSwitch Then
        'EP2_1915
        If vobjCommon.IsAdditionalBorrowingTOE And Not vobjCommon.IsProductSwitch Then
            strLegalFeeText = "ABORTOE"
        Else
            'Additional Borrowing or Product Switch
            strLegalFeeText = "ADDTNLORSWITCH"
        End If
    Else
        'Not additional Borrowing or Product Switch
        If vobjCommon.Data.selectSingleNode(gcstrPANELLEGALREP_PATH) Is Nothing Then
            'No panel solicitor
            strLegalFeeText = "NOTADDTNLORSWITCHOWNSOL"
        Else
            'Panel solicitor
            strLegalFeeText = "NOTADDTNLORSWITCHPNLSOL"
        End If
        Set xmlLegal = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'LEG')]")
        If vobjCommon.IsFreeLegal() Then
            Set xmlHasFee = vobjCommon.CreateNewElement("ISFREE", vxmlNode)
        Else
            Set xmlHasFee = vobjCommon.CreateNewElement("HASFEE", vxmlNode)
            If legalFeeAmount > 0 Then
                xmlSetAttributeValue xmlHasFee, "SOLFEEAMOUNT", CStr(legalFeeAmount)
            Else
                xmlSetAttributeValue xmlHasFee, "SOLFEEAMOUNT", "250"
            End If
        End If
    
    End If
    Set xmlCost = vobjCommon.CreateNewElement(strLegalFeeText, vxmlNode)

        
    '*-add the ESTIMATEDLEGALCOSTS element
    '*-Amended to check if it is a PANEL solicitor and return relevant element
'    Set xmlCost = vobjCommon.CreateNewElement("ESTIMATEDLEGALCOSTS", vxmlNode)
    
    'PM 22/06/2006 EP826 START (There is no requirement for an 'in-house' statement)
'    Set xmlPanelLegal = vobjCommon.Data.selectSingleNode(gcstrPANELLEGALREP_PATH)
'    If xmlPanelLegal Is Nothing Then
'        Set xmlCost = vobjCommon.CreateNewElement("NONPANELESTIMATEDLEGALCOSTS", vxmlNode)
'    Else
'        Set xmlCost = vobjCommon.CreateNewElement("PANELESTIMATEDLEGALCOSTS", vxmlNode)
'    End If
    
    'Do not need to extract amounts as they have a default amount of £250
    'strTemp = ""
    'If Not xmlLegal Is Nothing Then
    '    strTemp = xmlGetAttributeText(xmlLegal, "AMOUNT")
    'End If
    'strTemp = IIf(Len(strTemp) > 0, strTemp, "0")
    'Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", strTemp)
    'SAB 01/05/2006 - EPSOM EP479 - END
    
'    If vblnIsOfferDocument Then
'        '*-add the PROVIDER attribute
'        Call xmlSetAttributeValue(xmlCost, "PROVIDER", vobjCommon.Provider)
'    End If
    'PM 22/06/2006 EP826 END
    
    Set xmlList = Nothing
    Set xmlCost = Nothing
    Set xmlPayable = Nothing
    Set xmlDetail = Nothing
    Set xmlLender = Nothing
    Set xmlNotAdded = Nothing
    Set xmlMIGCost = Nothing
    Set xmlHLC = Nothing
    Set xmlNoFees = Nothing
    Set xmlLegal = Nothing
    Set xmlBroker = Nothing
    Set xmlTPVFee = Nothing
    Set xmlNode = Nothing
    Set xmlValFee = Nothing
    Set xmlDetail2 = Nothing
    Set xmlDetail3 = Nothing
    Set xmlDetail4 = Nothing
    'SAB 01/05/2006 - EPSOM EP479 - START
    Set xmlPanelLegal = Nothing
    'SAB 01/05/2006 - EPSOM EP479 - END
    'EP2_718
    Set xmlHasFee = Nothing

    
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlCost = Nothing
    Set xmlPayable = Nothing
    Set xmlDetail = Nothing
    Set xmlLender = Nothing
    Set xmlNotAdded = Nothing
    Set xmlMIGCost = Nothing
    Set xmlHLC = Nothing
    Set xmlNoFees = Nothing
    Set xmlLegal = Nothing
    Set xmlBroker = Nothing
    Set xmlTPVFee = Nothing
    Set xmlNode = Nothing
    Set xmlValFee = Nothing
    Set xmlDetail2 = Nothing
    Set xmlDetail3 = Nothing
    Set xmlDetail4 = Nothing
    'SAB 01/05/2006 - EPSOM EP479 - START
    Set xmlPanelLegal = Nothing
    'SAB 01/05/2006 - EPSOM EP479 - END
    'EP2_718
    Set xmlHasFee = Nothing
    
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       GetFeeDetailDescription
'** Created by:     Andy Maggs
'** Date:           18/05/2004
'** Description:    Gets the detail description of fees for the specified one
'**                 one off cost. IMHO this text should be built in the template
'**                 not in this component!
'** Parameters:     vxmlCost - the cost element.
'** Returns:        The fee detail description.
'** Errors:         None Expected
'********************************************************************************
Private Function GetFeeDetailDescription(ByVal vxmlCost As IXMLDOMElement, _
                                    ByVal vobjCommon As CommonDataHelper, _
                                    ByRef strDescription2 As String, _
                                    ByRef strDescription3 As String, _
                                    ByRef strDescription4 As String, _
                                    ByVal vblnIsOfferDocument As Boolean) As String
    Const cstrFunctionName As String = "GetFeeDetailDescription"
    'Const cstrPOUND_SYMBOL As String = "&#163;" 'code for £
    Dim xmlfees As IXMLDOMNodeList
    Dim xmlFee As IXMLDOMNode
    Dim lngAmountOutstanding As Long
    Dim lngAmountPayable As Long
    Dim lngAmountPaid As Long
    Dim lngRefundAmount As Long
    Dim lngAmountWaived As Long
    Dim lngNotRefunded As Long
    
    Dim dtRefundDate As Date
    Dim strDescription As String
    Dim strEvent As String
    Dim strCostType As String
    Dim bAddedToLoan As Boolean
    Dim xmlNewProperty As IXMLDOMNode, xmlItem As IXMLDOMNode 'SR 08/10/2004 : BBG1545
    'Dim dblValAdmFee As Double
    Dim strValuationType As String 'SR 08/10/2004 : BBG1545
    Dim xmlAppFactFindNode  As IXMLDOMNode
    'Dim strSpecialGroupValidationType As String
    'Dim strPackgedValuationRebate As String 'AW  15/05/2006  EP520
    'PB 22/06/2006 EP821 Begin
    'Dim xmlLender As IXMLDOMNode
    'Dim blnAllowArrgemtFeeAdded As Boolean
    'Dim blnAllowPortingFeeAdded As Boolean
    'Dim blnAllowValnFeeAdded As Boolean
    'Dim blnAllowReinsptFeeAdded As Boolean
    'Dim blnAllowAdminFeeAdded As Boolean
    'Dim blnAllowTTFeeAdded As Boolean
    'Dim blnAllowMIGFeeAdded As Boolean
    'PB EP821 End
    'PB 22/06/2006 EP827 Begin
    Dim xmlBrokerContact As IXMLDOMNode
    Dim strBrokerName As String
    'PB EP827 End
    'PB 26/06/2007 EP841 Begin
    Dim lngTotalAmountPaid As Long
    Dim lngTotalAmountOutstanding As Long
    'PB EP841 End
    
    'PB 27/06/2006 EP827
    Dim strBrokerRebate As String
    'PB EP827 End
        
    'Peter Edney - EP1100 - 01/09/06
    Dim blnRetype As Boolean
    Dim xmlValuation As IXMLDOMNode
    Dim lngBorrow As Long   'EP2_1930 GHun
        
On Error GoTo ErrHandler
    
    'PB 22/06/2006 EP821 Begin
    'Set xmlLender = vobjCommon.Data.selectSingleNode(gcstrMORTGAGELENDER_PATH)
    'blnAllowArrgemtFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWARRGEMTFEEADDED", False)
    'blnAllowPortingFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWPORTINGFEEADDED", False)
    'blnAllowValnFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWVALNFEEADDED", False)
    'blnAllowReinsptFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWREINSPTFEEADDED", False)
    'blnAllowAdminFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWADMINFEEADDED", False)
    'blnAllowTTFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWTTFEEADDED", False)
    'blnAllowMIGFeeAdded = xmlGetAttributeAsBoolean(xmlLender, "ALLOWMIGFEEADDED", False)
    'PB EP821 End
    
    '*-get the amount of the one off cost
    lngAmountPayable = xmlGetAttributeAsLong(vxmlCost, "AMOUNT")
    'PB 27/06/2006 EP841 Begin
    lngTotalAmountOutstanding = lngAmountPayable
    'PB EP841 End
    '*-set the amount outstanding to 0
    lngAmountOutstanding = 0
    lngAmountWaived = 0
    lngNotRefunded = 0
    
    '*-now get the list of fee payments for this cost if there are any
    Set xmlfees = vxmlCost.selectNodes("FEEPAYMENT")
    For Each xmlFee In xmlfees
        '*-get the amount 'paid'
        lngAmountPaid = xmlGetAttributeAsLong(xmlFee, "AMOUNTPAID")
        '*-get any refunded amount
        lngRefundAmount = xmlGetAttributeAsLong(xmlFee, "REFUNDAMOUNT")
        
        '*-is this fee payment record a payment, a deduction, a rebate
        '*-or a return of funds
        strEvent = xmlGetAttributeText(xmlFee, "PAYMENTEVENT")
        If InStr(strEvent, "P") > 0 Then
            '*-this is an amount paid
            'PB 27/06/2006 EP841 Begin
            lngTotalAmountPaid = lngTotalAmountPaid + lngAmountPaid
            lngTotalAmountOutstanding = lngTotalAmountOutstanding - Abs(lngAmountPaid) ' - might be negative so use Abs
            'PB EP841 End
            lngAmountOutstanding = lngAmountOutstanding + lngAmountPaid
            If lngRefundAmount > 0 Then
                dtRefundDate = xmlGetAttributeAsDate(xmlFee, "REFUNDDATE")
            End If
            
        '* MAR774 - BC 05/Dec/05 - Begin
        ElseIf (InStr(strEvent, "O") > 0 Or InStr(strEvent, "RA") > 0) And _
                lngRefundAmount > 0 Then
            '*-this is a waived Fee
            lngAmountWaived = lngAmountWaived + lngRefundAmount
            
        ElseIf (InStr(strEvent, "NTR")) > 0 And lngRefundAmount > 0 Then
            '*-this is a non-refundable Fee
            lngNotRefunded = lngNotRefunded + lngRefundAmount
        '* MAR774 - BC 05/Dec/05 - Begin
        ElseIf InStr(strEvent, "D") > 0 Or InStr(strEvent, "RFV") > 0 _
                Or InStr(strEvent, "F") > 0 Then
            '*-this is a deduction, a rebate or a return of funds
            lngAmountOutstanding = lngAmountOutstanding + lngAmountPaid
        End If
    Next xmlFee
    
    'CORE82
    bAddedToLoan = xmlGetAttributeAsBoolean(vxmlCost, "ADDTOLOAN")
    strCostType = xmlGetAttributeText(vxmlCost, "MORTGAGEONEOFFCOSTTYPE")
    
    If Not bAddedToLoan Then
        'if MORTGAGEONEOFFCOSTTYPE is SEA or DEE
        If ((InStr(strCostType, "SEA") > 0) Or (InStr(strCostType, "DEE") > 0)) Then
            strDescription = "Payable with final repayment of the mortgage and will be added to the amount to be repaid."
        'EP2_1930 GHun "BRK" will match "BRK", "ABRK" and "BRKP"
        ElseIf (InStr(strCostType, "BRK") > 0) Then  'SR 07/10/2004 : BBG1545
            'PB 22/06/2006 EP827 Begin
            'strDescription = "A Fee paid to " & vobjCommon.IntermediaryOrganisationName & " for arranging the mortgage."  'SR 07/10/2004 : BBG1545
            'EP2_1860 Get the correct Broker name
            strBrokerName = vobjCommon.BrokerName()
            If Len(strBrokerName) = 0 Then
                strBrokerName = "your broker" ' default value if none supplied
            End If
            'EP2_1382
            If xmlAttributeValueExists(vxmlCost, "REFUNDAMOUNT") Then
                strBrokerRebate = set2DP(xmlGetAttributeText(vxmlCost, "REFUNDAMOUNT"))
            End If
            
            strDescription2 = "This fee is payable to " & strBrokerName & " for arranging this mortgage. "
            
            'EP2_1382
            If (InStr(strCostType, "ABRK") > 0) Then
                strDescription = "This fee is payable on completion. "
            'EP2_1930 GHun
            ElseIf (InStr(strCostType, "BRKP") > 0) Then
                strDescription = "This fee has already been paid."
            ElseIf (InStr(strCostType, "BRK") > 0) Then
                If vblnIsOfferDocument Then
                    strDescription = "This fee has already been paid. "
                Else
                    strDescription = "This fee is payable on application."
                End If
            'EP2_1930 End
            End If
            
            'PB 17/07/2006 EP821 Begin
            'If strBrokerRebate = "" Then
            If Len(strBrokerRebate) = 0 Or strBrokerRebate = "0.00" Then
            'PB 17/07/2006 EP821 End
                strDescription3 = "This fee is non-refundable."
            Else
                strDescription3 = "£" & strBrokerRebate & " will be refundable on completion."
            End If
            'PB 27/06/2006 EP827 End
            'PB EP827 End
                        
        'PB 27/06/2006 EP841 Begin
        'EP2_1930 GHun Commented out as values set here are overwritten later
        'ElseIf InStr(strCostType, "ARR") > 0 Then
        '    ' Arrangement fee
        '    If lngTotalAmountOutstanding > 0 Then
        '        strDescription = "Payable on completion."
        '    Else
        '        strDescription = "This fee has already been paid."
        '    End If
        'EP2_1930 End
        'PB EP841 End
        ElseIf lngAmountOutstanding = 0 Then
            'If ((InStr(strCostType, "TTF") > 0) Or (InStr(strCostType, "IAF") > 0)) Then
            If (InStr(strCostType, "IAF") > 0) Then
                strDescription = "Payable on completion."
            Else
                'PB 21/06/2006 EP821
                'strDescription = "Payable on application."
                strDescription = "This fee has already been paid."
                'PB EP821 End
            End If
        ElseIf lngAmountWaived > 0 Then
            If lngAmountWaived = lngAmountPayable Then
                strDescription = "Fee Amount has been waived."
            Else
                strDescription = "£" & lngAmountWaived & " of the Fee has been waived."
            End If
        ElseIf lngAmountOutstanding = lngAmountPayable Then
            strDescription = "Payable at the start of the mortgage/additional borrowing."
        ElseIf lngAmountOutstanding <> lngAmountPayable Then
            strDescription = "£" & CStr(lngAmountPayable - lngAmountOutstanding) & " payable on application." & _
                " " & "£" & CStr(lngAmountOutstanding) & " payable at the start of the mortgage/additional borrowing."
        End If
    Else
        'bAddedToLoan is True
        strDescription = strDescription & "Payable at the start of the mortgage and is added to the loan."  'SR 20/09/2004 : CORE82
    End If
    
    'EP2_1930 GHun Commented out as the description set here will always be overwritten
    ''MV 21/02/2005 : BBG1932 - Start
    'If InStr(strCostType, "ARR") > 0 Then
    '    Set xmlAppFactFindNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    '    strSpecialGroupValidationType = xmlGetAttributeText(xmlAppFactFindNode, "SPECIALGROUP")
    '    If InStr(strSpecialGroupValidationType, "STLP") > 0 Then
    '        strDescription = strDescription & " " & "The Fee amount includes the cost of obtaining bank references."
    '    End If
    'End If
    ''End
    'EP2_1930 End
    
    If lngRefundAmount > 0 Then
        strDescription = strDescription & " " & "£" & CStr(lngRefundAmount) & " is refundable"
        
        'BC MAR774 02Dec05 Begin
        If dtRefundDate > 0 Then
            strDescription = strDescription & " on the " & CStr(dtRefundDate)
        End If
        
        strDescription = strDescription & "."
    End If

    'SAB 03/05/2006 - EPSOM EP479 - START
    'The foolowing is not required as this text appears in a separate bulletpoint.
    'BC MAR774 02Dec05 Begin
'    If InStr(strCostType, "VAL") > 0 Then
'        If lngNotRefunded > 0 Then
'            strDescription = strDescription & " The Valuation Fee is a non-refundable fee."
'        Else
'            strDescription = strDescription & " The basic valuation fee is refundable upon completion."
'        End If
    
'    Else
'        If InStr(strCostType, "BRK") = 0 Then 'SR 17/10/2004 : BBG1643
'            strDescription = strDescription & " Not refundable." 'SR 24/09/2004 : CORE82
'        End If 'SR 17/10/2004 : BBG1643
    
    'BC MAR774 02Dec05 End
    
'    End If
    
    'If InStr(strCostType, "BRK") = 0 Then 'SR 17/10/2004 : BBG1643
        'PB 21/06/2006 EP821 Begin
        'If InStr(strCostType, "XAC") > 0 Then
        '    strDescription = strDescription & " Fee is based on current rates and may be subject to change in the future."
        'End If
        'PB EP821 End
    'End If 'SR 17/10/2004 : BBG1643
    'SR 08/10/2004 : BBG1545
'    If InStr(strCostType, "VAL") > 0 Then
'        Set xmlNewProperty = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/NEWPROPERTY")
'        strValuationType = xmlGetAttributeText(xmlNewProperty, "VALUATIONTYPE")
'        If InStr(strValuationType, "H,") > 0 Then
'            Set xmlItem = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "KFIHomeValAdmFee" & Chr$(34) & "]")
'        Else
'            Set xmlItem = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "KFIMortValAdmFee" & Chr$(34) & "]")
'        End If
'        dblValAdmFee = xmlGetAttributeAsDouble(xmlItem, "AMOUNT")
'        strDescription = strDescription & " The fee amount includes an administration fee of £" + CStr(dblValAdmFee)
'    End If
    'SR 08/10/2004 : BBG1545 - End
    
    'EPSOM EP454/EP479 - 27/04/2006 SAB - START
    'Set the text for the 2nd and 3rd bullet points
    
    'TPV
    'EP2_1753 Need to deal with "VAL" & "TPV" seperately
    If CheckForValidationType(strCostType, "TPV") Then
        Dim rebateAmount As Double
        Dim refundAmount As Double
        rebateAmount = 0
        refundAmount = 0
        'Using the amounts that are in TPV and in TPVA
        rebateAmount = xmlGetAttributeAsDouble(vxmlCost, "COMBINEDREBATE")
        refundAmount = xmlGetAttributeAsDouble(vxmlCost, "COMBINEDFEE")
        
        strDescription2 = "This fee is to carry out the valuation on the security property."
        'EP2_1930/EP2_2477 GHun
        If vblnIsOfferDocument Then
            strDescription = "This fee has already been paid."
        Else
            strDescription = "This fee is payable on application."
        End If
        'EP2_1930/EP2_2477 End
        'EP2_2534 take into account refund as well
        If rebateAmount > 0 Then
            strDescription3 = "In the event of the valuation not taking place £" & CStr(rebateAmount) & " will be refunded."
        End If
        
        If refundAmount > 0 Then
            If vobjCommon.IsFreevaluation Then
                strDescription4 = "On completion of this mortgage, you will receive a refund of your valuation fee £" & CStr(refundAmount) & " which will be included in the mortgage advance sent to your solicitor."
            End If
        Else
            strDescription4 = "This fee is non-refundable."
        End If

   ' VAL
    'EP2_1753 Need to deal with "VAL" & "TPV" seperately
    ElseIf CheckForValidationType(strCostType, "VAL") Then
        Dim refundOfValuation As String
        refundOfValuation = 0
        refundOfValuation = xmlGetAttributeText(vxmlCost, "AMOUNT")
        
        'Peter Edney - EP1100 - 01/09/06
        'If the validation type is Retype, then display different text.
        strValuationType = vbNullString
        blnRetype = False
        Set xmlValuation = vobjCommon.Data.selectSingleNode(gcstrVALUATIONINSTRUCTIONS_PATH & "[last()]")
        If Not (xmlValuation Is Nothing) Then
            strValuationType = UCase(xmlGetAttributeText(xmlValuation, "VALUATIONTYPE_VALIDID", ""))
            If Len(strValuationType) > 0 Then
                blnRetype = CheckForValidationType(strValuationType, "RT") 'AW EP1185 - 04/10/06
            End If
        End If
        
        strDescription2 = strDescription2 & "This fee is to carry out the valuation on the security property."  'EP2_2477 GHun
        If Not blnRetype Then
            If vobjCommon.IsFreevaluation Then
                strDescription3 = strDescription3 & "On completion of this mortgage, you will receive a refund of your valuation fee £" & refundOfValuation & " which will be included in the  mortgage advance sent to your solicitor."
            Else
                strDescription3 = strDescription3 & "This fee is non-refundable."
            End If
        Else
            If vobjCommon.IsFreevaluation Then
                strDescription3 = strDescription3 & "On completion of this mortgage, you will receive a refund of your valuation fee £" & refundOfValuation & " which will be included in the  mortgage advance sent to your solicitor."
            Else
                strDescription3 = strDescription3 & "This fee is non-refundable."
            End If
        End If
        
        'PM 06/06/2006 EPSOM EP610   Start
        'PB 21/06/2006 EP821 Begin
        'EP2_1980
        'EP2_1930 GHun
        If vblnIsOfferDocument Then
            strDescription = "This fee has already been paid."
        Else
            strDescription = "This fee is payable on application."
        End If
        'EP2_1930 End
        'You have already paid the fee for the valuation.
        'PB EP821 End
        'PM 06/06/2006 EPSOM EP610   End
        'PB EP1861 Begin
    ' ARR
    ElseIf InStr(strCostType, "ARR") > 0 Then
        'EP2_1980
        strDescription2 = "This fee is to cover our administration expenses."
        'PM 06/06/2006 EPSOM EP610   Start
        'PB 22/06/2006 EP821 Begin
        'EP2_1980
        If vblnIsOfferDocument Then
            'EP2_1930 GHun
            'If lngAmountOutstanding > 0 Then
            '    strDescription = "This fee is added to the loan on completion."
            'Else
            '    strDescription = "This fee has already been paid."
            'End If
            strDescription = "This fee is added to your loan on completion."
        Else
            'strDescription = "This fee is payable on application or can be added to your loan on completion."
            If bAddedToLoan Then
                strDescription = "This fee is added to your loan on completion."
            Else
                If vobjCommon.TotalLoanAmount > 0 Then
                    lngBorrow = vobjCommon.TotalLoanAmount
                Else
                    lngBorrow = vobjCommon.LoanAmount + vobjCommon.FeesAddedToLoanAmount
                End If
                lngBorrow = lngBorrow + lngAmountPayable
                strDescription = "If you wish you can add this Arrangement fee to the mortgage." _
                    & " This would increase the amount you borrow to " & FormatCurrency(lngBorrow, 0) & " and would increase the payments shown in Section 6." _
                    & " If you want to do this, you should ask for another illustration that shows the effect of this on your monthly payments."
            End If
            'EP2_1930 End
        End If
'        If bAddedToLoan And blnAllowAdminFeeAdded Then
'            strDescription = "This fee is added to the loan on the completion date."
'        End If
        'PB EP821 End
        'This fee is added to the loan on the completion date.
        'PM 06/06/2006 EPSOM EP610   End
        'EP2_1980
        'EP2_1930 GHun
        'If vblnIsOfferDocument And lngAmountOutstanding = 0 Then
        '    strDescription3 = strDescription3 & "This fee is non-refundable once completion has taken place."
        'Else
        '    strDescription3 = strDescription3 & "This fee is non-refundable."
        'End If
        strDescription3 = "This fee is non-refundable once completion has taken place."
        'EP2_1930 End
    ' TTF
    ElseIf InStr(strCostType, "TTF") > 0 Then
        strDescription2 = "This fee is for sending the completion monies electronically to your solicitor."
        'PM 06/06/2006 EPSOM EP610   Start
        'PB 22/06/2006 EP821 Begin
        'EP2_1980
        If vblnIsOfferDocument Then
            'EP2_1930 GHun
            'If lngAmountOutstanding > 0 Then
            '    strDescription = "This fee is added to the loan on the completion date."
            'Else
            '    strDescription = "This fee has already been paid."
            'End If
            strDescription = "This fee is added to your loan on completion."
        Else
            'strDescription = "This fee is payable on application or can be added to your loan on completion."
            If bAddedToLoan Then
                strDescription = "This fee is added to your loan on completion."
            Else
                If vobjCommon.TotalLoanAmount > 0 Then
                    lngBorrow = vobjCommon.TotalLoanAmount
                Else
                    lngBorrow = vobjCommon.LoanAmount + vobjCommon.FeesAddedToLoanAmount
                End If
                lngBorrow = lngBorrow + lngAmountPayable
                strDescription = "If you wish you can add this Telegraphic Transfer fee to the mortgage." _
                    & " This would increase the amount you borrow to " & FormatCurrency(lngBorrow, 0) & " and would increase the payments shown in Section 6." _
                    & " If you want to do this, you should ask for another illustration that shows the effect of this on your monthly payments."
            End If
            'EP2_1930 End
        End If
'        If bAddedToLoan And blnAllowTTFeeAdded Then
'            strDescription = "This fee is added to the loan on the completion date."
'        End If
        'PB EP821 End
        'This fee is added from the loan on the completion date.
        'PM 06/06/2006 EPSOM EP610   End
        'EP2_1980
        'EP2_1930 GHun
        'If vblnIsOfferDocument And lngAmountOutstanding = 0 Then
        '    strDescription3 = strDescription3 & "This fee is non-refundable once completion has taken place."
        'Else
        '    strDescription3 = strDescription3 & "This fee is non-refundable."
        'End If
        strDescription3 = "This fee is non-refundable once completion has taken place."
        'EP2_1930 End
    ' REI
    ElseIf InStr(strCostType, "REI") > 0 Then
        
        strDescription2 = "This fee is to carry out a re-inspection of the security property."

        'EP2_1980
        If vblnIsOfferDocument Then
            'EP2_1930 GHun Cannot be added to loan
            'If lngAmountOutstanding > 0 Then
            '    strDescription = "This fee is added to the loan on the completion date."
            'Else
                strDescription = "This fee has already been paid."
            'End If
            'EP2_1930 End
        Else
            strDescription = "This fee is payable before the re-inspection is carried out." 'EP2_1930 GHun
        End If
    
        strDescription3 = "This fee is non-refundable."
    
    ' SRF
    ElseIf InStr(strCostType, "SRF") > 0 Then
        'Telegraphic Transfer Fee - Retention/Stage Advance'
        strDescription2 = strDescription2 & "This fee is for sending the retention monies to you/your solicitor electronically."
        
        'PE - 27/06/2006 - EP842
        strDescription = "This fee will be deducted from the retention monies"
        
        'This fee is payable before we release your retention.
        strDescription3 = strDescription3 & "This fee is non-refundable."
    'DRC 21/06/2006 EP801 start
    
    ' DEE
    'PB EP2_1931 30/03/2007
    ElseIf InStr(strCostType, "DEE") > 0 And Not vobjCommon.IsAdditionalBorrowing Then
        'Redemption Admin Fee
        'EP2_1930 GHun
        'strDescription2 = "This fee is for the administration of closing your mortgage account and processing any legal and title documentation when your mortgage is redeemed."
        ''PM 06/06/2006 EPSOM EP610   Start
        'strDescription = "This fee is added to your loan on redemption."
        ''The fee is added to your loan on redemption.
        ''PM 06/06/2006 EPSOM EP610   End
        'strDescription3 = strDescription3 & "This fee is non-refundable."
        
        strDescription2 = "This fee is the current fee."
        strDescription = "This fee is for the administration of closing your mortgage account and processing any legal and title documentation when your mortgage is redeemed."
        strDescription3 = "This fee is added to your loan on redemption."
        strDescription4 = "This fee is non-refundable."
        'EP2_1930 End
        
    'DRC 21/06/2006 EP801 - End
    'AW  15/05/2006  EP520
    
'EP2_1980 TPV dealt with above
    ' TPV - will never be called tho...
'    ElseIf InStr(strCostType, "TPV") > 0 Then
'        strDescription2 = strDescription2 & "This fee was to carry out the valuation on the security property."
'
'        Set xmlAppFactFindNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
'        strPackgedValuationRebate = xmlGetAttributeText(xmlAppFactFindNode, "PACKAGEDVALUATIONFEEREBATE")
'
'        If Len(strPackgedValuationRebate) > 0 Then
'            strDescription3 = strDescription3 & "In the event of the valuation not taking place £" & strPackgedValuationRebate & " will be refunded."
'        End If
    
    'AW  15/05/2006  EP520 - End
    ' BUI
    ElseIf CheckForValidationType(strCostType, "BUI") Then
        
        'Need to add entry for BuildingsInsurance fee. At the moment there does not appear to be a validation
        'Need to check whether it has been paid, as the text changes accordingly.
        
        strDescription2 = "This fee is for arranging insurance to protect our interests as Lender when a Borrower fails to insure the mortgaged property in accordance with the Mortgage Conditions."
        strDescription3 = "This fee is non-refundable."
        'EP2_1980
        If vblnIsOfferDocument Then
            If lngAmountOutstanding > 0 Then
                strDescription = "This fee is added to the loan on the completion date."
            Else
                strDescription = "This fee has already been paid."
            End If
        Else
            strDescription = "This fee is payable on application or can be added to your loan on completion."
        End If

        
    ' REV
    ElseIf CheckForValidationType(strCostType, "REV") Then
        
        'Revaluation fee
        
        strDescription2 = "This fee is to carry out a revaluation of the security property."
        strDescription3 = "This fee is non-refundable."
        'EP2_1980
        If vblnIsOfferDocument Then
            strDescription = "This fee has already been paid."
        Else
            strDescription = "This fee is payable on application."
        End If
    'EP2_1980
    ' AB
    ElseIf CheckForValidationType(strCostType, "AB") Then
        
        'Additional Borrowing fee
        
        strDescription2 = "This fee is to cover our administration expenses."
        
        'EP2_1930 GHun
        'If vblnIsOfferDocument Then
        '    If lngAmountOutstanding > 0 Then
        '        strDescription = "This fee is added to your account on completion."
        '    Else
        '        strDescription = "This fee has already been paid."
        '    End If
        'Else
        '    strDescription = "This fee is payable on application or can be added to your loan on completion."
        'End If

        strDescription = "This fee is added to your loan on completion of the additional borrowing."
        'EP2_1930 End
               
        strDescription3 = "This fee is non-refundable."
    'EP2_1980
    ' TOE
    ElseIf CheckForValidationType(strCostType, "TOE") Then
        
        'Transfer Of Equity fee
        strDescription2 = "This fee is charged when one or more people transfer their share of the interest in a property to others."
        strDescription3 = "This fee is non-refundable."
        
        If vblnIsOfferDocument Then
            strDescription = "This fee has already been paid."
        Else
            strDescription = "This fee is payable on application."
        End If
    'EP2_1980
    ' PSF
    ElseIf CheckForValidationType(strCostType, "PSF") Then
        
        'Product Switch fee
        strDescription2 = "This fee is to cover our administration expenses."
        
        'EP2_1930 GHun
        'If vblnIsOfferDocument Then
        '    If lngAmountOutstanding > 0 Then
        '        strDescription = "This fee is added to your account on completion of the Product Switch."
        '    Else
        '        strDescription = "This fee has already been paid."
        '    End If
        'Else
        '    strDescription = "This fee is payable on application or can be added to your loan on completion."
        'End If
        
        strDescription = "This fee is added to your loan on completion of the product switch."
        'EP2_1930 End
        
        strDescription3 = "This fee is non-refundable."
        
    'EP2_1930 GHun
    ElseIf CheckForValidationType(strCostType, "ADM") Then
        'Admin fee
        strDescription2 = "This fee is to cover our administration expenses."
    
        If vblnIsOfferDocument Then
            strDescription = "This fee has already been paid."
        Else
            strDescription = "This fee is payable on application."
        End If
        
        strDescription3 = "This fee is non-refundable."
       
    ElseIf CheckForValidationType(strCostType, "PK") Then
        'Packager fee
        strBrokerName = vobjCommon.PackagerName
        If Len(strBrokerName) = 0 Then
            strBrokerName = "your packager"
        End If
        
        strDescription2 = "This fee is payable to " & strBrokerName & " for arranging this mortgage."
        'EP2_2477 GHun
        If vblnIsOfferDocument Then
            strDescription = "This fee has already been paid."
        Else
            strDescription = "This fee is payable on application."
        End If
        'EP2_2477 End
                
        If xmlAttributeValueExists(vxmlCost, "REFUNDAMOUNT") Then
            strBrokerRebate = set2DP(xmlGetAttributeText(vxmlCost, "REFUNDAMOUNT"))
        End If
                      
        If Len(strBrokerRebate) = 0 Or strBrokerRebate = "0.00" Then
            strDescription3 = "This fee is non-refundable."
        Else
            strDescription3 = "£" & strBrokerRebate & " will be refundable on completion."
        End If
        
    'EP2_1930 End
    End If
    
    'EPSOM EP454/EP479 - 27/04/2006 SAB - END
    
    GetFeeDetailDescription = strDescription
    
    Set xmlfees = Nothing
    Set xmlFee = Nothing
    Set xmlNewProperty = Nothing
    Set xmlItem = Nothing
    Set xmlAppFactFindNode = Nothing
    'Set xmlLender = Nothing 'PB EP821
    Set xmlBrokerContact = Nothing 'PB EP827
Exit Function
ErrHandler:
    Set xmlfees = Nothing
    Set xmlFee = Nothing
    Set xmlNewProperty = Nothing
    Set xmlItem = Nothing
    Set xmlAppFactFindNode = Nothing
    'Set xmlLender = Nothing 'PB EP821
    Set xmlBrokerContact = Nothing 'PB EP827
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildInsuranceSection
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the insurance section
'**                 of all documents.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section9 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildInsuranceSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection9Common"
    Dim blnHasBC As Boolean
    Dim blnHasPP As Boolean
    Dim strBCCover As String
    Dim strPPCover As String
    Dim dblBCCost As Double
    Dim dblPPCost As Double
    Dim strContactName As String
    Dim strIntermediaryContactName As String
    Dim strIntermediaryCompanyName As String
    Dim blnIsIntermediary As Boolean
    Dim xmlTiedInsurance As IXMLDOMNode
    Dim xmlLenderOrIntermed As IXMLDOMNode
    Dim xmlPolicyDetails As IXMLDOMNode
    Dim xmlOptionalInsurance As IXMLDOMNode
    Dim xmlRequiredInsurance As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strPropertyTenureType As String  'SR 22/10/2004 : BBG1687
    Dim xmlPurposeNode As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    '*-get the tied insurance data
    Call GetAllInsuranceData(vobjCommon, True, True, blnHasBC, strBCCover, _
            dblBCCost, blnHasPP, strPPCover, dblPPCost)
     
    '*-add the TIEDINSURANCE element
    Set xmlTiedInsurance = vobjCommon.CreateNewElement("TIEDINSURANCE", vxmlNode)
    
    'EP2_139 Use company instead of Introducer name
    blnIsIntermediary = vobjCommon.IsIntroducedByIntermediary(strIntermediaryContactName, False, strIntermediaryCompanyName, True)
    
    If blnIsIntermediary Then
        Set xmlLenderOrIntermed = vobjCommon.CreateNewElement("INTERMEDIARY", xmlTiedInsurance)
        xmlSetAttributeValue xmlLenderOrIntermed, "INTERMEDIARYNAME", strIntermediaryCompanyName



    Else
        Set xmlLenderOrIntermed = vobjCommon.CreateNewElement("LENDER", xmlTiedInsurance)
    End If
            
    If blnHasBC = False And blnHasPP = False Then
        Set xmlNode = vobjCommon.CreateNewElement("NONEREQUIRED", xmlLenderOrIntermed)
        'EP2_139 Use Company Name instead of Introducer
        xmlSetAttributeValue xmlNode, "INTERMEDIARYNAME", strIntermediaryCompanyName


        'SR 22/10/2004 : BBG1687
        strPropertyTenureType = GetPropertyTenureType(vobjCommon)
        If InStr(strPropertyTenureType, "C,") > 0 Or InStr(strPropertyTenureType, "L,") > 0 Then
            vobjCommon.CreateNewElement "LANDLORDINSURANCE", xmlNode
        End If
        'SR 22/10/2004 : BBG1687 - End
    Else
        If blnHasBC Then
            'Add policy details for B&C policy
            Set xmlPolicyDetails = vobjCommon.CreateNewElement("POLICYDETAILS", xmlLenderOrIntermed)
            xmlSetAttributeValue xmlPolicyDetails, "PREMIUM", dblBCCost
            xmlSetAttributeValue xmlPolicyDetails, "DESCRIPTION", strBCCover
            
            Set xmlNode = vobjCommon.CreateNewElement("DETAILS", xmlPolicyDetails)
            xmlSetAttributeValue xmlNode, "DESCRIPTION", "Reviewed Annually"
        End If
        
        If blnHasPP Then
            'Add policy details for PP policy
            Set xmlPolicyDetails = vobjCommon.CreateNewElement("POLICYDETAILS", xmlLenderOrIntermed)
            xmlSetAttributeValue xmlPolicyDetails, "PREMIUM", dblPPCost
            xmlSetAttributeValue xmlPolicyDetails, "DESCRIPTION", strPPCover
            
            Set xmlNode = vobjCommon.CreateNewElement("DETAILS", xmlPolicyDetails)
            xmlSetAttributeValue xmlNode, "DESCRIPTION", "Reviewed Annually"
        End If
    End If
    
    Set xmlRequiredInsurance = vobjCommon.CreateNewElement("REQUIREDINSURANCE", vxmlNode)
    'EP2_1449
    Set xmlPurposeNode = vobjCommon.AddInsuranceLoanPurposeElement(xmlRequiredInsurance)
    '*-add the MORTGAGETYPE attribute
    Call xmlSetAttributeValue(xmlPurposeNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)


    'EP610 - DRC
    If blnIsIntermediary Then
        Set xmlLenderOrIntermed = vobjCommon.CreateNewElement("INTERMEDIARY", xmlRequiredInsurance)
        'EP2_139 Use Company Name instead of Introducer
        xmlSetAttributeValue xmlLenderOrIntermed, "INTERMEDIARYNAME", strIntermediaryCompanyName
    End If
    'EP610 - End
    '*-add the OPTIONALINSURANCE element
    If blnHasBC = False And dblBCCost <> 0 Then
        Set xmlOptionalInsurance = vobjCommon.CreateNewElement("OPTIONALINSURANCE", vxmlNode)
        Set xmlPolicyDetails = vobjCommon.CreateNewElement("POLICYDETAILS", xmlOptionalInsurance)
        
        'Add policy details for B&C policy
        xmlSetAttributeValue xmlPolicyDetails, "PREMIUM", dblBCCost
        xmlSetAttributeValue xmlPolicyDetails, "DESCRIPTION", strBCCover
            
        Set xmlNode = vobjCommon.CreateNewElement("DETAILS", xmlPolicyDetails)
        xmlSetAttributeValue xmlNode, "DESCRIPTION", "Reviewed Annually"
    End If
    
    If blnHasPP = False And dblPPCost <> 0 Then
        If xmlOptionalInsurance Is Nothing Then
            Set xmlOptionalInsurance = vobjCommon.CreateNewElement("OPTIONALINSURANCE", vxmlNode)
        End If
        Set xmlPolicyDetails = vobjCommon.CreateNewElement("POLICYDETAILS", xmlOptionalInsurance)
        
        'Add policy details for PP policy
        xmlSetAttributeValue xmlPolicyDetails, "PREMIUM", dblPPCost
        xmlSetAttributeValue xmlPolicyDetails, "DESCRIPTION", strPPCover
        
        Set xmlNode = vobjCommon.CreateNewElement("DETAILS", xmlPolicyDetails)
        xmlSetAttributeValue xmlNode, "DESCRIPTION", "Reviewed Annually"
    End If

    Set xmlTiedInsurance = Nothing
    Set xmlLenderOrIntermed = Nothing
    Set xmlRequiredInsurance = Nothing
    Set xmlPolicyDetails = Nothing
    Set xmlOptionalInsurance = Nothing
    Set xmlNode = Nothing
    Set xmlPurposeNode = Nothing
Exit Sub
ErrHandler:
    Set xmlTiedInsurance = Nothing
    Set xmlLenderOrIntermed = Nothing
    Set xmlRequiredInsurance = Nothing
    Set xmlPolicyDetails = Nothing
    Set xmlOptionalInsurance = Nothing
    Set xmlNode = Nothing
    Set xmlPurposeNode = Nothing

    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'********************************************************************************
'** Function:       GetAllInsuranceData
'** Created by:     Srini Rao
'** Date:           19/05/2004
'** Description:    Adds the INSURANCE elements for any insurance (compulsory or optional).
'** Parameters:     vxmlLoanComponent - the loan component.
'**                 vxmlNode - the node to add the elements to.
'** Returns:        N/A
'** Errors:         None Expected
''********************************************************************************
Private Sub GetAllInsuranceData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vblnGetBC As Boolean, ByVal vblnGetPP As Boolean, _
        ByRef rblnHasBC As Boolean, ByRef rstrBCCoverType As String, _
        ByRef rdblBCMonthlyCost As Double, ByRef rblnHasPP As Boolean, _
        ByRef rstrPPCoverType As String, ByRef rdblPPMonthlyCost As Double)
    Const cstrFunctionName As String = "GetAllInsuranceData"
    Dim xmlParams As IXMLDOMNode
    Dim strCover As String
    Dim xmlComponents As IXMLDOMNodeList
    Dim xmlComponent As IXMLDOMNode
    Dim xmlQuote As IXMLDOMNode
    Dim xmlDetails As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Set xmlComponents = vobjCommon.LoanComponents
    For Each xmlComponent In xmlComponents
        Set xmlParams = xmlComponent.selectSingleNode("//MORTGAGEPRODUCTPARAMETERS")
        If Not xmlParams Is Nothing Then
            '*-only get the buildings and contents quote data once
            If vblnGetBC Then
                rblnHasBC = xmlGetAttributeAsBoolean(xmlParams, "COMPULSORYBC")
                vblnGetBC = False
                    
                '*-get the quote
                Set xmlQuote = vobjCommon.Data.selectSingleNode("//BUILDINGSANDCONTENTSSUBQUOTE")
                If Not xmlQuote Is Nothing Then
                    rdblBCMonthlyCost = xmlGetAttributeAsDouble(xmlQuote, "TOTALBCMONTHLYCOST")
                        
                    '*-get the type of insurance
                    Set xmlDetails = vobjCommon.Data.selectSingleNode("//BUILDINGSANDCONTENTSDETAILS")
                    strCover = xmlGetAttributeText(xmlDetails, "COVERTYPE")
                    Select Case strCover
                        Case "BC,"  'SR 28/09/2004 : CORE82
                            rstrBCCoverType = "Buildings & Contents"
                        Case "B,"  'SR 28/09/2004 : CORE82
                            rstrBCCoverType = "Buildings Only Cover"
                        Case Else
                            rstrBCCoverType = "Contents Only Cover"
                    End Select
                End If
            End If
            
            '*-only get the payment protection quote data once
            If vblnGetPP Then
                rblnHasPP = xmlGetAttributeAsBoolean(xmlParams, "COMPULSORYPP")
                vblnGetPP = False
                    
                Set xmlQuote = vobjCommon.Data.selectSingleNode("//PAYMENTPROTECTIONSUBQUOTE")
                If Not xmlQuote Is Nothing Then
                    rdblPPMonthlyCost = xmlGetAttributeAsDouble(xmlQuote, "TOTALPPMONTHLYCOST")
                    '*-get the type of insurance
                    rstrPPCoverType = "Mortgage Payment Protection"
                End If
            End If
        End If
    Next xmlComponent

    Set xmlParams = Nothing
    Set xmlComponents = Nothing
    Set xmlComponent = Nothing
    Set xmlQuote = Nothing
    Set xmlDetails = Nothing
Exit Sub
ErrHandler:
    Set xmlParams = Nothing
    Set xmlComponents = Nothing
    Set xmlComponent = Nothing
    Set xmlQuote = Nothing
    Set xmlDetails = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'Reworked for EP2_422
'********************************************************************************
'** Function:       BuildUsingIntermediarySection
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section13
'**                 (Using a mortgage intermediary) element.
'** Parameters:     vxmlNode - the section13 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildUsingIntermediarySection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildUsingIntermediarySection"
    Dim dblTPFee As Double
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode 'SR 17/10/2004 : BBG1643
    
    On Error GoTo ErrHandler

    'SR 13/10/2004 : BBG1593
    'ik_EP2_2158
    If Len(vobjCommon.FeePayableToThirdPartyAmount()) <> 0 Then
        dblTPFee = vobjCommon.FeePayableToThirdPartyAmount()
    End If
    'ik_EP2_2158_ends
    
    If dblTPFee > 0 Then 'SR 23/09/2004 : CORE82
        Set xmlNode = vobjCommon.CreateNewElement("INTERMEDIARYPAYMENT", vxmlNode)
        'EP2_704
        xmlSetAttributeValue xmlNode, "INTERMEDIARYNAMES", vobjCommon.IntermediaryNames
        xmlSetAttributeValue xmlNode, "FEEAMOUNT", CStr(dblTPFee)
        
        If vobjCommon.IsProcFeeRefund Then
            Set xmlTempNode = vobjCommon.CreateNewElement("BROKERREFUND", xmlNode)
            xmlSetAttributeValue xmlTempNode, "BROKERREFUNDAMOUNT", vobjCommon.ProcFeeRefundAmount
            xmlSetAttributeValue xmlTempNode, "INTERMEDIARYNAME", vobjCommon.IntroducerName
            Call AddApplicantNameAttribute(vobjCommon.Data, xmlTempNode)
        
        End If
            
        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call xmlSetAttributeValue(xmlNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
        End If
    Else
        Set xmlNode = vobjCommon.CreateNewElement("NOINTERMEDIARYPAYMENT", vxmlNode)
        xmlSetAttributeValue xmlNode, "INTERMEDIARYNAMES", vobjCommon.IntermediaryNames
    End If
    
    'EP2_583 PROVIDER Required for KFI as well
    '*-add the PROVIDER attribute
    Call xmlSetAttributeValue(xmlNode, "PROVIDER", vobjCommon.Provider)

    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'Moved FeePayableToThirdParty to CommonDataHelper for EP2_422
'SR 13/10/2004 : BBG1593
'Private Function FeePayableToThirdParty(ByVal vobjCommon As CommonDataHelper) As Double
'
'    Const cstrFunctionName As String = "FeePayableToThirdParty"
'    Dim xmlMortgageOneOffCostList As IXMLDOMNodeList
'    Dim xmlNode As IXMLDOMNode
'    Dim dblTPPayment As Double, strValidationType As String
'
'On Error GoTo ErrHandler
'
''AW     22/05/2006  EP590
''    Set xmlMortgageOneOffCostList = vobjCommon.Data.selectNodes(gcstrMORTGAGEONEOFFCOSTS_PATH)
''    For Each xmlNode In xmlMortgageOneOffCostList
''        strValidationType = xmlGetAttributeText(xmlNode, "MORTGAGEONEOFFCOSTTYPE")
''        If InStr(strValidationType, "PRC") > 0 Or _
''           InStr(strValidationType, "PAK") > 0 Or _
''           InStr(strValidationType, "MKT") > 0 Then
''            'TK 10/11/2004 BBG1791 value types from MortgageOneOff table are long whereas value types from ECROS are double.
''            dblTPPayment = dblTPPayment + xmlGetAttributeAsLong(xmlNode, "AMOUNT")
''        End If
''    Next xmlNode
'
'    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
'
'    If Not xmlNode Is Nothing Then
'        dblTPPayment = xmlGetAttributeAsDouble(xmlNode, "TOTALPROCURATIONFEE")
'    End If
''AW     22/05/2006  EP590 - End
'
'   FeePayableToThirdParty = dblTPPayment
'
'    Set xmlMortgageOneOffCostList = Nothing
'    Set xmlNode = Nothing
'ErrHandler:
'    Set xmlMortgageOneOffCostList = Nothing
'    Set xmlNode = Nothing
'    errCheckError cstrFunctionName, mcstrModuleName
'End Function
'SR 13/10/2004 : BBG1593 - End

'********************************************************************************
'** Function:       AddPeriodAttribute
'** Created by:     Andy Maggs
'** Date:           15/06/2004
'** Description:    Adds the period attribute to the destination node. This method
'**                 uses the period defined in the source node if it exists,
'**                 otherwise it computes it by taking the difference in months
'**                 between the end date and the supplied start date.
'** Parameters:     vxmlSrcNode - the source node.
'**                 vstrSrcPeriodAttribName - the name of the attribute on the
'**                 source node containing the period value.
'**                 vstrSrcEndDateAttribName - the name of the attribute on the
'**                 source node containing the end date value.
'**                 vdtStartDate - the start date of the period.
'**                 vstrDestPeriodAttribName - the name of the attribute to create
'**                 on the destination node containing the period value.
'**                 vxmlDestNode - the node to add the period attribute to.
'** Returns:        The actual period added.
'** Errors:         None Expected
'********************************************************************************

Public Function AddPeriodAttribute(ByVal vxmlSrcNode As IXMLDOMNode, _
        ByVal vstrSrcPeriodAttribName As String, _
        ByVal vstrSrcEndDateAttribName As String, _
        ByVal vdtStartDate As Date, _
        ByVal vstrDestPeriodAttribName As String, _
        ByVal vxmlDestNode As IXMLDOMNode, _
        Optional ByVal vxmlPeriodText As String = "") As Integer 'IK EP531
        
    Const cstrFunctionName As String = "AddPeriodAttribute"
    Dim intPeriod As Integer
    Dim dtEndDate As Date
    
    On Error GoTo ErrHandler

    '*-add the period attribute in months - get the period or get the end date
    If xmlAttributeValueExists(vxmlSrcNode, vstrSrcPeriodAttribName) Then
        intPeriod = xmlGetAttributeAsInteger(vxmlSrcNode, vstrSrcPeriodAttribName)
    Else
        dtEndDate = xmlGetAttributeAsDate(vxmlSrcNode, vstrSrcEndDateAttribName)
        ' PB 24/03/2006 EP503/MAR1777
        'intPeriod = DateDiff("m", vdtStartDate, dtEndDate)
        intPeriod = MonthDiff(vdtStartDate, dtEndDate)  'MAR1777 GHun
        ' EP503/MAR1777 End
    End If
    Call xmlSetAttributeValue(vxmlDestNode, vstrDestPeriodAttribName, CStr(intPeriod))
    
    '*-return the actual period for information purposes
    AddPeriodAttribute = intPeriod
    
Exit Function
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       BuildMortgageNoMoreSection
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the (What happens if you
'**                 do not want this mortgage any more?) element that appears in
'**                 all documents albeit with different section numbers.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section10 element to set the attributes on.
'**                 vblnIsOfferDocument - whether the section is being created
'**                 for the slightly different offer document.
'**                 vblnIsLifetimeDocument - whether this is being called for a
'**                 lifetime mortgage document (KFI or Offer).
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildMortgageNoMoreSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean, _
        ByVal vblnIsLifetimeDocument As Boolean)
    Const cstrFunctionName As String = "BuildMortgageNoMoreSection"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlFeesList As IXMLDOMNodeList
    Dim xmlFee As IXMLDOMNode
    Dim intNumber As Integer
    Dim xmlEarly As IXMLDOMNode
    Dim xmlCharges As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlRepay As IXMLDOMNode
    Dim xmlBandList As IXMLDOMNodeList
    Dim xmlBand As IXMLDOMNode
    Dim intStepNum As Integer
    Dim xmlPeriod As IXMLDOMNode
    Dim dblPercentage As Double
    Dim strBasis As String
    Dim blnPortable As Boolean
    Dim xmlProduct As IXMLDOMNode
    Dim xmlSummary As IXMLDOMNode
    Dim xmlClient As IXMLDOMNode
    Dim dblMaxCharge As Double
    Dim xmlRepaymentCharges As IXMLDOMNode
    Dim xmlPlusFee As IXMLDOMNode
    Dim xmlPlusValuation As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    'CORE82
    Dim lngRepaymentFee As Long
    Dim dblChargeAmount As Double
    Dim dblRedemptionFeeAmount As Double 'SR 23/09/2004 : CORE82
    Dim dblTotalRedemptionFeeAmount As Double  'SR 13/09/2004 : CORE82
    Dim dblRedAdminAmount As Double 'PM 26/05/2006 : EP628
    Dim bCalcDatePlus2Years As Boolean 'MAR88 BC
    Dim xmlPortable As IXMLDOMNode 'MAR1535 BC
    Dim bRatetype As Boolean 'MAR1535 BC
    ' PB 05/06/2006 EP561/MAR1590/MAR1818/etc begin
    Dim xmlProductLanguage As IXMLDOMNode
    Dim blnSingleComponent As Boolean
    Dim strProductName As String
    Dim xmlMultiComponent As IXMLDOMNode
    ' PB 05/06/2006 EP561/MAR1590 End
    ' PB 21/06/2006 EP796 Begin
    Dim intCurrentRepaymentCharge As Integer
    Dim strPeriodText As String
    Dim intRepaymentCharges As Integer
    Dim dblChargeAmountNoAdmin As Double
    ' PB EP796 End
    Dim xmlPeriodText As IXMLDOMAttribute
    'EP2_139
    Dim blnCashback As Boolean
    Dim cashBack As Long
    Dim totalCashBack As Long
    Dim xmlCashback As IXMLDOMNode
    Dim xmlMortgageIncentive As IXMLDOMNode
    Dim xmlMortgageIncentiveList As IXMLDOMNodeList
    Dim seqLCNum As Integer
    
    Dim dblTotalCashBackClawBack As Double
    Dim dblgTotalMaxChargeNoAdmin As Double
'    Dim dblgTotalMaxCharge As Double
    
    Dim xmlVariable As IXMLDOMNode
    Dim xmlMultiCharge As IXMLDOMNode
    Dim xmlMultiChargeTemp As IXMLDOMNode
    
    bCalcDatePlus2Years = False 'MAR88 BC
    blnSingleComponent = False 'PB EP796
    
    On Error GoTo ErrHandler
'SR 22/09/2004 : CORE82
'    If (Not vblnIsOfferDocument) Or vblnIsLifetimeDocument Then
'        '*-add the CLIENTSPECIFIC element
'        Set xmlClient = vobjCommon.CreateNewElement("CLIENTSPECIFIC", vxmlNode)
'        '*-add the CLIENTWORDING attribute
'        Call xmlSetAttributeValue(xmlClient, "CLIENTWORDING", "")
'    End If
'SR 22/09/2004 : CORE82 - End
        
    '*-get the list of loan components and from each component, get the list of
    '*-redemption fees
    Set xmlList = vobjCommon.LoanComponents
    
    'PB 07/06/2006 EP651 MAR1818/1590/1738/etc
    If xmlList.length = 1 Then
        blnSingleComponent = True
    End If
    'PB EP651 MAR1818/1590/1738/etc End
    totalCashBack = 0
    For Each xmlLoanComponent In xmlList
        Set xmlFeesList = xmlLoanComponent.selectNodes("LOANCOMPONENTREDEMPTIONFEE")
        intNumber = intNumber + xmlFeesList.length
        
        'EP2_2478 need to know the total cashback
        seqLCNum = xmlGetAttributeAsInteger(xmlLoanComponent, "LOANCOMPONENTSEQUENCENUMBER")
        Set xmlMortgageIncentiveList = xmlLoanComponent.selectNodes("//LOANCOMPONENT/MORTGAGEINCENTIVE[@LOANCOMPONENTSEQUENCENUMBER=" & seqLCNum & "][INCLUSIVEINCENTIVE]")
        If xmlMortgageIncentiveList.length > 0 Then
            For Each xmlMortgageIncentive In xmlMortgageIncentiveList
                totalCashBack = totalCashBack + xmlGetAttributeAsDouble(xmlMortgageIncentive, "INCENTIVEAMOUNT")
            Next xmlMortgageIncentive
        End If
        Set xmlMortgageIncentiveList = Nothing
        
    Next xmlLoanComponent
    
    'SR 22/10/2004 : BBG1687
    '*-is the mortgage product portable?
    Set xmlProduct = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEPRODUCT_PATH)
    If Not xmlProduct Is Nothing Then
        If xmlAttributeValueExists(xmlProduct, "MORTGAGEPRODUCTPORTABLEIND") Then
            If xmlGetAttributeAsBoolean(xmlProduct, "MORTGAGEPRODUCTPORTABLEIND") Then
                blnPortable = True
            End If
        End If
    End If
    'SR 22/10/2004 : BBG1687 - End
    
    'EP2_139
    If vobjCommon.IsCashback Then
        blnCashback = True
    End If
    
    'PB EP2_1931 20/03/2007  INR Still want ERCs if vobjCommon.IsAdditionalBorrowing
    If intNumber = 0 Then
        '*-add the NOREPAYMENTCHARGES element
        Set xmlTemp = vobjCommon.CreateNewElement("NOREPAYMENTCHARGES", vxmlNode)
        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, xmlTemp)
        End If
    Else
        '*-add the EARLYREPAYMENT element
        Set xmlEarly = vobjCommon.CreateNewElement("EARLYREPAYMENT", vxmlNode)
        '*-add the CHARGES element
        Set xmlCharges = vobjCommon.CreateNewElement("CHARGES", xmlEarly)
        
        intNumber = 0
        dblTotalRedemptionFeeAmount = 0 'SR 13/09/2004 : CORE82
        For Each xmlLoanComponent In xmlList
            intCurrentRepaymentCharge = 0
            '*-increment the count
            intNumber = intNumber + 1
            '*-add the COMPONENT element
            Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", xmlCharges)
            
            ' PB 07/06/2006 EP651 MAR1818/1590/1738/etc Begin
            ''*-add the mandatory PART attribute
            'Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intNumber))
            '
            If blnSingleComponent = False Then
                Set xmlMultiComponent = vobjCommon.CreateNewElement("MULTICOMPONENT", xmlComponent)
                Set xmlProduct = xmlLoanComponent.selectSingleNode("MORTGAGEPRODUCT")
                Set xmlProductLanguage = xmlProduct.selectSingleNode("MORTGAGEPRODUCTLANGUAGE")
                strProductName = xmlGetAttributeText(xmlProductLanguage, "PRODUCTNAME")
                Call xmlSetAttributeValue(xmlMultiComponent, "PART", CStr(intNumber))
                Call xmlSetAttributeValue(xmlMultiComponent, "PRODUCTNAME", strProductName)
            End If
            'PB EP651 MAR1818/1590/1738/etc End
            
            '*-get the redemption fees list
            Set xmlFeesList = xmlLoanComponent.selectNodes("LOANCOMPONENTREDEMPTIONFEE")
            If xmlFeesList.length = 0 Then
                '*-add the NOREPAYMENTCHARGES element
                Call vobjCommon.CreateNewElement("NOREPAYMENTCHARGES", xmlComponent)
            Else
                '*-add the REPAYMENTCHARGES element
                Set xmlRepay = vobjCommon.CreateNewElement("REPAYMENTCHARGES", xmlComponent)
                '*-add the mandatory LOANAMOUNT attribute
                'Call xmlCopyAttributeValue(xmlLoanComponent, xmlRepay, "LOANAMOUNT", "LOANAMOUNT") 'SR 21/09/2004 : CORE82
                'SR 11/04/2007 : EP2_2341
                'xmlSetAttributeValue xmlRepay, "LOANAMOUNT", _
                '        CStr(xmlGetAttributeAsLong(xmlLoanComponent, "LOANAMOUNT") + vobjCommon.FeesAddedToLoanAmount) 'BC MAR1367 06/03/2006
                xmlSetAttributeValue xmlRepay, "LOANAMOUNT", CStr(xmlGetAttributeAsLong(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT"))
                'SR 11/04/2007 : EP2_2341 - End

                AddCalculationDateAttribute vobjCommon, xmlComponent, bCalcDatePlus2Years 'MAR88 BC 22/09/2004
                
                 'EP2_1931 If single component & not Add borrowing, Need Redemption Admin Fee
                Set xmlTemp = vobjCommon.Data.selectSingleNode(gcstrMORTGAGELENDER_PATH)
                dblRedAdminAmount = 0
                If Not xmlTemp Is Nothing Then
                    dblRedAdminAmount = xmlGetAttributeAsDouble(xmlTemp, "DEEDSRELEASEFEE")
                    Set xmlTemp = Nothing
                End If
                
                'EP2_139 do we have CASHBACK
                If blnCashback Then
                    seqLCNum = xmlGetAttributeAsInteger(xmlLoanComponent, "LOANCOMPONENTSEQUENCENUMBER")
                    'EP2_1931 could have multiple incentives on a loan component
                    'or multiple loan components with multiple incentives
                    Set xmlMortgageIncentiveList = xmlLoanComponent.selectNodes("//LOANCOMPONENT/MORTGAGEINCENTIVE[@LOANCOMPONENTSEQUENCENUMBER=" & seqLCNum & "][INCLUSIVEINCENTIVE]")
                    
                    If xmlMortgageIncentiveList.length > 0 Then
                        For Each xmlMortgageIncentive In xmlMortgageIncentiveList
                            cashBack = cashBack + xmlGetAttributeAsDouble(xmlMortgageIncentive, "INCENTIVEAMOUNT")
                        Next xmlMortgageIncentive
                        'Only want 1 cashback node, could be with any of the LCs
                        Set xmlCashback = xmlEarly.selectSingleNode("CASHBACK")
                        If xmlCashback Is Nothing Then
                            Set xmlCashback = vobjCommon.CreateNewElement("CASHBACK", xmlEarly)
                            
                            Set xmlMultiChargeTemp = xmlCashback.selectSingleNode("NOTADDBORROW")
                            If Not vobjCommon.IsAdditionalBorrowing And xmlMultiChargeTemp Is Nothing Then
                                Set xmlMultiCharge = vobjCommon.CreateNewElement("NOTADDBORROW", xmlCashback)
                                Call xmlSetAttributeValue(xmlMultiCharge, "REDADMINAMOUNT", set2DP(CStr(dblRedAdminAmount)))
                            End If
                        End If
                    'PB 03/01/2007 EP2_648 Begin
                    Else
                        'No Cashback
                        'EP2_2478 Only want a NOCASHBACK NODE if we aren't expecting a CASHBACK
                        If totalCashBack = 0 Then
                            'Only want 1 cashback node, could be with any of the LCs
                            Set xmlCashback = xmlEarly.selectSingleNode("NOCASHBACK")
                            If xmlCashback Is Nothing Then
                                Set xmlCashback = vobjCommon.CreateNewElement("NOCASHBACK", xmlEarly)
                            End If
                            
                            Set xmlMultiChargeTemp = xmlCashback.selectSingleNode("NOTADDBORROW")
                            If Not vobjCommon.IsAdditionalBorrowing And xmlMultiChargeTemp Is Nothing Then
                                Set xmlMultiCharge = vobjCommon.CreateNewElement("NOTADDBORROW", xmlCashback)
                                Call xmlSetAttributeValue(xmlMultiCharge, "REDADMINAMOUNT", set2DP(CStr(dblRedAdminAmount)))
                                'EP2_1931 If we have a freevaluation, we still need the cashback text
                                If vobjCommon.IsRefundOfValuation Then
                                    Set xmlMultiCharge = vobjCommon.CreateNewElement("CASHBACK", xmlCashback)
                                    Call xmlSetAttributeValue(xmlMultiCharge, "FREELEGAL", set2DP(CStr(vobjCommon.GetFreeValuationAmount)))
                                End If
    
                            End If
                        End If
                    'EP2_648 End
                    End If
                Else
                    cashBack = 0
                    Set xmlCashback = xmlEarly.selectSingleNode("NOCASHBACK")
                    If xmlCashback Is Nothing Then
                        Set xmlCashback = vobjCommon.CreateNewElement("NOCASHBACK", xmlEarly)
                    End If

                    Set xmlMultiChargeTemp = xmlCashback.selectSingleNode("NOTADDBORROW")
                    If Not vobjCommon.IsAdditionalBorrowing And xmlMultiChargeTemp Is Nothing Then
                        Set xmlMultiCharge = vobjCommon.CreateNewElement("NOTADDBORROW", xmlCashback)
                        Call xmlSetAttributeValue(xmlMultiCharge, "REDADMINAMOUNT", set2DP(CStr(dblRedAdminAmount)))
                        'EP2_1931 If we have a freevaluation, we still need the cashback text
                        If vobjCommon.IsRefundOfValuation Then
                            Set xmlMultiCharge = vobjCommon.CreateNewElement("CASHBACK", xmlCashback)
                            Call xmlSetAttributeValue(xmlMultiCharge, "FREELEGAL", set2DP(CStr(vobjCommon.GetFreeValuationAmount)))
                        End If
                    End If
                End If

                'EP2_1931
                If blnSingleComponent = False And Not vobjCommon.IsAdditionalBorrowing Then
                    Set xmlMultiCharge = xmlEarly.selectSingleNode("MULTIREPAYCHARGE")
                    If xmlMultiCharge Is Nothing Then
                        Set xmlMultiCharge = vobjCommon.CreateNewElement("MULTIREPAYCHARGE", xmlEarly)
                    End If
                End If
                
                '*-add the VARIABLERATES element
                ' SR 22/09/2004 : CORE82 - Do not add for LifeTime docs
                If Not vblnIsLifetimeDocument Then
                    Call vobjCommon.CreateNewElement("VARIABLERATES", xmlRepay)
                End If
                
                lngRepaymentFee = vobjCommon.MortgageCompletionFee
                dblRedemptionFeeAmount = GetMaxDblAttribValue(xmlFeesList, "REDEMPTIONFEEAMOUNT")  'SR 23/09/2004 : CORE82
                dblTotalRedemptionFeeAmount = dblTotalRedemptionFeeAmount + dblRedemptionFeeAmount 'SR 13/09/2004 : CORE82
                
                '*-get the redemption fee band list
'               MAR54 add extra qualification to retrieve REDEMPTIONFEEBAND
                Set xmlBandList = xmlLoanComponent.selectNodes(".//MORTGAGEPRODUCT/REDEMPTIONFEEBAND")
                For Each xmlBand In xmlBandList
                     
                     'PB 21/06/2006 EP796 Start
                     intCurrentRepaymentCharge = intCurrentRepaymentCharge + 1
                     If intCurrentRepaymentCharge = 1 Then
                        strPeriodText = "the first "
                     Else
                        If intCurrentRepaymentCharge = intRepaymentCharges Then
                            strPeriodText = "the final "
                        Else
                            strPeriodText = "the next "
                        End If
                    End If
                    'PB EP796 End

                    '*-get the step number for this band item
                    intStepNum = xmlGetAttributeAsInteger(xmlBand, "REDEMPTIONFEESTEPNUMBER")
                    '*-and get the corresponding redemption fee item
                    Set xmlFee = xmlLoanComponent.selectSingleNode("LOANCOMPONENTREDEMPTIONFEE[@REDEMPTIONFEESTEPNUMBER=" & intStepNum & "]")
                    
                    '*-add the PERIOD element
                    Set xmlPeriod = vobjCommon.CreateNewElement("PERIOD", xmlRepay)
                    
                    'EP2_1931
                    ' 1 LC ChargeAmount is ERC + Admin Fee + clawbacks (Cashback + Valuation)
                    ' >1 LC ChargeAmount is ERC + clawbacks (Cashback + Valuation) for the LC
                    If Not xmlFee Is Nothing Then
                        dblChargeAmount = xmlGetAttributeAsDouble(xmlFee, "REDEMPTIONFEEAMOUNT")
                    End If
                    
                    dblChargeAmount = dblChargeAmount + cashBack
                    'EP2_2478 Only add into the first component
                    If vobjCommon.IsRefundOfValuation And intNumber = 1 Then
                        cashBack = cashBack + vobjCommon.GetFreeValuationAmount
                        dblChargeAmount = dblChargeAmount + vobjCommon.GetFreeValuationAmount
                    End If
                   
                    'PB 21/06/2006 EP796
                    dblChargeAmountNoAdmin = dblChargeAmount
                    'PB EP796 End
                    'EP2_2448
                    If Not vobjCommon.IsAdditionalBorrowing And blnSingleComponent = True Then
                        dblChargeAmount = dblChargeAmount + dblRedAdminAmount
                    End If

                    Call xmlSetAttributeValue(xmlPeriod, "CHARGEAMOUNT", CStr(set2DP(dblChargeAmount)))
                    'PB EP796 End
                            
                    'PM 25/05/2006 EPSOM EP628   Start
                    'PB 21/06/2006 EP796 Begin
                    'If intStepNum = 1 Then Call xmlSetAttributeValue(xmlRepay, "MAXCHARGE", CStr(dblChargeAmount))
'                    If intStepNum = 1 Then
'                        'Call xmlSetAttributeValue(xmlRepay, "MAXCHARGE", set2DP(CStr(dblChargeAmount)))
'                        dblgTotalMaxChargeNoAdmin = dblgTotalMaxChargeNoAdmin + dblChargeAmount
'                        'Call xmlSetAttributeValue(xmlCashback, "MAXCHARGENOADMIN", set2DP(CStr(dblChargeAmountNoAdmin)))
''                        dblgTotalMaxCharge = dblgTotalMaxCharge + dblChargeAmountNoAdmin
'                    End If
                    'PB EP796 End
                    'PM 25/05/2006 EPSOM EP628   End
                    'EP2_1931 This should be just the ERC from each LC
                    If Not xmlFee Is Nothing Then
                        dblgTotalMaxChargeNoAdmin = dblgTotalMaxChargeNoAdmin + xmlGetAttributeAsDouble(xmlFee, "REDEMPTIONFEEAMOUNT")
                    End If
                                                       
                                                        
                    '*-add the BASIS attribute
                    If xmlAttributeValueExists(xmlBand, "FEEPERCENTAGE") Then
                        dblPercentage = xmlGetAttributeAsDouble(xmlBand, "FEEPERCENTAGE")
                    Else
                        dblPercentage = 0
                    End If
                    
                    If dblPercentage > 0 Then
                        'PB 21/06/2006 EP697 Start
                        'EP2_139
                        If blnCashback Then
                            strBasis = set2DP(CStr(dblPercentage)) & "% on the outstanding balance plus cashback £" & CStr(cashBack)
                        Else
                            strBasis = set2DP(CStr(dblPercentage)) & "% on the outstanding balance"
                        End If
                    Else
                        strBasis = CStr(xmlGetAttributeAsInteger(xmlBand, "FEEMONTHSINTEREST")) & " months interest"
                    End If
                    Call xmlSetAttributeValue(xmlPeriod, "BASIS", strBasis)
                    
                    '*-add the PERIOD attribute - this is the number of months the
                    '*-charge applies - get the period or get the end date
                    'MAR88 add defensive code
                    
                    If Not xmlFee Is Nothing Then
                        Call AddPeriodAttribute(xmlFee, "REDEMPTIONFEEPERIOD", _
                                "REDEMPTIONFEEPERIODENDDATE", _
                                vobjCommon.ExpectedCompletionDate, "PERIOD", xmlPeriod, strPeriodText)
                        ' PB 21/06/2006 EP796 Begin
                        Call xmlSetAttributeValue(xmlPeriod, "PERIOD", strPeriodText & CStr(xmlGetAttributeAsInteger(xmlPeriod, "PERIOD")))
                        ' PB EP796 End
                    End If
                Next xmlBand
                
                'PB EP529 / MAR1716
                '*-get the maximum charge
                dblMaxCharge = dblMaxCharge + dblRedemptionFeeAmount
                'EP529 / MAR1716 End
                            
                'PM 25/05/2006 EPSOM EP628   Start
                '*-add the DEEDSRELEASEFEE element
                If Not xmlCashback Is Nothing Then
                    'EP2_139
                    Dim temp As String
                    temp = xmlGetAttributeText(xmlCashback, "CASHBACKCLAWBACKAMOUNT")
                    If temp = "" Then
                        Call xmlSetAttributeValue(xmlCashback, "CASHBACKCLAWBACKAMOUNT", CStr(cashBack))
                    End If
                End If
                dblTotalCashBackClawBack = dblTotalCashBackClawBack + cashBack
                'EP2_1931 reinitialise the cashBack
                cashBack = 0

                If blnPortable Then
                    If xmlPortable Is Nothing Then
                        '*-add the PORTABLE element
                        Set xmlPortable = vobjCommon.CreateNewElement("PORTABLE", xmlEarly)
                        bRatetype = CheckIfVariableMortgageInterestRateType(vobjCommon, xmlLoanComponent)
                        If bRatetype = True Then
                            ' Variable
                            Set xmlVariable = vobjCommon.CreateNewElement("VARIABLE", xmlPortable)
                        Else
                            ' Fixed
                            Set xmlVariable = vobjCommon.CreateNewElement("FIXED", xmlPortable)
                        End If
                    End If
                Else
                    If xmlPortable Is Nothing Then
                        '*-add the NOTPORTABLE element
                        Set xmlPortable = vobjCommon.CreateNewElement("NOTPORTABLE", xmlEarly)
                    End If
                End If

                xmlSetAttributeValue xmlRepay, "SECTIONNUMBER", "8"
                
            End If
            
            
        Next xmlLoanComponent
        
               
        'Add total MAXCHARGENOADMIN and CASHBACKCLAWBACKAMOUNT here
        If Not xmlCashback Is Nothing Then
            xmlSetAttributeValue xmlCashback, "MAXCHARGENOADMIN", set2DP(CStr(dblgTotalMaxChargeNoAdmin))
            xmlSetAttributeValue xmlCashback, "CASHBACKCLAWBACKAMOUNT", set2DP(CStr(dblTotalCashBackClawBack))
            xmlSetAttributeValue xmlCashback, "REDADMINAMOUNT", set2DP(CStr(dblRedAdminAmount))
        End If
        If Not xmlMultiCharge Is Nothing Then
            xmlSetAttributeValue xmlMultiCharge, "REDADMINAMOUNT", set2DP(CStr(dblRedAdminAmount))
        End If
        
        ' PB 16/05/2006 EP529 / MAR1731
        If dblMaxCharge > 0 Then
        ' EP529 / MAR1731 End
        
            '*-add the SUMMARY element
            Set xmlSummary = vobjCommon.CreateNewElement("SUMMARY", xmlCharges)
            'SR 13/09/2004 : CORE82
            If vblnIsLifetimeDocument Then
                xmlSetAttributeValue xmlSummary, "MAXCHARGE", set2DP(CStr(dblTotalRedemptionFeeAmount))
                 'CORE82 *-add the REPAYMENTFEE attribute if > 0
                If lngRepaymentFee > 0 Then
                    Set xmlPlusFee = vobjCommon.CreateNewElement("PLUSFEE", xmlSummary)
                    Call xmlSetAttributeValue(xmlPlusFee, "REPAYMENTFEE", CStr(lngRepaymentFee))
                End If
               
               If dblTotalRedemptionFeeAmount > 0 Then
                    Set xmlTemp = vobjCommon.CreateNewElement("EARLYREPAYMENTCHARGESAPPLY", xmlSummary)
                    xmlSetAttributeValue xmlTemp, "EARLYREPAYMENTFEE", lngRepaymentFee
                    vobjCommon.CreateNewElement "SINGLECOMPONENT", xmlTemp
                End If
            Else 'SR 13/09/2004 : CORE82 - End
                If dblMaxCharge > 0 Then
                    'CORE82 NOT ON DTD*-add the EARLYREPAYMENTCHARGESAPPLY element mcf>0
        '                Set xmlRepaymentCharges = vobjCommon.CreateNewElement("EARLYREPAYMENTCHARGESAPPLY", xmlSummary)
                    Call AddRepaymentChargesInfo(vobjCommon, xmlSummary)
                    'CORE82 *-add the MAXCHARGENOINTRATE element
                    Set xmlRepaymentCharges = vobjCommon.CreateNewElement("MAXCHARGENOINTRATE", xmlSummary)
                    Call xmlSetAttributeValue(xmlRepaymentCharges, "MAXCHARGE", set2DP(CStr(dblMaxCharge)))
                    'CORE82 *-add the REPAYMENTFEE attribute if > 0
                    If lngRepaymentFee > 0 Then
                        Set xmlPlusFee = vobjCommon.CreateNewElement("PLUSFEE", xmlRepaymentCharges)
                        Call xmlSetAttributeValue(xmlPlusFee, "REPAYMENTFEE", CStr(lngRepaymentFee))
                    End If
                End If
             End If
        End If
    ' PB 16/05/2006 EP529 / MAR1731
    End If
    ' EP529 / MAR1731 End
        
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE element to the PORTABLE/NOTPORTABLE element
        'DS EP2_1931
        If Not xmlPortable Is Nothing Then
            Call xmlSetAttributeValue(xmlPortable, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
        End If
            '*-and to this section
        'DS EP2_1931
        If Not vxmlNode Is Nothing Then
            Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
        End If
    End If

    Set xmlList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlFeesList = Nothing
    Set xmlFee = Nothing
    Set xmlEarly = Nothing
    Set xmlCharges = Nothing
    Set xmlComponent = Nothing
    Set xmlRepay = Nothing
    Set xmlBandList = Nothing
    Set xmlBand = Nothing
    Set xmlPeriod = Nothing
    Set xmlProduct = Nothing
    Set xmlSummary = Nothing
    Set xmlClient = Nothing
    Set xmlRepaymentCharges = Nothing
    Set xmlPlusFee = Nothing
    Set xmlTemp = Nothing
    Set xmlProductLanguage = Nothing ' PB 05/06/2006 EP651/MAR1590
    'EP2_139
    Set xmlCashback = Nothing
    Set xmlMortgageIncentive = Nothing
    Set xmlMultiCharge = Nothing
    Set xmlMultiChargeTemp = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlFeesList = Nothing
    Set xmlFee = Nothing
    Set xmlEarly = Nothing
    Set xmlCharges = Nothing
    Set xmlComponent = Nothing
    Set xmlRepay = Nothing
    Set xmlBandList = Nothing
    Set xmlBand = Nothing
    Set xmlPeriod = Nothing
    Set xmlProduct = Nothing
    Set xmlSummary = Nothing
    Set xmlClient = Nothing
    Set xmlRepaymentCharges = Nothing
    Set xmlPlusFee = Nothing
    Set xmlTemp = Nothing
    Set xmlProductLanguage = Nothing ' PB 05/06/2006 EP651/MAR1590
    'EP2_139
    Set xmlCashback = Nothing
    Set xmlMortgageIncentive = Nothing
    Set xmlMultiCharge = Nothing
    Set xmlMultiChargeTemp = Nothing

    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Private Sub AddRepaymentChargesInfo(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddRepaymentChargesInfo"
    Dim xmlComponentNode As IXMLDOMNode
    'CORE82 Rework all of this - Elements changed.

    On Error GoTo ErrHandler
    
    'set up a dummy node for AddComponentsTypeElement call
    Set xmlComponentNode = vxmlNode.cloneNode(False)
    
    '*-add the SINGLECOMPONENTERC or MULTICOMPONENTERC element
    Set xmlComponentNode = vobjCommon.AddComponentsTypeElement(xmlComponentNode)
    
    'MAR88 - BC 05 Oct
    If vobjCommon.MortgageCompletionFee > 0 Then
        If xmlComponentNode.baseName = "SINGLECOMPONENT" Then
            '*-add the SINGLECOMPONENTERC element
            Set xmlComponentNode = vobjCommon.CreateNewElement("SINGLECOMPONENTERC", vxmlNode)
    
    '        Set xmlComponentNode = CreateNewElement("SINGLECOMPONENTERC", vxmlNode)
        Else
            '*-add the MULTICOMPONENTERC element
            Set xmlComponentNode = vobjCommon.CreateNewElement("MULTICOMPONENTERC", vxmlNode)
    '        Set xmlComponentNode = CreateNewElement("MULTICOMPONENTERC", vxmlNode)
        End If
        
        '*-add the EARLYREPAYMENTFEE attribute
        Call xmlSetAttributeValue(xmlComponentNode, "EARLYREPAYMENTFEE", vobjCommon.MortgageCompletionFee)
    End If
'    Call vobjCommon.AddComponentsTypeElement(vxmlNode)
    Set xmlComponentNode = Nothing
Exit Sub
ErrHandler:
    Set xmlComponentNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

' EP2_139 Offer and KFI changes.
'********************************************************************************
'** Function:       BuildAdditionalFeaturesSection
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the Additional Features
'**                 section.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section element to set the attributes on.
'**                 vblnIsOfferDocument - whether an offer document is being
'**                 built or not.
'** Returns:        N/A
'** Errors:         None Expected
'***** MAR88 BC - Section revampled to cater for the requirements of Project MARS
'********************************************************************************
Public Sub BuildAdditionalFeaturesSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean, _
        ByVal vblnIsLifetimeMortgage As Boolean)
    Const cstrFunctionName As String = "BuildAdditionalFeaturesSection"
    
    Dim xmlIncentivesList As IXMLDOMNodeList
    Dim xmlLoanComponentList As IXMLDOMNodeList 'MAR1535 BC
    Dim xmlMultiComponent As IXMLDOMNode
    Dim xmlItem As IXMLDOMNode
    Dim xmlIncentive As IXMLDOMNode
    Dim xmlIncent As IXMLDOMNode
    Dim intValue As Integer
    Dim xmlValue As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    Dim xmlDrawDown As IXMLDOMNode, dblDrawDownAmount As Double 'BBG1643
    Dim bCalcDatePlus2Years As Boolean 'MAR88 BC
    Dim bAdditionalFeatures As Boolean 'MAR88 BC
    Dim blnFixedRateElement As Boolean 'BC MAR1535
    'TK 04/02/2005 E2EM00001894 End
    'E2EM00001894
    Dim intProductNum As Integer
            
    Dim xmlSingleComponent As IXMLDOMNode
    Dim xmlAllowed As IXMLDOMNode
    Dim xmlIRNode As IXMLDOMNode
    Dim xmlIRList As IXMLDOMNodeList
    Dim xmlPart As IXMLDOMNode
    Dim xmlIncentives  As IXMLDOMNode
    
    'EP2_1861 Begin
    Dim xmlCostNode As IXMLDOMNode
    Dim dblRefundAmount As Double
    'EP2_1861 End
    
    On Error GoTo ErrHandler
    
    '*-get the incentives list
     Set xmlIncentivesList = vobjCommon.Data.selectNodes("//MORTGAGEINCENTIVE")
    
    bCalcDatePlus2Years = True 'MAR88 BC
    bAdditionalFeatures = False 'MAR88 BC
    blnFixedRateElement = False 'BC MAR1535
    intProductNum = 0 'PB EP529 / MAR1720 12/05/2006
    
    'SR 13/10/2004 : BBG1592
    
    If vblnIsLifetimeMortgage Then
        'SR 19/10/2004 : BBG1662 - removed code that is not required for BBG. Check V 25 for the code removed
        'SR 19/10/2004 : BBG1662
        'Get the DrawDown amount from table MortgageSubQuote
        Set xmlTemp = vobjCommon.Data.selectSingleNode(gcstrMORTGAGESUBQUOTE_PATH)

'IK 01/12/2004 E2EM00003126
'        dblDrawDownAmount = xmlGetAttributeAsDouble(xmlTemp, "DRAWDOWN")
'IK 01/12/2004 E2EM00003126 ends
        
        Set xmlTemp = vobjCommon.CreateNewElement("ADDITIONALSECUREDBORROWING", vxmlNode)
'IK 01/12/2004 E2EM00003126
'       If dblDrawDownAmount > 0 Then
'           Set xmlDrawDown = vobjCommon.CreateNewElement("DRAWDOWN", xmlTemp)
'           xmlSetAttributeValue xmlDrawDown, "DRAWDOWNAMOUNT", set2DP(CStr(dblDrawDownAmount))
        If vobjCommon.HasDrawDown Then
            Set xmlDrawDown = vobjCommon.CreateNewElement("DRAWDOWN", xmlTemp)
            xmlSetAttributeValue xmlDrawDown, "DRAWDOWNAMOUNT", CStr(vobjCommon.DrawDownAmount)
'IK 01/12/2004 E2EM00003126 ends
        Else
            Set xmlDrawDown = vobjCommon.CreateNewElement("NODRAWDOWN", xmlTemp)
        End If
        'SR 19/10/2004 : BBG1662 - End
        
        For Each xmlIncentive In xmlIncentivesList
            '*-add the INCENTIVES element
            Set xmlIncent = vobjCommon.CreateNewElement("INCENTIVES", vxmlNode)
            
            intValue = xmlGetAttributeAsInteger(xmlIncentive, "INCENTIVEBENEFITTYPE")
            If intValue = 10 Then
                '*-add the MONETARYVALUE element
                Set xmlValue = vobjCommon.CreateNewElement("MONETARYVALUE", xmlIncent)
                '*-add the INCENTIVEAMOUNT attribute
                Call xmlCopyAttribute(xmlIncentive, xmlValue, "INCENTIVEAMOUNT")
            End If
        Next xmlIncentive
    Else  'SR 17/10/2004 : BBG1643
    
        Set xmlLoanComponentList = vobjCommon.LoanComponents 'MAR1535 BC
        
        If xmlLoanComponentList.length = 0 Then 'MAR1535 BC
            Exit Sub
        End If
        
        If vobjCommon.IsFlexible Then
            bAdditionalFeatures = True

            Set xmlTemp = vobjCommon.CreateNewElement("UNDERPAYMENTS", vxmlNode)
            Set xmlSingleComponent = vobjCommon.CreateNewElement("SINGLECOMPONENT", xmlTemp)
            Set xmlAllowed = vobjCommon.CreateNewElement("ALLOWED", xmlSingleComponent)
       
            Set xmlTemp = vobjCommon.CreateNewElement("PAYMENTHOLIDAYS", vxmlNode)
            Set xmlSingleComponent = vobjCommon.CreateNewElement("SINGLECOMPONENT", xmlTemp)
            Set xmlAllowed = vobjCommon.CreateNewElement("ALLOWED", xmlSingleComponent)
       
            Set xmlTemp = vobjCommon.CreateNewElement("BORROWBACK", vxmlNode)
            Set xmlSingleComponent = vobjCommon.CreateNewElement("SINGLECOMPONENT", xmlTemp)
            
            'PB EP2_1861 07/03/2007 Start
            If vobjCommon.IsRefundOfValuation Then
                For Each xmlCostNode In vobjCommon.Data.selectNodes("//MORTGAGEONEOFFCOST")
                    If CheckForValidationType(xmlGetAttributeText(xmlCostNode, "MORTGAGEONEOFFCOSTTYPE"), "VAL") Then
                        dblRefundAmount = dblRefundAmount + xmlGetAttributeAsDouble(xmlCostNode, "REFUNDAMOUNT")
                    End If
                    If CheckForValidationType(xmlGetAttributeText(xmlCostNode, "MORTGAGEONEOFFCOSTTYPE"), "TPV") Or _
                       CheckForValidationType(xmlGetAttributeText(xmlCostNode, "MORTGAGEONEOFFCOSTTYPE"), "TPVA") Then
                        dblRefundAmount = dblRefundAmount + xmlGetAttributeAsDouble(xmlCostNode, "REFUNDAMOUNT")
                    End If
                Next
                If dblRefundAmount > 0 Then
                    Set xmlTemp = vobjCommon.CreateNewElement("REFUNDOFVALUATION", vxmlNode)
                    xmlSetAttributeValue xmlTemp, "REFUNDAMOUNT", dblRefundAmount
                End If
            End If
            'PB EP2_1861 07/03/2007 End
            
            Set xmlTemp = vobjCommon.CreateNewElement("ADDITIONALSECUREDBORROWING", vxmlNode)

            If vobjCommon.HasDrawDown Then
                Dim totalLCAmount As Long
                Dim totalMSQAmount As Long
                Dim numOfLC
                'Array of the fee amounts that are added to loan
                Dim unDrawnCreditLimit() As Long

                'ADDITIONALBORROWING Totals from LoanComponent
                If xmlLoanComponentList.length > 1 Then
                    'Array of the unDrawnCreditLimit amounts to allocate across loan components
                    'It is allocated bottom up and can only be allocated to each loan component
                    'up to the totalloancomponentamount
                    ReDim unDrawnCreditLimit(xmlLoanComponentList.length - 1)
                    Dim i As Integer
                    Dim undrawnCredit As Long
                    undrawnCredit = vobjCommon.DrawDownAmount
                    i = xmlLoanComponentList.length - 1
                    Do While i >= 0
                      Dim LoanAmount As Long
                      LoanAmount = xmlGetAttributeText(xmlLoanComponentList.Item(i), "TOTALLOANCOMPONENTAMOUNT")
                      If i > 0 Then
                        If undrawnCredit > LoanAmount Then
                            'can only allocate part of the balance
                            unDrawnCreditLimit(i) = LoanAmount
                            undrawnCredit = undrawnCredit - LoanAmount
                        Else
                            'can allocate the balance
                            unDrawnCreditLimit(i) = undrawnCredit
                            i = 0
                        End If
                      Else
                        unDrawnCreditLimit(i) = undrawnCredit
                      End If
                      i = i - 1
                    Loop
        
                    Set xmlTemp = vobjCommon.CreateNewElement("ADDITIONALBORROWING", vxmlNode)
                    Set xmlMultiComponent = vobjCommon.CreateNewElement("MULTICOMPONENT", xmlTemp)
                    numOfLC = 0
                    totalMSQAmount = vobjCommon.TotalLoanAmount + vobjCommon.DrawDownAmount
                    For Each xmlItem In xmlLoanComponentList
                        'GetTypeOfMortgage()
                        '*-get all the interest rates for this loan component
                        Set xmlIRNode = xmlItem.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER = '1']/BASERATESETDATA/BASERATE")
                        'EP2_1742
                        totalLCAmount = xmlGetAttributeAsLong(xmlItem, "TOTALLOANCOMPONENTAMOUNT") + unDrawnCreditLimit(numOfLC)
                        If (numOfLC + 1) = xmlLoanComponentList.length Then
                            Set xmlPart = vobjCommon.CreateNewElement("FINALPART", xmlMultiComponent)
                            xmlSetAttributeValue xmlPart, "FINALPART", CStr(numOfLC + 1)
                            xmlSetAttributeValue xmlPart, "UNDRAWNCREDITLIMIT", CStr(set2DP(unDrawnCreditLimit(numOfLC)))
                            xmlSetAttributeValue xmlPart, "CURRENTINTRATE", set2DP(xmlGetAttributeText(xmlIRNode, "BASEINTERESTRATE"))
                         ElseIf numOfLC > 0 Then
                            Set xmlPart = vobjCommon.CreateNewElement("PART", xmlMultiComponent)
                            xmlSetAttributeValue xmlPart, "PART", CStr(numOfLC + 1)
                            xmlSetAttributeValue xmlPart, "UNDRAWNCREDITLIMIT", CStr(set2DP(unDrawnCreditLimit(numOfLC)))
                            xmlSetAttributeValue xmlPart, "CURRENTINTRATE", set2DP(xmlGetAttributeText(xmlIRNode, "BASEINTERESTRATE"))
                         ElseIf numOfLC = 0 Then
                            xmlSetAttributeValue xmlMultiComponent, "UNDRAWNCREDITLIMIT", CStr(set2DP(unDrawnCreditLimit(numOfLC)))
                            xmlSetAttributeValue xmlMultiComponent, "CURRENTINTRATE", set2DP(xmlGetAttributeText(xmlIRNode, "BASEINTERESTRATE"))
                            xmlSetAttributeValue xmlMultiComponent, "TOTALDEBT", CStr(set2DP(totalMSQAmount))
                            xmlSetAttributeValue xmlMultiComponent, "TOTALMONTHLYPAYMENT", CStr(set2DP(vobjCommon.TotalNetMonthlyCost))
                        End If
                      
                        numOfLC = numOfLC + 1

                    Next xmlItem
                Else
                    Set xmlIRNode = xmlLoanComponentList.Item(0).selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER = '1']/BASERATESETDATA/BASERATE")

                    Set xmlTemp = vobjCommon.CreateNewElement("ADDITIONALBORROWING", vxmlNode)
                    Set xmlSingleComponent = vobjCommon.CreateNewElement("SINGLECOMPONENT", xmlTemp)
                    'EP2_1742 should be xmlSingleComponent
                    xmlSetAttributeValue xmlSingleComponent, "UNDRAWNCREDITLIMIT", CStr(set2DP(vobjCommon.DrawDownAmount))
                    xmlSetAttributeValue xmlSingleComponent, "CURRENTINTRATE", set2DP(xmlGetAttributeText(xmlIRNode, "BASEINTERESTRATE"))
                    'EP2_1742
                    totalLCAmount = xmlGetAttributeAsLong(xmlLoanComponentList.Item(0), "TOTALLOANCOMPONENTAMOUNT") + vobjCommon.DrawDownAmount
                    xmlSetAttributeValue xmlSingleComponent, "TOTALDEBT", CStr(set2DP(totalLCAmount))
                    xmlSetAttributeValue xmlSingleComponent, "TOTALMONTHLYPAYMENT", CStr(set2DP(vobjCommon.TotalNetMonthlyCost))
                End If
                'TOTALADDITIONALBORROWING Totals from MortgageSubQuote
                Set xmlTemp = vobjCommon.CreateNewElement("TOTALADDITIONALBORROWING", vxmlNode)
                xmlSetAttributeValue xmlTemp, "TOTALDEBT", CStr(set2DP(totalMSQAmount))
                xmlSetAttributeValue xmlTemp, "TOTALMONTHLYPAYMENT", CStr(set2DP(vobjCommon.TotalNetMonthlyCost))
            End If
       
        End If
        
        If vobjCommon.IsIncentive Then
           Set xmlIncentives = vobjCommon.CreateNewElement("INCENTIVES", vxmlNode)
           bAdditionalFeatures = True

            If vobjCommon.IsCashback Then
                Dim totalCashBack As Long
                Dim lcCashBack As Long
                Dim seqLCNum As Integer
                Dim xmlMortgageIncentive As IXMLDOMNode
                
                For Each xmlItem In xmlLoanComponentList
                    lcCashBack = 0
                    seqLCNum = xmlGetAttributeAsInteger(xmlItem, "LOANCOMPONENTSEQUENCENUMBER")
                    
                    Set xmlMortgageIncentive = xmlItem.selectSingleNode("//LOANCOMPONENT/MORTGAGEINCENTIVE[@LOANCOMPONENTSEQUENCENUMBER=" & seqLCNum & "][INCLUSIVEINCENTIVE]")
                    If Not xmlMortgageIncentive Is Nothing Then
                        lcCashBack = xmlGetAttributeAsDouble(xmlMortgageIncentive, "INCENTIVEAMOUNT")
                    End If
                    totalCashBack = totalCashBack + lcCashBack
                Next xmlItem
            
                'EP2_1931 only show if > 0
                If totalCashBack > 0 Then
                    Set xmlTemp = vobjCommon.CreateNewElement("CASHBACK", xmlIncentives)
                    xmlSetAttributeValue xmlTemp, "CASHBACKAMOUNT", totalCashBack
                    'PB EP2_1629 05/03/2007
                    'If vobjCommon.IsTransferOfEquity Or vobjCommon.IsProductSwitch Then
                    '    xmlSetAttributeValue xmlTemp, "SECTIONNUMBER", "9"
                    'Else
                        xmlSetAttributeValue xmlTemp, "SECTIONNUMBER", "10"
                    'End If
                End If
            End If
            If vobjCommon.IsFreeLegal Then
                Set xmlTemp = vobjCommon.CreateNewElement("FREELEGAL", xmlIncentives)
            End If
            If vobjCommon.IsFreevaluation Then
                Set xmlTemp = vobjCommon.CreateNewElement("FREEVALUATION", xmlIncentives)
                xmlSetAttributeValue xmlTemp, "VALFEEAMOUNT", CStr(vobjCommon.GetFreeValuationAmount)
                'PB EP2_1629 05/03/2007
                'If vobjCommon.IsTransferOfEquity Or vobjCommon.IsProductSwitch Then
                '    xmlSetAttributeValue xmlTemp, "SECTIONNUMBER", "9"
                'Else
                    xmlSetAttributeValue xmlTemp, "SECTIONNUMBER", "10"
                'End If
            End If
        End If
        
        If bAdditionalFeatures = False Then
            Set xmlTemp = vobjCommon.CreateNewElement("NONE", vxmlNode)
        End If
        
    End If
   
    Set xmlSingleComponent = Nothing
    Set xmlAllowed = Nothing
    Set xmlIRNode = Nothing
    Set xmlIRList = Nothing
    Set xmlPart = Nothing
       
    Set xmlIncentivesList = Nothing
    Set xmlIncentive = Nothing
    Set xmlIncent = Nothing
    
    Set xmlValue = Nothing
    Set xmlTemp = Nothing
    Set xmlDrawDown = Nothing
    
Exit Sub
ErrHandler:
    Set xmlSingleComponent = Nothing
    Set xmlAllowed = Nothing
    Set xmlIRNode = Nothing
    Set xmlIRList = Nothing
    Set xmlPart = Nothing
       
    Set xmlIncentivesList = Nothing
    Set xmlIncentive = Nothing
    Set xmlIncent = Nothing
    
    Set xmlValue = Nothing
    Set xmlTemp = Nothing
    Set xmlDrawDown = Nothing
    
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'********************************************************************************
'** Function:       AddFeatureInformation
'** Created by:     Andy Maggs
'** Date:           20/05/2004
'** Description:    Adds the feature information to the specified node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - The node to add the information to.
'**                 vblnIsOfferDocument - whether the offer or illustration
'**                 document is being created.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub AddFeatureInformation(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, _
        ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "AddFeatureInformation"
    
    On Error GoTo ErrHandler

    '*-add the ATANYTIME element
    Call vobjCommon.CreateNewElement("ATANYTIME", vxmlNode)
    
    '*-add the OVERPAYMENTSCONDITION element
    Call vobjCommon.CreateNewElement("OVERPAYMENTSCONDITION", vxmlNode)

    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
    End If

Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddOverpayFeatureInformation
'** Created by:     Andy Maggs
'** Date:           20/05/2004
'** Description:    Adds the feature information to the specified node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - The node to add the information to.
'**                 vblnIsOfferDocument - whether the offer or illustration
'**                 document is being created.
'**                 vintRestPeriod -
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub AddOverpayFeatureInformation(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean, _
        ByVal vblnIsSingleComponent As Boolean, ByVal vintRestPeriod As Integer)
    Const cstrFunctionName As String = "AddOverpayFeatureInformation"
    Dim blnIsSingleOffer As Boolean
    Dim xmlNoRestrict As IXMLDOMNode
    Dim xmlLump As IXMLDOMNode
    Dim xmlRecalc As IXMLDOMNode
    Dim strValue As String
    Dim strFrequency As String
    
    On Error GoTo ErrHandler
    
    blnIsSingleOffer = vblnIsOfferDocument And vblnIsSingleComponent
    
    '*-add the NORESTRICTIONS element
    Set xmlNoRestrict = vobjCommon.CreateNewElement("NORESTRICTIONS", vxmlNode)
    If blnIsSingleOffer Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlNoRestrict)
        '*-add the LUMPSUM element
        Set xmlLump = vobjCommon.CreateNewElement("LUMPSUM", vxmlNode)
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlLump)
    End If
    
    If vintRestPeriod = 2 Then
        '*-add the RECALCULATIONSIMMEDIATELY element
        Call vobjCommon.CreateNewElement("RECALCULATIONSIMMEDIATELY", vxmlNode)
    Else
        '*-add the RECALCULATIONSLATER element
        Set xmlRecalc = vobjCommon.CreateNewElement("RECALCULATIONSLATER", vxmlNode)
        '*-add the RECALCULATIONDATE attribute
        Select Case vintRestPeriod
            Case 0
                strValue = "at next year end"
                strFrequency = "annually"
            Case 1
                strValue = "at next month end"
                strFrequency = "monthly"
            Case 3
                strValue = "at next half-year"
                strFrequency = "half-yearly"
            Case 4
                strValue = "at next quarter end"
                strFrequency = "quarterly"
        End Select
        Call xmlSetAttributeValue(xmlRecalc, "RECALCULATIONDATE", strValue)
        
        If vblnIsOfferDocument And Not vblnIsSingleComponent Then
            '*-add the FREQUENCY attribute
            Call xmlSetAttributeValue(xmlRecalc, "FREQUENCY", strFrequency)
        End If
    End If
    
    If vblnIsSingleComponent Then
        '*-add the CHARGESAPPLY element
        Call vobjCommon.CreateNewElement("CHARGESAPPLY", vxmlNode)
    End If

    Set xmlNoRestrict = Nothing
    Set xmlLump = Nothing
    Set xmlRecalc = Nothing
Exit Sub
ErrHandler:
    Set xmlNoRestrict = Nothing
    Set xmlLump = Nothing
    Set xmlRecalc = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Public Sub AddLoanComponentInterestRateData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode, ByVal vxmlComponentNode As IXMLDOMNode, _
        ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "AddLoanComponentInterestRateData"
    Dim xmlRate As IXMLDOMNode
    Dim intTotalNumPayments As Integer
    Dim xmlList As IXMLDOMNodeList
    Dim intIndex As Integer
    Dim intTotalMonths As Integer
    Dim xmlTemp As IXMLDOMNode
    
    Dim xmlLCPScheduleList As IXMLDOMNodeList, xmlLCPSchedule As IXMLDOMNode, xmlNextLCPSchedule As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strPaymentType As String
    Dim dblCurrentMonthlyCost As Double, dblInterestRate As Double
    Dim dtCurrentDate As Date, dtNextDate As Date, dtTemp As Date
    Dim intStepNo As Integer, intTotalSteps As Integer
    Dim intNoOfPayments As Integer, intTotalNoOfPayments As Integer, intNoOfPaymentsConsidered As Integer
    Dim blnFirstMonthlyPayment As Boolean
    Dim bCalcDatePlus2Years As Boolean 'MAR88 BC
    
    blnFirstMonthlyPayment = True
    bCalcDatePlus2Years = False 'MAR88 BC
    
    On Error GoTo ErrHandler

    '*-get the first rate
    Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(vxmlLoanComponent)
    If Not xmlRate Is Nothing Then
        'CORE82 Get NUMBEROFPAYMENTS first in case it returns -1 i.e. only one rate
        '*-add the mandatory NUMBEROFPAYMENTS attribute
'        intTotalNumPayments = AddNumberOfPaymentsAttribute(vobjCommon, xmlRate, vxmlComponentNode) SR 12/09/2004 : CORE82
        
        If intTotalNumPayments = -1 Then
            '*-Modify the mandatory NUMBEROFPAYMENTS attribute
            '*-this is the last rate and applies to the end of the loan
            '*-component, so calculate the actual number of payments that would
            '*-be required and change NUMBEROFPAYMENTS from -1 to the correct number
            intTotalMonths = xmlGetAttributeAsInteger(vxmlLoanComponent, "TERMINYEARS") * 12
            intTotalMonths = intTotalMonths + xmlGetAttributeAsInteger(vxmlLoanComponent, "TERMINMONTHS")
            Call xmlSetAttributeValue(vxmlComponentNode, "NUMBEROFPAYMENTS", CStr(intTotalMonths))
        
        End If
        
        '*-add the mandatory RATE attribute
        Call AddRateAttribute(vobjCommon, vxmlLoanComponent, xmlRate, _
                vxmlComponentNode, "RATE")
        '*-add the mandatory MORTGAGECOMMENCEMENTDATE attribute
        'Call AddMortgageCommencementDateAttribute(vobjCommon, vxmlComponentNode) 'SR 22/09/2004
        AddCalculationDateAttribute vobjCommon, vxmlComponentNode, bCalcDatePlus2Years 'MAR88 BC
        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, vxmlComponentNode)
        End If
        
        '*-add the appropriate Fees and/or Premiums element
        Set xmlTemp = AddFeesPremiumsElement(vobjCommon, vxmlComponentNode)
        If vblnIsOfferDocument Then
            If xmlTemp.baseName <> "NOFEESORPREMIUMS" Then
                '*-add the MORTGAGETYPE attribute
                Call AddMortgageTypeAttribute(vobjCommon, xmlTemp)
            End If
        End If
'** SR 30/09/2004 : CORE82 - Removed commented code. See V 68 for commented code
    End If
    
    'SR 30/09/2004 : CORE82 - Build the xml based on LoanComponentPaymentSchedule
    intTotalNoOfPayments = (vobjCommon.TermInYears * 12) + (vobjCommon.TermInMonths)
    Set xmlLCPScheduleList = vxmlLoanComponent.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
    intTotalSteps = xmlLCPScheduleList.length
    For intStepNo = 0 To intTotalSteps - 1 Step 1
        Set xmlLCPSchedule = xmlLCPScheduleList.Item(intStepNo)
        If intStepNo <> intTotalSteps - 1 Then
            Set xmlNextLCPSchedule = xmlLCPScheduleList.Item(intStepNo + 1)
            dtNextDate = xmlGetAttributeAsDate(xmlNextLCPSchedule, "STARTDATE")
        End If
        
        strPaymentType = xmlGetAttributeText(xmlLCPSchedule, "PAYMENTTYPE")
        dblCurrentMonthlyCost = xmlGetAttributeAsDouble(xmlLCPSchedule, "MONTHLYCOST")
        dblInterestRate = xmlGetAttributeAsDouble(xmlLCPSchedule, "INTERESTRATE")
        dtCurrentDate = xmlGetAttributeAsDate(xmlLCPSchedule, "STARTDATE")
        
        Select Case UCase$(strPaymentType)
            Case "ACCRUEDINTEREST" 'If exists, it will always be the first record in PaymentSchedule
                Set xmlNode = vobjCommon.CreateNewElement("ACCRUED", vxmlComponentNode)
                xmlSetAttributeValue xmlNode, "ACCRUEDINTEREST", set2DP(CStr(dblCurrentMonthlyCost))
                xmlSetAttributeValue xmlNode, "RATE", set2DP(CStr(dblInterestRate))
            Case "MONTHLYPAYMENT"
                ' PB 24/05/2006 EP603/MAR1777
                'intNoOfPayments = IIf(intStepNo < intTotalSteps - 1, DateDiff("M", dtCurrentDate, dtNextDate), _
                '                                                intTotalNoOfPayments - intNoOfPaymentsConsidered)
                intNoOfPayments = IIf(intStepNo < intTotalSteps - 1, MonthDiff(dtCurrentDate, dtNextDate), _
                                                                intTotalNoOfPayments - intNoOfPaymentsConsidered)   'MAR1777 GHun
                ' EP603/MAR1777 End
                
                If blnFirstMonthlyPayment Then
                    xmlSetAttributeValue vxmlComponentNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                    xmlSetAttributeValue vxmlComponentNode, "RATE", set2DP(CStr(dblInterestRate))
                    xmlSetAttributeValue vxmlComponentNode, "NUMBEROFPAYMENTS", CStr(intNoOfPayments)
                    blnFirstMonthlyPayment = False
                Else
                    Set xmlNode = vobjCommon.CreateNewElement("FOLLOWEDBY", vxmlComponentNode)
                    xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                    xmlSetAttributeValue xmlNode, "RATE", set2DP(CStr(dblInterestRate))
                    xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", CStr(intNoOfPayments)
                End If
                intNoOfPaymentsConsidered = intNoOfPaymentsConsidered + intNoOfPayments
            Case "FINALPAYMENT"
                Set xmlNode = vobjCommon.CreateNewElement("REMAINDEROFTERM", vxmlComponentNode)
                xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                xmlSetAttributeValue xmlNode, "RATE", set2DP(CStr(dblInterestRate))
        End Select
    Next intStepNo
    'SR 30/09/2004 : CORE82 - End
            
    Set xmlRate = Nothing
    Set xmlList = Nothing
    Set xmlTemp = Nothing
    Set xmlLCPScheduleList = Nothing
    Set xmlLCPSchedule = Nothing
    Set xmlNextLCPSchedule = Nothing
    Set xmlNode = Nothing

Exit Sub
ErrHandler:
    Set xmlRate = Nothing
    Set xmlList = Nothing
    Set xmlTemp = Nothing
    Set xmlLCPScheduleList = Nothing
    Set xmlLCPSchedule = Nothing
    Set xmlNextLCPSchedule = Nothing
    Set xmlNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'BBG1589 New Function
Public Sub AddPostMDayLoanComponentInterestRateData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode, ByVal vxmlComponentNode As IXMLDOMNode, _
        ByVal vblnIsOfferDocument As Boolean, ByVal vblnInterestOnly As Boolean)
    Const cstrFunctionName As String = "AddLoanComponentInterestRateData"
    Dim xmlRate As IXMLDOMNode
    Dim intTotalNumPayments As Integer
    Dim xmlList As IXMLDOMNodeList
    Dim intIndex As Integer
    Dim intTotalMonths As Integer
    Dim xmlTemp As IXMLDOMNode
    Dim xmlLCPScheduleList As IXMLDOMNodeList, xmlLCPSchedule As IXMLDOMNode, xmlNextLCPSchedule As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strPaymentType As String
    Dim dblCurrentMonthlyCost As Double, dblInterestRate As Double
    Dim dblFinalPayment As Double
    Dim dtCurrentDate As Date, dtNextDate As Date, dtTemp As Date, dtEndDate As Date
    Dim intStepNo As Integer, intTotalSteps As Integer
    Dim intNoOfPayments As Integer, intTotalNoOfPayments As Integer, intNoOfPaymentsConsidered As Integer
    Dim blnFirstMonthlyPayment As Boolean
    Dim int2ndEntryNoOfPayments As Integer
    Dim dblAccruedCost As Double
    Dim dblFirstMonthCost As Double
    Dim xmlIntRateTypeList As IXMLDOMNodeList
    Dim xmlIntRateType As IXMLDOMNode
    Dim xmlAssumes As IXMLDOMNode 'BC MAR907
    Dim intNumPeriods As Integer, intTotNumPeriods As Integer
    Dim intToEndDatePeriods As Integer, intTotNumRates As Integer, intRateType As Integer
    Dim audtInterestRateData() As InterestRateDataType
    Dim strIntRateType As String
    Dim xmlRateNode As IXMLDOMNode
    Dim intNoOfPaymentsToCompare As Integer
    Dim intStepNo2 As Integer
    Dim intCumulativePeriod As Integer
    Dim intFirstNoOfPayments As Integer
    Dim bCalcDatePlus2Years As Boolean 'MAR88 BC
    Dim intMonthAdjustment As Integer 'PB EP529 / MAR1720
    Dim strAssumedStartDate As String 'PB EP529 / MAR1731
    
    On Error GoTo ErrHandler
    blnFirstMonthlyPayment = True
    bCalcDatePlus2Years = False 'MAR88 BC

    '*-get the first rate
    Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(vxmlLoanComponent)
    If Not xmlRate Is Nothing Then
        'Get NUMBEROFPAYMENTS first in case it returns -1 i.e. only one rate
        '*-add the mandatory NUMBEROFPAYMENTS attribute
        
        If intTotalNumPayments = -1 Then
            '*-Modify the mandatory NUMBEROFPAYMENTS attribute
            '*-this is the last rate and applies to the end of the loan
            '*-component, so calculate the actual number of payments that would
            '*-be required and change NUMBEROFPAYMENTS from -1 to the correct number
            intTotalMonths = xmlGetAttributeAsInteger(vxmlLoanComponent, "TERMINYEARS") * 12
            intTotalMonths = intTotalMonths + xmlGetAttributeAsInteger(vxmlLoanComponent, "TERMINMONTHS")
            Call xmlSetAttributeValue(vxmlComponentNode, "NUMBEROFPAYMENTS", CStr(intTotalMonths))
        End If
        
        '*-add the mandatory RATE attribute
        Call AddRateAttribute(vobjCommon, vxmlLoanComponent, xmlRate, _
                vxmlComponentNode, "RATE")
                
        'BC MAR907 Begin
        If vobjCommon.blnIsProductSwitch Then
        '*-add the PRODUCTSWITCH element
            Set xmlAssumes = vobjCommon.CreateNewElement("PRODUCTSWITCH", vxmlComponentNode)
        Else
            '*-add the NONPRODUCTSWITCH element
            Set xmlAssumes = vobjCommon.CreateNewElement("NONPRODUCTSWITCH", vxmlComponentNode)
        End If
        'BC MAR907 End

        'PB 16/05/2006 EP529 / MAR1731
        'PB 24/05/2006 EP603/MAR1777
        'strAssumedStartDate = Format$(Now, "dd/mm/yyyy")
        'strAssumedStartDate = "01" & Mid(strAssumedStartDate, 3, 8)
        'strAssumedStartDate = Format$(DateAdd("m", 1, strAssumedStartDate), "dd/mm/yyyy")
        strAssumedStartDate = Format$(CalcExpectedCompletionDate(Date), "dd/mm/yyyy")
        ' EP603/MAR1777 End
        
        xmlSetAttributeValue vxmlComponentNode, "ASSUMEDSTARTDATE", strAssumedStartDate
        ' EP529 / MAR1731 End
        
        'Peter Edney - 05/09/2006 - EP1114
        'AddCalculationDateAttribute vobjCommon, vxmlComponentNode, bCalcDatePlus2Years  'MAR88 BC
        AddCalculationDateAttribute2 vobjCommon, vxmlComponentNode
        
        
        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, vxmlComponentNode)
        End If
        
        '*-add the appropriate Fees and/or Premiums element
        Set xmlTemp = AddFeesPremiumsElement(vobjCommon, vxmlComponentNode)
        If vblnIsOfferDocument Then
            If xmlTemp.baseName <> "NOFEESORPREMIUMS" Then
                '*-add the MORTGAGETYPE attribute
                Call AddMortgageTypeAttribute(vobjCommon, xmlTemp)
            End If
        End If
    End If
    
    ' Build the xml based on LoanComponentPaymentSchedule

    
    intTotalNoOfPayments = (vobjCommon.TermInYears * 12) + (vobjCommon.TermInMonths)
    
    'Need to match the interest rate type to the periods returned from ACE
    intNumPeriods = 0
    intTotNumPeriods = 0
    Set xmlIntRateTypeList = vxmlLoanComponent.selectNodes(".//INTERESTRATETYPE")
    intTotNumRates = xmlIntRateTypeList.length
    ReDim Preserve audtInterestRateData(1 To intTotNumRates)
    
    Set xmlLCPScheduleList = vxmlLoanComponent.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
    If xmlLCPScheduleList.length > 0 Then
        dtCurrentDate = xmlGetAttributeAsDate(xmlLCPScheduleList.Item(0), "STARTDATE")
    End If
    For intStepNo = 0 To intTotNumRates - 1 Step 1
        Set xmlIntRateType = xmlIntRateTypeList.Item(intStepNo)
        'Could be a mixture of Periods and enddates specified
        intNumPeriods = xmlGetAttributeAsDouble(xmlIntRateType, "INTERESTRATEPERIOD")
        If intNumPeriods <> 0 Then
            If intNumPeriods = -1 Then
                'This period runs to the end & is therefore the last period
                audtInterestRateData(intStepNo + 1).intCumulativePeriods = intNumPeriods
            Else
                intTotNumPeriods = intTotNumPeriods + intNumPeriods
                audtInterestRateData(intStepNo + 1).intCumulativePeriods = intTotNumPeriods
            End If
        Else

            'TK 07/11/2004 BBG1782
            Dim strIntRateEndDate As String
            strIntRateEndDate = xmlGetAttributeText(xmlIntRateType, "INTERESTRATEENDDATE")
            
            If InStr(1, strIntRateEndDate, "T") Then
                strIntRateEndDate = Left$(strIntRateEndDate, InStr(1, strIntRateEndDate, "T") - 1)

                If IsDate(strIntRateEndDate) Then
                    dtEndDate = CDate(strIntRateEndDate)
                End If
            Else
                dtEndDate = xmlGetAttributeAsDate(xmlIntRateType, "INTERESTRATEENDDATE")
            End If
            'TK 07/11/2004 BBG1782 End
            intToEndDatePeriods = DateDiff("M", dtCurrentDate, dtEndDate)
            intNumPeriods = IIf(intTotNumPeriods > 0, intToEndDatePeriods - intTotNumPeriods, _
                                                            intToEndDatePeriods)
        
            intTotNumPeriods = intTotNumPeriods + intNumPeriods
            audtInterestRateData(intStepNo + 1).intCumulativePeriods = intTotNumPeriods
            
        End If
        
        audtInterestRateData(intStepNo + 1).strIntRateType = _
                        xmlGetAttributeText(xmlIntRateType, "RATETYPE")
    
    Next intStepNo
   
 '  MAR44 BC 12 Sep 05 Save AccruedInterest from LoanComponent
    dblAccruedCost = xmlGetAttributeAsDouble(vxmlLoanComponent, "ACCRUEDINTEREST")
    dblFinalPayment = xmlGetAttributeAsDouble(vxmlLoanComponent, "FINALPAYMENT")
    
    'PB EP529 / MAR1731 Removed MAR1720 merge
    ''PB EP529 / MAR1720 Begin
    'If dblFinalPayment > 0 Then
    '    intMonthAdjustment = 1
    'Else
    '    intMonthAdjustment = 0
    'End If
    ''PB EP529 / MAR1720 End
    'PB EP529 / MAR1731 End
    
    intTotalSteps = xmlLCPScheduleList.length
    For intStepNo = 0 To intTotalSteps - 1 Step 1
        Set xmlLCPSchedule = xmlLCPScheduleList.Item(intStepNo)
        If intStepNo <> intTotalSteps - 1 Then
            Set xmlNextLCPSchedule = xmlLCPScheduleList.Item(intStepNo + 1)
            dtNextDate = xmlGetAttributeAsDate(xmlNextLCPSchedule, "STARTDATE")
        End If

'       MAR44 BC 12 Sep 05
'       strPaymentType = xmlGetAttributeText(xmlLCPSchedule, "PAYMENTTYPE")

        dblCurrentMonthlyCost = xmlGetAttributeAsDouble(xmlLCPSchedule, "MONTHLYCOST")
        dblInterestRate = xmlGetAttributeAsDouble(xmlLCPSchedule, "INTERESTRATE")
        dtCurrentDate = xmlGetAttributeAsDate(xmlLCPSchedule, "STARTDATE")
       
'       MAR44 BC 12 Sep 05
'       Select Case UCase$(strPaymentType)
'            Case "ACCRUEDINTEREST" 'If exists, it will always be the first record in PaymentSchedule
'                'We need to save the amount to add it to the first months payment
'                dblAccruedCost = dblCurrentMonthlyCost
'            Case "MONTHLYPAYMENT"

' EP529 / MAR1731 - Removed MAR1720 and reverted to original
'' EP529 / MAR1720
''                intNoOfPayments = IIf(intStepNo < intTotalSteps - 1, DateDiff("M", dtCurrentDate, dtNextDate), _
''                                                                intTotalNoOfPayments - intNoOfPaymentsConsidered)
'                intNoOfPayments = IIf(intStepNo < intTotalSteps - 1, _
'                                      DateDiff("M", dtCurrentDate, dtNextDate), _
'                                      intTotalNoOfPayments - intNoOfPaymentsConsidered - intMonthAdjustment) 'BC MAR1720
'' EP529 / MAR1720 End
                intNoOfPayments = IIf(intStepNo < intTotalSteps - 1, DateDiff("M", dtCurrentDate, dtNextDate), _
                                                                intTotalNoOfPayments - intNoOfPaymentsConsidered)
' EP529 / MAR1731 End

                If blnFirstMonthlyPayment Then
                    'Add in the Accrued Interest to the first payment
                    'Second lot of payments is less one as a result.
                    'MAR88 - BC 06Oct following commented as MARS will calculate Accrued Interest
'                    dblFirstMonthCost = dblAccruedCost + dblCurrentMonthlyCost
'                    Set xmlNode = vobjCommon.CreateNewElement("ACCRUED", vxmlComponentNode)
'                    xmlSetAttributeValue xmlNode, "FIRSTPAYMENTAMOUNT", set2DP(CStr(dblFirstMonthCost))
'                    'SR 11/01/2005 : E2EM00003136
'                    If StrComp(audtInterestRateData(1).strIntRateType, "F") = 0 Then
'                        Set xmlNode = vobjCommon.CreateNewElement("FIXED", xmlNode)
'                    Else
'                        Set xmlNode = vobjCommon.CreateNewElement("VARIABLE", xmlNode)
'                    End If
'                    'SR 11/01/2005 : E2EM00003136 - End
'                    xmlSetAttributeValue xmlNode, "RATE", set2DP(CStr(dblInterestRate))
'                    int2ndEntryNoOfPayments = intNoOfPayments - 1
                    int2ndEntryNoOfPayments = intNoOfPayments
                    If int2ndEntryNoOfPayments > 1 Then
                        Set xmlNode = vobjCommon.CreateNewElement("MONTHLYPAYMENT", vxmlComponentNode)
                        xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                        ' work out which rate period to use
                        For intStepNo2 = 1 To intTotNumRates Step 1
                            intCumulativePeriod = audtInterestRateData(intStepNo2).intCumulativePeriods
                            If intCumulativePeriod = -1 Then
                                'last period
                                intRateType = intStepNo2
                            Else
                                'should always be the first period here
                                If intNoOfPayments <= intCumulativePeriod Then
                                    'found the correct intRateType
                                    intRateType = intStepNo2
                                    'lets leave the loop
                                    intStepNo2 = intTotNumRates
                                End If
                                
                            End If
                        Next intStepNo2
                        
                        strIntRateType = audtInterestRateData(intRateType).strIntRateType
                        If StrComp(strIntRateType, "F") = 0 Then
                            Set xmlRateNode = vobjCommon.CreateNewElement("FIXED", xmlNode)
                            xmlSetAttributeValue xmlRateNode, "FIXEDRATE", set2DP(CStr(dblInterestRate))
                        Else
                            Set xmlRateNode = vobjCommon.CreateNewElement("VARIABLE", xmlNode)
                            xmlSetAttributeValue xmlRateNode, "RESOLVEDRATE", set2DP(CStr(dblInterestRate))
                        End If
                        xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", CStr(int2ndEntryNoOfPayments)
                    End If
                    blnFirstMonthlyPayment = False
                Else
                    intNoOfPaymentsToCompare = intNoOfPaymentsConsidered + intNoOfPayments
                    Set xmlNode = vobjCommon.CreateNewElement("FOLLOWEDBY", vxmlComponentNode)
                    xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                        ' work out which rate period to use
                    For intStepNo2 = 1 To intTotNumRates Step 1
                        intCumulativePeriod = audtInterestRateData(intStepNo2).intCumulativePeriods
                        If intCumulativePeriod = -1 Then
                            'last period
                            intRateType = intStepNo2
                        Else
                            If intNoOfPaymentsToCompare <= intCumulativePeriod Then
                                'found the correct intRateType
                                intRateType = intStepNo2
                                'lets leave the loop
                                intStepNo2 = intTotNumRates
                            End If
                            
                        End If
                    Next intStepNo2
                    
                    strIntRateType = audtInterestRateData(intRateType).strIntRateType
                    If StrComp(strIntRateType, "F") = 0 Then
                        Set xmlRateNode = vobjCommon.CreateNewElement("FIXED", xmlNode)
                        xmlSetAttributeValue xmlRateNode, "FIXEDRATE", set2DP(CStr(dblInterestRate))
                    Else
                        Set xmlRateNode = vobjCommon.CreateNewElement("VARIABLE", xmlNode)
                        xmlSetAttributeValue xmlRateNode, "RESOLVEDRATE", set2DP(CStr(dblInterestRate))
                    End If
                    xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", CStr(intNoOfPayments)
                End If
                intNoOfPaymentsConsidered = intNoOfPaymentsConsidered + intNoOfPayments

        
    Next intStepNo
    
'   MAR44 BC 12 Sep 05  moved this "If" statement out of "Next intStepNo" loop
    If dblFinalPayment > 0 Then
        Set xmlNode = vobjCommon.CreateNewElement("REMAINDEROFTERM", vxmlComponentNode)
            If vblnInterestOnly Then
                dblFinalPayment = dblFinalPayment - _
                        (vobjCommon.LoanAmount + vobjCommon.FeesAddedToLoanAmount)
            End If
        xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblFinalPayment))
        'SR 11/01/2005 : E2EM00003136
            If StrComp(audtInterestRateData(intTotNumRates).strIntRateType, "F") = 0 Then
                Set xmlNode = vobjCommon.CreateNewElement("FIXED", xmlNode)
            Else
                Set xmlNode = vobjCommon.CreateNewElement("VARIABLE", xmlNode)
            End If
        'SR 11/01/2005 : E2EM00003136 - End
        xmlSetAttributeValue xmlNode, "RATE", set2DP(CStr(dblInterestRate))
    End If
    
    'TK 07/11/2004 BBG1782
    Set xmlRate = Nothing
    Set xmlList = Nothing
    Set xmlTemp = Nothing
    Set xmlAssumes = Nothing 'BC MAR907
    Set xmlLCPScheduleList = Nothing
    Set xmlLCPSchedule = Nothing
    Set xmlNextLCPSchedule = Nothing
    Set xmlNode = Nothing
    Set xmlIntRateTypeList = Nothing
    Set xmlIntRateType = Nothing
    'TK 07/11/2004 BBG1782 End
    Set xmlRateNode = Nothing
Exit Sub
ErrHandler:
    Set xmlRate = Nothing
    Set xmlList = Nothing
    Set xmlTemp = Nothing
    Set xmlAssumes = Nothing 'BC MAR907
    Set xmlLCPScheduleList = Nothing
    Set xmlLCPSchedule = Nothing
    Set xmlNextLCPSchedule = Nothing
    Set xmlNode = Nothing
    Set xmlIntRateTypeList = Nothing
    Set xmlIntRateType = Nothing
    Set xmlRateNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
    
End Sub

'********************************************************************************
'** Function:       BuildOverallCostsSection
'** Created by:     Andy Maggs
'** Date:           31/03/2004
'** Description:    Sets the elements and attributes for the overall costs
'**                 section that is common to standard mortgage (KFI and Offer)
'**                 and to transfer of equity mortgage offer documents.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildOverallCostsSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildOverallCostsSection"
    Dim xmlNewNode As IXMLDOMNode
    Dim intYears As Integer
    Dim intMonths As Integer
    Dim xmlMortgageType As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    Dim xmlPurposeNode As IXMLDOMNode
    
    On Error GoTo ErrHandler

    '*-add the APR attribute
    Call AddAPRAttribute(vobjCommon, vxmlNode)
    '*-add the PAYBACKPERPOUND attribute
'CORE82 Get this from the loan component    Call AddPaybackPerPoundAttribute(vobjCommon, vxmlNode)
    ' PB 25/05/2006
    'Call xmlSetAttributeValue(vxmlNode, "PAYBACKPERPOUND", _
    '        CStr(vobjCommon.AmountPerUnitBorrowed))
    Call xmlSetAttributeValue(vxmlNode, "PAYBACKPERPOUND", _
            CStr(vobjCommon.AmountPerPound))
    ' EP603/MAR1788 End
    '*-add the TOTALREPAYMENTAMOUNT attribute
    Call xmlSetAttributeValue(vxmlNode, "TOTALREPAYMENTAMOUNT", _
            CStr(vobjCommon.LoanAmountRepayable))
        
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, vxmlNode)
    End If
        
    '*-add the PAYMENTCOMMENT1 element
    Set xmlNewNode = vobjCommon.CreateNewElement("PAYMENTCOMMENT1", vxmlNode)
    '*-add the type of mortgage mandatory (REPAYMENT, INTERESTONLY or
    '*-PARTANDPART) elements
    Set xmlTemp = vobjCommon.AddMortgageRepaymentTypeElement(xmlNewNode)
    'INR CORE82 Need LoanPurpose
    Set xmlPurposeNode = vobjCommon.AddLoanPurposeElement(xmlTemp)
    
    If vblnIsOfferDocument Then
        '*-add the MORTGAGETYPE attribute
        Call AddMortgageTypeAttribute(vobjCommon, xmlTemp)
    End If
    
    '*-add the VARIABLE element
    If vobjCommon.IsVariableRate Then
        Set xmlNewNode = vobjCommon.CreateNewElement("VARIABLE", vxmlNode)
        'CORE82 Add TermYears & Months from WhatYouHaveToldUs
'        Call GetFixedRateTerm(vobjCommon, intYears, intMonths)
'        Call xmlSetAttributeValue(xmlNewNode, "TERMYEARS", CStr(intYears))
'        Call xmlSetAttributeValue(xmlNewNode, "TERMMONTHS", CStr(intMonths))
        '* Add TERMYEARS and TERMMONTHS as attribs
        Call AddMortgageTermAttributes(vobjCommon, xmlNewNode, "TERMYEARS", "TERMMONTHS")

        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, xmlNewNode)
        End If
    End If
        
    '*-add the PAYMENTCOMMENT2 element
    Set xmlNewNode = vobjCommon.CreateNewElement("PAYMENTCOMMENT2", vxmlNode)
    If vobjCommon.IsAdditionalBorrowing Then
        '*-add the ADDITIONALBORROWING element
        Set xmlMortgageType = vobjCommon.CreateNewElement("ADDITIONALBORROWING", xmlNewNode)
    Else
        '*-add the MORTGAGE element
        Set xmlMortgageType = vobjCommon.CreateNewElement("MORTGAGE", xmlNewNode)
    End If
    '*-add the REPAYMENT, INTERESTONLY or PARTANDPART element as appropriate
    Call vobjCommon.AddMortgageRepaymentTypeElement(xmlMortgageType)

    'PB 05/03/2007 EP2_1629
    'If vobjCommon.IsProductSwitch Or vobjCommon.IsTransferOfEquity Then
    '    xmlSetAttributeValue vxmlNode, "SECTIONNUMBER", "5 and 7"
    'Else
        xmlSetAttributeValue vxmlNode, "SECTIONNUMBER", "6 and 8"
    'End If
    
    Set xmlNewNode = Nothing
    Set xmlMortgageType = Nothing
    Set xmlTemp = Nothing
    Set xmlPurposeNode = Nothing
Exit Sub
ErrHandler:
    Set xmlNewNode = Nothing
    Set xmlMortgageType = Nothing
    Set xmlTemp = Nothing
    Set xmlPurposeNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Public Function set2DP(ByVal strValue As String) As String
'Checks for existance of decimal point and two digits after decimal.
'If no, creates them (adds decimal and zero)
    
    ' Peter Edney - 25/09/2006 - EP1160
    set2DP = FormatNumber(strValue, 2, vbTrue, vbFalse, vbFalse)
    
'    Dim lngDecimalPosition As Long, lngStringLength As String
'    Dim strReturn As String
'
'    lngStringLength = Len(strValue)
'    lngDecimalPosition = InStr(strValue, ".")
'    If lngDecimalPosition = 0 Then
'        strReturn = strValue & ".00"
'    ElseIf lngStringLength - lngDecimalPosition = 1 Then
'        strReturn = strValue & "0"
'    ElseIf lngStringLength - lngDecimalPosition > 2 Then
'        strReturn = Left$(strValue, lngDecimalPosition) + Mid$(strValue, lngDecimalPosition + 1, 2)
'    Else
'        strReturn = strValue
'    End If
'
'    set2DP = strReturn

End Function

'BBG1565 New Function
'********************************************************************************
'** Function:       CheckIfVariableMortgageInterestRateType
'** Created by:     Ian Ross
'** Date:           06/10/2004
'** Description:    Checks if at least one of the mortgage interest rate types
'**                 on the loan component is variable.
'** Parameters:     vxmlLoanComponent - the loan component XML to check.
'** Returns:        The mortgage interest rate type. Fixed if no variable else non fixed.
'** Errors:         None Expected
'********************************************************************************
Public Function CheckIfVariableMortgageInterestRateType(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode) As MortgageInterestRateType
    Const cstrFunctionName As String = "CheckIfVariableMortgageInterestRateType"
    
    Dim eResult As MortgageInterestRateType
    Dim eType As MortgageInterestRateType
    Dim xmlList As IXMLDOMNodeList
    Dim xmlRate As IXMLDOMNode
    Dim intIndex As Integer
    Dim intTo As Integer

    On Error GoTo ErrHandler

    vobjCommon.blnFixedRateMortgage = False
    eResult = mrtFixedRate
    
    '*-get all the interest rates for this loan component
    Set xmlList = vobjCommon.GetLoanComponentInterestRates(vxmlLoanComponent)
    
    intTo = xmlList.length - 1
    
    '*-add the rate periods
    For intIndex = 0 To intTo
        Set xmlRate = xmlList.Item(intIndex)
               
        '*-get the rate type for this item
        eType = GetInterestRateType(xmlRate)
    
        If eType <> mrtFixedRate Then
            '*-there is at least one component that is not fixed
            eResult = eType
        Else
            vobjCommon.blnFixedRateMortgage = True
        End If
    
    Next intIndex
    
    '*-return the result
    CheckIfVariableMortgageInterestRateType = eResult

    Set xmlList = Nothing
    Set xmlRate = Nothing
Exit Function
ErrHandler:
    Set xmlList = Nothing
    Set xmlRate = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'BBG1489 New NonRegulated template Sub
'********************************************************************************
'** Sub:            BuildNonRegFeesYouMustPaySection
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the elements and attributes for the Non Regulated
'**                 KFI document.
'** Parameters:     vxmlNode - the section element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildNonRegFeesYouMustPaySection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode, ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "BuildNonRegFeesYouMustPaySection"
    Dim xmlCost As IXMLDOMNode
    Dim xmlFee As IXMLDOMNode
    Dim lngFeeAmount As Long
    Dim xmlApplication As IXMLDOMNode
    
'PB 04/12/2006 EP2_139
    Dim strLegalFeeText As String
    
    On Error GoTo ErrHandler
    
    '*-add the VALUATIONFEE element
    'SR 18/10/2004 : BBG1645 - First check for Third Party Valuation Fee and then Valuation Fee
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'TPV')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
        If lngFeeAmount > 0 Then
            Set xmlCost = vobjCommon.CreateNewElement("VALUATIONFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    
    If xmlFee Is Nothing Or lngFeeAmount = 0 Then
        Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'VAL')]")
        If Not xmlFee Is Nothing Then
            lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
            Set xmlCost = vobjCommon.CreateNewElement("VALUATIONFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    'SR 18/10/2004 : BBG1645 - End
    
    '*-add the COMPLETIONFEE element
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'ARR')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
        If lngFeeAmount > 0 Then
            Set xmlCost = vobjCommon.CreateNewElement("COMPLETIONFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    
    '*-add the LEGALFEE element
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'LEG')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
    Else
        lngFeeAmount = 0
    End If
    strLegalFeeText = GetLegalFeeText(vobjCommon)
    '
    
    Call xmlSetAttributeValue(xmlCost, "LEGALCOSTSVALUE", lngFeeAmount)
    'Check for Additional Borrowing ("ABO") or Product Switch ("PSW")
    Set xmlApplication = vobjCommon.Data.selectSingleNode(gcstrAPPLICATION_PATH)
    Debug.Print xmlGetAttributeText(xmlApplication, "TYPEOFBUYER")
    
    
    '*-add the TTFEE element
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'TTF')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
        If lngFeeAmount > 0 Then
            Set xmlCost = vobjCommon.CreateNewElement("TTFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    
    '*-add the SEALINGFEE element
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'SEA')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
        If lngFeeAmount > 0 Then
            Set xmlCost = vobjCommon.CreateNewElement("SEALINGFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    
    '*-add the DEEDSRELEASEFEE element
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'DEE')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
        If lngFeeAmount > 0 Then
            Set xmlCost = vobjCommon.CreateNewElement("DEEDSRELEASEFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    
    '*-add the BROKERFEE element
    Set xmlFee = vobjCommon.Data.selectSingleNode(gcstrMORTGAGEONEOFFCOSTS_PATH & "[contains(@MORTGAGEONEOFFCOSTTYPE,'BRK')]")
    If Not xmlFee Is Nothing Then
        lngFeeAmount = xmlGetAttributeAsLong(xmlFee, "AMOUNT")
        If lngFeeAmount > 0 Then
            Set xmlCost = vobjCommon.CreateNewElement("BROKERFEE", vxmlNode)
            Call xmlSetAttributeValue(xmlCost, "FEEAMOUNT", lngFeeAmount)
        End If
    End If
    
    Set xmlCost = Nothing
    Set xmlFee = Nothing
Exit Sub
ErrHandler:
    Set xmlCost = Nothing
    Set xmlFee = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'BBG1489 New NonRegulated template Sub
'********************************************************************************
'** Sub:       NRBuildEarlyRepayCharges
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the elements and attributes for the (What happens if you
'**                 do not want this mortgage any more ?) element that appears in
'**                 the Non Regulated KFI.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section10 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub NRBuildEarlyRepayCharges(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "NRBuildEarlyRepayCharges"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlFeesList As IXMLDOMNodeList
    Dim xmlFee As IXMLDOMNode
    Dim intNumber As Integer
    Dim xmlEarly As IXMLDOMNode
    Dim xmlCharges As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlRepay As IXMLDOMNode
    Dim xmlBandList As IXMLDOMNodeList
    Dim xmlBand As IXMLDOMNode
    Dim intStepNum As Integer
    Dim xmlPeriod As IXMLDOMNode
    Dim dblPercentage As Double
    Dim strBasis As String
    Dim lngRepaymentFee As Long
    Dim dblChargeAmount As Double
    Dim dblRedemptionFeeAmount As Double
    Dim dblTotalRedemptionFeeAmount As Double
        
    On Error GoTo ErrHandler
    
    '*-get the list of loan components and from each component, get the list of
    '*-redemption fees
    Set xmlList = vobjCommon.LoanComponents
    For Each xmlLoanComponent In xmlList
        Set xmlFeesList = xmlLoanComponent.selectNodes("LOANCOMPONENTREDEMPTIONFEE")
        intNumber = intNumber + xmlFeesList.length
    Next xmlLoanComponent
    
    If intNumber > 0 Then
    
        '*-add the EARLYREPAYMENT element
        Set xmlEarly = vobjCommon.CreateNewElement("EARLYREPAYMENT", vxmlNode)
        '*-add the CHARGES element
        Set xmlCharges = vobjCommon.CreateNewElement("CHARGES", xmlEarly)
        
        intNumber = 0
        dblTotalRedemptionFeeAmount = 0
        For Each xmlLoanComponent In xmlList
            '*-increment the count
            intNumber = intNumber + 1
            '*-add the COMPONENT element
            Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", xmlCharges)
            '*-add the mandatory PART attribute
            Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intNumber))
            
            '*-get the redemption fees list
            ' already got above, but get again because of MS Parser for each bug
            Set xmlFeesList = xmlLoanComponent.selectNodes("LOANCOMPONENTREDEMPTIONFEE")
                
            '*-add the REPAYMENTCHARGES element
            Set xmlRepay = vobjCommon.CreateNewElement("REPAYMENTCHARGES", xmlComponent)
            
            lngRepaymentFee = vobjCommon.MortgageCompletionFee
            dblRedemptionFeeAmount = GetMaxDblAttribValue(xmlFeesList, "REDEMPTIONFEEAMOUNT")  'SR 23/09/2004 : CORE82
            dblTotalRedemptionFeeAmount = dblTotalRedemptionFeeAmount + dblRedemptionFeeAmount 'SR 13/09/2004 : CORE82
            
            '*-get the redemption fee band list
            Set xmlBandList = xmlLoanComponent.selectNodes("//REDEMPTIONFEEBAND")
            For Each xmlBand In xmlBandList
                '*-get the step number for this band item
                intStepNum = xmlGetAttributeAsInteger(xmlBand, "REDEMPTIONFEESTEPNUMBER")
                '*-and get the corresponding redemption fee item
                Set xmlFee = xmlLoanComponent.selectSingleNode("LOANCOMPONENTREDEMPTIONFEE[@REDEMPTIONFEESTEPNUMBER=" & intStepNum & "]")
                
                '*-add the PERIOD element
                Set xmlPeriod = vobjCommon.CreateNewElement("PERIOD", xmlRepay)
                '-add the CHARGEAMOUNT attribute
                dblChargeAmount = xmlGetAttributeAsDouble(xmlFee, "REDEMPTIONFEEAMOUNT")
                dblChargeAmount = dblChargeAmount + lngRepaymentFee
                Call xmlSetAttributeValue(xmlPeriod, "CHARGEAMOUNT", CStr(dblChargeAmount))
                        
                '*-add the BASIS attribute
                If xmlAttributeValueExists(xmlBand, "FEEPERCENTAGE") Then
                    dblPercentage = xmlGetAttributeAsDouble(xmlBand, "FEEPERCENTAGE")
                Else
                    dblPercentage = 0
                End If
                
                If dblPercentage > 0 Then
                    'PB 21/06/2006 EP697
                    'strBasis = CStr(dblPercentage) & "% of outstanding loan amount"
                    strBasis = CStr(dblPercentage) & "% of capital balance"
                    'PB EP697 End
                Else
                    strBasis = CStr(xmlGetAttributeAsInteger(xmlBand, "FEEMONTHSINTEREST")) & " months interest"
                End If
                Call xmlSetAttributeValue(xmlPeriod, "BASIS", strBasis)
                
                '*-add the PERIOD attribute - this is the number of months the
                '*-charge applies - get the period or get the end date
                ' PB 21/06/2006 EP697 Begin
                'Call AddPeriodAttribute(xmlFee, "REDEMPTIONFEEPERIOD", _
                '       "REDEMPTIONFEEPERIODENDDATE", _
                '       vobjCommon.ExpectedCompletionDate, "PERIOD", xmlPeriod)
                Call AddPeriodAttribute(xmlFee, "REDEMPTIONFEEPERIOD", _
                       "REDEMPTIONFEEPERIODENDDATE", _
                       vobjCommon.ExpectedCompletionDate, "PERIOD", xmlPeriod, "")
                ' PB EP697 End
            Next xmlBand

        Next xmlLoanComponent
        
    End If
    
    Set xmlList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlFeesList = Nothing
    Set xmlFee = Nothing
    Set xmlEarly = Nothing
    Set xmlCharges = Nothing
    Set xmlComponent = Nothing
    Set xmlRepay = Nothing
    Set xmlBandList = Nothing
    Set xmlBand = Nothing
    Set xmlPeriod = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlFeesList = Nothing
    Set xmlFee = Nothing
    Set xmlEarly = Nothing
    Set xmlCharges = Nothing
    Set xmlComponent = Nothing
    Set xmlRepay = Nothing
    Set xmlBandList = Nothing
    Set xmlBand = Nothing
    Set xmlPeriod = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'BBG1489 New NonRegulated template Sub
'********************************************************************************
'** Sub:            AddSpecialGroupNameAttribute
'** Created by:     Ian Ross
'** Date:           30/09/2004
'** Description:    Gets the SpecialGroup text from ApplicationFactFind.
'**
'** Parameters:     vxmlData - the full KFI input data.
'**                 vblnPurposePurchase - whether the mortgage is for a new
'**                 purchase or not.
'**                 rblnIsEstimated - True if a valuation is returned.
'**                 rlngValue - the valuation or purchase price as appropriate.
'** Returns:        rblnEstimated, rlngValue.
'** Errors:         None Expected
'********************************************************************************
Public Sub AddSpecialGroupNameAttribute(ByVal vxmlData As IXMLDOMNode, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddSpecialGroupNameAttribute"
    Dim xmlItem As IXMLDOMNode
    Dim strSpecialGroup As String
    
    On Error GoTo ErrHandler

    '*-get the APPLICATIONFACTFIND Node
    Set xmlItem = vxmlData.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        strSpecialGroup = xmlGetAttributeText(xmlItem, "SPECIALGROUP_TEXT")
        Call xmlSetAttributeValue(vxmlNode, "SPECIALGROUP", strSpecialGroup)
    End If
    
    Set xmlItem = Nothing
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'BBG1489 New NonRegulated template Sub
Public Sub AddNRLoanComponentInterestRateData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode, ByVal vxmlComponentNode As IXMLDOMNode, _
        ByVal vblnIsOfferDocument As Boolean)
    Const cstrFunctionName As String = "AddNRLoanComponentInterestRateData"
    Dim xmlRate As IXMLDOMNode
    Dim intTotalNumPayments As Integer
    Dim xmlList As IXMLDOMNodeList
    Dim intIndex As Integer
    Dim intTotalMonths As Integer
    Dim xmlTemp As IXMLDOMNode
    Dim bCalcDatePlus2Years As Boolean 'MAR88 BC
        
    On Error GoTo ErrHandler

    bCalcDatePlus2Years = False 'MAR88 BC
    
    '*-get the first rate
    Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(vxmlLoanComponent)
    If Not xmlRate Is Nothing Then
        
        '*-add the mandatory RATE attribute
        Call AddRateAttribute(vobjCommon, vxmlLoanComponent, xmlRate, _
                vxmlComponentNode, "RATE")
        AddCalculationDateAttribute vobjCommon, vxmlComponentNode, bCalcDatePlus2Years 'MAR88 BC
        If vblnIsOfferDocument Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, vxmlComponentNode)
        End If
        
        '*-add the appropriate Fees and/or Premiums element
        Set xmlTemp = AddFeesPremiumsElement(vobjCommon, vxmlComponentNode)
        If vblnIsOfferDocument Then
            If xmlTemp.baseName <> "NOFEESORPREMIUMS" Then
                '*-add the MORTGAGETYPE attribute
                Call AddMortgageTypeAttribute(vobjCommon, xmlTemp)
            End If
        End If
    End If
    
    ' Build the xml based on LoanComponentPaymentSchedule
    Dim xmlLCPScheduleList As IXMLDOMNodeList, xmlLCPSchedule As IXMLDOMNode, xmlNextLCPSchedule As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strPaymentType As String
    Dim dblCurrentMonthlyCost As Double, dblInterestRate As Double
    Dim dblFirstMonthCost As Double, dblAccruedCost As Double
    Dim dtCurrentDate As Date, dtNextDate As Date, dtTemp As Date
    Dim intStepNo As Integer, intTotalSteps As Integer
    Dim intNoOfPayments As Integer, intTotalNoOfPayments As Integer, intNoOfPaymentsConsidered As Integer
    Dim blnFirstMonthlyPayment As Boolean
    Dim int2ndEntryNoOfPayments As Integer
    blnFirstMonthlyPayment = True
    
    intTotalNoOfPayments = (vobjCommon.TermInYears * 12) + (vobjCommon.TermInMonths)
    Set xmlLCPScheduleList = vxmlLoanComponent.selectNodes(".//LOANCOMPONENTPAYMENTSCHEDULE")
    intTotalSteps = xmlLCPScheduleList.length
    For intStepNo = 0 To intTotalSteps - 1 Step 1
        Set xmlLCPSchedule = xmlLCPScheduleList.Item(intStepNo)
        If intStepNo <> intTotalSteps - 1 Then
            Set xmlNextLCPSchedule = xmlLCPScheduleList.Item(intStepNo + 1)
            dtNextDate = xmlGetAttributeAsDate(xmlNextLCPSchedule, "STARTDATE")
        End If
        
        strPaymentType = xmlGetAttributeText(xmlLCPSchedule, "PAYMENTTYPE")
        dblCurrentMonthlyCost = xmlGetAttributeAsDouble(xmlLCPSchedule, "MONTHLYCOST")
        dblInterestRate = xmlGetAttributeAsDouble(xmlLCPSchedule, "INTERESTRATE")
        dtCurrentDate = xmlGetAttributeAsDate(xmlLCPSchedule, "STARTDATE")
        
        Select Case UCase$(strPaymentType)
            Case "ACCRUEDINTEREST"
                'If exists, it will always be the first record in PaymentSchedule
                'We need to save the amount and add it to the first months payment
                dblAccruedCost = dblCurrentMonthlyCost
            Case "MONTHLYPAYMENT"
                intNoOfPayments = IIf(intStepNo < intTotalSteps - 1, DateDiff("M", dtCurrentDate, dtNextDate), _
                                                                intTotalNoOfPayments - intNoOfPaymentsConsidered)
                                                                
                If blnFirstMonthlyPayment Then
                    'Add in the Accrued Interest
                    dblFirstMonthCost = dblAccruedCost + dblCurrentMonthlyCost
                    Set xmlNode = vobjCommon.CreateNewElement("FOLLOWEDBY", vxmlComponentNode)
                    xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblFirstMonthCost))
                    xmlSetAttributeValue xmlNode, "INTRATE", set2DP(CStr(dblInterestRate))
                    xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", "1"
                    int2ndEntryNoOfPayments = intNoOfPayments - 1
                    If int2ndEntryNoOfPayments > 1 Then
                        Set xmlNode = vobjCommon.CreateNewElement("FOLLOWEDBY", vxmlComponentNode)
                        xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                        xmlSetAttributeValue xmlNode, "INTRATE", set2DP(CStr(dblInterestRate))
                        xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", CStr(int2ndEntryNoOfPayments)
                    End If
                    blnFirstMonthlyPayment = False
                Else
                    Set xmlNode = vobjCommon.CreateNewElement("FOLLOWEDBY", vxmlComponentNode)
                    xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                    xmlSetAttributeValue xmlNode, "INTRATE", set2DP(CStr(dblInterestRate))
                    xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", CStr(intNoOfPayments)
                End If
                intNoOfPaymentsConsidered = intNoOfPaymentsConsidered + intNoOfPayments
            Case "FINALPAYMENT"
                Set xmlNode = vobjCommon.CreateNewElement("FOLLOWEDBY", vxmlComponentNode)
                xmlSetAttributeValue xmlNode, "PAYMENT", set2DP(CStr(dblCurrentMonthlyCost))
                xmlSetAttributeValue xmlNode, "INTRATE", set2DP(CStr(dblInterestRate))
                xmlSetAttributeValue xmlNode, "NUMBEROFPAYMENTS", "1"
        End Select
    Next intStepNo
    
    Set xmlRate = Nothing
    Set xmlList = Nothing
    Set xmlTemp = Nothing
    Set xmlLCPScheduleList = Nothing
    Set xmlLCPSchedule = Nothing
    Set xmlNextLCPSchedule = Nothing
    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlRate = Nothing
    Set xmlList = Nothing
    Set xmlTemp = Nothing
    Set xmlLCPScheduleList = Nothing
    Set xmlLCPSchedule = Nothing
    Set xmlNextLCPSchedule = Nothing
    Set xmlNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'BBG1587 New Function
'********************************************************************************
'** Function:       IsSpecialDeals
'** Created by:     Ian Ross
'** Date:           11/10/2004
'** Description:    Gets whether the mortgage is has SpecialDeals.
'** Parameters:     vxmlData - the XML data to search.
'** Returns:        True if it is, else False.
'** Errors:         None Expected
'********************************************************************************
Public Function IsSpecialDeals(ByVal vobjCommon As CommonDataHelper) As Boolean
    Const cstrFunctionName As String = "IsSpecialDeals"
    Dim xmlNode As IXMLDOMNode
    Dim blnIsSpecialDeal As Boolean
    Dim strSpecialGroup As String

    On Error GoTo ErrHandler
    
    blnIsSpecialDeal = False
    
    ' *-get the APPLICATIONFACTFIND Node
    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlNode Is Nothing Then
        'BBG1637 change from SPECIALGROUP_VALIDID to SPECIALGROUP
        'didn't always seem to work
        strSpecialGroup = xmlGetAttributeText(xmlNode, "SPECIALGROUP")
        If (InStr(1, strSpecialGroup, "HPC", vbTextCompare) _
        Or InStr(1, strSpecialGroup, "FLX", vbTextCompare) _
        Or InStr(1, strSpecialGroup, "STD", vbTextCompare)) Then
            blnIsSpecialDeal = True
        End If
    End If
    
    IsSpecialDeals = blnIsSpecialDeal
    
    Set xmlNode = Nothing
Exit Function
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'BBG1545 New Function
'********************************************************************************
'** Function:       AddSpecialDealElements
'** Created by:     Ian Ross
'** Date:           11/10/2004
'** Description:    Adds SPECIALDEALBANDS (and children) elements.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmLoanComponent - the input data representing a single
'**                 LOANCOMPONENT
'**                 vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddSpecialDealElements(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlLoanComponent As IXMLDOMNode, ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddSpecialDealElements"
    
    Dim xmlList As IXMLDOMNodeList
    Dim xmlBaseRateBandList As IXMLDOMNodeList

    Dim xmlLC As IXMLDOMNode
    Dim xmlBaseRateSet As IXMLDOMNode
    Dim xmlBaseRateBand As IXMLDOMNode
    Dim xmlBaseRate As IXMLDOMNode 'SR 25/10/2004 : BBG1684
    Dim xmlSpecialDealBandNode As IXMLDOMNode
    Dim xmlLTVBandNode As IXMLDOMNode
    Dim strNameOfSVR As String
    Dim strText As String
    Dim intIndex As Integer
    Dim dtStartDate As Date
    Dim dtMaxDate As Date

    On Error GoTo ErrHandler

    Set xmlLC = vobjCommon.SingleLoanComponent
    'LOANCOMPONENT/MORTGAGEPRODUCT/INTERESTRATETYPE/RATETYPE[@INTERESTRATEPERIOD]/BASERATESETDATA
    Set xmlBaseRateSet = xmlLC.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATEPERIOD = -1]/BASERATESETDATA")
    
    If Not xmlBaseRateSet Is Nothing Then
        Set xmlBaseRate = xmlBaseRateSet.selectSingleNode("BASERATE") 'SR 25/10/2004 : BBG1684
        strNameOfSVR = xmlGetAttributeText(xmlBaseRate, "RATEDESCRIPTION") 'SR 25/10/2004 : BBG1684
        Set xmlBaseRateBandList = xmlBaseRateSet.selectNodes(".//BASERATEBAND")
        
        'E2EM00002929 Find the most recent BASERATEBANDSTARTDATE
        If xmlBaseRateBandList.length > 0 Then
            dtMaxDate = xmlGetAttributeAsDate(xmlBaseRateBandList.Item(0), "BASERATEBANDSTARTDATE")
            For intIndex = 0 To xmlBaseRateBandList.length - 1
                dtStartDate = xmlGetAttributeAsDate(xmlBaseRateBandList.Item(intIndex), "BASERATEBANDSTARTDATE")
                If dtStartDate > dtMaxDate Then
                    dtMaxDate = dtStartDate
                End If
            Next intIndex
            
            Set xmlSpecialDealBandNode = vobjCommon.CreateNewElement("SPECIALDEALBANDS", vxmlNode)
             
            '*-add the BaseRateBands
            For intIndex = 0 To xmlBaseRateBandList.length - 1
                'E2EM00002929 only want bands associated with the latest BASERATEBANDSTARTDATE
                Set xmlBaseRateBand = xmlBaseRateBandList.Item(intIndex)
                dtStartDate = xmlGetAttributeAsDate(xmlBaseRateBand, "BASERATEBANDSTARTDATE")
                
                If dtStartDate = dtMaxDate Then
                    'Only add the LTVBANDS which have dtMaxDate
                    Set xmlLTVBandNode = vobjCommon.CreateNewElement("LTVBANDS", xmlSpecialDealBandNode)
                    '*-add the NAMEOFSVR attribute
                    Call xmlSetAttributeValue(xmlLTVBandNode, "NAMEOFSVR", strNameOfSVR)
            
                    strText = xmlGetAttributeText(xmlBaseRateBand, "MAXIMUMLTV")
                    '*-add the LTVBAND attribute
                    Call xmlSetAttributeValue(xmlLTVBandNode, "LTVBAND", strText)
                    strText = xmlGetAttributeText(xmlBaseRateBand, "RATEDIFFERENCE")
                    '*-add the PERCENTABOVE attribute
                    Call xmlSetAttributeValue(xmlLTVBandNode, "PERCENTABOVE", strText)
                End If
        
            Next intIndex
        End If 'xmlBaseRateBandList.length > 0
    End If

    Set xmlList = Nothing
    Set xmlBaseRateBandList = Nothing
    Set xmlLC = Nothing
    Set xmlBaseRateSet = Nothing
    Set xmlBaseRateBand = Nothing
    Set xmlBaseRate = Nothing
    Set xmlSpecialDealBandNode = Nothing
    Set xmlLTVBandNode = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlBaseRateBandList = Nothing
    Set xmlLC = Nothing
    Set xmlBaseRateSet = Nothing
    Set xmlBaseRateBand = Nothing
    Set xmlBaseRate = Nothing
    Set xmlSpecialDealBandNode = Nothing
    Set xmlLTVBandNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'E2EM00001894 New Function
'********************************************************************************
'** Function:       IsBuyToLet
'** Created by:     AW
'** Date:           18/01/2005
'** Description:    Gets whether the mortgage is a Buy To Let.
'** Parameters:     vxmlData - the XML data to search.
'** Returns:        True if it is, else False.
'** Errors:         None Expected
'********************************************************************************
Public Function IsBuyToLet(ByVal vobjCommon As CommonDataHelper) As Boolean
    Const cstrFunctionName As String = "IsBuyToLet"
    Dim xmlNode As IXMLDOMNode
    Dim blnIsBuyToLet As Boolean
    Dim strSpecialGroup As String

    On Error GoTo ErrHandler
    
    blnIsBuyToLet = False
    
    ' *-get the APPLICATIONFACTFIND Node
    Set xmlNode = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlNode Is Nothing Then
        strSpecialGroup = xmlGetAttributeText(xmlNode, "SPECIALGROUP")
        If (InStr(1, strSpecialGroup, "BTL", vbTextCompare)) Then
            blnIsBuyToLet = True
        End If
    End If
    
    IsBuyToLet = blnIsBuyToLet
    
    Set xmlNode = Nothing
Exit Function
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function
'EP603 - DRC
'MAR1777 GHun
'Calculate the number of whole months difference between 2 dates (i.e. rounding down)
'The VB DataDiff function ignores the day of the month in the dates, so it sometimes
'rounds this value up, and other times rounds it down.
Public Function MonthDiff(ByVal vdteDate1 As Date, ByVal vdteDate2 As Date) As Long
    Dim lngMonths As Long
    
    lngMonths = DateDiff("m", vdteDate1, vdteDate2)
    'If the difference has been rounded up then decrement the number of months
    If DateDiff("d", DateAdd("m", lngMonths, vdteDate1), vdteDate2) < 0 Then
        lngMonths = lngMonths - 1
    End If
    
    MonthDiff = lngMonths
End Function
'MAR1777 End
'EP603 - End
'EP603 - DRC
'MAR1777 GHun
Public Function CalcExpectedCompletionDate(ByVal vdteDate As Date) As Date
    Dim dteCompDate As Date
    Dim intDayOfMonth As Integer
    
    'Add one month
    dteCompDate = DateAdd("m", 1, vdteDate)
    
    'Set the day of the month to AssumedCompletionDayOfMonth
    intDayOfMonth = GetGlobalParamAmount("AssumedCompletionDayOfMonth")
    If intDayOfMonth = 0 Then   'Default the value to 1 in case the globalParameter does not exist
        intDayOfMonth = 1
    End If
    
    dteCompDate = CSafeDate(Format$(intDayOfMonth, "00") & "/" & Format$(Month(dteCompDate), "00") & "/" & CStr(Year(dteCompDate)))
    
    CalcExpectedCompletionDate = dteCompDate
    
End Function
'MAR1777 End
'EP603 - DRC

'PB 29/06/2006 EP845/EP846 Begin
Function CheckForValidationType(strListOfTypes As String, strTypeToSearchFor As String) As Boolean
    '
    Dim a_strList() As String
    Dim intLoop As Integer
    '
    a_strList = Split(strListOfTypes, ",")
    For intLoop = LBound(a_strList) To UBound(a_strList)
        If a_strList(intLoop) = strTypeToSearchFor Then
            CheckForValidationType = True
            Exit For
        End If
    Next intLoop
    '
End Function
'PB 29/06/2006 EP845/EP846 End

'********************************************************************************
'** Sub:            FormatNames
'** Created by:     Peter Edney (EP892)
'** Date:           25/08/2006
'** Description:    Takes the paramarray of names and returns the names in a
'**                 comma delimted format and with an "and" before the last
'**                 name
'** Parameters:     strNames - paramarray of names
'** Returns:        List of names: examples
'**                 FormatNames("Harry") : Harry
'**                 FormatNames("Harry", "Barry") : Harry and Barry
'**                 FormatNames("Harry", "Barry", "Larry") : Harry, Barry and Larry
'** Errors:         None Expected
'********************************************************************************
Private Function FormatNames(ParamArray strNames()) As String

Dim intCount As Integer
Dim intIndex As Integer
Dim strResult As String
Dim strNewNames() As String
    
    'Remove blank names
    strResult = ""
    ReDim strNewNames(UBound(strNames))
    intCount = 0
    For intIndex = 0 To UBound(strNames)
        If strNames(intIndex) <> "" Then
            strNewNames(intCount) = strNames(intIndex)
            intCount = intCount + 1
        End If
    Next
    
    'Join the names togethor
    If intCount > 0 Then
        intCount = intCount - 1
        ReDim Preserve strNewNames(intCount)
        
        strResult = ""
        For intIndex = 0 To intCount
            If strNewNames(intIndex) <> "" Then
                If intIndex = 0 Then
                    strResult = strNewNames(intIndex)
                Else
                    If (intIndex = intCount) And intCount > 0 Then
                        strResult = strResult & " and " & strNewNames(intIndex)
                    Else
                        strResult = strResult & ", " & strNewNames(intIndex)
                    End If
                End If
            End If
        Next
    End If
    
    FormatNames = strResult
    
End Function

'Peter Edney - EP1114 - 05/09/2006
'The original function (AddCalculationDateAttribute) generates the wrong
'calculation date for section 6, but the function is used for sections 10 and 11.
'So a new function was requried just for section 6.
Public Sub AddCalculationDateAttribute2(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
        
    Const cstrFunctionName As String = "AddCalculationDateAttribute2"
    Dim dtValue As Date
    Dim blnDateFound As Boolean
    Dim xmlNode As IXMLDOMNode
    Dim strValue As String
    
    On Error GoTo ErrHandler
    
    'Retrieve the issue date from the generated xml
    Set xmlNode = vxmlNode.ownerDocument.selectSingleNode("/TEMPLATEDATA")
    If Not xmlNode Is Nothing Then
        If xmlAttributeValueExists(xmlNode, "ISSUEDATE") Then
            dtValue = xmlGetAttributeAsDate(xmlNode, "ISSUEDATE")
            dtValue = DateAdd("m", 1, dtValue)
            blnDateFound = True
        End If
    End If
    
    'If there is no issue date, use todays date
    If Not blnDateFound Then
        dtValue = Date
        dtValue = DateAdd("m", 1, dtValue)
        blnDateFound = True
    End If
        
    strValue = Format$(dtValue, "dd/mm/yyyy")
    Call xmlSetAttributeValue(vxmlNode, "CALCULATIONDATE", strValue)

    Set xmlNode = Nothing
    
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Public Function GetLegalFeeText(vobjCommon As CommonDataHelper) As String
    '
    'Add legal fee if applicable
    'If type of mortgage<>'ABO' and type of mortgage <> 'PSW' then
    '   If own solicitors then
    '       {TEXT1}
    '   ElseIf Panel solicitors then
    '       {TEXT2}
    '   End If
    'Else
    '   {HeadingWithNoText}
    'End If
    '
    'EP2_718 Switched these around
    If vobjCommon.IsAdditionalBorrowing Or vobjCommon.IsProductSwitch Then
        'Additional Borrowing or Product Switch
        GetLegalFeeText = "LEGALCOSTS3"
    Else
        'Not additional Borrowing or Product Switch
        If vobjCommon.Data.selectSingleNode(gcstrPANELLEGALREP_PATH) Is Nothing Then
            'No panel solicitor
            GetLegalFeeText = "LEGALCOSTS1"
        Else
            'Panel solicitor
            GetLegalFeeText = "LEGALCOSTS2"
        End If
    End If
    '
End Function

