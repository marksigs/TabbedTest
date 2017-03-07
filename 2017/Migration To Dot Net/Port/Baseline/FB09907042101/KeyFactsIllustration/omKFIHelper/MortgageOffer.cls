VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         MortgageOffer
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Contains the functionality to create the appropriate KFI XML
'**                 for a standard mortgage offer document.
'
'BBG Specific History
'
'Prog   Date        Description
'
'AW     18/01/2005  E2EM00001894 - Additions to BuilsSection3.
'HMA    03/02/2005  E2EM00001894 - Move Buy To Let drawdown from Section 3 to Section 12
'BC     17/03/2006  MAR1425      - Move Conditions Section after Contacts Section
'SAB    03/05/2006  EP479        - Amended BuildSection3 for MEMOPAD element
'SAB    03/05/2006  EP489        - Amended IMortgage_FormatData to only create SECTION14 if it is a Regulated mortgage
'PB     16/05/2006  EP529        - MAR1731 - Fulfilment - Offer Document - oncorrect validity period shown
'PB     09/06/2006  EP651        - MAR1650/MAR1818 etc. Fix bad text in section 3 or offer document
'PB     11/12/2006  EP2_139      - Changes for Epsom phase 2
'PB     12/12/2006  EP2_422      - KFI changes for phase 2
'PB     28/02/2007  EP2_1627     - KFI fixes - missing element for section 3
'INR    05/03/2007  EP2_1667    TOE and Ported product needs adding section3
'INR    20/03/2007  EP2_1977  Deal with REGULATIONINDICATOR_VALIDID
'INR    26/03/2007  EP2_2057    use CheckForValidationType(strRegulationIndicatorValidId, "R") for section 14
'INR    29/03/2007  EP2_1881     Never entered section 3 TOE stuff
'********************************************************************************
Option Explicit

Implements IMortgage

Const mcstrModuleName As String = "MortgageOffer.cls"
Private mobjCommon As CommonDataHelper

Private Sub Class_Terminate()
    Set mobjCommon = Nothing 'TK 18/11/2004 BBG1672
End Sub

'********************************************************************************
'** Function:       IMortgage_FormatData
'** Created by:     Andy Maggs
'** Date:           21/05/2004
'** Description:    The method responsible for formatting the data into the XML
'**                 format required for the standard mortgage offer document
'**                 template.
'** Parameters:     vobjCommon - the object containing the data to format.
'** Returns:        N/A.
'** Errors:         None Expected
'********************************************************************************
Private Sub IMortgage_FormatData(ByVal vobjCommon As CommonDataHelper)
    Const cstrFunctionName As String = "IMortgage_FormatData"
    Dim xmlElem As IXMLDOMElement
    Dim xmlTemplate As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strRegulationIndicatorValidId As String     'SAB 03/05/2006 - EPSOM EP489

    On Error GoTo ErrHandler

    '*-store the common data helper object
    Set mobjCommon = vobjCommon
    
    '*-create the TEMPLATEDATA element and set the attributes
    Set xmlElem = mobjCommon.Document.createElement("TEMPLATEDATA")
    Set xmlTemplate = mobjCommon.Document.appendChild(xmlElem)
    Call BuildTemplateData(xmlTemplate)
    
    '*-create the mandatory AUTONUMBERING element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate)
    Call AddAutonumberingAttributes(xmlNode)
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate) 'SR 24/09/2004 - CORE82
    Call AddAutonumberingAttributes(xmlNode, "NUMBER")  'BC MAR1535

    '*-Add REGULATED element
    'EP2_1977 REGULATIONINDICATOR now holds ValidationId
    If xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "REGULATIONINDICATOR") = "R" Then
        'Regulated
        vobjCommon.CreateNewElement "REGULATED", xmlElem
    End If
    
    '*-Add APPLICANTS element with APPLICANTNAMELIST attribute
    AddFormattedApplicantList ByVal vobjCommon, xmlElem
    
    '*-create the mandatory SECTION1 (About this illustration) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION1", xmlTemplate)
    Call BuildSection1(xmlNode)

    '*-create the mandatory SECTION2 (Which service are we providing you with?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION2", xmlTemplate)
    Call BuildSection2(xmlNode)
     
    '*-create the mandatory SECTION3 (What you have told us) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION3", xmlTemplate)
    Call BuildSection3(xmlNode)

    '*-create the mandatory SECTION4 (Description of this mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION4", xmlTemplate)
    Call BuildSection4(xmlNode)

    '*-create the mandatory SECTION5 (Overall Costs) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION5", xmlTemplate)
    Call BuildSection5(xmlNode)

    '*-create the mandatory SECTION6 (What you will need to pay each month) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION6", xmlTemplate)
    Call BuildSection6(xmlNode)

    '*-create the SECTION6A (What you will need to pay in the future) element
    If mobjCommon.IsMultiComponent Then
        Set xmlNode = mobjCommon.CreateNewElement("SECTION6A", xmlTemplate)
        Call BuildSection6A(xmlNode)
    End If
    
    '*-create the mandatory SECTION7 (Are you comfortable with the risks) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION7", xmlTemplate)
    Call BuildSection7(xmlNode)
    
    '*-create the SECTION7A (Total Borrowing) element
    If mobjCommon.IsAdditionalBorrowing Then
        Set xmlNode = mobjCommon.CreateNewElement("SECTION7A", xmlTemplate)
        Call BuildSection7A(xmlNode)
    End If
    
    '*-create the mandatory SECTION8 (What fees must you pay?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION8", xmlTemplate)
    Call BuildSection8(xmlNode)

    '*-create the mandatory SECTION9 (Insurance) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION9", xmlTemplate)
    Call BuildSection9(xmlNode)

    '*-create the mandatory SECTION10 (What happens if you do not want this mortgage any more?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION10", xmlTemplate)
    Call BuildSection10(xmlNode)

    '*-create the mandatory SECTION11 (What happens if you want to make overpayments?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION11", xmlTemplate)
    Call BuildSection11(xmlNode)

    '*-create the mandatory SECTION12 (Additional Features) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION12", xmlTemplate)
    Call BuildSection12(xmlNode)

    '*-create the mandatory SECTION13 (Using a mortgage intermediary) element
    If vobjCommon.IntroducedByIntermediaryProperty Then  ' SR 23/10/2004 : CORE82
        Set xmlNode = mobjCommon.CreateNewElement("SECTION13", xmlTemplate)
        Call BuildSection13(xmlNode)
    End If ' SR 23/10/2004 : CORE82

    'SAB 03/05/2006 - EPSOM EP489 - START
    '*-create the mandatory SECTION14 (Where can you get more information on mortgages?) element
    'Call mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    'Amended to only display if it is a Regulated mortgage
    'Set xmlNode = mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    'EP2_1977 REGULATIONINDICATOR now holds ValidationId
    strRegulationIndicatorValidId = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "REGULATIONINDICATOR")
    'EP2_2057
    If CheckForValidationType(strRegulationIndicatorValidId, "R") Then     'R=Regulated
        Set xmlNode = mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    End If
    'SAB 03/05/2006 - EPSOM EP489 - END
    
    '*-create the CONTACTSECTION element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("CONTACTSECTION", xmlTemplate)
    Call BuildContactSection(xmlNode)
    
    'BC MAR1425 17/03/2006 Begin
    '*-create the CONDITIONSECTION element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("CONDITIONSECTION", xmlTemplate)
    Call BuildConditionSection(xmlNode)
    'BC MAR1425 17/03/2006 End
    
    'EP2_139
    '*-create the RISKWARNINGSECTION element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("RISKSECTION", xmlTemplate)
    Call BuildRiskSection(xmlNode)
   
    
    Set xmlElem = Nothing
    Set xmlTemplate = Nothing
    Set xmlNode = Nothing
Exit Sub
ErrHandler:
    Set xmlElem = Nothing
    Set xmlTemplate = Nothing
    Set xmlNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildTemplateData
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Builds the template data section of the document.
'** Parameters:     vxmlNode - the node to add the template data to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildTemplateData(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildTemplateData"

    On Error GoTo ErrHandler

    Call BuildCommonOfferTemplateData(mobjCommon, vxmlNode)
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText) 'SR 23/09/2004 : CORE82
    xmlSetAttributeValue vxmlNode, "PROVIDER", mobjCommon.Provider 'SR 24/09/2004 : CORE82
    xmlSetAttributeValue vxmlNode, "TOTALLOANAMOUNT", mobjCommon.TotalLoanAmount 'PB 16/05/2006 EP529 / MAR1731

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

' BC 05/Jan/2004 - E2EM00003121 - Remove COPYTO element
'********************************************************************************
'** Function:       BuildCopyTo
'** Created by:     Andy Maggs
'** Date:           24/05/2004
'** Description:    Builds the CopyTo element containing the addresses of
'**                 applicants to whom the Offer document should be copied to.
'** Parameters:     vxmlNode - the node to add the document details to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
'Private Sub BuildCopyTo(ByVal vxmlNode As IXMLDOMNode)
'    Const cstrFunctionName As String = "BuildCopyTo"
'
'    On Error GoTo ErrHandler
'
'    Call BuildCommonOfferCopyTo(mobjCommon, vxmlNode)
'
'Exit Sub
'ErrHandler:
'    '*-check the error and throw it on
'    errCheckError cstrFunctionName, TypeName(Me)
'End Sub

'********************************************************************************
'** Function:       BuildSection1
'** Created by:     Andy Maggs
'** Date:           25/04/2004
'** Description:    Sets the attributes for the compulsory section1 (About this
'**                 illustration) element.
'** Parameters:     vxmlNode - the section1 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection1(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection1"

    On Error GoTo ErrHandler

    Call BuildCommonOfferSection1(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection2
'** Created by:     Andy Maggs
'** Date:           25/05/2004
'** Description:    Sets the attributes for the compulsory section2 (Which
'**                 service are we providing you with?) element.
'** Parameters:     vxmlNode - the section2 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection2(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection2"

    On Error GoTo ErrHandler

    Call BuildCommonOfferSection2(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection3
'** Created by:     Andy Maggs
'** Date:           25/05/2004
'** Description:    Sets the attributes for the compulsory section3 (What you
'**                 have told us) element.
'** Parameters:     vxmlNode - the section3 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection3"
    Dim xmlPurposeNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim blnValuationUsed As Boolean
    'SR 17/01/2005 : BBG082/E2EM00002927
    Dim lngValue As Long
    Dim rblnIsValuation As Boolean
    'SR 17/01/2005 : BBG082/E2EM00002927 - End
        
    'SAB 26/04/2006 - EPSOM EP479 : START
    Dim strMemoPadText As String
    'SAB 26/04/2006 - EPSOM EP479 : END
    Dim xmlPartandPartNode As IXMLDOMNode 'PB EP651/MAR1590
    'PB 22/11/2006 EP2_139
    Dim objTOE As IXMLDOMNode
    Dim objPurpose As IXMLDOMNode
    Dim strValuationPurchasePrice As String
    Dim strLTV As String
    Dim objValuation As IXMLDOMNode
    Dim strPurchasePrice As String
    Dim strValuation As String
    Dim dblValuation As Double
    Dim xmlComponents As IXMLDOMNodeList
    Dim xmlComponent As IXMLDOMNode
    Dim xmlProduct As IXMLDOMNode
    Dim strProductName As String
    Dim xmlPortedProduct As IXMLDOMNode
    Dim xmlProducts As IXMLDOMNode
    'EP2_1667
    Dim selectRemoveTOE As String
    
    On Error GoTo ErrHandler

    '*-add the ESTIMATEDVALUE attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, False, blnValuationUsed)
    '*-add TERMYEARS and TERMMONTHS attributes
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS")
    '*-add the TOTALLOANAMOUNT attribute
    Call xmlSetAttributeValue(vxmlNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    '*-add the MORTGAGETYPE attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)

    '*-now add one of the mandatory PURPOSEMORTGAGE, PURPOSEREMORTGAGE or
    '*-PURPOSEADDITIONALBORROWING elements to the PURPOSE node
    Select Case mobjCommon.LoanPurposeText
        Case "PURPOSEPURCHASE"
            'EP_1881 If we don't have a TOE
            If Not mobjCommon.IsTransferOfEquity Then
                Set objPurpose = mobjCommon.CreateNewElement("PURPOSEMORTGAGEHEADING", vxmlNode)
            End If
        Case "PURPOSEREMORTGAGE"  'SR 23/09/2004 : CORE82
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEREMORTGAGE", vxmlNode)
        Case "PURPOSEFURTHERADVANCE"
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEADDITIONALBORROWING", vxmlNode)
        Case "PURPOSEPRODUCTSWITCH"
            'PB 28/02/2007 EP2_1627 Begin
            xmlSetAttributeValue vxmlNode, "TOESwitchTerm", "You require the mortgage to continue over a term of "
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEPRODUCTSWITCH", vxmlNode)
            mobjCommon.CreateNewElement "ProdSwitchProducts", vxmlNode
            'PB EP2_1617 End
'PB 21/11/2006 EP2_139 Begin
'EP_1881 Will Never get in here
'        Case "PURPOSETOE"
'            'PB 28/02/2007 EP2_1627
'            xmlSetAttributeValue vxmlNode, "TOESwitchTerm", "You require the mortgage to continue over a term of "
'            Set objPurpose = mobjCommon.CreateNewElement("PURPOSETOE", vxmlNode)
'            'EP2_1667
'            'AddToENames objPurpose
'            selectRemoveTOE = gcstrREMOVEDTOECUSTOMER_PATH & "[contains(@TYPE,'A')]"
'            '*-add the ADDNAMES element and NAMESTOADD attribute if applicable
'            Call mobjCommon.AddCustomerAddedRemovedData(selectRemoveTOE, "ADDNAMES", _
'                    "NAMESTOADD", vxmlNode)
'
'            selectRemoveTOE = gcstrREMOVEDTOECUSTOMER_PATH & "[contains(@TYPE,'R')]"
'            '*-add the REMOVENAMES element and NAMESTOREMOVE attribute if applicable
'            Call mobjCommon.AddCustomerAddedRemovedData(selectRemoveTOE, "REMOVENAMES", _
'                    "NAMESTOREMOVE", vxmlNode)

'PB 21/11/2006 EP2_139 End
    End Select
    'EP_1881 If we have a TOE
    If mobjCommon.IsTransferOfEquity Then
        xmlSetAttributeValue vxmlNode, "TOESwitchTerm", "You require the mortgage to continue over a term of "
        Set objPurpose = mobjCommon.CreateNewElement("PURPOSETOE", vxmlNode)
        'EP2_1667
        'AddToENames objPurpose
        selectRemoveTOE = gcstrREMOVEDTOECUSTOMER_PATH & "[contains(@TYPE,'A')]"
        '*-add the ADDNAMES element and NAMESTOADD attribute if applicable
        Call mobjCommon.AddCustomerAddedRemovedData(selectRemoveTOE, "ADDNAMES", _
                "NAMESTOADD", vxmlNode)

        selectRemoveTOE = gcstrREMOVEDTOECUSTOMER_PATH & "[contains(@TYPE,'R')]"
        '*-add the REMOVENAMES element and NAMESTOREMOVE attribute if applicable
        Call mobjCommon.AddCustomerAddedRemovedData(selectRemoveTOE, "REMOVENAMES", _
                "NAMESTOREMOVE", vxmlNode)
    End If
    
'PB 21/11/2006 EP2_139 Begin
    ' Add <Sec3LoanAm> element
    xmlSetAttributeValue vxmlNode, "Sec3LoanAm", mobjCommon.LoanAmount
    ' Add <Sec3TotalFee1> value
    xmlSetAttributeValue vxmlNode, "Sec3TotalFee1", mobjCommon.FeesAddedToLoanAmount
    '*-add the appropriate (1-16) mandatory FEESANDPREMIUMS element (exactly one)
    Call AddFeesAndPremiumsElement(mobjCommon, vxmlNode)
    '
    'Now get one-off costs, and assign text for KFIFeeStatement1
    
    
'PB 21/11/2006 EP2_139 End

    '*-add the type of mortgage mandatory REPAYMENT, INTERESTONLY or
    '*-PARTANDPART elements
    Call mobjCommon.AddMortgageRepaymentTypeElement(vxmlNode)
        
    'PB EP651/MAR1590 Begin
    If mobjCommon.mstrMortgageRepaymentTypeElement = "PARTANDPART" Then
        Set xmlPartandPartNode = vxmlNode.selectSingleNode("PARTANDPART")
        Call xmlSetAttributeValue(xmlPartandPartNode, "TOTALREPAYMENT", mobjCommon.mlngTotalCapitalandInterest)
        Call xmlSetAttributeValue(xmlPartandPartNode, "TOTALINTERESTONLY", mobjCommon.mlngTotalInterestOnly)
    End If
    'PB EP651/MAR1590 End

    If mobjCommon.LoanPurposeText = "PURPOSEPURCHASE" Then
        Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(vxmlNode)

        '*-add the PURCHASEPRICE attribute
        'E2EM00002842 This will always be purchase price
       Call AddPurchPriceValueAttribute(mobjCommon.Data, xmlPurposeNode)
'        Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, True, _
'                blnValuationUsed)
    End If

    '*-add the optional FIRSTTIMEBUYER element
    If IsFirstTimeBuyer(mobjCommon.Data) Then
        Call mobjCommon.CreateNewElement("FIRSTTIMEBUYER", vxmlNode)
    End If


    'SAB 26/04/2006 - EPSOM EP479 : START
    '*-add the optional MEMOPAD element
    strMemoPadText = ""
    If IsSection3MemoPad(mobjCommon.Data, strMemoPadText) Then
        Set xmlNewNode = mobjCommon.CreateNewElement("MEMOPAD", vxmlNode)
        Call xmlSetAttributeValue(xmlNewNode, "MEMOENTRY", strMemoPadText)
    End If
    'SAB 26/04/2006 - EPSOM EP479 : END




    '*-add the LOANTOVALUE element
    Set xmlNewNode = mobjCommon.CreateNewElement("LOANTOVALUE", vxmlNode)
    'CORE82 Need LoanPurposeElement
    Set xmlNewNode = mobjCommon.AddLoanPurposeElement(xmlNewNode)
'PB 22/11/2006 EP2_139 Begin
    Set xmlPurposeNode = xmlNewNode
'PB 22/11/2006 EP2_139 End
    '*-add the LTV attribute
    'SR 17/01/2005 : BBG082/E2EM00002927
    
'PB 22/11/2006 EP2_139 Begin
    strPurchasePrice = CStr(xmlGetAttributeAsDouble(mobjCommon.Data.selectSingleNode(".//APPLICATIONFACTFIND"), "PURCHASEPRICEORESTIMATEDVALUE"))
    Set objValuation = mobjCommon.Data.selectSingleNode(".//VALNREPVALUATION")
    dblValuation = xmlGetAttributeAsDouble(objValuation, "PRESENTVALUATION")
    If xmlPurposeNode.nodeName = "PURPOSEADDITIONALBORROWING" Then
        strValuationPurchasePrice = "valuation"
    Else
        If dblValuation > CDbl(strPurchasePrice) Then
            strValuationPurchasePrice = "purchase price"
        Else
            strValuationPurchasePrice = "valuation"
        End If
    End If
    strLTV = mobjCommon.LTV
    xmlSetAttributeValue xmlPurposeNode, "VALUATIONPURCHASEPRICE", strValuationPurchasePrice
    xmlSetAttributeValue xmlPurposeNode, "LTV", strLTV
'PB 22/11/2006 EP2_139 End
    
'PB 22/11/2006 EP2_139 (Ported Products) Begin
    If mobjCommon.GetMainMortgageTypeGroup = "NP" Then
        'Ported product
        'EP2_1667
        Set xmlComponents = mobjCommon.Data.selectNodes(".//MORTGAGEPRODUCT[@MORTGAGEPRODUCTPORTABLEIND=1]")
        If xmlComponents.length = 1 Then
            Set xmlPortedProduct = mobjCommon.CreateNewElement("PORTEDPRODUCT", vxmlNode)
            Set xmlProduct = xmlComponents.Item(0).selectSingleNode(".//MORTGAGEPRODUCTLANGUAGE")
            strProductName = xmlGetAttributeText(xmlProduct, "PRODUCTNAME")
            xmlSetAttributeValue xmlPortedProduct, "PRODUCTNAME", strProductName
        Else
            If xmlComponents.length > 1 Then
                Set xmlPortedProduct = mobjCommon.CreateNewElement("PORTEDPRODUCTS", vxmlNode)
            End If
            For Each xmlComponent In xmlComponents
                Set xmlProducts = mobjCommon.CreateNewElement("PRODUCT", xmlPortedProduct)
                Set xmlProduct = xmlComponent.selectSingleNode(".//MORTGAGEPRODUCTLANGUAGE")
                strProductName = xmlGetAttributeText(xmlProduct, "PRODUCTNAME")
                xmlSetAttributeValue xmlProducts, "PRODUCTNAME", strProductName
            Next
        End If
    End If
'PB 22/11/2006 EP2_139 End
            
    If mobjCommon.GetMainMortgageTypeGroup = "N" Then
        Call PurchasePriceOrEstimatedValue(mobjCommon.Data, True, rblnIsValuation, lngValue)
        If rblnIsValuation Then
            Set xmlNewNode = mobjCommon.CreateNewElement("LTVCALCULATIONUSEVALUATION", xmlNewNode)
        Else
            Set xmlNewNode = mobjCommon.CreateNewElement("LTVCALCULATIONUSEPURCHASEPRICE", xmlNewNode)
        End If
    End If
    'SR 17/01/2005 : BBG082/E2EM00002927 - End
    
    'E2EM00001894 Move the optional BUYTOLET element to Section 12
    
    Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))

    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlPartandPartNode = Nothing 'PB EP651/MAR1590
Exit Sub
ErrHandler:
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlPartandPartNode = Nothing 'PB EP651/MAR1590
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection4
'** Created by:     Andy Maggs
'** Date:           26/05/2004
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element.
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection4(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection4"
    
    On Error GoTo ErrHandler
    
'PM 09/06/2006 EP697 Start [Added PB 12/06/2006]
'    Call BuildCommonOfferSection4(mobjCommon, vxmlNode)
'PB 22/11/2006 EP2_139 Begin
    If mobjCommon.LoanComponents.length = 1 Then
        Call BuildCommonOfferSection4Single(mobjCommon, vxmlNode)
    Else
        Call BuildCommonOfferSection4Multiple(mobjCommon, vxmlNode)
    End If
'PB 22/11/2006 EP2_139 End
'PM 09/06/2006 EP697 End

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection5
'** Created by:
'** Date:
'** Description:    Sets the elements and attributes for the compulsory section5
'**                 (Overall Costs) element.
'** Parameters:     vxmlNode - the section5 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection5(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection5"

    On Error GoTo ErrHandler
    
    Call BuildOverallCostsSection(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6
'** Created by:     Andy Maggs
'** Date:           27/05/2004
'** Description:    Sets the elements and attributes for the compulsory section6
'**                 (What you will need to pay each month) element.
'** Parameters:     vxmlNode - the section6 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6"
    
    On Error GoTo ErrHandler
    
    Call BuildCommonOfferSection6(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6A
'** Created by:
'** Date:
'** Description:    Sets the elements and attributes for the compulsory section6A
'**                 (What you will need to pay in the future) element.
'** Parameters:     vxmlNode - the section6A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6A(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6A"

    On Error GoTo ErrHandler
    
    Call BuildCommonOfferSection6A(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7
'** Created by:     Andy Maggs
'** Date:           27/05/2004
'** Description:    Sets the elements and attributes for the compulsory section7
'**                 (Are you comfortable with the risks?) element.
'** Parameters:     vxmlNode - the section7 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7"

    On Error GoTo ErrHandler

    Call BuildCommonOfferSection7(mobjCommon, vxmlNode)
        
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7A
'** Created by:
'** Date:
'** Description:    Sets the elements and attributes for the compulsory section7A
'**                 (Total Borrowing) element.
'** Parameters:     vxmlNode - the section7A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7A(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7A"

    On Error GoTo ErrHandler

    Call BuildCommonOfferSection7A(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection8
'** Created by:     Andy Maggs
'** Date:           01/06/2004
'** Description:    Sets the elements and attributes for the compulsory section8
'**                 (What fees must you pay?) element.
'** Parameters:     vxmlNode - the section8 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection8(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection8"

    On Error GoTo ErrHandler

    Call BuildWhatFeesYouMustPaySection(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection9
'** Created by:
'** Date:
'** Description:    Sets the elements and attributes for the compulsory section9
'**                 (Insurance) element.
'** Parameters:     vxmlNode - the section9 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection9(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection9"

    On Error GoTo ErrHandler

    Call BuildCommonOfferInsuranceSection(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection10
'** Created by:     Andy Maggs
'** Date:           02/06/2004
'** Description:    Sets the elements and attributes for the compulsory section10
'**                 (What happens if you do not want this mortgage any more?) element.
'** Parameters:     vxmlNode - the section10 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection10(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection10"

    On Error GoTo ErrHandler

    Call BuildMortgageNoMoreSection(mobjCommon, vxmlNode, True, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection11
'** Created by:
'** Date:
'** Description:    Sets the elements and attributes for the compulsory section11
'**                 (What happens if you want to make overpayments?) element.
'** Parameters:     vxmlNode - the section11 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection11(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection11"

    On Error GoTo ErrHandler

    Call BuildSection11Common(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection12
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Sets the elements and attributes for the compulsory section11
'**                 (Additional Features) element.
'** Parameters:     vxmlNode - the section12 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection12(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection12"

    On Error GoTo ErrHandler

    Call BuildAdditionalFeaturesSection(mobjCommon, vxmlNode, True, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection13
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Sets the elements and attributes for the compulsory section13
'**                 (Using a mortgage intermediary) element.
'** Parameters:     vxmlNode - the section13 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection13(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection13"

    On Error GoTo ErrHandler

    Call BuildUsingIntermediarySection(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildConditionSection
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Sets the elements and attributes for the compulsory Condition
'**                 Section.
'** Parameters:     vxmlNode - the condition section element to set the elements
'**                 and attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildConditionSection(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildConditionSection"

    On Error GoTo ErrHandler

    Call BuildCommonOfferConditionSection(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildContactSection
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Adds the appropriate contact attributes to the node.
'** Parameters:     vxmlNode - the contact section node to add the attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildContactSection(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildContactSection"

    On Error GoTo ErrHandler

    Call BuildCommonOfferContactSection(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       AddToENames
'** Created by:     Paul Buck
'** Date:           22/11/2006
'** Description:    Adds the Transfer of Equity names to the node.
'** Parameters:     objXML - the xml node to add the attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub AddToENames(objXML As IXMLDOMNode)
    Const cstrFunctionName As String = "AddToENames"
    
    On Error GoTo ErrHandler
    
    
    
    
    
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildRiskSection
'** Created by:     Ian Ross
'** Date:           01/12/2006
'** Description:    Adds the appropriate risk attributes to the node.
'** Parameters:     vxmlNode - the risk section node to add the attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildRiskSection(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildRiskSection"

    On Error GoTo ErrHandler

    Call BuildCommonOfferRiskSection(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       AddFormattedApplicantList
'** Created by:     Paul Buck
'** Date:           11/12/2006
'** Description:    Adds an APPLICANTS element with an APPLICANTNAMESLIST attribute
'** Parameters:     vxmlNode - the node to add the element to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub AddFormattedApplicantList(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddApplicantNameAttribute"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode, xmlCustomerVersion As IXMLDOMNode 'SR 30/09/2004 : BBG1509
    Dim strName As String
    Dim strTitle As String
    Dim strValue As String, strCustomerRoleType As String 'SR 30/09/2004 : BBG1509
    Dim strValidationToSearchFor As String
    Dim intApplicantCount As Integer
    Dim intCount As Integer
    Dim strDelimiter As String
    Dim xmlData As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode

    On Error GoTo ErrHandler
    
    Set xmlData = vobjCommon.Data
    'Check whether a company is mentioned in customer list
    
    Set xmlList = xmlData.selectNodes(gcstrCUSTOMERROLE_PATH & "[contains(@CUSTOMERROLETYPE,'C,')]")
    strValidationToSearchFor = IIf(xmlList.length > 0, "C,", "A,")
    
    Set xmlList = xmlData.selectNodes(gcstrCUSTOMERROLE_PATH)
    '
    intApplicantCount = xmlList.length
    For Each xmlItem In xmlList
        intCount = intCount + 1
        strCustomerRoleType = xmlGetAttributeText(xmlItem, "CUSTOMERROLETYPE")
        If InStr(UCase$(strCustomerRoleType), strValidationToSearchFor) > 0 Then
            Set xmlCustomerVersion = xmlItem.selectSingleNode("./CUSTOMERVERSION")
            strTitle = xmlGetAttributeText(xmlCustomerVersion, "TITLEOTHER")
            If strTitle = "" Then
                strTitle = xmlGetAttributeText(xmlCustomerVersion, "TITLE_TEXT")
            End If
            strName = BuildName(strTitle, _
                                GetAllForeNames(xmlCustomerVersion), _
                                xmlGetAttributeText(xmlCustomerVersion, "SURNAME"))
            'Add delimiter
            If intCount < intApplicantCount Then
                If intCount + 1 = intApplicantCount Then
                    strDelimiter = " and "
                Else
                    strDelimiter = ", "
                End If
            Else
                strDelimiter = ""
            End If
            strValue = strValue & strName
        End If
    Next xmlItem
    '
    Set xmlNewNode = vobjCommon.CreateNewElement("APPLICANTS", vxmlNode)
    xmlSetAttributeValue xmlNewNode, "APPLICANTNAMELIST", strValue
    Call xmlSetAttributeValue(vxmlNode, "APPLICANTNAMES", strValue)
    Set xmlList = Nothing
    Set xmlItem = Nothing

Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


