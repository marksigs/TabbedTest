VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StandardMortgage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         StandardMortgage
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Contains the functionality to create the appropriate KFI XML
'**                 for a standard mortgage illustration.
'Prog   Date        Description
'
'BC     05/10/2005  MAR88 - MARS Project.
'BC     20/12/2005  MAR907 - Processing of Transfer of Equity (ToE) KFI
'BC     20/01/2006  MAR1091- Set REFERENCE attribute inside NONDISPOSABLE element
'BC     02/03/2006  MAR1347  - se first Validation Type of RepaymentMethod, rather than the second
'PB     06/06/2006  EP651/MAR1590 - modified BuildSection6 - show two rows if a loan is Part & Part
'INR    07/02/2007  EP2_583 KFI changes
'INR    07/02/2007  EP2_1450 added repayment vehicles .
'INR    01/03/2007  EP2_1449 Use FirmName rather than ContactName for Section6 Intermediary
'PB     02/03/2007  EP2_1627 Add missing text for TOE and PSW
'INR    06/03/2007  EP2_1667 Use REMOVEDTOECUSTOMER for determining both Add and Remove TOE Customers
'PB     09/03/2007  EP2_1627 Added text to section 3
'INR    09/03/2007  EP2_1882 removed "PRODSWITCHPRODUCTS" from BuildStandardSection3 duplication of PURPOSEPRODUCTSWITCH
'PB     13/03/2007  EP2_1930 Added check so that repay veh text only appears if there is a repay veh.
'INR    20/03/2007  EP2_1977  Deal with REGULATIONINDICATOR_VALIDID
'INR    21/03/2007  EP2_1982  Sect3 ValPurchase text
'PB     21/03/2007  EP2_1861 Added section 2 for TOE and PSW applications
'PB     23/03/2007  EP2_1952 Missing text from 'What you have told us' section - added default PURPOSE
'INR    26/03/2007  EP2_1983/1984 Call Offer section4 code as its the same
'INR    26/03/2007  EP2_2057    use CheckForValidationType(strRegulationIndicatorValidId, "R") for section 14
'INR    29/03/2007  EP2_2103 PURPOSEFURTHERADVANCE node will already exist, need LTV Nodes for TOE
'INR    29/03/2007  EP2_1881 PURPOSETOE required
'INR    01/04/2007  EP2_2042 New repayment types for buildsection6
'INR    04/04/2007  EP2_1952 If Additional Borrowing & TOE treat as just Additional Borrowing
'TW     09/04/2007  EP2_1743 - kfi - not appearing correctly in a case with drawdown
'SR     11/04/2007  EP2_2270 modified BuildStdSection3
'INR    15/04/2007  EP2_2395 Call new AddInterestRateTypeAttributeSect6
'********************************************************************************
Option Explicit

Implements IMortgage

Private mobjCommon As CommonDataHelper

Private Sub Class_Terminate()
    Set mobjCommon = Nothing 'TK 18/11/2004 BBG1672
End Sub

'********************************************************************************
'** Function:       IMortgage_FormatData
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    The method responsible for formatting the data into the XML
'**                 format required for the standard mortgage KFI document
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
    Dim blnIsIntermediary As Boolean
    Dim strContactName As String
    Dim intAddedCustomers As Integer
    Dim intRemovedCustomers As Integer
    Dim strAdditionOrRemoval As String
    'EP2_2057
    Dim strRegulationIndicatorValidId As String

    On Error GoTo ErrHandler

    '*-store the common data helper object
    Set mobjCommon = vobjCommon
    
    '*-create the TEMPLATEDATA element and set the attributes
    Set xmlElem = mobjCommon.Document.createElement("TEMPLATEDATA")
    Set xmlTemplate = mobjCommon.Document.appendChild(xmlElem)
    Call BuildTemplateData(xmlTemplate)
    
    '*-create the mandatory TIMELIMIT element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("TIMELIMIT", xmlTemplate)
    '*-add the mandatory COMPLETEDBYDATE attribute
    Call AddCompletedByDateAttribute(mobjCommon, xmlNode)
    
    '*-create the mandatory AUTONUMBERING element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate)
    Call AddAutonumberingAttributes(xmlNode)
    
    '*-create the mandatory SECTION1 (About this illustration) element
    '*-NB there are no attributes on this element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION1", xmlTemplate)
    Call BuildSection1(xmlNode)

    '*-create the mandatory SECTION2 (Which service are we providing you with?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION2", xmlTemplate)
    'EP2_583 if not PRODUCTSWITCH and not Transfer of Equity
    
    'EP2_1952 If Additional Borrowing & TOE treat as just Additional Borrowing
    If vobjCommon.IsTransferOfEquity And Not vobjCommon.IsAdditionalBorrowing() Then
        
        gblnSection2 = True
        
        intAddedCustomers = vobjCommon.Data.selectNodes("//APPLICATION/REMOVEDTOECUSTOMER[@TYPE='A']").length
        intRemovedCustomers = vobjCommon.Data.selectNodes("//APPLICATION/REMOVEDTOECUSTOMER[@TYPE='R']").length
        
        strAdditionOrRemoval = "removal"
        If intAddedCustomers > 0 Then
            If intRemovedCustomers > 0 Then
                strAdditionOrRemoval = "addition and removal"
            Else
                strAdditionOrRemoval = "addition"
            End If
        End If
        
        Set xmlNode = vobjCommon.CreateNewElement("TOEORPSW", xmlNode)
        Set xmlNode = vobjCommon.CreateNewElement("TOE", xmlNode)
        xmlSetAttributeValue xmlNode, "ADDITIONREMOVAL", strAdditionOrRemoval
        'EP2_1881
'        xmlSetAttributeValue xmlNode, "ADDITIONREMOVALPARTIES", CStr(intAddedCustomers + intRemovedCustomers)
        'EP2_1952
        If (intAddedCustomers + intRemovedCustomers) > 1 Then
            xmlSetAttributeValue xmlNode, "ADDITIONREMOVALPARTIES", CStr(intAddedCustomers + intRemovedCustomers) & " parties"
        Else
            xmlSetAttributeValue xmlNode, "ADDITIONREMOVALPARTIES", "a party"
        End If
    
    ElseIf mobjCommon.IsProductSwitch Then
        
        gblnSection2 = True
        
        Set xmlNode = vobjCommon.CreateNewElement("TOEORPSW", xmlNode)
        Set xmlNode = vobjCommon.CreateNewElement("PSW", xmlNode)
        
    Else
        'If (Not (vobjCommon.blnIsProductSwitch) And Not (vobjCommon.blnIsTransferofEquity)) Then
        Call BuildStandardSection2(xmlNode)
    End If
    'BC MAR907 Begin
'    If (vobjCommon.blnIsTransferofEquity) Then
'        Set xmlNode = mobjCommon.CreateNewElement("TOEKFI", xmlNode)
'        Call BuildToESection2(xmlNode)
'    ElseIf (vobjCommon.blnIsProductSwitch) Then
'        Set xmlNode = mobjCommon.CreateNewElement("PSWKFI", xmlNode)
'        Call BuildStandardSection2(xmlNode)
'    Else
'        Set xmlNode = mobjCommon.CreateNewElement("STANDARDKFI", xmlNode)
'        Call BuildStandardSection2(xmlNode)
'    End If
    'BC MAR907 End
    
    '*-create the mandatory SECTION3 (What you have told us) element
    
    
    'BC MAR907 Begin
    If Not (vobjCommon.blnIsTransferofEquity) Then
'EP2_583        Set xmlNode = mobjCommon.CreateNewElement("STANDARDKFI", xmlTemplate)
        Set xmlNode = mobjCommon.CreateNewElement("SECTION3", xmlTemplate)
        'Call BuildStandardSection3(xmlNode)
        Call BuildStdSection3(xmlNode)

    Else
'EP2_583        Set xmlNode = mobjCommon.CreateNewElement("TOEKFI", xmlTemplate)
        Set xmlNode = mobjCommon.CreateNewElement("SECTION3", xmlTemplate)
        Call BuildToESection3(xmlNode)
    End If
    'BC MAR907 End

    '*-create the mandatory SECTION4 (Description of this mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION4", xmlTemplate)
    Call BuildSection4(xmlNode)

    '*-create the mandatory SECTION5 (Overall Costs) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION5", xmlTemplate)
    Call BuildSection5(xmlNode)

    '*-create the mandatory SECTION6 (What you will need to pay each month) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION6", xmlTemplate)
    Call BuildSection6(xmlNode)

    '*-create the mandatory SECTION6A (What you will need to pay in the future) element
    If mobjCommon.IsMultiComponent Then
        Set xmlNode = mobjCommon.CreateNewElement("SECTION6A", xmlTemplate)
        'SR EP2_2159
        'Call BuildSection6A(xmlNode)
        Call BuildCommonOfferSection6A(vobjCommon, xmlNode)
        'SR EP2_2159 - End
    End If
    
    '*-create the mandatory SECTION7 (Are you comfortable with the risks) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION7", xmlTemplate)
    Call BuildSection7(xmlNode)
    
    '*-create the mandatory SECTION7A (Total Borrowing) element
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
    blnIsIntermediary = vobjCommon.IsIntroducedByIntermediary(strContactName, False)
    If blnIsIntermediary Then
        Set xmlNode = mobjCommon.CreateNewElement("SECTION13", xmlTemplate)
        Call BuildSection13(xmlNode)
    End If

    'EP2_2057
    '*-If this a regulated mortgage create section14
    strRegulationIndicatorValidId = xmlGetAttributeText(mobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "REGULATIONINDICATOR")
    If CheckForValidationType(strRegulationIndicatorValidId, "R") Then     'R=Regulated
        Set xmlNode = mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    End If
'    Call BuildSection14(xmlNode)
    
    'EP2_583
    '*-create the CONTACTSECTION element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("CONTACTSECTION", xmlTemplate)
    Call BuildContactSection(xmlNode)
    
    'EP2_583
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

    Dim xmlNode As IXMLDOMNode 'BC MAR1091
    Dim xmlAFF As IXMLDOMNode
    Dim strRegInd As String
    
    On Error GoTo ErrHandler

    '*-add the mandatory PROVIDER attribute
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", mobjCommon.Provider)
    '*-add the mandatory VALIDTODATE attribute
    Call AddValidToDateAttribute(mobjCommon.Data, vxmlNode)
    '*-add the mandatory APPLICANTNAME attribute
    Call AddApplicantNameAttribute(mobjCommon.Data, vxmlNode)
    '*-add the mandatory REFERENCE attribute
    Set xmlNode = mobjCommon.CreateNewElement("NONDISPOSABLE", vxmlNode)
    Call AddReferenceAttribute(mobjCommon.Data, vxmlNode)
    '*-add the mandatory TODAY attribute
    Call AddTodayAttribute(vxmlNode)
    '*-add the mandatory MORTGAGETYPE  attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)
    '*-add the DOCUMENTNAME attribute
    Call xmlSetAttributeValue(vxmlNode, "DOCUMENTNAME", "Key Facts Illustration")
    'EP2_583
    '*-If this a regulated mortgage create the REGULATED element for key facts logo
    Set xmlAFF = mobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlAFF Is Nothing Then
        'EP2_1977 REGULATIONINDICATOR now holds ValidationId
        strRegInd = xmlGetAttributeText(xmlAFF, "REGULATIONINDICATOR")
        If CheckForValidationType(strRegInd, "R") Then
            Call mobjCommon.CreateNewElement("REGULATED", vxmlNode)
        End If
    End If


    Set xmlNode = Nothing 'BC MAR1091
    Set xmlAFF = Nothing
    
Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    Set xmlAFF = Nothing

    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection1
'** Created by:     Andy Maggs
'** Date:           30/03/2004
'** Description:    Sets the attributes for the compulsory section1 (About this
'**                 illustration) element.
'** Parameters:     vxmlNode - the section1 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection1(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection1"

    On Error GoTo ErrHandler

    'EP2_583 Need to know if its regulated or not change to common call
    Call BuildCommonOfferSection1(mobjCommon, vxmlNode)


Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildStandardSection2
'** Created by:     Andy Maggs
'** Date:           25/03/2004
'** Description:    Sets the attributes for Section2 - Standard KFI
'** Parameters:     vxmlNode - the section2 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildStandardSection2(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildStandardSection2"

    On Error GoTo ErrHandler

    Call AddServicesProvided(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'********************************************************************************
'** Function:       BuildToESection2
'** Created by:     Bill Coates
'** Date:           20/121/2005
'** Description:    Sets the attributes for Section2 - Transfer of Equity KFI
'** Parameters:     vxmlNode - the section1 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildToESection2(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildToESection2"
    Dim xmlAdded As IXMLDOMNodeList
    Dim xmlRemoved As IXMLDOMNodeList
    Dim iAdded As Integer
    Dim iRemoved As Integer
    
    On Error GoTo ErrHandler
    
    '*-NB BEFORE THIS SECTION WILL WORK CORRECTLY, THE DATABASE SCHEMA WILL NEED TO
    '*-UPDATED WITH THE NEW TABLES CUSTOMERADDED AND CUSTOMERREMOVED AND THE
    '*-omSchema.xml IS UPDATED WITH THE FOLLOWING (UNDER THE APPLICATIONFACTFIND NODE):
    '<CUSTOMERADDED DATASRCE="CUSTOMERADDED">
    '   <APPLICATIONNUMBER DATATYPE="STRING" KEYTYPE="PRIMARY"/>
    '   <APPLICATIONFACTFINDNUMBER DATATYPE="SHORT" KEYTYPE="PRIMARY"/>
    '   <FORENAME DATATYPE="STRING"/>
    '   <SURNAME DATATYPE="STRING"/>
    '   <TITLE DATATYPE="COMBO" COMBOGROUP="Title"/>
    '</CUSTOMERADDED>
    '
    '<REMOVEDTOECUSTOMER DATASRCE="REMOVEDTOECUSTOMER">
    '   <APPLICATIONNUMBER DATATYPE="STRING" KEYTYPE="PRIMARY"/>
    '   <OMIGACUSTOMERNUMBER DATATYPE="STRING" KEYTYPE="PRIMARY"/>
    '   <FIRSTFORENAME DATATYPE="STRING"/>
    '   <SURNAME DATATYPE="STRING"/>
    '   <TITLE DATATYPE="COMBO" COMBOGROUP="Title"/>
    '</REMOVEDTOECUSTOMER>
    
    Set xmlAdded = mobjCommon.Data.selectNodes(gcstrCUSTOMERVERSION_PATH & "[contains(@NEWTOECUSTOMERIND,'1')]")
    If xmlAdded Is Nothing Then
        iAdded = 0
    Else
        iAdded = xmlAdded.length
    End If
    
    Set xmlRemoved = mobjCommon.Data.selectNodes(gcstrREMOVEDTOECUSTOMER_PATH)
    If xmlRemoved Is Nothing Then
        iRemoved = 0
    Else
        iRemoved = xmlRemoved.length
    End If
    
    '*-create one of ADDITION, REMOVAL or BOTH element as appropriate
    If iAdded > 0 And iRemoved > 0 Then
        Call mobjCommon.CreateNewElement("BOTH", vxmlNode)
    ElseIf iAdded = 1 Then
        Call mobjCommon.CreateNewElement("ADDITION", vxmlNode)
    ElseIf iRemoved = 1 Then
        Call mobjCommon.CreateNewElement("REMOVAL", vxmlNode)
    End If
    
    '*-create the PARTY or PARTIES element as appropriate
    If iAdded + iRemoved = 1 Then
        Call mobjCommon.CreateNewElement("PARTY", vxmlNode)
    ElseIf iAdded + iRemoved > 1 Then
        Call mobjCommon.CreateNewElement("PARTIES", vxmlNode)
    End If
    
    Set xmlAdded = Nothing
    Set xmlRemoved = Nothing
Exit Sub
ErrHandler:
    Set xmlAdded = Nothing
    Set xmlRemoved = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildStandardSection3
'** Created by:     Andy Maggs
'** Date:           25/03/2004
'** Description:    Sets the attributes for Section3 - Standard KFIWhat you
'** Parameters:     vxmlNode - the section3 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildStandardSection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildStandardSection3"
    Dim xmlMainPurposeNode As IXMLDOMNode
    Dim xmlPurposeNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim blnValuationUsed As Boolean
    'PB 06/06/06 EP351/MAR1590
    Dim xmlPartandPartNode As IXMLDOMNode
    'PB EP351/MAR1590 End

    On Error GoTo ErrHandler

    '* Add TERMYEARS and TERMMONTHS as attribs to SECTION3
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS")

    '*-create the mandatory PURPOSE element
    'Set xmlMainPurposeNode = mobjCommon.CreateNewElement("SECTION3", vxmlNode)
    'INR CORE82
    Call xmlSetAttributeValue(vxmlNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    
    '*-now add one of the mandatory PURPOSEPURCHASE, PURPOSEREMORTGAGE or
    '*-PURPOSEFURTHERADVANCE elements to the PURPOSE node
    Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(vxmlNode)
    
    'BC MAR907
    Call xmlSetAttributeValue(xmlPurposeNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    
    '*-add the appropriate (1-16) mandatory FEESANDPREMIUMS element (exactly one)
    Call AddFeesAndPremiumsElement(mobjCommon, vxmlNode)
    
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
    
    '*-now add one of the mandatory PURPOSEPURCHASE, PURPOSEREMORTGAGE or
    '*-PURPOSEFURTHERADVANCE elements to the SECTION3 NODE
    Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(vxmlNode)
    Select Case xmlPurposeNode.baseName
        Case "PURPOSEPURCHASE"
            '*-add the mandatory PURCHASEPRICE attribute
            'E2EM00002842 This will always be purchase price
            Call AddPurchPriceValueAttribute(mobjCommon.Data, xmlPurposeNode)
            'EP2_583
            Set xmlMainPurposeNode = mobjCommon.CreateNewElement("PURPOSEMORTGAGE", vxmlNode)
        
        Case "PURPOSEREMORTGAGE"
            '*-add the mandatory INSUREDVALUE attribute - set to 0 as per the
            '*-mapping document
            Call xmlSetAttributeValue(xmlPurposeNode, "INSUREDVALUE", CStr(0))
            
            '*-add the mandatory ESTIMATEDVALUE attribute
            Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, False, _
                    blnValuationUsed)
             'EP2_583
            Set xmlMainPurposeNode = mobjCommon.CreateNewElement("PURPOSEREMORTGAGE", vxmlNode)
       
        Case "PURPOSEFURTHERADVANCE"
            '*-add the mandatory ESTIMATEDVALUE attribute
            Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, False, _
                    blnValuationUsed)
            'EP2_583
            Set xmlMainPurposeNode = mobjCommon.CreateNewElement("PURPOSEADDITIONALBORROWING", vxmlNode)

                    
        Case "PURPOSEPRODUCTSWITCH"
            '*-add the mandatory ESTIMATEDVALUE attribute
            Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, False, _
                    blnValuationUsed)
            'EP2_583
            Set xmlMainPurposeNode = mobjCommon.CreateNewElement("PURPOSEPRODUCTSWITCH", vxmlNode)

    
    End Select
    
    '*-add the optional FIRSTTIMEBUYER element
    If IsFirstTimeBuyer(mobjCommon.Data) Then
        Call mobjCommon.CreateNewElement("FIRSTTIMEBUYER", vxmlNode)
    End If

    
    'CORE82 Creating this on the wrong node
    '*-add the LOANTOVALUE element
    Set xmlNewNode = mobjCommon.CreateNewElement("LOANTOVALUE", xmlPurposeNode)
    '*-add the LTV attribute
    Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))
    If blnValuationUsed Then
        '*-add the ESTIMATEDVALUATION element
        Call mobjCommon.CreateNewElement("ESTIMATEDVALUATION", xmlNewNode)
    Else
        '*-add the PROPERTYPURCHASEPRICE element
        Call mobjCommon.CreateNewElement("PROPERTYPURCHASEPRICE", xmlNewNode)
    End If
    
    '*-add the FURTHERADVANCELOANTOVALUE element
    If mobjCommon.IsAdditionalBorrowing Then
        Set xmlNewNode = mobjCommon.CreateNewElement("FURTHERADVANCELOANTOVALUE", vxmlNode)
        '*-add the LTV attribute
        Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))
    End If
    
    Set xmlMainPurposeNode = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    'PB 06/06/06 EP651/MAR1590
    Set xmlPartandPartNode = Nothing
    'PB EP651/MAR1590 End
Exit Sub
ErrHandler:
    Set xmlMainPurposeNode = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    'PB 06/06/06 EP651/MAR1590
    Set xmlPartandPartNode = Nothing
    'PB EP651/MAR1590 End
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub BuildStdSection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildStdSection3"
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
    'EP2_1982
    Dim xmlItem As IXMLDOMNode
    Dim typeOfAppValue As String
    Dim sec3ValPurchText As String
    
    
    On Error GoTo ErrHandler

    'PB 01/03/2007 EP2_1627 Begin
    'If TOE or PSW add the relevent text:
    If mobjCommon.IsProductSwitch Or mobjCommon.IsTransferOfEquity Then
        xmlSetAttributeValue vxmlNode, "TOESWITCHTERM", "You require the mortgage to continue over a term of "
    End If
    'PB 01/03/2007 EP2_1627 End
    
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
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEMORTGAGE", vxmlNode)
        Case "PURPOSEREMORTGAGE"  'SR 23/09/2004 : CORE82
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEREMORTGAGE", vxmlNode)
        Case "PURPOSEFURTHERADVANCE"
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEADDITIONALBORROWING", vxmlNode)
        Case "PURPOSEPRODUCTSWITCH"
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEPRODUCTSWITCH", vxmlNode)
'PB 21/11/2006 EP2_139 Begin
        Case "PURPOSETOE"
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSETOE", vxmlNode)
'            AddToENames objPurpose
        Case Else
            Set objPurpose = mobjCommon.CreateNewElement("PURPOSEMORTGAGE", vxmlNode)
'PB 21/11/2006 EP2_139 End
    End Select
    
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

' TW 09/04/2007 EP2_1743 - Temporary fix
    If mobjCommon.DrawDownAmount > 0 Then
        Set xmlNewNode = mobjCommon.CreateNewElement("DRAWDOWN", vxmlNode)
        Call xmlSetAttributeValue(xmlNewNode, "DRAWDOWN", CStr(mobjCommon.DrawDownAmount))
        If mobjCommon.GetMainMortgageTypeGroup = "F" Then
            'Additional borrowing
            xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE1", "additional borrowing"
        Else
            xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE1", "mortgage"
        End If
        Call xmlSetAttributeValue(xmlNewNode, "MONTHLYCOSTLESSDRAWDOWN", CStr(mobjCommon.MonthlyCostLessDrawDown))
    
    End If
' TW 09/04/2007 EP2_1743 End

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
    If Not objValuation Is Nothing Then  ' SR EP2_2270 - valuation data might not exist for KFI
        dblValuation = xmlGetAttributeAsDouble(objValuation, "PRESENTVALUATION")
    End If
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
        'EP2_1667 get data from correct place and sort typos
        Set xmlComponents = mobjCommon.Data.selectNodes(".//MORTGAGEPRODUCT")
        For Each xmlComponent In xmlComponents
            Set xmlPortedProduct = mobjCommon.CreateNewElement("PORTEDPRODUCT", vxmlNode)
            Set xmlProduct = xmlComponent.selectSingleNode(".//MORTGAGEPRODUCTLANGUAGE")
            strProductName = xmlGetAttributeText(xmlProduct, "PRODUCTNAME")
            xmlSetAttributeValue xmlPortedProduct, "PRODUCTNAME", strProductName
        Next
    End If
'PB 22/11/2006 EP2_139 End

    'EP2_1982
    Set xmlItem = mobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        '*-get the purpose of the loan
        typeOfAppValue = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
    End If
    If CheckForValidationType(typeOfAppValue, "R") _
        Or CheckForValidationType(typeOfAppValue, "TOE") _
        Or CheckForValidationType(typeOfAppValue, "PSW") _
        Or CheckForValidationType(typeOfAppValue, "RTB") Then
        sec3ValPurchText = "estimated value"
    Else
        sec3ValPurchText = "purchase price"
    End If

    If CheckForValidationType(typeOfAppValue, "F") Then
        'EP2_2103 PURPOSEFURTHERADVANCE node will already exist
        Call xmlSetAttributeValue(xmlNewNode, "SEC3VALPURCH", sec3ValPurchText)
    Else
        Set xmlNewNode = mobjCommon.CreateNewElement("LTVCALCULATIONUSEPURCHASEPRICE", xmlNewNode)
        Call xmlSetAttributeValue(xmlNewNode, "SEC3VALPURCH", sec3ValPurchText)
    End If
    'SR 17/01/2005 : BBG082/E2EM00002927 - End
    
    'E2EM00001894 Move the optional BUYTOLET element to Section 12
    
    Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))

    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlPartandPartNode = Nothing 'PB EP651/MAR1590
    Set xmlItem = Nothing
    Set xmlPartandPartNode = Nothing
    Set objTOE = Nothing
    Set objPurpose = Nothing
    Set objValuation = Nothing
    Set xmlComponents = Nothing
    Set xmlComponent = Nothing
    Set xmlProduct = Nothing
    Set xmlPortedProduct = Nothing
Exit Sub
ErrHandler:
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlPartandPartNode = Nothing 'PB EP651/MAR1590
    Set xmlItem = Nothing
    Set xmlPartandPartNode = Nothing
    Set objTOE = Nothing
    Set objPurpose = Nothing
    Set objValuation = Nothing
    Set xmlComponents = Nothing
    Set xmlComponent = Nothing
    Set xmlProduct = Nothing
    Set xmlPortedProduct = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildToESection3
'** Created by:     Andy Maggs
'** Date:           25/05/2004
'** Description:    Sets the attributes for Section3 - Transfer of Equity KFI
'** Parameters:     vxmlNode - the section3 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildToESection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildToESection3"
    Dim xmlPurposeNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim blnValuationUsed As Boolean
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim blnIsValuation As Boolean
    'PB 06/06/06 EP651/MAR1590
    Dim xmlPartandPartNode As IXMLDOMNode
    'PB EP651/MAR1590 End
    'EP2_1667
    Dim selectRemoveTOE As String
    'EP2_2103
    Dim typeOfAppValue As String
    Dim sec3ValPurchText As String
    
    On Error GoTo ErrHandler
    
    'EP2_1881
    'EP2_1952 If Additional Borrowing & TOE treat as just Additional Borrowing
    If mobjCommon.IsTransferOfEquity And Not mobjCommon.IsAdditionalBorrowing() Then
        Set xmlNewNode = mobjCommon.CreateNewElement("PURPOSETOE", vxmlNode)
    End If

    'PB 01/03/2007 EP2_1627 Begin
    If mobjCommon.IsProductSwitch Or mobjCommon.IsTransferOfEquity Then
        xmlSetAttributeValue vxmlNode, "TOESWITCHTERM", "You require the mortgage to continue over a term of "
        If mobjCommon.IsProductSwitch Then
            mobjCommon.CreateNewElement "PRODSWITCHPRODUCTS", vxmlNode
        End If
    End If
    'PB 01/03/2007 EP2_1627 End
    
    'PB 29/03/2007 EP2_1952 begin
    '-*add purpose node
    If mobjCommon.IsAdditionalBorrowing Then
        mobjCommon.CreateNewElement "PURPOSEADDITIONALBORROWING", vxmlNode
    Else
        'EP2_1952 Don't add if TOE
        If Not mobjCommon.IsTransferOfEquity Then
            mobjCommon.CreateNewElement "PURPOSEMORTGAGE", vxmlNode
        End If
    End If
    'PB EP2_1952 End
    
    '*-add the ESTIMATEDVALUE attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, False, blnValuationUsed)
    '*-add TERMYEARS and TERMMONTHS attributes
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS")
    '*-add the TOTALLOANAMOUNT attribute
    Call xmlSetAttributeValue(vxmlNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    '*-add the MORTGAGETYPE attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)
    
    '*-add the PROPERTYVALUATIONAMOUNT attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, _
            Not mobjCommon.IsAdditionalBorrowing, blnIsValuation, _
            "PROPERTYVALUATIONAMOUNT")
    
    'EP2_2103
    '*-add the LOANTOVALUE element
    Set xmlNewNode = mobjCommon.CreateNewElement("LOANTOVALUE", vxmlNode)

    Set xmlItem = mobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    If Not xmlItem Is Nothing Then
        '*-get the purpose of the loan
        typeOfAppValue = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
    End If
    If CheckForValidationType(typeOfAppValue, "R") _
        Or CheckForValidationType(typeOfAppValue, "TOE") _
        Or CheckForValidationType(typeOfAppValue, "PSW") _
        Or CheckForValidationType(typeOfAppValue, "RTB") Then
        sec3ValPurchText = "estimated value"
    Else
        sec3ValPurchText = "purchase price"
    End If

    If CheckForValidationType(typeOfAppValue, "F") Then
        Set xmlNewNode = mobjCommon.CreateNewElement("PURCHASEFURTHERADVANCE", xmlNewNode)
        Call xmlSetAttributeValue(xmlNewNode, "SEC3VALPURCH", sec3ValPurchText)
        '*-add the LTV attribute
        Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))
    Else
        Set xmlNewNode = mobjCommon.CreateNewElement("PURPOSEPURCHASE", xmlNewNode)
        Set xmlNewNode = mobjCommon.CreateNewElement("LTVCALCULATIONUSEPURCHASEPRICE", xmlNewNode)
        Call xmlSetAttributeValue(xmlNewNode, "SEC3VALPURCH", sec3ValPurchText)
        '*-add the LTV attribute
        Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))
    End If
    
'    If mobjCommon.IsAdditionalBorrowing Then
'        '*-now add the ADDITIONALBORROWING element to the PURPOSE node
'        Call mobjCommon.CreateNewElement("ADDITIONALBORROWING", vxmlNode)
'    End If
    
    '*-add the appropriate (1-16) mandatory FEESANDPREMIUMS element (exactly one)
    Call AddFeesAndPremiumsElement(mobjCommon, vxmlNode)
    
    '******************************
    '*-NB ADDNAMES AND REMOVENAMES WILL NOT WORK CORRECTLY UNTIL THE NEW
    '*-CUSTOMERADDED AND CUSTOMERREMOVED TABLES ARE ADDED TO THE DATABASE AND THE
    '*-APPROPRIATE SCHEMA IS ADDED TO omSchema.xml (SEE SECTION 2 FOR DETAILS)
    'EP2_1667
    selectRemoveTOE = gcstrREMOVEDTOECUSTOMER_PATH & "[contains(@TYPE,'A')]"
    '*-add the ADDNAMES element and NAMESTOADD attribute if applicable
    Call mobjCommon.AddCustomerAddedRemovedData(selectRemoveTOE, "ADDNAMES", _
            "NAMESTOADD", vxmlNode)
    'EP2_1667
    selectRemoveTOE = gcstrREMOVEDTOECUSTOMER_PATH & "[contains(@TYPE,'R')]"
    '*-add the REMOVENAMES element and NAMESTOREMOVE attribute if applicable
    Call mobjCommon.AddCustomerAddedRemovedData(selectRemoveTOE, "REMOVENAMES", _
            "NAMESTOREMOVE", vxmlNode)
    '******************************
    
    '*-add the type of mortgage mandatory REPAYMENT, INTERESTONLY or
    '*-PARTANDPART elements
    Call mobjCommon.AddMortgageRepaymentTypeElement(vxmlNode)
    
    'PB 06/06/06 EP651/MAR1590 Begin
    If mobjCommon.mstrMortgageRepaymentTypeElement = "PARTANDPART" Then
        Set xmlPartandPartNode = vxmlNode.selectSingleNode("PARTANDPART")
        Call xmlSetAttributeValue(xmlPartandPartNode, "TOTALREPAYMENT", mobjCommon.mlngTotalCapitalandInterest)
        Call xmlSetAttributeValue(xmlPartandPartNode, "TOTALINTERESTONLY", mobjCommon.mlngTotalInterestOnly)
    End If
    'PB EP651/MAR1590 End
    
     If mobjCommon.IsAdditionalBorrowing Then
        Set xmlPurposeNode = mobjCommon.CreateNewElement("PURPOSEFURTHERADVANCE", _
                vxmlNode)
        '*-add the ADDITIONALBORROWINGAMOUNT attribute
        Call xmlSetAttributeValue(xmlPurposeNode, "ADDITIONALBORROWINGAMOUNT", _
                CStr(mobjCommon.LoanAmount))
    End If
    
    '*-add the LOANTOVALUE element
    Set xmlNewNode = mobjCommon.CreateNewElement("FURTHERADVANCELOANTOVALUE", vxmlNode)
    '*-add the LTV attribute
    Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))
    
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlList = Nothing
    Set xmlItem = Nothing
Exit Sub
ErrHandler:
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlList = Nothing
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'********************************************************************************
'** Function:       BuildSection4
'** Created by:     Andy Maggs
'** Date:           25/03/2004
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element.
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection4(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection4"

    On Error GoTo ErrHandler
    'EP2_1983/1984
'    Call AddMortgageComponentsDetails(mobjCommon, vxmlNode)
    If mobjCommon.LoanComponents.length = 1 Then
        Call BuildCommonOfferSection4Single(mobjCommon, vxmlNode)
    Else
        Call BuildCommonOfferSection4Multiple(mobjCommon, vxmlNode)
    End If

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection5
'** Created by:     Andy Maggs
'** Date:           31/03/2004
'** Description:    Sets the elements and attributes for the compulsory section5
'**                 (Overall Costs) element.
'** Parameters:     vxmlNode - the section5 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection5(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection5"

    On Error GoTo ErrHandler
    
    Call BuildOverallCostsSection(mobjCommon, vxmlNode, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6
'** Created by:     Andy Maggs
'** Date:           31/03/2004
'** Description:    Sets the elements and attributes for the compulsory section6
'**                 (What you will need to pay each month) element.
'** Parameters:     vxmlNode - the section6 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6"
    Dim xmlComponentNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlLCList As IXMLDOMNodeList
    Dim xmlList As IXMLDOMNodeList
    Dim xmlRate As IXMLDOMNode
    Dim strRepayMethod As String
    Dim strRepayChar As String
    Dim blnInterestOnly As Boolean
    Dim bPartandPart As Boolean
    Dim lngIOAmount As Long
    Dim intIndex As Integer
    'PB 06/06/06 EP651/MAR1590
    Dim intPart As Integer
    'PB EP651/MAR1590 End
    Dim xmlPurposeNode As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlMDay As IXMLDOMNode
    Dim dteMDay As Date
    Dim intPostMday As Integer
    Dim bCalcDatePlus2Years As Boolean
    Dim strTerminMonths As String 'BC MAR1510
    Dim intTerminMonths As Integer 'BC MAR1510
    Dim intYears As Integer  'BC MAR1510
    Dim intMonths As Integer  'BC MAR1510
    Dim dtInterestRateEndDate As Date  'BC MAR1510
    Dim strAssumedStartDate As String
    'PB 06/06/06 EP651/MAR1590
    Dim strValue As String
    Dim strPart As String
    'PB EP651/MAR1590 End
    'EP2_1450
    Dim strRepayVehicle As String
    Dim xmlRepayVehicle As IXMLDOMNode
    Dim xmlPlan As IXMLDOMNode
    Dim strRepayVehicleType As String
    Dim repayVehicleMnthlyPayment As Long
    Dim xmlTemp As IXMLDOMNode
    Dim xmlLoanComponentPaySched As IXMLDOMNode 'BC MAR1831
    Dim strIntOnlyValue As String
    Dim strCapIntValue As String
    Dim xmlNewNode2 As IXMLDOMNode
    Dim xmlKnownList As IXMLDOMNodeList
    Dim xmlMixedList As IXMLDOMNodeList
    Dim xmlUnKnownList As IXMLDOMNodeList
    Dim xmlIntOnlyRepays As IXMLDOMNode
    Dim xmlIntermedNode As IXMLDOMNode
    Dim xmlKnownNode As IXMLDOMNode
    Dim numberIntOnlyRepays As Integer
    Dim costNotKnown As Boolean
    Dim strContactName As String
    Dim strFirmName As String
    Dim isIntermediary As Boolean
    Dim xmlDoc As FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    Dim totalRepayVehicleMnthlyPayment As Double
    Dim xmlRepayVehicleCost As IXMLDOMNode
    Dim xmlModifyCost As IXMLDOMNode
    Dim originalMonthlyCost As Double
    Dim singleComponent As Boolean
    Dim xmlRepay As IXMLDOMNode
    Dim repayText As String
    Dim xmlComponentList As IXMLDOMNodeList
    Dim multiRepayVehicleMnthlyPayment As Double
    Dim xmlComponent As IXMLDOMNode
    Dim totalMonthlyCost As Double

    xmlDoc.async = False
    
    On Error GoTo ErrHandler

    blnInterestOnly = False
    bPartandPart = False
    bCalcDatePlus2Years = False 'MAR88 - BC 05 Oct
    lngIOAmount = 0
    'EP2_1450
    strRepayVehicle = ""
    repayVehicleMnthlyPayment = 0
    originalMonthlyCost = 0
    
    Set xmlIntOnlyRepays = xmlDoc.createElement("REPAYS")
    
    Set xmlComponentNode = mobjCommon.AddComponentsTypeElement(vxmlNode)
    
    If xmlComponentNode.baseName = "SINGLECOMPONENT" Then
        'EP2_2042
        singleComponent = True
        '*-get the single loan component
        Set xmlLoanComponent = mobjCommon.SingleLoanComponent
        If xmlLoanComponent Is Nothing Then
            '*-if there is no loan component we can't really do a lot here!
            Exit Sub
        End If
        
        '*-get the repayment method
        strRepayMethod = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD")
        strRepayChar = Mid(strRepayMethod, 1, 1)
        
        If strRepayChar = "1" Then
            blnInterestOnly = True
            lngIOAmount = xmlGetAttributeAsLong(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT") 'SR EP2_2341-use TOTALLOANCOMPONENTAMOUNT (not LOANAMOUNT)
        End If
                
        If strRepayChar = "3" Then
            bPartandPart = True
            lngIOAmount = xmlGetAttributeAsLong(xmlLoanComponent, "INTERESTONLYELEMENT")
        End If
        
        '*-add the mandatory PAYMENT attribute
        Call AddPaymentAttribute(xmlLoanComponent, xmlComponentNode, "PAYMENT")
        
        'BBG1589 Are we Pre or Post MDay?
'        MAR44 BC 12 Sep 05
'        Set xmlMDay = mobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "KFIMDay" & Chr$(34) & "]")
'        dteMDay = xmlGetAttributeAsDate(xmlMDay, "STRING")
'        intPostMday = DateDiff("d", Now, dteMDay)
'
'        If intPostMday <= 0 Then
            '*-add the loan component interest rate data
            Call AddPostMDayLoanComponentInterestRateData(mobjCommon, xmlLoanComponent, _
                    xmlComponentNode, False, blnInterestOnly)
'        Else
'            '*-add the loan component interest rate data
'            Call AddLoanComponentInterestRateData(mobjCommon, xmlLoanComponent, _
'                    xmlComponentNode, False)
'        End If
        'EP2_1450
        If blnInterestOnly Or bPartandPart Then
            '*-get the repayment vehicle
            strRepayVehicle = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_TEXT")
            
            strRepayVehicleType = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_VALIDID")
            'EP2_2042 should be cost not payment
            repayVehicleMnthlyPayment = xmlGetAttributeAsLong(xmlLoanComponent, "REPAYMENTVEHICLEMONTHLYCOST")
            
            Select Case True
            Case CheckForValidationType(strRepayVehicleType, "U")
                Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlIntOnlyRepays)
            Case CheckForValidationType(strRepayVehicleType, "M")
                Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES_MIXED", xmlIntOnlyRepays)
                Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", mobjCommon.Provider)
            'EP2_2042 New repayment types
            Case CheckForValidationType(strRepayVehicleType, "SRP")
                Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                Set xmlRepay = mobjCommon.CreateNewElement("SALEOFRESPROPERTY", xmlRepayVehicle)
                Call xmlSetAttributeValue(xmlRepay, "PROVIDER", mobjCommon.Provider)
            Case CheckForValidationType(strRepayVehicleType, "INH")
                Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                Set xmlRepay = mobjCommon.CreateNewElement("INHERIT", xmlRepayVehicle)
                Call xmlSetAttributeValue(xmlRepay, "PROVIDER", mobjCommon.Provider)
            Case CheckForValidationType(strRepayVehicleType, "SOB")
                Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                Set xmlRepay = mobjCommon.CreateNewElement("SALEOFBUSINESSSINVEST", xmlRepayVehicle)
                Call xmlSetAttributeValue(xmlRepay, "PROVIDER", mobjCommon.Provider)
            Case Else
                If repayVehicleMnthlyPayment > 0 Then
                    If CheckForValidationType(strRepayVehicleType, "E") Then
                        repayText = "an endowment policy/policies"
                    ElseIf CheckForValidationType(strRepayVehicleType, "I") Then
                        repayText = "an ISA/ISAs"
                    ElseIf CheckForValidationType(strRepayVehicleType, "P") Then
                        repayText = "a Pension/Pensions"
                    End If
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                    Set xmlRepay = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlRepayVehicle)
                Else
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                    Set xmlRepay = mobjCommon.CreateNewElement("REPAYMENTVEHICLESNOCOST", xmlRepayVehicle)
                End If
                Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", mobjCommon.Provider)
                '*-add the PLAN element to REPAYMENTVEHICLES element
                If Len(strRepayVehicle) > 0 Then
                    Set xmlPlan = mobjCommon.CreateNewElement("PLAN", xmlRepay)
                    Call xmlSetAttributeValue(xmlPlan, "PRODUCTNAME", strRepayVehicle)
                    If Len(repayText) > 0 Then
                        Call xmlSetAttributeValue(xmlPlan, "PLANTEXT", repayText)
                    End If
                    Call xmlSetAttributeValue(xmlPlan, "PLANCOST", set2DP(CStr(repayVehicleMnthlyPayment)))
                End If
            End Select
       
        End If
    Else
        'MAR88 - BC 05 Oct
'        strAssumedStartDate = Format$(Now, "dd/mm/yyyy")
'        strAssumedStartDate = "01" & Mid(strAssumedStartDate, 3, 8)
'        strAssumedStartDate = Format$(DateAdd("m", 1, strAssumedStartDate), "dd/mm/yyyy")
        'EP2_2042
        singleComponent = False
        
        'EP2_1450
        strAssumedStartDate = Format$(CalcExpectedCompletionDate(Date), "dd/mm/yyyy")
        
        xmlSetAttributeValue xmlComponentNode, "ASSUMEDSTARTDATE", strAssumedStartDate
        
        AddCalculationDateAttribute mobjCommon, xmlComponentNode, bCalcDatePlus2Years
        '*-add the appropriate Fees and/or Premiums element
        'EP2_1450
        Set xmlTemp = AddFeesPremiumsElement(mobjCommon, xmlComponentNode)
        If xmlTemp.baseName <> "NOFEESORPREMIUMS" Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(mobjCommon, xmlTemp)
        End If
        
        '*-get the list of LoanComponents
        Set xmlLCList = mobjCommon.LoanComponents
        '*-add a COMPONENT element for each LoanComponent
        For Each xmlLoanComponent In xmlLCList
            strRepayMethod = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD")
            strRepayChar = Mid(strRepayMethod, 1, 1)
            
            If strRepayChar = "1" Then
                blnInterestOnly = True
                lngIOAmount = lngIOAmount + xmlGetAttributeAsLong(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT") 'SR EP2_2341-use TOTALLOANCOMPONENTAMOUNT (not LOANAMOUNT)
            End If
            
            If strRepayChar = "3" Then
                blnInterestOnly = True
                lngIOAmount = lngIOAmount + xmlGetAttributeAsLong(xmlLoanComponent, "INTERESTONLYELEMENT")
            End If
            
            '*-add the COMPONENT element
            Set xmlNewNode = mobjCommon.CreateNewElement("COMPONENT", xmlComponentNode)
            '*-add the mandatory INITIALPAYMENT attribute
            Call AddPaymentAttribute(xmlLoanComponent, xmlNewNode, "INITIALPAYMENT")
            
            '*-and get the first interestratetype record
            Set xmlRate = mobjCommon.GetLoanComponentFirstInterestRate(xmlLoanComponent)
            'EP2_2395
            '*-add the mandatory RATETYPE attribute
            Call AddInterestRateTypeAttributeSect6(xmlRate, xmlNewNode)
            '*-add the mandatory INITIALRATE attribute
            Call AddRateAttribute(mobjCommon, xmlLoanComponent, _
                    xmlRate, xmlNewNode, "INITIALRATE")
            '*-add the mandatory REPAYMENTMETHOD attribute
            Call AddRepaymentMethodAttribute(xmlLoanComponent, xmlNewNode)
            
            'BC MAR1510 Begin
            '*-add the mandatory TERMINMONTHS attribute
            '*-get the first Interest Rate Type record
            
            If Not xmlRate Is Nothing Then
                If xmlAttributeValueExists(xmlRate, "INTERESTRATEPERIOD") Then
                    strTerminMonths = xmlGetAttributeText(xmlRate, "INTERESTRATEPERIOD")
                    If strTerminMonths = "-1" Then
                        intYears = xmlGetAttributeAsInteger(xmlLoanComponent, "TERMINYEARS")
                        intMonths = xmlGetAttributeAsInteger(xmlLoanComponent, "TERMINMONTHS")
                        intMonths = (12 * intYears) + intMonths
                        Call xmlSetAttributeValue(xmlNewNode, "TERMINMONTHS", CStr(intMonths))
                    Else
                        Call xmlSetAttributeValue(xmlNewNode, "TERMINMONTHS", strTerminMonths)
                    End If
                ElseIf xmlAttributeValueExists(xmlRate, "INTERESTRATEENDDATE") Then
                    dtInterestRateEndDate = xmlGetAttributeText(xmlRate, "INTERESTRATEENDDATE")
                    'EP2_1450
'                    intTerminMonths = DateDiff("m", mobjCommon.ExpectedCompletionDate, dtInterestRateEndDate)
                    intTerminMonths = MonthDiff(mobjCommon.ExpectedCompletionDate, dtInterestRateEndDate)   'MAR1777 GHun
                    Call xmlSetAttributeValue(xmlNewNode, "TERMINMONTHS", CStr(intTerminMonths))
                End If
            End If
            'BC MAR1510 End
    
            'Call AddTermInMonthsAttribute(xmlLoanComponent, xmlNewNode)
            '*-add the mandatory LOANCOMPONENTAMOUNT attribute
            Call AddLoanComponentAmountAttribute(xmlLoanComponent, xmlNewNode)
            '*-add the mandatory PART attribute
            Call AddComponentPartAttribute(xmlLoanComponent, xmlNewNode)
            
            'PB 06/06/06 EP651/MAR1590
            If (InStr(1, xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD"), "3", vbTextCompare)) Then
                'EP2_1450
                Set xmlLoanComponentPaySched = xmlLoanComponent.selectSingleNode("LOANCOMPONENTPAYMENTSCHEDULE")
                strIntOnlyValue = xmlGetAttributeText(xmlLoanComponentPaySched, "INTONLYMONTHLYCOST")
                strCapIntValue = xmlGetAttributeText(xmlLoanComponentPaySched, "CAPINTMONTHLYCOST")


                Set xmlNewNode2 = xmlNewNode.cloneNode(True)
                xmlComponentNode.appendChild xmlNewNode2
                intPart = xmlGetAttributeText(xmlNewNode, "PART")
                Call xmlSetAttributeValue(xmlNewNode, "PART", CStr(intPart) + "(a)")
                Call xmlSetAttributeValue(xmlNewNode, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlLoanComponent, "CAPITALANDINTERESTELEMENT"))
                
                Call xmlSetAttributeValue(xmlNewNode, "REPAYMENTMETHOD", "Repayment (i.e. capital and interest)")
                Call xmlSetAttributeValue(xmlNewNode, "INITIALPAYMENT", strCapIntValue)
                Call xmlSetAttributeValue(xmlNewNode2, "PART", CStr(intPart) + "(b)")
                Call xmlSetAttributeValue(xmlNewNode2, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlLoanComponent, "INTERESTONLYELEMENT"))
                Call xmlSetAttributeValue(xmlNewNode2, "REPAYMENTMETHOD", "Interest Only")
                Call xmlSetAttributeValue(xmlNewNode2, "INITIALPAYMENT", strIntOnlyValue)
            End If
            'PB 06/06/06 EP651/MAR1590 End
            'EP2_1450
            If blnInterestOnly Or bPartandPart Then
    
                '*-get the repayment vehicle
                strRepayVehicle = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_TEXT")
                
                strRepayVehicleType = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_VALIDID")
                'EP2_2042 should be cost not payment
                repayVehicleMnthlyPayment = xmlGetAttributeAsLong(xmlLoanComponent, "REPAYMENTVEHICLEMONTHLYCOST")
                
                Select Case True
                Case CheckForValidationType(strRepayVehicleType, "U")
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlIntOnlyRepays)
                Case CheckForValidationType(strRepayVehicleType, "M")
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES_MIXED", xmlIntOnlyRepays)
                    Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", mobjCommon.Provider)
                'EP2_2042 New repayment types
                Case CheckForValidationType(strRepayVehicleType, "SRP")
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                    Set xmlRepay = mobjCommon.CreateNewElement("SALEOFRESPROPERTY", xmlRepayVehicle)
                    Call xmlSetAttributeValue(xmlRepay, "PROVIDER", mobjCommon.Provider)
                Case CheckForValidationType(strRepayVehicleType, "INH")
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                    Set xmlRepay = mobjCommon.CreateNewElement("INHERIT", xmlRepayVehicle)
                    Call xmlSetAttributeValue(xmlRepay, "PROVIDER", mobjCommon.Provider)
                Case CheckForValidationType(strRepayVehicleType, "SOB")
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                    Set xmlRepay = mobjCommon.CreateNewElement("SALEOFBUSINESSSINVEST", xmlRepayVehicle)
                    Call xmlSetAttributeValue(xmlRepay, "PROVIDER", mobjCommon.Provider)
                Case Else
                    If repayVehicleMnthlyPayment > 0 Then
                        If CheckForValidationType(strRepayVehicleType, "E") Then
                            repayText = "an endowment policy/policies"
                        ElseIf CheckForValidationType(strRepayVehicleType, "I") Then
                            repayText = "an ISA/ISAs"
                        ElseIf CheckForValidationType(strRepayVehicleType, "P") Then
                            repayText = "a Pension/Pensions"
                        End If
                        Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                        Set xmlRepay = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlRepayVehicle)
                    Else
                        Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlIntOnlyRepays)
                        Set xmlRepay = mobjCommon.CreateNewElement("REPAYMENTVEHICLESNOCOST", xmlRepayVehicle)
                    End If
                    Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", mobjCommon.Provider)
                    '*-add the PLAN element to REPAYMENTVEHICLES element
                    If Len(strRepayVehicle) > 0 Then
                        Set xmlPlan = mobjCommon.CreateNewElement("PLAN", xmlRepay)
                        Call xmlSetAttributeValue(xmlPlan, "PRODUCTNAME", strRepayVehicle)
                        If Len(repayText) > 0 Then
                            Call xmlSetAttributeValue(xmlPlan, "PLANTEXT", repayText)
                        End If
                        Call xmlSetAttributeValue(xmlPlan, "PLANCOST", set2DP(CStr(repayVehicleMnthlyPayment)))
                    End If
                End Select
                multiRepayVehicleMnthlyPayment = multiRepayVehicleMnthlyPayment + repayVehicleMnthlyPayment
            End If
        Next xmlLoanComponent
    End If
   
    If blnInterestOnly Or bPartandPart Then

        costNotKnown = False
        numberIntOnlyRepays = xmlIntOnlyRepays.childNodes.length
        'EP2_1449 Should use the FirmName
        isIntermediary = mobjCommon.IsIntroducedByIntermediary(strContactName, False, strFirmName, True)
        
        '*-add the INTERESTONLY element
        Set xmlNewNode = mobjCommon.CreateNewElement("INTERESTONLY", vxmlNode)
        '*-add the mandatory INTERESTONLYTOTAL attribute
        'SR ep2_2341 : DO not add Fees as it has already been added to lngIOAmount
        Call xmlSetAttributeValue(xmlNewNode, "INTERESTONLYTOTAL", _
                    CStr(lngIOAmount))  'SR 21/09/2004 : CORE82

        Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE4", mobjCommon.MortgageTypeText4)
        'CORE82
        Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(xmlNewNode)
        If isIntermediary Then
            Set xmlIntermedNode = mobjCommon.CreateNewElement("INTERMEDIARY", xmlNewNode)
            'EP2_1449 Should use the FirmName
            Call xmlSetAttributeValue(xmlIntermedNode, "INTERMEDIARYFIRMNAME", strFirmName)
        End If
        If (numberIntOnlyRepays > 0) Then

            Set xmlKnownList = xmlIntOnlyRepays.selectNodes("REPAYMENTVEHICLES")
            Set xmlMixedList = xmlIntOnlyRepays.selectNodes("REPAYMENTVEHICLES_MIXED")
            Set xmlUnKnownList = xmlIntOnlyRepays.selectNodes("REPAYMENTVEHICLES_UNKNOWN")
            'If some repayment vehicles are Mixed or we have known and unknown,
            'show as a single repayment vehicle entry of mixed
            If ((xmlMixedList.length > 0) Or ((xmlKnownList.length > 0) And (xmlUnKnownList.length > 0))) Then
                If (xmlMixedList.length > 0) Then
                    xmlNewNode.appendChild xmlMixedList.Item(0)
                Else
                    Set xmlRepayVehicle = mobjCommon.CreateNewElement("REPAYMENTVEHICLES_MIXED", xmlNewNode)
                    Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", mobjCommon.Provider)
                End If
            ElseIf xmlUnKnownList.length = numberIntOnlyRepays Then
            'if all repayment vehicles are unknown, single repayment vehicle of unknown
                xmlNewNode.appendChild xmlUnKnownList.Item(0)
            ElseIf xmlKnownList.length = numberIntOnlyRepays Then
            'If all repayment vehicles are known repayment vehicles for each
                For Each xmlKnownNode In xmlKnownList
'                    If Not xmlGetAttributeAsDouble(xmlKnownNode, "PLANCOST") > 0 Then
'                        costNotKnown = True
'                    End If
                    xmlNewNode.appendChild xmlKnownNode.childNodes(0).cloneNode(True)

                Next xmlKnownNode
            End If
            
        Else
            Call mobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
        End If

        'EP2_2042
'        If costNotKnown = True Then
'            Set xmlRepayVehicle = mobjCommon.CreateNewElement("MISSINGCOST", xmlNewNode)
'        End If
            
        If singleComponent Then
        'We only show this bit if we have items with a cost
            If repayVehicleMnthlyPayment > 0 Then
                'We have singlecomponent
                Set xmlModifyCost = xmlComponentNode.cloneNode(True)
                'MONTHLYPAYMENT only one
                Set xmlTemp = xmlModifyCost.selectSingleNode("//MONTHLYPAYMENT")
                originalMonthlyCost = xmlGetAttributeAsDouble(xmlTemp, "PAYMENT")
                originalMonthlyCost = originalMonthlyCost + repayVehicleMnthlyPayment
                xmlSetAttributeValue xmlTemp, "PAYMENT", set2DP(originalMonthlyCost)
                'FOLLOWEDBY could be multiple
                Set xmlComponentList = xmlModifyCost.selectNodes("//FOLLOWEDBY")
                For Each xmlComponent In xmlComponentList
                    originalMonthlyCost = 0
                    originalMonthlyCost = xmlGetAttributeAsDouble(xmlComponent, "PAYMENT")
                    originalMonthlyCost = originalMonthlyCost + repayVehicleMnthlyPayment
                    xmlSetAttributeValue xmlComponent, "PAYMENT", set2DP(originalMonthlyCost)
                    xmlNewNode.appendChild xmlModifyCost
                Next
            End If
        Else
            'We only show this bit if we have items with a cost
            If multiRepayVehicleMnthlyPayment > 0 Then
                'We have multicomponent
                Set xmlModifyCost = xmlComponentNode.cloneNode(True)
                
                Set xmlComponentList = xmlModifyCost.selectNodes("COMPONENT")
                For Each xmlComponent In xmlComponentList
                    originalMonthlyCost = xmlGetAttributeAsDouble(xmlComponent, "INITIALPAYMENT")
                    totalMonthlyCost = totalMonthlyCost + originalMonthlyCost
                Next
                totalMonthlyCost = totalMonthlyCost + multiRepayVehicleMnthlyPayment

                For Each xmlKnownNode In xmlKnownList
                    Dim planCost As Double
                    Set xmlTemp = xmlKnownNode.childNodes(0).selectSingleNode("PLAN")
                    If Not xmlTemp Is Nothing Then
                        planCost = xmlGetAttributeAsDouble(xmlTemp, "PLANCOST")
                        If (planCost > 0) Then
                            xmlModifyCost.appendChild xmlKnownNode.childNodes(0).cloneNode(True)
                        End If
                    End If
                Next
                Set xmlTemp = mobjCommon.CreateNewElement("TOTALWITHRPV", xmlModifyCost)
                Call xmlSetAttributeValue(xmlTemp, "TOTALMONTHLYCOST", totalMonthlyCost)
                
                xmlNewNode.appendChild xmlModifyCost
                
            End If
        End If
  
    End If
    
    Set xmlComponentNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLCList = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNode = Nothing
    Set xmlMDay = Nothing
    Set xmlTemp = Nothing
    Set xmlLoanComponentPaySched = Nothing
    Set xmlKnownList = Nothing
    Set xmlMixedList = Nothing
    Set xmlUnKnownList = Nothing
    Set xmlIntOnlyRepays = Nothing
    Set xmlIntermedNode = Nothing
    Set xmlKnownNode = Nothing
    Set xmlDoc = Nothing
    Set xmlRepayVehicleCost = Nothing
    Set xmlModifyCost = Nothing
    Set xmlRepay = Nothing
    Set xmlComponentList = Nothing
    Set xmlComponent = Nothing
    
Exit Sub
ErrHandler:
    Set xmlComponentNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLCList = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNode = Nothing
    Set xmlMDay = Nothing
    Set xmlTemp = Nothing
    Set xmlLoanComponentPaySched = Nothing
    Set xmlKnownList = Nothing
    Set xmlMixedList = Nothing
    Set xmlUnKnownList = Nothing
    Set xmlIntOnlyRepays = Nothing
    Set xmlIntermedNode = Nothing
    Set xmlKnownNode = Nothing
    Set xmlDoc = Nothing
    Set xmlRepayVehicleCost = Nothing
    Set xmlModifyCost = Nothing
    Set xmlRepay = Nothing
    Set xmlComponentList = Nothing
    Set xmlComponent = Nothing
    
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6A
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section6A
'**                 (What you will need to pay in the future) element.
'** Parameters:     vxmlNode - the section6A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6A(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6A"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlBalance As IXMLDOMNode
    Dim xmlRate As IXMLDOMNode
    
    On Error GoTo ErrHandler

    Set xmlList = mobjCommon.LoanComponents
    If xmlList.length > 1 Then
        For Each xmlItem In xmlList
            'MAR54
            '*-get the LOANCOMPONENTPAYMENTSCHEDULE element for this loan component
            Set xmlBalance = xmlItem.selectSingleNode(".//LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=2]")
            'MAR54 - do not create entry for this component if there is only 1 payment schedule
            If Not xmlBalance Is Nothing Then
            '*-add the COMPONENT element
                Set xmlComponent = mobjCommon.CreateNewElement("COMPONENT", vxmlNode)
                            
                '*-and get the first interestratetyperecord
                Set xmlRate = mobjCommon.GetLoanComponentFirstInterestRate(xmlItem)
                
                '*-add the RATEPERIOD attribute
                Call AddRatePeriodAttribute(mobjCommon, xmlRate, xmlComponent)
            
                '*-add the NEWAMOUNT attribute
                Call AddNewAmountAttribute(xmlBalance, xmlComponent)
                
                '*-add the PART attribute
                Call AddComponentPartAttribute(xmlItem, xmlComponent)
                      
                '*-add the OFFERTYPE attribute
                Call AddOfferTypeAttribute(xmlRate, xmlComponent)
                
                Set xmlBalance = Nothing
            End If
            
        Next xmlItem
    End If

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlComponent = Nothing
    Set xmlBalance = Nothing
    Set xmlRate = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlComponent = Nothing
    Set xmlBalance = Nothing
    Set xmlRate = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section7
'** Parameters:     vxmlNode - the section7 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7"

    On Error GoTo ErrHandler

    Call BuildSection7Common(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7A
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section7A
'**                 (Total Borrowing) element.
'** Parameters:     vxmlNode - the section7A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7A(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7A"

    On Error GoTo ErrHandler

    Call BuildSection7ACommon(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection8
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section8
'**                 (What fees must you pay?) element.
'** Parameters:     vxmlNode - the section8 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection8(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection8"

    On Error GoTo ErrHandler

    Call BuildWhatFeesYouMustPaySection(mobjCommon, vxmlNode, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection9
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section9
'**                 (Insurance) element.
'** Parameters:     vxmlNode - the section9 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection9(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection9"

    On Error GoTo ErrHandler

    Call BuildInsuranceSection(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection10
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section10
'**                 (What happens if you do not want this mortgage any more?) element.
'** Parameters:     vxmlNode - the section10 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection10(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection10"

    On Error GoTo ErrHandler

    Call BuildMortgageNoMoreSection(mobjCommon, vxmlNode, False, False)
    
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection11
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section11
'**                 (What happens if you want to make overpayments?) element.
'** Parameters:     vxmlNode - the section11 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection11(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection11"

    On Error GoTo ErrHandler

    Call BuildSection11Common(mobjCommon, vxmlNode, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection12
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section11
'**                 (Additional Features) element.
'** Parameters:     vxmlNode - the section12 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection12(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection12"

    On Error GoTo ErrHandler

    Call BuildAdditionalFeaturesSection(mobjCommon, vxmlNode, False, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection13
'** Created by:     Srini Rao
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section13
'**                 (Using a mortgage intermediary) element.
'** Parameters:     vxmlNode - the section13 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection13(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection13"

    On Error GoTo ErrHandler

    Call BuildUsingIntermediarySection(mobjCommon, vxmlNode, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection14
'** Created by:     Andy Maggs
'** Date:           01/04/2004
'** Description:    Sets the elements and attributes for the compulsory section14
'**                 (Where can you get more information on mortgages?) element.
'** Parameters:     vxmlNode - the section14 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection14(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection14"

    On Error GoTo ErrHandler

    BuildLifeTimeContactsSection mobjCommon, vxmlNode  'SR 27/10/2004 : BBG1717
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'********************************************************************************
'** Function:       AddCustomerAddedRemovedData
'** Created by:     Andy Maggs
'** Date:           06/07/2004
'** Description:    Adds the customers added/removed details as specified.
'** Parameters:     vstrXPath - the XPath to select the customer added or
'**                 customer removed nodes.
'**                 vstrNodeName - the name of the node to create.
'**                 vstrAttribName - the name of the attribute to create.
'**                 vxmlNode - the node to create them on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
'EP2_1667 Moved to CommonDataHelper
'Private Sub AddCustomerAddedRemovedData(ByVal vstrXPath As String, _
'        ByVal vstrNodeName As String, ByVal vstrAttribName As String, _
'        ByVal vxmlNode As IXMLDOMNode)
'    Dim xmlList As IXMLDOMNodeList
'    Dim xmlItem As IXMLDOMNode
'    Dim strNamesList As String
'    Dim xmlNames As IXMLDOMNode
'
'    Set xmlList = mobjCommon.Data.selectNodes(vstrXPath)
'    If Not xmlList Is Nothing Then
'        For Each xmlItem In xmlList
'            If Len(strNamesList) > 0 Then
'                strNamesList = strNamesList & vbCrLf
'            End If
'            strNamesList = strNamesList & BuildName(xmlGetAttributeText(xmlItem, "TITLE_TEXT"), _
'                    xmlGetAttributeText(xmlItem, "FORENAME"), _
'                    xmlGetAttributeText(xmlItem, "SURNAME"))
'        Next xmlItem
'
'        If Len(strNamesList) > 0 Then
'            '*-add the ADDNAMES element
'            Set xmlNames = mobjCommon.CreateNewElement(vstrNodeName, vxmlNode)
'            '*-add the NAMESTOADD attribute
'            Call xmlSetAttributeValue(xmlNames, vstrAttribName, strNamesList)
'        End If
'    End If
'
'    Set xmlList = Nothing
'    Set xmlItem = Nothing
'    Set xmlNames = Nothing
'End Sub

'EP2_583
Private Sub BuildContactSection(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection14"

    On Error GoTo ErrHandler

    Call BuildCommonOfferContactSection(mobjCommon, vxmlNode)
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'EP2_583
Private Sub BuildRiskSection(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildRiskSection"

    On Error GoTo ErrHandler

    Call BuildCommonOfferRiskSection(mobjCommon, vxmlNode)
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub


