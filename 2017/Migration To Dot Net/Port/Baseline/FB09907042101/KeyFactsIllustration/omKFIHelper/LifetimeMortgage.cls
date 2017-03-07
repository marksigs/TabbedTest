VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LifetimeMortgage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         LifetimeMortgage
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Contains the functionality to create the appropriate KFI XML
'**                 for a lifetime mortgage illustration.
'********************************************************************************
Option Explicit

Implements IMortgage

Private mobjCommon As CommonDataHelper

Private Sub Class_Terminate()
    Set mobjCommon = Nothing 'TK 18/11/2004 BBG1672
End Sub

'********************************************************************************
'** Function:       IMortgage_FormatData
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    The method responsible for formatting the data into the XML
'**                 format required for the lifetime mortgage KFI document
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

    On Error GoTo ErrHandler

    '*-store the common data helper object
    Set mobjCommon = vobjCommon

    '*-create the TEMPLATEDATA element and set the attributes
    Set xmlElem = mobjCommon.Document.createElement("TEMPLATEDATA")
    Set xmlTemplate = mobjCommon.Document.appendChild(xmlElem)
    Call BuildTemplateData(xmlTemplate)

    '*-create the mandatory SECTION1 (About this illustration) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION1", xmlTemplate)
    Call BuildSection1(xmlNode)
    
    '*-create the mandatory SECTION2 (Which service are we providing you with?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION2", xmlTemplate)
    Call BuildSection2(xmlNode)
    
    '*-create the mandatory SECTION3 (What is a life time mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION3", xmlTemplate)
    Call BuildSection3(xmlNode)
    
    '*-create the mandatory SECTION4 (your mortgage requirements) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION4", xmlTemplate)
    Call BuildSection4(xmlNode)
    
    '*-create the mandatory SECTION5 (description of mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION5", xmlTemplate)
    Call BuildSection5(xmlNode)
    
    '*-create the mandatory SECTION6 (benefits) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION6", xmlTemplate)
    Call BuildSection6(xmlNode)
    
    '*-create the mandatory SECTION7 (Risks) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION7", xmlTemplate)
    Call BuildSection7(xmlNode)
    
    '*-create the mandatory SECTION8 (What will you owe and when) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION8", xmlTemplate)
    Call BuildSection8(xmlNode)
    
    '*-create the mandatory SECTION9 (Will the interest rate change) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION9", xmlTemplate)
    Call BuildSection9(xmlNode)

    '*-create the mandatory SECTION10 (How the value of your home could change) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION10", xmlTemplate)
    Call BuildSection10(xmlNode)

    '*-create the mandatory SECTION11 (what fee must you pay?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION11", xmlTemplate)
    Call BuildSection11(xmlNode)

    '*-create the mandatory SECTION12 (insurance) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION12", xmlTemplate)
    Call BuildSection12(xmlNode)

    '*-create the mandatory SECTION13 (what happens if you do not want this any more) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION13", xmlTemplate)
    Call BuildSection13(xmlNode)

    '*-create the mandatory SECTION14 (Additional Features) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    Call BuildSection14(xmlNode)
    
    '*-create the mandatory SECTION15 (Overall cost of the mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION15", xmlTemplate)
    Call BuildSection15(xmlNode)
    
    '*-create the mandatory SECTION16 (Using a mortgage intermediary) element
    '* Add this section only if Intermediary services are used
    blnIsIntermediary = vobjCommon.IsIntroducedByIntermediary(strContactName, False)
    If blnIsIntermediary Then
        Set xmlNode = mobjCommon.CreateNewElement("SECTION16", xmlTemplate)
        Call BuildSection16(xmlNode)
    End If
    
    '*-create the mandatory CONTACTSECTION element
    Set xmlNode = mobjCommon.CreateNewElement("CONTACTSECTION", xmlTemplate) 'SR 11/09/2004 : CORE82
    BuildContactsSection xmlNode

    Set xmlElem = Nothing
    Set xmlTemplate = Nothing
    Set xmlNode = Nothing
    Exit Sub
ErrHandler:
    Set xmlElem = Nothing
    Set xmlTemplate = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildTemplateData
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:
'** Parameters:
'** Returns:
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildTemplateData(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildTemplateData"
    Dim blnValuationUsed As Boolean
    Dim xmlNode As IXMLDOMNode

    On Error GoTo ErrHandler

    '*-add the mandatory PROVIDER attribute
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", mobjCommon.Provider)
    
    '*-add the mandatory REFERENCE attribute
    Call AddReferenceAttribute(mobjCommon.Data, vxmlNode)
    
    '*-add the mandatory APPLICANTNAME attribute
    Call AddApplicantNameAttribute(mobjCommon.Data, vxmlNode)
    
    '*-add the mandatory VALIDTODATE attribute
    Call AddValidToDateAttribute(mobjCommon.Data, vxmlNode)
    
    '*-add the mandatory TOTALLOANAMOUNT attribute
    xmlSetAttributeValue vxmlNode, "TOTALLOANAMOUNT", mobjCommon.TotalLoanAmount
    
    '*-add the mandatory TOTALFEE attribute
    xmlSetAttributeValue vxmlNode, "TOTALFEE", mobjCommon.TotalFees
    
    '*-add the mandatory MORTGAGETYPE  attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)
    
    '*-add the mandatory MORTGAGETYPE2  attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE2", mobjCommon.MortgageTypeText2)
    
    '*-Add Term of the Mortgage (TERMINYEARS, TERMINMONTHS)
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode)
        
    '*-Add the mandatory PURCHASEPRICE attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, True, _
            blnValuationUsed)
    
    '*-Add LifeTime MortageType as element
    If mobjCommon.IsAdditionalBorrowing Then
         Call mobjCommon.CreateNewElement("LIFETIMEADDITIONALBORROWING", vxmlNode)
    Else
        Call mobjCommon.CreateNewElement("LIFETIMEMORTGAGE", vxmlNode)
    End If
    
    '*-Add 'TIMELIMIT' element
    Set xmlNode = mobjCommon.CreateNewElement("TIMELIMIT", vxmlNode)
    Call AddCompletedByDateAttribute(mobjCommon, xmlNode)
    
    '*-create the mandatory AUTONUMBERING element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", vxmlNode)
    Call AddAutonumberingAttributes(xmlNode)
    
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", vxmlNode)
    Call AddAutonumberingAttributes(xmlNode, "RISK")
    
    'SR 12/09/2004 : CORE82 - Set SVR as attribute
    Call SetSVR(vxmlNode, mobjCommon)
    
    Set xmlNode = Nothing
    Exit Sub
ErrHandler:
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection1
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section1 node (About this illustration).
'** Parameters:     vxmlNode - the section1 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection1(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection1"
    
    On Error GoTo ErrHandler
    
    If Not mobjCommon.IsAdditionalBorrowing Then
        '*-this is not a further advance, so add the empty NOTFURTHERADVANCE
        '*-element
        Call mobjCommon.CreateNewElement("NOTFURTHERADVANCE", vxmlNode)
    End If

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection2
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section2 node (Which service are we providing you
'**                 with?).
'** Parameters:     vxmlNode - the section2 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection2(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection2"
    
    On Error GoTo ErrHandler

    AddServicesProvided mobjCommon, vxmlNode, True

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection3
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section3 node (What is a life time mortgage?).
'** Parameters:     vxmlNode - the section3 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection3"
    
    On Error GoTo ErrHandler
    
    BuildLifeTimeSection3 mobjCommon, vxmlNode

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection4
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section4 node (Your mortgage requirements).
'** Parameters:     vxmlNode - the section4 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection4(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection4"
    Dim xmlMainPurposeNode As IXMLDOMNode
    Dim xmlPurposeNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlIncomeRequirements As IXMLDOMNode
    Dim xmlLTV As IXMLDOMNode
    Dim blnValuationUsed As Boolean
    Dim blnPurposePurchase As Boolean
    Dim blnIsValuation As Boolean

    On Error GoTo ErrHandler

    '*- Add the term (highest amongst the term for all loan components) as attributes
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS") 'SR 10/09/2004 : CORE82
    
    '*-add the Purpose node
    Set xmlMainPurposeNode = mobjCommon.CreateNewElement("PURPOSE", vxmlNode)
    Set xmlNode = mobjCommon.AddLoanPurposeElement(xmlMainPurposeNode)
    xmlSetAttributeValue xmlNode, "TOTALLOANAMOUNT", mobjCommon.LoanAmount()
    
    '*-add the appropriate (1-16) mandatory FEESANDPREMIUMS element (exactly one)
    Call AddFeesAndPremiumsElement(mobjCommon, xmlMainPurposeNode)
    
    '*- add element CALCULATEDTERM
    Set xmlNode = mobjCommon.CreateNewElement("CALCULATEDTERM", xmlMainPurposeNode)
    Call AddMortgageTermAttributes(mobjCommon, xmlNode)
   
    '*-now add one of the mandatory PURPOSEPURCHASE, PURPOSEREMORTGAGE or
    '*-PURPOSEFURTHERADVANCE elements to the SECTION4 NODE
    Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(vxmlNode)
    Select Case xmlPurposeNode.baseName
        Case "PURPOSEPURCHASE"
            '*-add the mandatory PURCHASEPRICE attribute
            'E2EM00002842 This will always be purchase price
            Call AddPurchPriceValueAttribute(mobjCommon.Data, xmlPurposeNode)
'            Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, True, _
'                   blnValuationUsed)

        Case "PURPOSEREMORTGAGE"
            'TODO: add the mandatory INSUREDVALUE attribute - temporarily set to 0
            '*-as per the mapping document
            'SR 10/09/2004 : CORE82  - attribute 'INSUREDVALUE' is not required
            'Call xmlSetAttributeValue(xmlPurposeNode, "INSUREDVALUE", CStr(0))
            
            '*-add the mandatory ESTIMATEDVALUE attribute
            Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, False, _
                    blnValuationUsed)
        
        Case "PURPOSEFURTHERADVANCE"
            '*-add the mandatory ESTIMATEDVALUE attribute
            Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, False, _
                    blnValuationUsed)
    End Select
    
    '*-add element OWNHOUSE
    If IsPropertyFreeHold(mobjCommon) Then
        Set xmlNode = mobjCommon.CreateNewElement("OWNHOUSE", vxmlNode)
    End If
        
    '* - Add CLIENTAGE element to SECTION4 node
    Set xmlNode = mobjCommon.CreateNewElement("CLIENTAGE", vxmlNode)
    xmlSetAttributeValue xmlNode, "AGE", GetYoungestApplicantsAge(mobjCommon)
        
    '* - Add NOOTHERMORTGAGE/OTHERMORTGAGE element to SECTION4 node.
    '* This adds the details of mortgage loans (if any existing) on this property.
    Call AddOtherMortgagesData(mobjCommon, vxmlNode)

    '* - Add INCOMEREQUIREMENTS element to SECTION4 node
    Set xmlIncomeRequirements = mobjCommon.CreateNewElement("INCOMEREQUIREMENTS", vxmlNode)
    
'IK 01/12/2004 E2EM00003126 LUMPSUMWITHDRAWDOWN mods
    If mobjCommon.HasDrawDown Then
        Set xmlNode = mobjCommon.CreateNewElement("LUMPSUMWITHDRAWDOWN", xmlIncomeRequirements)
        xmlSetAttributeValue xmlNode, "LUMPSUMAMOUNT", mobjCommon.AmountRequestedLessDrawDown
        xmlSetAttributeValue xmlNode, "DRAWDOWNAMOUNT", mobjCommon.DrawDownAmount
    Else
        Set xmlNode = mobjCommon.CreateNewElement("LUMPSUM", xmlIncomeRequirements)
        xmlSetAttributeValue xmlNode, "LUMPSUMAMOUNT", mobjCommon.LoanAmount
    End If
'IK 01/12/2004 E2EM00003126 ends
    
    '* - Add LOANTOVALUE element to SECTION4 node
    Set xmlLTV = mobjCommon.CreateNewElement("LOANTOVALUE", vxmlNode)
    xmlSetAttributeValue xmlLTV, "LTV", mobjCommon.LTV
'SR 10/09/2004 : CORE82
    If mobjCommon.GetMainMortgageTypeGroup <> "R" Then 'SR 19/10/2004 : BBG1662
        mobjCommon.CreateNewElement "PROPERTYPRICE", xmlLTV  'SR 17/10/2004 : BBG1643
    Else
        mobjCommon.CreateNewElement "ESTIMATEDVALUE", xmlLTV
    End If
        
'    blnPurposePurchase = (mobjCommon.GetMainMortgageTypeGroup = "N")
'    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlLTV, blnPurposePurchase, _
'            blnIsValuation)
'SR 10/09/2004 : CORE82 - End
    Set xmlMainPurposeNode = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNode = Nothing
    Set xmlIncomeRequirements = Nothing
    Set xmlLTV = Nothing
    Exit Sub
ErrHandler:
    Set xmlMainPurposeNode = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNode = Nothing
    Set xmlIncomeRequirements = Nothing
    Set xmlLTV = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection5
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section5 node (description of mortgage).
'** Parameters:     vxmlNode - the section5 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection5(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection5"
    
    On Error GoTo ErrHandler
    
    Call BuildLifetimeSection5(mobjCommon, vxmlNode)
    
    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section6 node (benefits).
'** Parameters:     vxmlNode - the section6 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6"

    On Error GoTo ErrHandler

    BuildLifeTimeSection6 mobjCommon, vxmlNode, False

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section7 node (Risks).
'** Parameters:     vxmlNode - the section7 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7"

    On Error GoTo ErrHandler

    BuildLifeTimeSection7 mobjCommon, vxmlNode, False

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection8
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section8 node (What will you owe and when).
'** Parameters:     vxmlNode - the section8 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection8(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection8"

    On Error GoTo ErrHandler
    
    BuildLifeTimeSection8 mobjCommon, vxmlNode, False

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection9
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section9 node (Will the interest rate change?).
'** Parameters:     vxmlNode - the section9 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection9(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection9"

    On Error GoTo ErrHandler
    
    BuildLifeTimeSection9 mobjCommon, vxmlNode, False

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection10
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section10 node (How the value of your home could
'**                 change).
'** Parameters:     vxmlNode - the section10 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection10(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection10"

    On Error GoTo ErrHandler
    
    BuildLifeTimeSection10 mobjCommon, vxmlNode, False
    
    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection11
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section11 node (what fee must you pay?).
'** Parameters:     vxmlNode - the section11 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection11(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection11"

    On Error GoTo ErrHandler

    Call BuildWhatFeesYouMustPaySection(mobjCommon, vxmlNode, False)

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection12
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section12 node (insurance).
'** Parameters:     vxmlNode - the section12 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection12(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection12"

    On Error GoTo ErrHandler

    Call BuildInsuranceSection(mobjCommon, vxmlNode)

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection13
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section13 node (what happens if you do not want
'**                 this any more?).
'** Parameters:     vxmlNode - the section13 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection13(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection13"
    
    On Error GoTo ErrHandler

    Call BuildLifeTimeSection13(mobjCommon, vxmlNode, False)
    
    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection14
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section14 node (Additional Features).
'** Parameters:     vxmlNode - the section14 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection14(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection14"

    On Error GoTo ErrHandler

    Call BuildAdditionalFeaturesSection(mobjCommon, vxmlNode, False, True)

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection15
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section15 node (Overall cost of the mortgage).
'** Parameters:     vxmlNode - the section15 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection15(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection15"

    On Error GoTo ErrHandler
    
    BuildLifeTimeSection15 mobjCommon, vxmlNode, False
    
    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection16
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the Section16 node (Using a mortgage intermediary).
'** Parameters:     vxmlNode - the section16 element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection16(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection16"

    On Error GoTo ErrHandler

    Call BuildUsingIntermediarySection(mobjCommon, vxmlNode, False)

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildContactsSection
'** Created by:     Srini Rao
'** Date:           29/06/2004
'** Description:    Builds the contact section.
'** Parameters:     vxmlNode - the contact section element to build.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildContactsSection(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildContactsSection"

    On Error GoTo ErrHandler

    Call BuildLifeTimeContactsSection(mobjCommon, vxmlNode)

    Exit Sub
ErrHandler:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub


