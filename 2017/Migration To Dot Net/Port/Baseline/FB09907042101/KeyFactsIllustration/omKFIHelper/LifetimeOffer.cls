VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LifetimeOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMortgage

Private Const mcstrModuleName As String = "LifetimeOffer"
Private mobjCommon As CommonDataHelper

Private Sub Class_Terminate()
    Set mobjCommon = Nothing 'TK 18/11/2004 BBG1672
End Sub

'********************************************************************************
'** Function:       IMortgage_FormatData
'** Created by:     Andy Maggs
'** Date:           07/06/2004
'** Description:    The method responsible for formatting the data into the XML
'**                 format required for the lifetime mortgage offer document
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

    On Error GoTo ErrHandler
    
    '*-store the common data helper object
    Set mobjCommon = vobjCommon
    
    '*-create the TEMPLATEDATA element and set the attributes
    Set xmlElem = mobjCommon.Document.createElement("TEMPLATEDATA")
    Set xmlTemplate = mobjCommon.Document.appendChild(xmlElem)
    Call BuildTemplateData(xmlTemplate)
    
    ' BC 05/Jan/2004 - E2EM00003121 - Remove COPYTO element
    '*-create the COPYTO element and set the attributes
'    If GetNumberOfApplicants(vobjCommon) > 1 Then  'SR 16/09/2004 : CORE82
'        'Set xmlNode = mobjCommon.CreateNewElement("COPYTO", xmlTemplate) 'SR 23/09/2004 : CORE82
'        Call BuildCopyTo(xmlTemplate) 'SR 23/09/2004 : CORE82
'    End If  'SR 16/09/2004 : CORE82
    
    '*-create the mandatory AUTONUMBERING element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate)
    Call AddAutonumberingAttributes(xmlNode)
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate) 'SR 16/09/2004 : CORE82
    Call AddAutonumberingAttributes(xmlNode, "RISK") 'SR 16/09/2004 : CORE82
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate) 'SR 24/09/2004 : CORE82
    Call AddAutonumberingAttributes(xmlNode, "OFFERCONDITION") 'SR 24/09/2004 : CORE82


    '*-create the mandatory SECTION1 (About this illustration) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION1", xmlTemplate)
    Call BuildSection1(xmlNode)

    '*-create the mandatory SECTION2 (Which service did we provide you with?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION2", xmlTemplate)
    Call BuildSection2(xmlNode)

    '*-create the mandatory SECTION3 (What is a lifetime mortgage?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION3", xmlTemplate)
    Call BuildSection3(xmlNode)

    '*-create the mandatory SECTION4 (Your mortgage requirements) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION4", xmlTemplate)
    Call BuildSection4(xmlNode)

    '*-create the mandatory SECTION5 (Description of this mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION5", xmlTemplate)
    Call BuildSection5(xmlNode)

    '*-create the mandatory SECTION6 (Benefits) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION6", xmlTemplate)
    Call BuildSection6(xmlNode)

    '*-create the mandatory SECTION7 (Risks - important things you must consider) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION7", xmlTemplate)
    Call BuildSection7(xmlNode)
    
    '*-create the mandatory SECTION8 (What you will owe and when) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION8", xmlTemplate)
    Call BuildSection8(xmlNode)

    '*-create the mandatory SECTION9 (Will the interest rate change?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION9", xmlTemplate)
    Call BuildSection9(xmlNode)

    '*-create the mandatory SECTION10 (How the value of your home could change) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION10", xmlTemplate)
    Call BuildSection10(xmlNode)

    '*-create the mandatory SECTION11 (What fees you must pay) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION11", xmlTemplate)
    Call BuildSection11(xmlNode)

    '*-create the mandatory SECTION12 (Insurance) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION12", xmlTemplate)
    Call BuildSection12(xmlNode)

    '*-create the mandatory SECTION13 (What happens if you do not want this mortgage any more?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION13", xmlTemplate)
    Call BuildSection13(xmlNode)

    '*-create the mandatory SECTION14 (Additional Features) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    Call BuildSection14(xmlNode)
    
    '*-create the mandatory SECTION15 (Overall cost of this mortgage) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION15", xmlTemplate)
    Call BuildSection15(xmlNode)
    
    '*-create the mandatory SECTION16 (Using a mortgage intermediary) element
    If vobjCommon.IntroducedByIntermediaryProperty Then ' SR 23/10/2004 : CORE82
        Set xmlNode = mobjCommon.CreateNewElement("SECTION16", xmlTemplate)
        Call BuildSection16(xmlNode)
    End If ' SR 23/10/2004 : CORE82
    
    '*-create the CONDITIONSECTION element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("CONDITIONSECTION", xmlTemplate)
    Call BuildConditionSection(xmlNode)
    
    '*-create the CONTACTSECTION element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("CONTACTSECTION", xmlTemplate)
    Call BuildContactSection(xmlNode)

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
'** Date:           07/06/2004
'** Description:    Builds the template data section of the document.
'** Parameters:     vxmlNode - the node to add the template data to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildTemplateData(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildTemplateData"

    On Error GoTo ErrHandler

    Call BuildCommonOfferTemplateData(mobjCommon, vxmlNode)
    
    'SR 16/09/2004 : CORE82
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE2", mobjCommon.MortgageTypeText2)
    xmlSetAttributeValue vxmlNode, "PROVIDER", mobjCommon.Provider
    xmlSetAttributeValue vxmlNode, "MORTGAGETYPECAPITALS", UCase(mobjCommon.MortgageTypeText2)
    'SR 16/09/2004 : CORE82 - End
    
    SetSVR vxmlNode, mobjCommon 'SR 20/09/2004 : CORE82

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

' BC 05/Jan/2004 - E2EM00003121 - Remove COPYTO element
'********************************************************************************
'** Function:       BuildCopyTo
'** Created by:     Andy Maggs
'** Date:           07/06/2004
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
'    errCheckError cstrFunctionName, mcstrModuleName
'End Sub

'********************************************************************************
'** Function:       BuildSection1
'** Created by:     Andy Maggs
'** Date:           08/06/2004
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
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection2
'** Created by:     Andy Maggs
'** Date:           08/06/2004
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
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection3
'** Created by:     Andy Maggs
'** Date:           08/06/2004
'** Description:    Sets the elements and attributes for the compulsory section3
'**                 (What is a lifetime mortgage?)
'** Parameters:     vxmlNode - the section3 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection3"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection3(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection4
'** Created by:     Andy Maggs
'** Date:           08/06/2004
'** Description:    Sets the elements/attributes for the compulsory section4
'**                 (Your mortgage requirements) element.
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection4(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection4"
    Dim xmlPurposeNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim blnValuationUsed As Boolean
    Dim xmlIncomeRequirements As IXMLDOMNode
    
    'SR 17/01/2005 : BBG082/E2EM00002927
    Dim lngValue As Long
    Dim rblnIsValuation As Boolean
    'SR 17/01/2005 : BBG082/E2EM00002927 - End
    
    On Error GoTo ErrHandler

    '*-add the ESTIMATEDVALUE attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, False, blnValuationUsed)
    '*-add TERMYEARS and TERMMONTHS attributes
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS")
    '*-add the LTV attribute
    Call xmlSetAttributeValue(vxmlNode, "LTV", CStr(mobjCommon.LTV))

    '*-now add one of the mandatory PURPOSEMORTGAGE, PURPOSEREMORTGAGE or
    '*-PURPOSEADDITIONALBORROWING elements to the PURPOSE node
    Select Case mobjCommon.LoanPurposeText
        Case "PURPOSEPURCHASE"
            Call mobjCommon.CreateNewElement("PURPOSEMORTGAGE", vxmlNode)
        Case "PURCHASEREMORTGAGE"
            Call mobjCommon.CreateNewElement("PURPOSEREMORTGAGE", vxmlNode)
        Case "PURPOSEFURTHERADVANCE"
            Call mobjCommon.CreateNewElement("PURPOSEADDITIONALBORROWING", vxmlNode)
    End Select
    
    '*-add the mandatory LUMPSUM element
    Set xmlNewNode = mobjCommon.CreateNewElement("LUMPSUM", vxmlNode)
    '*-add the TOTALLOANAMOUNT attribute
    Call xmlSetAttributeValue(xmlNewNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    
    '*-add the appropriate (1-16) mandatory FEESANDPREMIUMS element (exactly one)
    Call AddFeesAndPremiumsElement(mobjCommon, xmlNewNode)

    If mobjCommon.LoanPurposeText = "PURPOSEPURCHASE" Then
        Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(vxmlNode)
        '*-add the PURCHASEPRICE attribute
        'E2EM00002842 This will always be purchase price
        Call AddPurchPriceValueAttribute(mobjCommon.Data, xmlPurposeNode)
'        Call AddPurchPriceEstValueAttribute(mobjCommon.Data, xmlPurposeNode, True, _
'                blnValuationUsed)
    End If

    'SR 18/01/2005 : BBG082/E2EM00002927
    If mobjCommon.GetMainMortgageTypeGroup = "N" Then
        Call PurchasePriceOrEstimatedValue(mobjCommon.Data, True, rblnIsValuation, lngValue)
        If rblnIsValuation Then
            Set xmlNewNode = mobjCommon.CreateNewElement("LTVCALCULATIONUSEVALUATION", vxmlNode)
        Else
            Set xmlNewNode = mobjCommon.CreateNewElement("LTVCALCULATIONUSEPURCHASEPRICE", vxmlNode)
        End If
        
        Call xmlSetAttributeValue(xmlNewNode, "LTV", CStr(mobjCommon.LTV))
    End If
    'SR 18/01/2005 : BBG082/E2EM00002927 - End

    '*-add OWNHOUSE element
    If IsPropertyFreeHold(mobjCommon) Then
        Call mobjCommon.CreateNewElement("OWNHOUSE", vxmlNode)
    End If
    
    '*-add the CLIENTAGE element and AGE attribute
    Set xmlNewNode = mobjCommon.CreateNewElement("CLIENTAGE", vxmlNode)
    Call xmlSetAttributeValue(xmlNewNode, "AGE", GetYoungestApplicantsAge(mobjCommon))
    
    '*-add the NOOTHERMORTGAGE/OTHERMORTGAGE element
    '* This adds the details of mortgage loans (if any existing)
    '*-on this property.
    Call AddOtherMortgagesData(mobjCommon, vxmlNode)

    '*-add the INCOMEREQUIREMENTS element
    Set xmlIncomeRequirements = mobjCommon.CreateNewElement("INCOMEREQUIREMENTS", vxmlNode)
'IK 01/12/2004 E2EM00003126 LUMPSUMWITHDRAWDOWN mods
    If mobjCommon.HasDrawDown Then
        Set xmlNewNode = mobjCommon.CreateNewElement("LUMPSUMWITHDRAWDOWN", xmlIncomeRequirements)
        xmlSetAttributeValue xmlNewNode, "LUMPSUMAMOUNT", mobjCommon.AmountRequestedLessDrawDown
        xmlSetAttributeValue xmlNewNode, "DRAWDOWNAMOUNT", mobjCommon.DrawDownAmount
    Else
        Set xmlNewNode = mobjCommon.CreateNewElement("LUMPSUM", xmlIncomeRequirements)
        xmlSetAttributeValue xmlNewNode, "LUMPSUMAMOUNT", mobjCommon.LoanAmount
    End If
'IK 01/12/2004 E2EM00003126 ends

    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlIncomeRequirements = Nothing
Exit Sub
ErrHandler:
    Set xmlPurposeNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlIncomeRequirements = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection5
'** Created by:     Andy Maggs
'** Date:           09/06/2004
'** Description:    Adds the elements and attributes for the compulsory Section 5
'**                 (Description of this mortgage).
'** Parameters:     vxmlNode - the section 5 node to add the elements and
'**                 attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection5(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection5"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlPurpose As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Call BuildLifetimeSection5(mobjCommon, vxmlNode)
    
    '*-add the offer document specific bits
    '*-add the MORTGAGETYPE and PROVIDER attribute to various elements,
    '*-if present
    Set xmlItem = vxmlNode.selectSingleNode("ROLLUP")
    Call AddMortgageTypeAttribute(mobjCommon, xmlItem)
    Call xmlSetAttributeValue(xmlItem, "PROVIDER", mobjCommon.Provider)
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode.selectSingleNode("ROLLUP/LUMPSUM"))
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode.selectSingleNode("MULTICOMPONENT"))
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode.selectSingleNode("INSURANCE"))
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode.selectSingleNode("RESTRICTIONS/FINANCIALSTATUS"))
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode.selectSingleNode("CATSTANDARD"))
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode.selectSingleNode("LTV"))
    Call AddMortgageTypeAttribute(mobjCommon, vxmlNode)
    
    Set xmlList = vxmlNode.selectNodes(".//FULLTERM")
    For Each xmlItem In xmlList
        Call AddMortgageTypeAttribute(mobjCommon, xmlItem)
    Next xmlItem
    
    Set xmlList = vxmlNode.selectNodes(".//REMAININGTERM")
    For Each xmlItem In xmlList
        Call AddMortgageTypeAttribute(mobjCommon, xmlItem)
    Next xmlItem
        
    '*-now add the PROVIDER attribute to the root section node
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", mobjCommon.Provider)
    
    '*-remove the PURPOSEPURCHASE element from the LTV element if present
    Set xmlItem = vxmlNode.selectSingleNode("LTV")
    If Not xmlItem Is Nothing Then
        Set xmlPurpose = xmlItem.selectSingleNode("PURPOSEPURCHASE")
        If Not xmlPurpose Is Nothing Then
            Call xmlItem.removeChild(xmlPurpose)
        End If
    End If
    
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlPurpose = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlPurpose = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section6
'**                 (Benefits)
'** Parameters:     vxmlNode - the section6 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection6(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section7
'**                 (Risks - important things you must consider)
'** Parameters:     vxmlNode - the section7 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection7(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection8
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section8
'**                 (What you will owe and when)
'** Parameters:     vxmlNode - the section8 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection8(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection8"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection8(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection9
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section9
'**                 (Will the interest rate change?)
'** Parameters:     vxmlNode - the section9 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection9(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection9"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection9(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection10
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section10
'**                 (How the value of your home could change)
'** Parameters:     vxmlNode - the section10 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection10(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection10"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection10(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection11
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section11
'**                 (What fees you must pay)
'** Parameters:     vxmlNode - the section11 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection11(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection11"
    Dim xmlNotAdded As IXMLDOMNode

    On Error GoTo ErrHandler

    Call BuildWhatFeesYouMustPaySection(mobjCommon, vxmlNode, True)
    '*-need to remove the FEENAMES attribute from the NOTADDEDTOLOAN element
    Set xmlNotAdded = vxmlNode.selectSingleNode("FEESPAYABLE/NOTADDEDTOLOAN")
    If Not xmlNotAdded Is Nothing Then
        Call xmlNotAdded.Attributes.removeNamedItem("FEENAMES")
    End If

    Set xmlNotAdded = Nothing
Exit Sub
ErrHandler:
    Set xmlNotAdded = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection12
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section12
'**                 (Insurance) element.
'** Parameters:     vxmlNode - the section12 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection12(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection12"

    On Error GoTo ErrHandler

    Call BuildCommonOfferInsuranceSection(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection13
'** Created by:     Andy Maggs
'** Date:           16/06/2004
'** Description:    Sets the elements and attributes for the compulsory section13
'**                 (What happens if you do not want this mortgage any more?) element.
'** Parameters:     vxmlNode - the section13 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection13(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection13"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection13(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection14
'** Created by:     Andy Maggs
'** Date:           18/06/2004
'** Description:    Sets the elements and attributes for the compulsory section14
'**                 (Additional Features) element.
'** Parameters:     vxmlNode - the section14 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection14(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection14"

    On Error GoTo ErrHandler

    Call BuildAdditionalFeaturesSection(mobjCommon, vxmlNode, True, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection15
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section15
'**                 (Overall cost of this mortgage)
'** Parameters:     vxmlNode - the section15 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection15(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection15"

    On Error GoTo ErrHandler

    Call BuildLifeTimeSection15(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildSection16
'** Created by:     Andy Maggs
'** Date:           10/06/2004
'** Description:    Sets the elements and attributes for the compulsory section16
'**                 (Using a mortgage intermediary)
'** Parameters:     vxmlNode - the section16 element to populate.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection16(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection16"

    On Error GoTo ErrHandler

    Call BuildUsingIntermediarySection(mobjCommon, vxmlNode, True)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildConditionSection
'** Created by:     Andy Maggs
'** Date:           11/06/2004
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
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildContactSection
'** Created by:     Andy Maggs
'** Date:           11/06/2004
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
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

