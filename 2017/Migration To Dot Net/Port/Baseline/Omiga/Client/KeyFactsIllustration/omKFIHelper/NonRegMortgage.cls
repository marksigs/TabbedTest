VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NonRegMortgage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         NonRegMortgage
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Contains the functionality to create the appropriate KFI XML
'**                 for a Quick Quote Non Regulated illustration.
'********************************************************************************
'
'   Prog    Date        AQR         Description
'   PB      20/03/2007  EP2_1931    No redemption fees for ABO mortgages
'
Option Explicit

Implements IMortgage

Private mobjCommon As CommonDataHelper

Private Sub Class_Terminate()
    Set mobjCommon = Nothing 'TK 18/11/2004 BBG1672
End Sub

'********************************************************************************
'** Function:       IMortgage_FormatData
'** Created by:     Ian Ross
'** Date:           29/09/2004
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

    '*-create the mandatory SECTION6A (What you will need to pay in the future) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION6A", xmlTemplate)
    
    '*-create the mandatory SECTION7 (Are you comfortable with the risks) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION7", xmlTemplate)

    '*-create the mandatory SECTION8 (What fees must you pay?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION8", xmlTemplate)
    Call BuildSection8(xmlNode)

    '*-create the mandatory SECTION9 (Insurance) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION9", xmlTemplate)

    '*-create the mandatory SECTION10 (What happens if you do not want this mortgage any more?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION10", xmlTemplate)
    Call BuildSection10(xmlNode)


    '*-create the mandatory SECTION12 (Additional Features) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION12", xmlTemplate)

    '*-create the mandatory SECTION14 (Where can you get more information on mortgages?) element
    Set xmlNode = mobjCommon.CreateNewElement("SECTION14", xmlTemplate)

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
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Builds the template data section of the document.
'** Parameters:     vxmlNode - the node to add the template data to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildTemplateData(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildTemplateData"

    On Error GoTo ErrHandler

    '*-add the mandatory PROVIDER attribute
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", mobjCommon.Provider)
    '*-add the mandatory VALIDTODATE attribute
    Call AddValidToDateAttribute(mobjCommon.Data, vxmlNode)
    '*-add the mandatory APPLICANTNAME attribute
    Call AddApplicantNameAttribute(mobjCommon.Data, vxmlNode)
    '*-add the mandatory REFERENCE attribute
    Call AddReferenceAttribute(mobjCommon.Data, vxmlNode)
    '*-add the mandatory TODAY attribute
    Call AddTodayAttribute(vxmlNode)
    '*-add the mandatory MORTGAGETYPE  attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)
    '*-add the DOCUMENTNAME attribute
    Call xmlSetAttributeValue(vxmlNode, "DOCUMENTNAME", "Key Facts Illustration")
    '*-add the SPECIALGROUP attribute from SPECIALGROUP_TEXT
    Call AddSpecialGroupNameAttribute(mobjCommon.Data, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection1
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the attributes for the compulsory section1 (About this
'**                 illustration) element.
'** Parameters:     vxmlNode - the section1 element to set the attributes on.
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
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection3
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the attributes for the compulsory section3 (What you
'**                 have told us) element.
'** Parameters:     vxmlNode - the section3 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection3(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection3"
    Dim blnValuationUsed As Boolean

    On Error GoTo ErrHandler

    '* Add TOTALLOANAMOUNT as attribs to SECTION3
    Call xmlSetAttributeValue(vxmlNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    
    '* Add TERMYEARS and TERMMONTHS as attribs to SECTION3
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS")
    
    '*-add the ESTIMATEDVALUE element
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, False, _
            blnValuationUsed)

    '*-add the type of mortgage mandatory REPAYMENT, INTERESTONLY or
    '*-PARTANDPART elements
    Call mobjCommon.AddMortgageRepaymentTypeElement(vxmlNode)
        
   
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection4
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element.
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection4(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection4"

    On Error GoTo ErrHandler
    Dim xmlItem As IXMLDOMNode
    Dim xmlItems As IXMLDOMNodeList
    Dim xmlProduct As IXMLDOMNode
    
    '*-get the loan components
    Set xmlItems = mobjCommon.LoanComponents
    If xmlItems.length = 1 Then
        '*-get the single loan component
        '*-Will only have single loan component
        Set xmlItem = xmlItems.Item(0)
        
        '*-add the mandatory PROVIDER attribute
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", mobjCommon.Provider)
        
        '*-add the SINGLECOMPONENT element
        Set xmlProduct = mobjCommon.CreateNewElement("SINGLECOMPONENT", vxmlNode)
        
        '*-add the SPECIALGROUP attribute from SPECIALGROUP_TEXT
        Call AddSpecialGroupNameAttribute(mobjCommon.Data, xmlProduct)
        
        '*-add the PRODUCT attribute
        Call AddProductNameAttribute(xmlItem, xmlProduct, mobjCommon)
    End If

    Set xmlItem = Nothing
    Set xmlItems = Nothing
    Set xmlProduct = Nothing
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlItems = Nothing
    Set xmlProduct = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection5
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the elements and attributes for the compulsory section5
'**                 (Overall Costs) element.
'** Parameters:     vxmlNode - the section5 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection5(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection5"

    On Error GoTo ErrHandler
    '*-add the TOTALREPAYMENTAMOUNT attribute
    Call xmlSetAttributeValue(vxmlNode, "TOTALREPAYMENTAMOUNT", _
            CStr(mobjCommon.LoanAmountRepayable))

    '*-add the APR attribute
    Call AddAPRAttribute(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6
'** Created by:     Ian Ross
'** Date:           29/09/2004
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
    Dim blnInterestOnly As Boolean
    Dim lngIOAmount As Long
    Dim intIndex As Integer
    Dim xmlPurposeNode As IXMLDOMNode, xmlNode As IXMLDOMNode
    
    On Error GoTo ErrHandler

    blnInterestOnly = False
    lngIOAmount = 0
    Set xmlComponentNode = mobjCommon.AddComponentsTypeElement(vxmlNode)
    
    If xmlComponentNode.baseName = "SINGLECOMPONENT" Then
        '*-get the single loan component
        Set xmlLoanComponent = mobjCommon.SingleLoanComponent
        If xmlLoanComponent Is Nothing Then
            '*-if there is no loan component we can't really do a lot here!
            Exit Sub
        End If
        
        '*-get the repayment method
        strRepayMethod = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD")
        
        '*-add the loan component interest rate data
        Call AddNRLoanComponentInterestRateData(mobjCommon, xmlLoanComponent, _
                xmlComponentNode, False)
    End If
    
    If blnInterestOnly Then
        '*-add the INTERESTONLY element
        Set xmlNewNode = mobjCommon.CreateNewElement("INTERESTONLY", vxmlNode)
        '*-add the mandatory INTERESTONLYTOTAL attribute
        Call xmlSetAttributeValue(xmlNewNode, "INTERESTONLYTOTAL", _
                    CStr(mobjCommon.LoanAmount + mobjCommon.FeesAddedToLoanAmount))  'SR 21/09/2004 : CORE82

        'CORE82
        Set xmlPurposeNode = mobjCommon.AddLoanPurposeElement(xmlNewNode)
    End If

    Set xmlComponentNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLCList = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlPurposeNode = Nothing
    Set xmlNode = Nothing
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
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection8
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the elements and attributes for the compulsory section8
'**                 (What fees must you pay?) element.
'** Parameters:     vxmlNode - the section8 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection8(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection8"

    On Error GoTo ErrHandler

    Call BuildNonRegFeesYouMustPaySection(mobjCommon, vxmlNode, False)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection10
'** Created by:     Ian Ross
'** Date:           29/09/2004
'** Description:    Sets the elements and attributes for the compulsory section10
'**                 (What happens if you do not want this mortgage any more?) element.
'** Parameters:     vxmlNode - the section10 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection10(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection10"

    On Error GoTo ErrHandler

    'PB EP2_1931 20/03/2007 Begin
    If Not mobjCommon.IsAdditionalBorrowing Then
        Call NRBuildEarlyRepayCharges(mobjCommon, vxmlNode)
    End If
    'PB EP2_1931 End
    
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub


