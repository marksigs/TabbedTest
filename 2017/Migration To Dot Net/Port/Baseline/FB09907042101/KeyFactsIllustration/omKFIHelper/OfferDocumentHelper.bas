Attribute VB_Name = "OfferDocumentHelper"
'********************************************************************************
'** Module:         OfferDocumentHelper
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    This module contains functions/procedures that are common to
'**                 all KFI offer document types.
'********************************************************************************
'
'BBG Specific History
'
'Prog   Date        Description
'
'AW     18/01/2005  E2EM00003187 - Removed MDay date processing.
'HMA    19/01/2005  E2EM00002920 - Correct Solicitor contact details.
'BC     14/09/2005  MARS054 - Changes for MARS project.
'BC     16/10/2005  MAR88 - Changes for MARS project.
'BC     16/11/2005  MAR589 - KFI Updates
'BC     04/01/2006  MAR907 - Product Switches and ToE
'HMA    02/02/2006  MAR1071  Use the global parameter for Maximum LTV allowed.
'BC     02/03/2006  MAR1347  Use first Validation Type from RepaymentMethod, rather than the second
'********************************************************************************
'
'Epsom Specific History
'
'Prog   Date        Description
'
'SAB    03/05/2006  EP479   Amended BuildCommonOfferContactSection to get Intermediary Contact Details
'SAB    10/05/2006  EP489   Amended BuildCommonOfferSection4 to change PROVIDER text based on the NatureOfLoan
'PB     16/05/2006  EP529   MAR1731 - Fulfilment - Offer Document - incorrect validity period shown
'PB     24/05/2006  EP603   MAR1777 Changed BuildCommonOfferSection6
'PB     24/05/2006  EP603   MAR1736 Show LC Seq Number correctly in BuildCommonOfferSection4
'DRC    25/05/2006  EP610   Epsom specific changes as noted
'PB     05/06/2006  EP651   MAR1590  modified sections BuildCommonOfferSection4 and BuildCommonOfferSection6
'                           - show two rows if a loan is Part & Part
'PB/PM  12/06/2006  EP697   Offer Document Changes
'PB     12/06/2006  EP730   MAR1831 Changes to Sections 4 and 6 for multicomponent, Part and Part loans
'PM     20/06/2006  EP809   Updated BuildCommonOfferInsuranceSection (Section 9) different wording for mortgage/remortgage
'PE     20/06/2006  EP773   Amended BuildCommonOfferTemplateData. Add PropertyLocation to template.
'PM     23/06/2006  EP824   Amended BuildCommonOfferSection6. Add RepaymentVehicle to INTERESTONLY element
'PM     23/06/2006  EP769   Amended BuildCommonOfferSection4Single
'PB     26/06/2006  EP773   Get property address and description from NewProperty if valuation details not available
'PB     30/06/2006  EP929   Amended code to cope with missing addresses better
'DC     12/07/2006  EP983   Reverse MAR1071
'PB     22/08/2006  EP1082  Ensure other title is shown correctly in offer doc
'PE     24/08/2006  EP1099  CC71 - Phase 1 - Offer Template Amendment
'PE     13/09/2006  EP1127  Offer Conditions on the Offer Template formatting not as per Epsom CC71
'PE     15/09/2006  EP1144  Changed AddConditions to allow paragraphs in offer conditions.
'LH     13/10/2006  EP1226 / CC132 Added trading name description
'AW     23/10/2006  EP1238 / CC203 Amended trading name to 'DB UK'
'PB     21/11/2006  EP2_139  Added TypeOfMortgage element to section 2
'PB     29/11/2006  EP2_139  Other offer doc changes.
'PB     13/12/2006  EP2_422 KFI changes
'INR    14/02/2007  EP2_1401    Removed using global KFIMDay
'PB     28/02/2007  EP2_1558    Added text for TOE
'INR    01/03/2007  EP2_1449    Section9 update
'INR    03/03/2007  EP2_1382    Sort Type Mismatch
'PB     13/03/2007  EP2_1932    Check lender name not id, as this could change across systems
'PB     15/03/2007  EP2_1931    Several small fixes
'INR    20/03/2007  EP2_1977  Deal with REGULATIONINDICATOR_VALIDID
'SR     25/03/2007  EP2_1938  - Changes to building Section6A
'INR    26/03/2007  EP2_1983/1984 FT section is required for single component as well
'SR     26/03/2007  EP2_1938 - modified method 'BuildCommonOfferSectino6A'
'INR    27/03/2007  EP2_2057    Contact Telephone details
'INR    04/04/2007  EP2_2042    rework section6
'SR     12/04/2007  EP2_2341
'INR    15/04/2007  EP2_2395    Baserate description change
'INR    17/04/2007  EP2_2448 if coming from pre-dip KFI, don't have the specialgroup stuff
'IK     18/04/2007  EP2_2458 (repayment) PLANTYPE qualified in section 6
'INR    18/04/2007  EP2_2478 Various
'********************************************************************************
Option Explicit

Private Const mcstrModuleName As String = "OfferDocumentHelper"

Private Type AddressDataType
    strNameOrNumber As String
    strPostcode As String
End Type

Public mdblMaxLTV As Double

'********************************************************************************
'** Function:       BuildCommonOfferTemplateData
'** Created by:     Andy Maggs
'** Date:           07/06/2004
'** Description:    Builds the template data that is common to all offer
'**                 documents.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferTemplateData(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferTemplateData"
    Dim xmlDetails As IXMLDOMNode
    Dim xmlPropertyLocation As IXMLDOMNode
    Dim strPropertyLocation As String
    Dim strCondition As String
    'PB 26/06/2006 EP773 begin
    Dim xmlNewPropDetails As IXMLDOMNode
    'PB EP773 End
    Dim strValidEndDate As String
    On Error GoTo ErrHandler

    '*-add the MORTGAGETYPE1 attribute
    AddMortgageType1 vobjCommon, vxmlNode
    '*-add the APPLICANTNAME attribute
    Call AddApplicantNameAttribute(vobjCommon.Data, vxmlNode)
    '*-add the applicant address information
    'APPLICANTADDRESSLINE1, APPLICANTADDRESSLINE2, APPLICANTADDRESSLINE3
    'APPLICANTPOSTCODE
    Call AddApplicantAddressAttributes(vobjCommon.Data, vxmlNode)
    '*-add the DOCUMENTNAME attribute
    Call xmlSetAttributeValue(vxmlNode, "DOCUMENTNAME", "Offer Document")
    '*-add the ISSUEDATE attribute
    Call AddTodayAttribute(vxmlNode, "ISSUEDATE")
    '*-add the OFFERTYPE attribute
    Call xmlSetAttributeValue(vxmlNode, "OFFERTYPE", vobjCommon.MortgageTypeText)
    xmlSetAttributeValue vxmlNode, "OFFERTYPEA", vobjCommon.MortgageTypeText_A 'SR 13/10/2004 : BBG1596
    'EP2_
    '*-add the KFIVALIDENDDATE attribute
    strValidEndDate = DateAdd("m", 2, xmlGetAttributeText(vxmlNode, "ISSUEDATE"))
    xmlSetAttributeValue vxmlNode, "KFIVALIDENDDATE", strValidEndDate
    
    '*-get the latest valuation report property details object
    Set xmlDetails = vobjCommon.GetLatestValnReportPropertyDetails()
    'PB 26/06/2006 EP773 Begin
    'If Not xmlDetails Is Nothing Then
    '    '*-add the PROPERTYTYPE attribute
    '    Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "TYPEOFPROPERTY_TEXT", "PROPERTYTYPE")
    '    '*-add the PROPERTYDESCRIPTION attribute
    '    Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "PROPERTYDESCRIPTION_TEXT", "PROPERTYDESCRIPTION")
    '    '*-add the TENURE attribute
    '    Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "TENURE_TEXT", "TENURE")
    'End If
    Set xmlNewPropDetails = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH & "/NEWPROPERTY")
    'PB 30/06/2006 EP929 Begin
    If Not xmlDetails Is Nothing Then
        If xmlAttributeValueExists(xmlDetails, "TYPEOFPROPERTY_TEXT") = True Then
            Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "TYPEOFPROPERTY_TEXT", "PROPERTYTYPE")
        Else
            Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "TYPEOFPROPERTY_TEXT", "PROPERTYTYPE")
        End If
        If xmlAttributeValueExists(xmlDetails, "PROPERTYDESCRIPTION_TEXT") = True Then
            Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "PROPERTYDESCRIPTION_TEXT", "PROPERTYDESCRIPTION")
        Else
            Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "PROPERTYDESCRIPTION_TEXT", "PROPERTYDESCRIPTION")
        End If
        If xmlAttributeValueExists(xmlDetails, "TENURE_TEXT") = True Then
            Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "TENURE_TEXT", "TENURE")
        Else
            Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "TENURETYPE_TEXT", "TENURE")
        End If
    ElseIf Not xmlNewPropDetails Is Nothing Then
        Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "TYPEOFPROPERTY_TEXT", "PROPERTYTYPE")
        Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "PROPERTYDESCRIPTION_TEXT", "PROPERTYDESCRIPTION")
        Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "TENURETYPE_TEXT", "TENURE")
    Else
        ' Unable to get valuation details OR new property details
    End If
    '
    'If (Not xmlDetails Is Nothing) Or (Not xmlNewPropDetails Is Nothing) Then
    '    '*-add the PROPERTYTYPE attribute
    '    If xmlAttributeValueExists(xmlDetails, "TYPEOFPROPERTY_TEXT") = True Then
    '        Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "TYPEOFPROPERTY_TEXT", "PROPERTYTYPE")
    '    Else
    '        ' Can't get property type
    '        Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "TYPEOFPROPERTY_TEXT", "PROPERTYTYPE")
    '    End If
    '    '*-add the PROPERTYDESCRIPTION attribute
    '    If xmlAttributeValueExists(xmlDetails, "PROPERTYDESCRIPTION_TEXT") = True Then
    '        Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "PROPERTYDESCRIPTION_TEXT", "PROPERTYDESCRIPTION")
    '    Else
    '        ' Can't find it
    '        Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "PROPERTYDESCRIPTION_TEXT", "PROPERTYDESCRIPTION")
    '    End If
    '    '*-add the TENURE attribute
    '    If xmlAttributeValueExists(xmlDetails, "TENURE_TEXT") = True Then
    '        Call xmlCopyAttributeValue(xmlDetails, vxmlNode, "TENURE_TEXT", "TENURE")
    '    Else
    '        Call xmlCopyAttributeValue(xmlNewPropDetails, vxmlNode, "TENURETYPE_TEXT", "TENURE")
    '    End If
    'End If
    'PB EP773 End
    
    '*-add the REFERENCE attribute
    Call AddReferenceAttribute(vobjCommon.Data, vxmlNode)
    '*-add the SECURITYADDRESS attribute
    Call AddSecurityAddress(vobjCommon.Data, vxmlNode)
    '*-add the TODAY attribute
    Call AddTodayAttribute(vxmlNode)
    '*-add the VALIDTODATE attribute
    Call AddOfferValidToDateAttribute(vobjCommon.Data, vxmlNode) 'SR 19/10/2004 : BBG1657


    'EPSOM EP417 - 21/04/2006 SAB - START
    '*-add the Solicitor Company Name and Address attribute
    Call AddLegalRepContactDetails(vobjCommon, vobjCommon.Data, vxmlNode)
    
    '*-add the Intermediary Company Name and Address attribute
    Call AddIntermediaryContactDetails(vobjCommon, vobjCommon.Data, vxmlNode, True)
    'EPSOM EP417 - 21/04/2006 SAB - END
    
    'EP773 - 20/06/2006 - Peter Edney - START
    strPropertyLocation = ""
    Set xmlPropertyLocation = vobjCommon.Data.selectSingleNode(gcstrPROPERTYLOCATION_PATH)
    If Not xmlPropertyLocation Is Nothing Then
        strPropertyLocation = xmlGetAttributeText(xmlPropertyLocation, "PROPERTYLOCATION")
    End If
    
    Select Case strPropertyLocation
    Case "10", ""
        strCondition = "This Offer incorporates the Mortgage Conditions 2006 Edition England and Wales."
    Case "20"
        strCondition = "This Offer incorporates the Mortgage Conditions 2006 Edition Scotland."
    Case "30"
        strCondition = "This Offer incorporates the Mortgage Conditions 2006 Edition Northern Ireland."
    End Select
    
    xmlSetAttributeValue vxmlNode, "MORTGAGECONDITION", strCondition
    'EP773 - 20/06/2006 - Peter Edney - END
    
    Set xmlDetails = Nothing
Exit Sub
ErrHandler:
    Set xmlDetails = Nothing

    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
' BC 05/Jan/2004 - E2EM00003121 - Remove COPYTO element
'********************************************************************************
'** Function:       BuildCommonOfferCopyTo
'** Created by:     Andy Maggs
'** Date:           24/05/2004
'** Description:    Builds the CopyTo element containing the addresses of
'**                 applicants to whom the Offer document should be copied to.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the node to add the document details to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
'Public Sub BuildCommonOfferCopyTo(ByVal vobjCommon As CommonDataHelper, _
'        ByVal vxmlNode As IXMLDOMNode)
'    Const cstrFunctionName As String = "BuildCommonOfferCopyTo"
'    Dim xmlApplicants As IXMLDOMNodeList
'    Dim xmlApplicant As IXMLDOMNode
'    Dim audtAddresses() As AddressDataType
'    Dim intNumAddresses As Integer
'    Dim intIndex As Integer
'    Dim blnFound As Boolean
'    Dim bCompanyFound As Boolean
'    Dim xmlAddressee As IXMLDOMNode
'    Dim xmlCustomer As IXMLDOMNode
'    Dim strName As String
'    Dim xmlAddress As IXMLDOMNode
'    Dim intApplicant As Integer
'    Dim strNameOrNum As String
'    Dim strPostcode As String
'    Dim strCustomerRoleType As String
'    Dim xmlCopyTo As IXMLDOMNode
'
'    On Error GoTo ErrHandler
'
'    '*-get the list of applicants
'    Set xmlApplicants = vobjCommon.Data.selectNodes("//CUSTOMERROLE")
'    If xmlApplicants.length > 1 Then
'        bCompanyFound = False
'        For Each xmlApplicant In xmlApplicants
'            strCustomerRoleType = xmlGetAttributeText(xmlApplicant, "CUSTOMERROLETYPE")
'            If strCustomerRoleType = "C," Then
'                bCompanyFound = True
'            End If
'        Next xmlApplicant
'        '*-iterate through the applicants (ignoring the first) and add the
'        '*-address details for each unique address in the list
'        For Each xmlApplicant In xmlApplicants
'            strCustomerRoleType = xmlGetAttributeText(xmlApplicant, "CUSTOMERROLETYPE")
'            If Not (strCustomerRoleType = "C,") Then
'                intApplicant = intApplicant + 1
'                '*-get the address with type 2 first, if one
'                Set xmlAddress = xmlApplicant.selectSingleNode("CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE=2]/ADDRESS")
'                If xmlAddress Is Nothing Then
'                    '*-there isn't so try to get the address with type 1
'                    Set xmlAddress = xmlApplicant.selectSingleNode("CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE=1]/ADDRESS")
'                End If
'
'                strNameOrNum = xmlGetAttributeText(xmlAddress, "BUILDINGORHOUSENUMBER")
'                If Len(strNameOrNum) = 0 Then
'                    strNameOrNum = xmlGetAttributeText(xmlAddress, "BUILDINGORHOUSENAME")
'                End If
'                strPostcode = xmlGetAttributeText(xmlAddress, "POSTCODE")
'
'                blnFound = False
'                For intIndex = 1 To intNumAddresses
'                    If audtAddresses(intIndex).strNameOrNumber = strNameOrNum _
'                            And audtAddresses(intIndex).strPostcode = strPostcode Then
'                        blnFound = True
'                        Exit For
'                    End If
'                Next intIndex
'
'                If Not blnFound Then
'                    '*-add this address to the list
'                    intNumAddresses = intNumAddresses + 1
'                    ReDim Preserve audtAddresses(1 To intNumAddresses)
'                    audtAddresses(intNumAddresses).strNameOrNumber = strNameOrNum
'                    audtAddresses(intNumAddresses).strPostcode = strPostcode
'
'                    '*-and add the applicant information to the node
'                    '*-add all individuals for Ltd Compant applications
'                    '*-but only second and subsequent applicants for all other applications
'                    If bCompanyFound = True Or intApplicant > 1 Then
'                        Set xmlCopyTo = vobjCommon.CreateNewElement("COPYTO", vxmlNode) 'SR 23/09/2004 : CORE82
'                        Set xmlAddressee = vobjCommon.CreateNewElement("COPYADDRESSEE", xmlCopyTo) 'SR 23/09/2004 : CORE82
'                        '*-add the name attribute
'                        Set xmlCustomer = xmlApplicant.selectSingleNode("CUSTOMERVERSION")
'                        strName = BuildName(xmlGetAttributeText(xmlCustomer, "TITLE_TEXT"), _
'                                xmlGetAttributeText(xmlCustomer, "FIRSTFORENAME"), _
'                                xmlGetAttributeText(xmlCustomer, "SURNAME"))
'                        Call xmlSetAttributeValue(xmlAddressee, "COPYNAME", strName)
'
'                        '*-add the address attributes
'                        Call AddAddressAttributes(xmlAddress, xmlAddressee, "COPYADDRESSLINE1", _
'                                "COPYADDRESSLINE2", "COPYADDRESSLINE3", "COPYPOSTCODE")
'                    End If
'                End If
'            End If
'        Next xmlApplicant
'    End If
'
'    Set xmlApplicants = Nothing
'    Set xmlApplicant = Nothing
'    Set xmlAddressee = Nothing
'    Set xmlCustomer = Nothing
'    Set xmlAddress = Nothing
'    Set xmlCopyTo = Nothing
'Exit Sub
'ErrHandler:
'    Set xmlApplicants = Nothing
'    Set xmlApplicant = Nothing
'    Set xmlAddressee = Nothing
'    Set xmlCustomer = Nothing
'    Set xmlAddress = Nothing
'    Set xmlCopyTo = Nothing
'    '*-check the error and throw it on
'    errCheckError cstrFunctionName, mcstrModuleName
'End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection1
'** Created by:     Andy Maggs
'** Date:           25/04/2004
'** Description:    Sets the attributes for the compulsory section1 (About this
'**                 illustration) element.
'** Parameters:     vxmlNode - the section1 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection1(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection1"
    Dim xmlQuote As IXMLDOMNode
    Dim xmlProvided As IXMLDOMNode
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
    Dim strConditions As String
    Dim xmlMortgageType As IXMLDOMNode
    Dim strMortgageType As String
    
    'EPSOM EP417 - 21/04/2006 SAB - START
    Dim xmlAppFF As IXMLDOMNode
    Dim xmlRegulated As IXMLDOMNode
    Dim strRegulationIndicatorValidId As String
    Dim intKFIReceivedByAllAppsInd As Integer
    'EPSOM EP417 - 21/04/2006 SAB - END
    
    Dim blnFurtherAdvance As Boolean ' PB 13/12/2006 EP2_422
    
    
    On Error GoTo ErrHandler
    
    '*-add the CONDITION attribute
'CORE82    Set xmlList = vobjCommon.Data.selectNodes(gcstrAPPLICATIONCONDITIONS_PATH)
'    For Each xmlItem In xmlList
'        strValue = xmlGetAttributeText(xmlItem, "CONDITIONDESCRIPTION")
'        If Len(strValue) > 0 Then
'            If Len(strConditions) = 0 Then
'                strConditions = strValue
'            Else
'                strConditions = strConditions & ", " & strValue
'            End If
'        End If
'    Next xmlItem
'    Call xmlSetAttributeValue(vxmlNode, "CONDITION", strConditions)

    '* Core82 -add the PURPOSEMORTGAGE or PURPOSEFURTHERADVANCE element
    strMortgageType = UCase(vobjCommon.MortgageTypeText)
    If (vobjCommon.blnIsProductSwitch) Then
        Set xmlMortgageType = vobjCommon.CreateNewElement("PURPOSEPRODUCTSWITCH", vxmlNode)
    ElseIf (StrComp(strMortgageType, "MORTGAGE") = 0) Then
        Set xmlMortgageType = vobjCommon.CreateNewElement("PURPOSEMORTGAGE", vxmlNode)
    ElseIf (StrComp(strMortgageType, "ADDITIONAL BORROWING") = 0) Then
        Set xmlMortgageType = vobjCommon.CreateNewElement("PURPOSEFURTHERADVANCE", vxmlNode)
        blnFurtherAdvance = True
    End If
    
    
    
    
    'EPSOM EP417 - 21/04/2006 SAB - START
    '*-the ILLUSTRATIONPROVIDED element is set dependant on the REGULATIONINDICATOR

'    '*-add the ILLUSTRATIONPROVIDED element
'    Set xmlQuote = vobjCommon.Data.selectSingleNode(gcstrQUOTATION_PATH)
'    If xmlGetAttributeAsBoolean(xmlQuote, "KFIPRINTEDINDICATOR") Then
'        Set xmlProvided = vobjCommon.CreateNewElement("ILLUSTRATIONPROVIDED", vxmlNode)
'        Call AddMortgageTypeAttribute(vobjCommon, xmlProvided)
'    End If

    Set xmlAppFF = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
    'EP2_1977 REGULATIONINDICATOR now holds ValidationId
    strRegulationIndicatorValidId = xmlGetAttributeText(xmlAppFF, "REGULATIONINDICATOR")
    'EP610 - DRC Separated this from illustration
    If CheckForValidationType(strRegulationIndicatorValidId, "R") Then     'R=Regulated
     '*-add the REGULATEDMORTGAGE element
        Set xmlRegulated = vobjCommon.CreateNewElement("REGULATEDMORTGAGE", vxmlNode)
        Call AddMortgageTypeAttribute(vobjCommon, xmlRegulated)
        'PB 13/12/2006 EP2_422
        If blnFurtherAdvance = False Then
            vobjCommon.CreateNewElement "NOTADDITIONALBORROWING", xmlRegulated
        End If
    'PB 13/12/2006 EP2_422
    Else
    '*-add the NONREGULATED element
        Set xmlRegulated = vobjCommon.CreateNewElement("NONREGULATED", vxmlNode)
        'PB 13/12/2006 EP2_422
        If blnFurtherAdvance = False Then
            vobjCommon.CreateNewElement "NOTADDITIONALBORROWING", xmlRegulated
        End If
    End If
    'EP610
    intKFIReceivedByAllAppsInd = xmlGetAttributeAsInteger(xmlAppFF, "KFIRECEIVEDBYALLAPPS")

    If intKFIReceivedByAllAppsInd = 1 Then              '1=YES 2=NO
        If CheckForValidationType(strRegulationIndicatorValidId, "R") Then     'R=Regulated
            '*-add the REGULATEDILLUSTRATIONPROVIDED element
            Set xmlQuote = vobjCommon.Data.selectSingleNode(gcstrQUOTATION_PATH)
'            If xmlGetAttributeAsBoolean(xmlQuote, "KFIPRINTEDINDICATOR") Then
                Set xmlProvided = vobjCommon.CreateNewElement("REGULATEDILLUSTRATIONPROVIDED", vxmlNode)
                Call AddMortgageTypeAttribute(vobjCommon, xmlProvided)
'            End If
        
           
        Else
            '*-add the ILLUSTRATIONPROVIDED element
            Set xmlQuote = vobjCommon.Data.selectSingleNode(gcstrQUOTATION_PATH)
'            If xmlGetAttributeAsBoolean(xmlQuote, "KFIPRINTEDINDICATOR") Then
                Set xmlProvided = vobjCommon.CreateNewElement("ILLUSTRATIONPROVIDED", vxmlNode)
                Call AddMortgageTypeAttribute(vobjCommon, xmlProvided)
'            End If
        End If
    End If
    'EPSOM EP417 - 21/04/2006 SAB - END



    Set xmlQuote = Nothing
    Set xmlProvided = Nothing
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlMortgageType = Nothing

    'EPSOM EP417 - 21/04/2006 SAB - START
    Set xmlAppFF = Nothing
    Set xmlRegulated = Nothing
    'EPSOM EP417 - 21/04/2006 SAB - END
Exit Sub

ErrHandler:
    Set xmlQuote = Nothing
    Set xmlProvided = Nothing
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlMortgageType = Nothing
    
    'EPSOM EP417 - 21/04/2006 SAB - START
    Set xmlAppFF = Nothing
    Set xmlRegulated = Nothing
    'EPSOM EP417 - 21/04/2006 SAB - END
    
    
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection2
'** Created by:     Andy Maggs
'** Date:           25/05/2004
'** Description:    Sets the attributes for the compulsory section2 (Which
'**                 service are we providing you with?) element.
'** Parameters:     vxmlNode - the section2 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection2(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection2"
    Dim xmlNewNode As IXMLDOMNode
    Dim xmlTOENodes As IXMLDOMNodeList
    Dim intTOECustomersAdded As Integer
    Dim intTOECustomersRemoved As Integer
    Dim strAddedOrRemoved As String
    
    Dim strIntermediaryContactName As String
    Dim strIntermediaryCompanyName As String
    
    'AW 18/01/2005  E2EM00003187
    'MDay date processing no longer required.
    'SR 21/10/2004: BBG1676
'    Dim xmlMDay As IXMLDOMNode, xmlKFIDocAuditDetails As IXMLDOMNode, xmlAFF As IXMLDOMNode
'    Dim dtMDay As Date, dtKFIPrintDate As Date
'    Dim xmlKFIDate As IXMLDOMNode, intRegulationIndicator As Integer
    'SR 21/10/2004: BBG1676 - End
    'AW 18/01/2005  E2EM00003187 - End
    
    On Error GoTo ErrHandler
    'SR 21/10/2004: BBG1676
    
    'AW 18/01/2005  E2EM00003187
'    Set xmlMDay = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "KFIMDay" & Chr$(34) & "]")
'    dtMDay = xmlGetAttributeAsDate(xmlMDay, "STRING")
'
'    Set xmlAFF = vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH)
'    intRegulationIndicator = xmlGetAttributeAsInteger(xmlAFF, "REGULATEDMORTGAGEINDICATOR")
'
'    Set xmlKFIDocAuditDetails = vobjCommon.Data.selectSingleNode(gcstrAPPLICATION_PATH & "/DOCUMENTAUDITDETAILS")
'    If (xmlKFIDocAuditDetails Is Nothing) Or (intRegulationIndicator <> 1) Then
'        Set xmlKFIDate = vobjCommon.CreateNewElement("KFIDATELT311004", vxmlNode)
'    Else
'        dtKFIPrintDate = xmlGetAttributeAsDate(xmlKFIDocAuditDetails, "PRINTDATE")
'        If dtKFIPrintDate >= dtMDay Then
'            Set xmlKFIDate = vobjCommon.CreateNewElement("KFIDATEGT301004", vxmlNode)
'        Else
'            Set xmlKFIDate = vobjCommon.CreateNewElement("KFIDATELT311004", vxmlNode)
'        End If
'    End If
     'SR 21/10/2004: BBG1676 - End
    'AW 18/01/2005  E2EM00003187 - End
    
   
' BBG1739  BCoates 04Nov04
'    If vobjCommon.IsIntroducedByIntermediary(strIntermediaryName) Then
    'BC MAR907
     If (vobjCommon.blnIsProductSwitch) Then
        '*-add the PRODUCTSWITCH element
        Set xmlNewNode = vobjCommon.CreateNewElement("PRODUCTSWITCH", vxmlNode)
     Else
        ' Check whether Transfer of Equity
        If vobjCommon.blnIsTransferofEquity Then
            ' Add TRANSFEROFEQUITY element
            Set xmlNewNode = vobjCommon.CreateNewElement("TRANSFEROFEQUITY", vxmlNode)
            ' Add TOEADDITIONREMOVAL and TOEADDITIONREMOVALPARTIES
            Set xmlTOENodes = vobjCommon.Data.selectNodes("//REMOVEDTOECUSTOMER[@TYPE='A']")
            intTOECustomersAdded = xmlTOENodes.length
            Set xmlTOENodes = vobjCommon.Data.selectNodes("//REMOVEDTOECUSTOMER[@TYPE='R']")
            intTOECustomersRemoved = xmlTOENodes.length
            '
            Select Case intTOECustomersAdded + intTOECustomersRemoved
                Case 1
                    xmlSetAttributeValue xmlNewNode, "TOEADDITIONREMOVALPARTIES", intTOECustomersAdded + intTOECustomersRemoved & " party"
                Case Else
                    xmlSetAttributeValue xmlNewNode, "TOEADDITIONREMOVALPARTIES", intTOECustomersAdded + intTOECustomersRemoved & " parties"
            End Select
            '
            If intTOECustomersAdded > 0 Then
                If intTOECustomersRemoved > 0 Then
                    strAddedOrRemoved = "removal and addition"
                Else
                    strAddedOrRemoved = "addition"
                End If
            Else
                If intTOECustomersRemoved > 0 Then
                    strAddedOrRemoved = "removal"
                Else
                    strAddedOrRemoved = "removal or addition"
                End If
            End If
            xmlSetAttributeValue xmlNewNode, "TOEADDITIONREMOVAL", strAddedOrRemoved
        Else
            'EP2_139 Use the intermediary firm name
             If vobjCommon.IsIntroducedByIntermediary(strIntermediaryContactName, False, strIntermediaryCompanyName, True) Then
                
                '*-add the INTERMEDIARY element
                Set xmlNewNode = vobjCommon.CreateNewElement("INTERMEDIARY", vxmlNode)  'AW 18/01/2005  E2EM00003187
                '*-add the MORTGAGETYPE attribute
                Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
                
                'Will be one of PURPOSEPURCHASE/PURPOSEFURTHERADVANCE/PURPOSEREMORTGAGE/PURPOSEPRODUCTSWITCH
                'Set xmlNewNode = vobjCommon.CreateNewElement(vobjCommon.LoanPurposeText, xmlNewNode)
                If vobjCommon.GetMainMortgageTypeGroup = "F" Then
                    'Additional borrowing
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE1", "additional borrowing"
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE3", "mortgage for the additional borrowing"
                Else
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE1", "mortgage"
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE3", "mortgage"
                End If

                'add the INTERMEDIARYNAME attribute
                
                'EP2_139 Use the intermediary firm name
                Call xmlSetAttributeValue(xmlNewNode, "INTERMEDIARYNAME", strIntermediaryCompanyName)
            
                '*-add a RECOMMENDS or NOTRECOMMENDS element to the INTERMEDIARY element
                Call AddRecommendsNotRecommendsElement(vobjCommon, xmlNewNode)
            Else
                '*-or add the LENDER element
                Set xmlNewNode = vobjCommon.CreateNewElement("LENDER", vxmlNode)    'AW 18/01/2005  E2EM00003187
                '*-add the MORTGAGETYPE attribute
                Call xmlSetAttributeValue(xmlNewNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
                
                'Will be one of PURPOSEPURCHASE/PURPOSEFURTHERADVANCE/PURPOSEREMORTGAGE/PURPOSEPRODUCTSWITCH
                'Set xmlNewNode = vobjCommon.CreateNewElement(vobjCommon.LoanPurposeText, xmlNewNode)
                If vobjCommon.GetMainMortgageTypeGroup = "F" Then
                    'Additional borrowing
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE1", "additional borrowing"
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE3", "mortgage for the additional borrowing"
                Else
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE1", "mortgage"
                    xmlSetAttributeValue xmlNewNode, "MORTGAGETYPE3", "mortgage"
                End If
                
                '*-add a RECOMMENDS or NOTRECOMMENDS element to the LENDER element
                Call AddRecommendsNotRecommendsElement(vobjCommon, xmlNewNode)
            End If
        End If
    End If

    Set xmlNewNode = Nothing
    'AW 18/01/2005  E2EM00003187
'    Set xmlMDay = Nothing
'    Set xmlKFIDocAuditDetails = Nothing
'    Set xmlAFF = Nothing
'    Set xmlKFIDate = Nothing
    'AW 18/01/2005  E2EM00003187 - End
Exit Sub
ErrHandler:
    Set xmlNewNode = Nothing
    'AW 18/01/2005  E2EM00003187
'    Set xmlMDay = Nothing
'    Set xmlKFIDocAuditDetails = Nothing
'    Set xmlAFF = Nothing
'    Set xmlKFIDate = Nothing
    'AW 18/01/2005  E2EM00003187 - End
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferInsuranceSection
'** Created by:     Andy Maggs
'** Date:           11/06/2004
'** Description:    Sets the elements and attributes for the compulsory Insurance
'**                 section element (Section 9 for Mortgage Offer and Section 12
'**                 for Lifetime Offer).
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferInsuranceSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection9"
    Dim xmlItem As IXMLDOMNode
    Dim xmlList As IXMLDOMNodeList
    
    On Error GoTo ErrHandler

    '*-build the insurance section that is common to the kfi and offer
    Call BuildInsuranceSection(vobjCommon, vxmlNode)
    
    '*-modify for Offer only attributes
    Set xmlItem = vxmlNode.selectSingleNode("//LENDER")
    If Not xmlItem Is Nothing Then
        Call xmlSetAttributeValue(xmlItem, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlItem = vxmlNode.selectSingleNode("//INTERMEDIARY")
    If Not xmlItem Is Nothing Then
        Call xmlSetAttributeValue(xmlItem, "PROVIDER", vobjCommon.Provider)
    End If
    
    Set xmlList = vxmlNode.selectNodes("//NONEREQUIRED")
    For Each xmlItem In xmlList
        Call xmlSetAttributeValue(xmlItem, "PROVIDER", vobjCommon.Provider)
    Next xmlItem
    
    Set xmlList = vxmlNode.selectNodes("//POLICYDETAILS")
    For Each xmlItem In xmlList
        Call xmlSetAttributeValue(xmlItem, "PROVIDER", vobjCommon.Provider)
    Next xmlItem
    
    Set xmlList = vxmlNode.selectNodes("//ADDTOMORTGAGE")
    For Each xmlItem In xmlList
        Call xmlSetAttributeValue(xmlItem, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
    Next xmlItem
    
    Set xmlList = vxmlNode.selectNodes("//INTERESTATMORTGAGERATE")
    For Each xmlItem In xmlList
        Call xmlSetAttributeValue(xmlItem, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
    Next xmlItem
    
    Set xmlItem = vxmlNode.selectSingleNode("//REQUIREDINSURANCE")
    If Not xmlItem Is Nothing Then
        Call xmlSetAttributeValue(xmlItem, "PROVIDER", vobjCommon.Provider)
        Call xmlSetAttributeValue(xmlItem, "MORTGAGETYPE", vobjCommon.MortgageTypeText)
        
        'EP2_1449  "PURPOSEREMORTGAGE" should include     'Remortgage' or 'Additional Borrowing'
        'or 'Product Switch' or 'Transfer of Equity', now done in BuildInsuranceSection call above
        'PM 20/06/2006 EP809 - START
        '*-now add one of the mandatory PURPOSEMORTGAGE, PURPOSEREMORTGAGE or
'        Select Case vobjCommon.LoanPurposeText
'            Case "PURPOSEPURCHASE"
'                Call vobjCommon.CreateNewElement("PURPOSEMORTGAGE", xmlItem)
'            Case "PURPOSEREMORTGAGE"  'SR 23/09/2004 : CORE82
'                Call vobjCommon.CreateNewElement("PURPOSEREMORTGAGE", xmlItem)
'        End Select
        'PM 20/06/2006 EP809 - END
    End If

    Set xmlItem = Nothing
    Set xmlList = Nothing

Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlList = Nothing

    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub



'PM 08/06/2006 EPSOM EP697   Start [Added PB 12/06/2006]
'********************************************************************************
'** Function:       BuildCommonOfferSection4Single
'** Created by:     Pat Morse
'** Date:           08/06/2006
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element. This currently processes for
'**                 mortgages with a single loan component
'**                 PLEASE NOTE: THIS ONLY WORKS FOR SINGLE COMPONENTS!!!
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection4Single(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection4Single"
    
    Dim xmlItem                 As IXMLDOMNode
    'input xml
    Dim xmlLC                   As IXMLDOMNode
    Dim xmlInitialIntRate       As IXMLDOMNode
    Dim xmlSubsequentIntRate    As IXMLDOMNode
    
    'output xml
    Dim xmlMainComp             As IXMLDOMNode
    Dim xmlProduct              As IXMLDOMNode
    Dim xmlMortProd             As IXMLDOMNode
    Dim xmlFixedNode            As IXMLDOMNode
    Dim xmlDiscountNode         As IXMLDOMNode

    Dim eType As MortgageInterestRateType
    Dim intYears As Integer
    Dim intMths As Integer
    Dim intPrefPeriod As Integer
    Dim intCurrentRateType As Integer
    Dim intBaseType As Integer
    Dim dblCurrentBaseInterestRate As Double
    Dim dblResolvedRate As Double
    Dim dblFixedRate As Double
    'PM EP769 23/06/2006 Start
'    Dim dblBaseRateLoad As Double
'    Dim dblReverBaseSet As Double
    Dim dblInitialRateDiff As Double
    Dim dblSubseqentRateDiff As Double
    Dim dblMaxPercent As Double
    'PM EP769 23/06/2006 End
    Dim dblBaseRate As Double

    Dim dblInitialRate As Double
    Dim dblSubsequentRate As Double
    Dim strNatureOfLoan As String
    Dim strPrefPeriod As String
    Dim strBaseType As String
    Dim strDirection As String
    Dim strMaxLTV As String
    Dim strSpecialGroup As String
'    Dim strPrimeCode As String
    Dim strTradingNameDescription As String 'LH 13/10/2006  EP1226/CC132
    'EP2_1983/1984
    Dim xmlFT As IXMLDOMNode
    Dim xmlInterestRateType As IXMLDOMNode
    Dim xmlReferenceRate As IXMLDOMNode
    Dim strBBR As String
    Dim intRateType As Integer
    Dim specialGroupTypeValidation As String
    
    On Error GoTo ErrHandler
    
    'LH 13/10/2006  EP1226/CC132: start
    'PB 13/03/2007  EP2_1932 Begin
    'If vobjCommon.ProviderCode = "001" Then 'db mortgage
    If vobjCommon.Provider = "db mortgages" Then
    'PB EP2_1932 End
        strTradingNameDescription = ", a trading name of DB UK Bank Ltd."
    End If
    'LH 13/10/2006  EP1226/CC132: end
    
    'Name of lender
    'EP2_2395
    strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN_VALIDID")
    If Not Len(strNatureOfLoan) > 0 Then
        strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN")
        strNatureOfLoan = GetValidationTypeForValueID("NatureOfLoan", strNatureOfLoan)
    End If
    
    If CheckForValidationType(strNatureOfLoan, "BI") Or CheckForValidationType(strNatureOfLoan, "BR") Then    'Buy To Let
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This Buy To Let mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    ElseIf CheckForValidationType(strNatureOfLoan, "LT") Then                          'Let To Buy
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This Let To Buy mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    ElseIf CheckForValidationType(strNatureOfLoan, "RS") Then                          'Residential
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    End If
    
    'get a handle on the single loan component
    Set xmlLC = vobjCommon.SingleLoanComponent
    
    'get a handle on Initial interest rate
    Set xmlInitialIntRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlLC)
    
    'get a handle on Subsequent interest rate
    Set xmlSubsequentIntRate = xmlLC.selectSingleNode( _
            ".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=2]")
    
'----------------set variables to be used by both fixed and discount mortgages-----------------

    'get all required data from Initial interest rate
    'PM EP769 23/06/2006 Start
    dblInitialRate = GetBaseInterestRate(vobjCommon, xmlInitialIntRate, _
                    "", dblInitialRateDiff)
'                    "", dblBaseRateLoad)
                    
    'get all required data from Subsequent interest rate
    dblSubsequentRate = GetBaseInterestRate(vobjCommon, xmlSubsequentIntRate, _
                    "", dblSubseqentRateDiff, dblBaseRate, intBaseType)
'                    "", dblReverBaseSet, dblBaseRate, intBaseType)
    'PM EP769 23/06/2006 End
        
    'PrefPeriod
    intPrefPeriod = xmlGetAttributeAsDouble(xmlInitialIntRate, "INTERESTRATEPERIOD")
    intYears = intPrefPeriod / 12
    intMths = intPrefPeriod Mod 12
    If intYears <> 0 Then
        'PM EP769 23/06/2006 Start
'        strPrefPeriod = CStr(intYears) & IIf(intYears > 1, " Years", " Year")
        strPrefPeriod = CStr(intYears) & IIf(intYears > 1, " years", " year")
        'PM EP769 23/06/2006 End
    End If
    If intMths <> 0 Then
        If strPrefPeriod <> "" Then
            strPrefPeriod = strPrefPeriod & " and "
        End If
        'PM EP769 23/06/2006 Start
'        strPrefPeriod = strPrefPeriod & CStr(intMths) & IIf(intMths > 1, " Months", "Month")
        strPrefPeriod = strPrefPeriod & CStr(intMths) & IIf(intMths > 1, " months", "month")
        'PM EP769 23/06/2006 End
    End If
    
    'BaseType
    If intBaseType = 2 Then     '### is it ok to test for = 2?
        strBaseType = "LIBOR"
    Else
        strBaseType = "Bank of England base"
    End If
    
    'PM EP769 23/06/2006 Start
'    'Direction
'    If dblBaseRateLoad > 0 Then
'        strDirection = "above"
'    Else
'        strDirection = "below"
'    End If
    'PM EP769 23/06/2006 End
'----------------------------------------------------------------------------------------------

    '*-add SINGLECOMPONENT element
    Set xmlMainComp = vobjCommon.AddComponentsTypeElement(vxmlNode)
            
    '*-add PRODUCT element
    Set xmlProduct = vobjCommon.CreateNewElement("PRODUCT", xmlMainComp)
    
    'ProductName
    Call AddProductNameAttribute(xmlLC, xmlProduct, vobjCommon)
        
    '*-determine if a discount or fixed mortgage
    eType = GetInterestRateType(xmlInitialIntRate)
    Select Case eType
        Case mrtFixedRate
            '*-add the FIXED element
            Set xmlFixedNode = vobjCommon.CreateNewElement("FIXED", xmlMainComp)

            'FixedRate
            'PM EP769 23/06/2006 Start
'            Call xmlSetAttributeValue(xmlFixedNode, "FIXEDRATE", CStr(set2DP(dblBaseRateLoad)))
            Call xmlSetAttributeValue(xmlFixedNode, "FIXEDRATE", CStr(set2DP(dblInitialRateDiff)))
            'PM EP769 23/06/2006 End

            'Direction
            'PM EP769 23/06/2006 Start
            Call xmlSetAttributeValue(xmlFixedNode, "DIRECTION", IIf(dblInitialRateDiff > 0, "above", "below"))
'            Call xmlSetAttributeValue(xmlFixedNode, "DIRECTION", strDirection)
            'PM EP769 23/06/2006 End
                        
            'PrefPeriod
            Call xmlSetAttributeValue(xmlFixedNode, "PREFPERIOD", strPrefPeriod)
            
            'ReverBaseSet
            'PM EP769 23/06/2006 Start
'            Call xmlSetAttributeValue(xmlFixedNode, "REVERBASESET", CStr(set2DP(dblReverBaseSet)))
            Call xmlSetAttributeValue(xmlFixedNode, "REVERBASESET", CStr(set2DP(dblSubseqentRateDiff)))
            'PM EP769 23/06/2006 End
            
            'PB 15/03/2007 EP2_1931 Begin
            xmlSetAttributeValue xmlFixedNode, "REVERRATE", CStr(set2DP(dblBaseRate + dblSubseqentRateDiff))
            'EP2_1931 End
            
            'BaseType
            Call xmlSetAttributeValue(xmlFixedNode, "BASETYPE", strBaseType)
            
            'BaseRate
            Call xmlSetAttributeValue(xmlFixedNode, "BASERATE", dblBaseRate)

        Case mrtDiscountedRate
            '*-add the DISCOUNT element
            Set xmlDiscountNode = vobjCommon.CreateNewElement("DISCOUNT", xmlMainComp)
            
            'ProductName
            Call AddProductNameAttribute(xmlLC, xmlDiscountNode, vobjCommon)
            
            'BaseRateLoad
            'PM EP769 23/06/2006 Start
'            Call xmlSetAttributeValue(xmlDiscountNode, "BASERATELOAD", CStr(dblBaseRateLoad))
            Call xmlSetAttributeValue(xmlDiscountNode, "BASERATELOAD", CStr(set2DP(dblSubseqentRateDiff)))
            'PM EP769 23/06/2006 End
            
            'Direction
            'PM EP769 23/06/2006 Start
'            Call xmlSetAttributeValue(xmlDiscountNode, "DIRECTION", strDirection)
            Call xmlSetAttributeValue(xmlDiscountNode, "DIRECTION", IIf(dblSubseqentRateDiff > 0, "above", "below"))
            'PM EP769 23/06/2006 End
            
            'BaseType
            Call xmlSetAttributeValue(xmlDiscountNode, "BASETYPE", strBaseType)
            
            'BaseRate
            Call xmlSetAttributeValue(xmlDiscountNode, "BASERATE", dblBaseRate)

            'Discount
            Call xmlSetAttributeValue(xmlDiscountNode, "DISCOUNT", CStr(set2DP(dblSubsequentRate - dblInitialRate)))
            
            'Rate
            Call xmlSetAttributeValue(xmlDiscountNode, "RATE", CStr(set2DP(dblInitialRate)))
            '### or is the following to be used?
'            If xmlAttributeValueExists(xmlLC, "RESOLVEDRATE") Then
'                dblResolvedRate = xmlGetAttributeAsDouble(xmlLC, "RESOLVEDRATE")
'                xmlSetAttributeValue xmlProduct, "RESOLVEDRATE", set2DP(dblResolvedRate)
'            End If
            
            'PrefPeriod
            Call xmlSetAttributeValue(xmlDiscountNode, "PREFPERIOD", strPrefPeriod)
            
            'ReverBaseSet
            'PM EP769 23/06/2006 Start
'            Call xmlSetAttributeValue(xmlDiscountNode, "REVERBASESET", CStr(set2DP(dblReverBaseSet)))
            Call xmlSetAttributeValue(xmlDiscountNode, "REVERBASESET", CStr(set2DP(dblSubseqentRateDiff)))
            'PM EP769 23/06/2006 End
            
            'PB 15/03/2007 EP2_1931 Begin
            xmlSetAttributeValue xmlDiscountNode, "REVERRATE", CStr(set2DP(dblBaseRate + dblSubseqentRateDiff))
            'EP2_1931 End
            '
        Case Else
            'do nothing
    End Select
    
    'Add <NOTADDITIONALBORROWING/> element (if not type 'F' - Further Advance)
    If vobjCommon.GetMainMortgageTypeGroup <> "F" Then
        vobjCommon.CreateNewElement "NOTADDITIONALBORROWING", xmlMainComp
    End If
    
    'EP2_1983/1984
    Set xmlFT = vobjCommon.CreateNewElement("FT", vxmlNode)
    '
    'Check for BBR rate
    'Will check combo validation when one exists!
    ' SR EP2_2159 - Check for first INTERESTRATETYPE record which is not fixed
    intRateType = 0
    For Each xmlInterestRateType In xmlLC.selectNodes(".//INTERESTRATETYPE")
        If xmlGetAttributeText(xmlInterestRateType, "RATETYPE") <> "F" Then
            intRateType = xmlGetAttributeAsInteger(xmlInterestRateType.selectSingleNode(".//BASERATE"), "RATEID")
            Exit For
        End If
    Next xmlInterestRateType
    ' SR EP2_2159 - End
    If intRateType = 1 Then
        ' BBR (Bank of England Base Rate)
        strBBR = "Following a change in the Bank of England base rate, the mortgage rate applicable " & _
                "will be adjusted from the following business day."
    ElseIf intRateType <> 0 Then ' SR EP2_2159
        strBBR = "The LIBOR rate is reviewed on the 15th March, June, September and December each year. " & _
                "Following any changes, the mortgage rate applicable will be adjusted from the following " & _
                "business day."
    End If
    '
    Set xmlReferenceRate = vobjCommon.CreateNewElement("ReferenceRate2", xmlFT)
    xmlSetAttributeValue xmlReferenceRate, "ReferenceRateChange", strBBR
    
    Set xmlMortProd = xmlLC.selectSingleNode("MORTGAGEPRODUCT")
    dblMaxPercent = xmlGetAttributeAsDouble(xmlMortProd, "MAXIMUMLTV")
    'MaxLTV
    'MortgageProduct/@MaximumLTV 0 decimal places
     'EP983 - reverse MAR1071
    'MAR1071 Use the Global Parameter for Maximum LTV percentage
'    Set xmlItem = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "MaximumLTVallowed" & Chr$(34) & "]")
'    strMaxLTV = xmlGetAttributeText(xmlItem, "PERCENTAGE")
'    Call xmlSetAttributeValue(vxmlNode, "MAXLTV", strMaxLTV)
    Set xmlMortProd = xmlLC.selectSingleNode("MORTGAGEPRODUCT")
    dblMaxPercent = xmlGetAttributeAsDouble(xmlMortProd, "MAXIMUMLTV")
    Call xmlSetAttributeValue(vxmlNode, "MAXLTV", CStr(dblMaxPercent))
    'EP983 - End
    'FinancialDiff (no text for this)
    'SpecialGroup/@GroupType ###??
    
    'strSpecialGroup = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "SPECIALGROUP")
    'If InStr(1, strSpecialGroup, "AC", vbTextCompare) Then
    'EP899 - 28/06/2006 Peter Edney
    'If any of the products have a special group that is not a "prime" group then we should note the
    'fact by creating a FINANCIALSTATUS node.
'    strPrimeCode = "ST"
'    If vobjCommon.Data.selectNodes(gcstrMORTGAGEPRODUCT_PATH & _
'        "/SPECIALGROUP[(not (starts-with(@GROUPTYPESEQUENCENUMBER_VALIDID,'" & strPrimeCode & "')))" & _
'        "and (not (contains(@GROUPTYPESEQUENCENUMBER_VALIDID,'," & strPrimeCode & "')))]").length > 0 Then
'        Call vobjCommon.CreateNewElement("FINANCIALSTATUS", vxmlNode)
'    End If
    'EP2_2448 Changed for pre-dip KFI, doesn't have SpecialGroup
    specialGroupTypeValidation = xmlGetAttributeText(xmlProduct.selectSingleNode(".//SPECIALGROUP"), "GROUPTYPESEQUENCENUMBER_VALIDID")
    If Not Len(specialGroupTypeValidation) > 0 Then
        specialGroupTypeValidation = vobjCommon.GetSpecialSchemeValidation
    End If
    
    If Not CheckForValidationType(specialGroupTypeValidation, "ST") Then
                Call vobjCommon.CreateNewElement("FINANCIALSTATUS", vxmlNode)
    End If

    'CAT (no text for this)
    'MortgageProduct/@CATIND
    '*-add the CATSTANDARD element
    If IsCatStandardMortgage(vobjCommon.Data) Then
        Call vobjCommon.CreateNewElement("CATSTANDARD", vxmlNode)
    End If

    
    Set xmlItem = Nothing
    Set xmlLC = Nothing
    Set xmlInitialIntRate = Nothing
    
    Set xmlSubsequentIntRate = Nothing
    Set xmlMainComp = Nothing
    Set xmlProduct = Nothing
    Set xmlFixedNode = Nothing
    Set xmlDiscountNode = Nothing
    Set xmlFT = Nothing
    Set xmlInterestRateType = Nothing
    Set xmlReferenceRate = Nothing
    
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlLC = Nothing
    Set xmlInitialIntRate = Nothing
    Set xmlSubsequentIntRate = Nothing
    Set xmlMainComp = Nothing
    Set xmlProduct = Nothing
    Set xmlFixedNode = Nothing
    Set xmlDiscountNode = Nothing
    Set xmlFT = Nothing
    Set xmlInterestRateType = Nothing
    Set xmlReferenceRate = Nothing
    
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub
'PM 08/06/2006 EPSOM EP697   End



'********************************************************************************
'** Function:       BuildCommonOfferConditionSection
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Sets the elements and attributes for the compulsory Condition
'**                 Section.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the condition section element to set the elements
'**                 and attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferConditionSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildConditionSection"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlMain As IXMLDOMNode

    On Error GoTo ErrHandler

    '*-create the GENERAL element
    Set xmlList = vobjCommon.Data.selectNodes(gcstrAPPLICATIONCONDITIONS_PATH & "[@CONDITIONTYPE=10]")
    Set xmlMain = vobjCommon.CreateNewElement("GENERAL", vxmlNode)
    Call AddConditions(vobjCommon, xmlList, xmlMain)
    
    '*-create the SPECIFIC element
    Set xmlList = vobjCommon.Data.selectNodes(gcstrAPPLICATIONCONDITIONS_PATH & "[@CONDITIONTYPE!=10]")
    Set xmlMain = vobjCommon.CreateNewElement("SPECIFIC", vxmlNode)
    Call AddConditions(vobjCommon, xmlList, xmlMain)

    Set xmlList = Nothing
    Set xmlMain = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlMain = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       AddConditions
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Adds any conditions from the list to the specified node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlList - the list of conditions.
'**                 vxmlNode - the node to add the details to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub AddConditions(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlList As IXMLDOMNodeList, ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "AddConditions"
    Dim xmlItem As IXMLDOMNode
    Dim xmlCondition As IXMLDOMNode
    Dim strDescription As String 'EP1127
    Dim intStart As Integer 'EP1127
    Dim arrBullet() As String 'EP1127
    Dim xmlBullet As IXMLDOMNode 'EP1127
    Dim strBullet As Variant 'EP1127
    Dim blnBullets As Boolean 'EP1127
    Const cstrBullet As String = "*" 'EP1127
    Dim arrParagraph() As String 'EP1144
    Dim strParagraph As String 'EP1144
    Dim xmlParagraph As IXMLDOMNode 'EP1144
    Dim intIndex As Integer 'EP1144
    
    On Error GoTo ErrHandler
    
    
    For Each xmlItem In vxmlList
                
        'Check for carriage-returns
        strDescription = xmlGetAttributeText(xmlItem, "CONDITIONDESCRIPTION", "")
        arrParagraph = Split(strDescription, Chr(10))
                   
        '*-create a CONDITION element
        If strDescription <> "" Then
            Set xmlCondition = vobjCommon.CreateNewElement("OFFERCONDITION", vxmlNode)
        End If
           
        For intIndex = 0 To UBound(arrParagraph)
            
            strParagraph = Trim(arrParagraph(intIndex))
            
            If strParagraph <> "" Then
                'Check for bullet point character
                blnBullets = False
                strDescription = strParagraph
                If Len(strParagraph) > 0 Then
                    intStart = InStr(1, strParagraph, cstrBullet)
                    If intStart > 0 Then
                        arrBullet = Split(Mid(strParagraph, intStart + 1), cstrBullet)
                        strDescription = Left(strParagraph, intStart - 1)
                        blnBullets = True
                    End If
                End If
                
                If strDescription <> "" Then
                    If intIndex = 0 Then
                        '*-add the DESCRIPTION attribute
                        'Call xmlCopyAttributeValue(xmlItem, xmlCondition, "CONDITIONDESCRIPTION", "DESCRIPTION")
                        Call xmlSetAttributeValue(xmlCondition, "DESCRIPTION", strDescription)
                    Else
                        Set xmlParagraph = vobjCommon.CreateNewElement("PARAGRAPH", xmlCondition)
                        Call xmlSetAttributeValue(xmlParagraph, "DESCRIPTION", strDescription)
                    End If
                End If
                    
                'Add subcondition nodes so we can create bullet points.
                If blnBullets Then
                    For Each strBullet In arrBullet
                        Set xmlBullet = vobjCommon.CreateNewElement("SUBCONDITION", xmlCondition)
                        Call xmlSetAttributeValue(xmlBullet, "DESCRIPTION", strBullet)
                    Next
                End If
            
            End If
        
        '*-add the NUMBER attribute
        Call xmlCopyAttributeValue(xmlItem, xmlCondition, "APPLNCONDITIONSSEQ", "NUMBER")
        
        Next
                
    Next xmlItem
    
    Set xmlItem = Nothing
    Set xmlCondition = Nothing
Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlCondition = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub


'********************************************************************************
'** Function:       BuildCommonOfferContactSection
'** Created by:     Andy Maggs
'** Date:           03/06/2004
'** Description:    Adds the appropriate contact attributes to the node.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the contact section node to add the attributes to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
'EP2_139   Clean up for Offer and KFI document changes.
Public Sub BuildCommonOfferContactSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferContactSection"
    Dim xmlAddress As IXMLDOMNode
    Dim xmlIntermediaryInfo As IXMLDOMNode
    Dim xmlProviderContact As IXMLDOMNode
    Dim xmlIntermedContact As IXMLDOMNode
    Dim providerTelephone As String
    Dim intermedTelephone As String

    On Error GoTo ErrHandler
    
    '*-add the CONTACT attribute
    Call xmlSetAttributeValue(vxmlNode, "CONTACTCOMPANY", vobjCommon.Provider)
    '*-add the CONTACT address attributes
 
    'EP2_2057
    Set xmlProviderContact = vobjCommon.Data.selectSingleNode(gcstrNANDADIRECTORY_PATH & "/CONTACTTELEPHONEDETAILS[@USAGE='10']")
    If Not xmlProviderContact Is Nothing Then
        providerTelephone = xmlGetAttributeText(xmlProviderContact, "AREACODE") & _
                            " " & xmlGetAttributeText(xmlProviderContact, "TELENUMBER")
    End If
    
    Set xmlAddress = vobjCommon.Data.selectSingleNode(gcstrNANDADIRECTORY_PATH & "/ADDRESS")
    Call AddAddressAttributes(xmlAddress, vxmlNode, "CONTACTADDRESS")
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", vobjCommon.Provider)

    '*-add the PROVIDER attribute
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", vobjCommon.Provider)
    '*-add the PROVIDER address attributes
    Set xmlAddress = vobjCommon.Data.selectSingleNode(gcstrNANDADIRECTORY_PATH & "/ADDRESS")
    Call AddAddressAttributes(xmlAddress, vxmlNode, "PROVIDERADDRESS")
    Call xmlSetAttributeValue(vxmlNode, "CONTACTTELEPHONE", providerTelephone)

    If vobjCommon.IndividualIntroducerProperty Then
        '*-add the Intermediary Company Name and Address attribute
        Call AddIntermediaryContactDetails(vobjCommon, vobjCommon.Data, vxmlNode, False, True)
    Else
        'No individual introducer
        Set xmlIntermediaryInfo = vobjCommon.CreateNewElement("NOINTERMEDIARY", vxmlNode)

        '*-add the PROVIDER attribute
        Call xmlSetAttributeValue(xmlIntermediaryInfo, "PROVIDER", vobjCommon.Provider)
        '*-add the PROVIDER address attributes
        Set xmlAddress = vobjCommon.Data.selectSingleNode(gcstrNANDADIRECTORY_PATH & "/ADDRESS")
        Call AddAddressAttributes(xmlAddress, xmlIntermediaryInfo, "PROVIDERADDRESS")
        Call xmlSetAttributeValue(xmlIntermediaryInfo, "CONTACTTELEPHONE", providerTelephone)
    
    End If
    
    Set xmlAddress = Nothing
    Set xmlIntermediaryInfo = Nothing
    Set xmlProviderContact = Nothing
    Set xmlIntermedContact = Nothing
    
Exit Sub
ErrHandler:
    
    Set xmlAddress = Nothing
    Set xmlIntermediaryInfo = Nothing
    Set xmlProviderContact = Nothing
    Set xmlIntermedContact = Nothing

    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection4
'** Created by:     Andy Maggs
'** Date:           26/05/2004
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element. NB this is only
'**                 common to standard mortgage offers and transfer of equity
'**                 offer documents.  Lifetime offers have their own method.
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection4(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection4"
    Dim xmlItems As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlMainComp As IXMLDOMNode
    ' PB 05/06/2006 EP651/MAR1590 begin
    'Dim xmlComponent As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode, xmlComp2 As IXMLDOMNode
    ' PB EP651/MAR1590 End
    Dim xmlProduct As IXMLDOMNode
    Dim blnHasMoreRateTypes As Boolean
    Dim xmlFT As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    Dim xmlRate As IXMLDOMNode
    Dim intLowestSeqNum As Integer
    Dim intPart As Integer
    Dim xmlLTV As IXMLDOMNode
    Dim xmlPurpose As IXMLDOMNode
    Dim xmlFinStatus As IXMLDOMNode
    Dim xmlCat As IXMLDOMNode
    Dim xmlMortProd As IXMLDOMNode
    Dim dblPercent As Double
    Dim dblMaxPercent As Double
    Dim xmlLC As IXMLDOMNode, dblResolvedRate As Double ' SR 22/09/2004 : CORE82
    'BBG1637
    Dim xmlSpecialDeals As IXMLDOMNode
    'EP983 reverse MAR1071
    Dim strMaxPercent As String                         ' MAR1071
    Dim xmlMonthsNode As IXMLDOMNode 'BC MAR1685 01/05/2006
    Dim strNatureOfLoan As String       'SAB 03/05/2006 - EPSOM EP489
    Dim strTradingNameDescription As String 'LH 13/10/2006  EP1226/CC132
    
    On Error GoTo ErrHandler
    
    'LH 13/10/2006  EP1226/CC132: start
    'If vobjCommon.ProviderCode = "001" Then 'db mortgage
    If vobjCommon.Provider = "db mortgages" Then 'PB 14/03/2007 EP2_1932 - check name not id
        strTradingNameDescription = ", a trading name of DB UK Bank Ltd."
    End If
    'LH 13/10/2006  EP1226/CC132: end
    
    'SAB 03/05/2006 - EPSOM EP489 - START
    '*-add the PROVIDER attribute
    'Call xmlSetAttributeValue(vxmlNode, "PROVIDER", vobjCommon.Provider)
    'EP2_2395
    strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN_VALIDID")
    If Not Len(strNatureOfLoan) > 0 Then
        strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN")
        strNatureOfLoan = GetValidationTypeForValueID("NatureOfLoan", strNatureOfLoan)
    End If
    
    If CheckForValidationType(strNatureOfLoan, "BI") Or CheckForValidationType(strNatureOfLoan, "BR") Then    'Buy To Let
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This Buy To Let mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    ElseIf CheckForValidationType(strNatureOfLoan, "LT") Then                          'Let To Buy
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This Let To Buy mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    ElseIf CheckForValidationType(strNatureOfLoan, "RS") Then                          'Residential
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    End If
    'SAB 03/05/2006 - EPSOM EP489 - END
    
    '*-add the MORTGAGETYPE attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", vobjCommon.MortgageTypeText)

    '*-add one of SINGLECOMPONENT or MULTICOMPONENT elements
    Set xmlMainComp = vobjCommon.AddComponentsTypeElement(vxmlNode)
    
    '*-create an element to temporarily store elements to be added to the
    '*-document later
    Set xmlTemp = vobjCommon.Document.createElement("TEMP")
    
    '*-get the loan components
    Set xmlItems = vobjCommon.LoanComponents
    If xmlItems.length = 1 Then
        '*-get the single loan component
        Set xmlItem = xmlItems.Item(0)
        
        '*-add the PRODUCT element
        Set xmlProduct = vobjCommon.CreateNewElement("PRODUCT", xmlMainComp)
        Call AddProductNameAttribute(xmlItem, xmlProduct, vobjCommon)
        
        ' SR 22/09/2004 : CORE82
        Set xmlLC = vobjCommon.SingleLoanComponent
        If xmlAttributeValueExists(xmlLC, "RESOLVEDRATE") Then
            dblResolvedRate = xmlGetAttributeAsDouble(xmlLC, "RESOLVEDRATE")
            xmlSetAttributeValue xmlProduct, "RESOLVEDRATE", set2DP(dblResolvedRate)
        End If
        ' SR 22/09/2004 : CORE82 - end
        
        '*-add the mortgage rate period elements
        Call AddRatePeriodElements(vobjCommon, xmlItem, xmlProduct, False, _
                False)
        '*-we now need to add the MORTGAGETYPE attribute to the FULLTERM and
        '*-REMAININGTERM elements
        Call AddMortgageTypeAttribute(vobjCommon, xmlProduct.selectSingleNode("RATEPERIOD/FULLTERM"))
        Call AddMortgageTypeAttribute(vobjCommon, xmlProduct.selectSingleNode("RATEPERIOD/REMAININGTERM"))
        'SR 16/09/2004 : CORE82
        Set xmlMortProd = xmlItem.selectSingleNode("MORTGAGEPRODUCT")
        dblMaxPercent = xmlGetAttributeAsDouble(xmlMortProd, "MAXIMUMLTV")
        'SR 16/09/2004 : CORE82 - End
        '*-add any Insurance elements
        Call AddTiedInsuranceElements(vobjCommon, xmlTemp)
        
    Else
        '*-get the lowest component sequence number and deduct 1 because we are going
        '*-to use this value to ensure that the PART attribute starts at 1
        intLowestSeqNum = GetMinIntAttribValue(xmlItems, "LOANCOMPONENTSEQUENCENUMBER") - 1
        
        '*-for each component, add a COMPONENT and an FT element
        For Each xmlItem In xmlItems
            Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", xmlMainComp)
            '*-add the INITIALRATE attribute
            Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlItem)
            Call AddRateAttribute(vobjCommon, xmlItem, xmlRate, xmlComponent, "INITIALRATE")
            '*-add the COMPONENTREPAYMENTMETHOD attribute
            Call AddRepaymentMethodAttribute(xmlItem, xmlComponent, "COMPONENTREPAYMENTMETHOD")
            
            'BC MAR1685 Begin
            If Not (xmlGetAttributeAsInteger(xmlItem, "TERMINMONTHS") = 0) Then
                Set xmlMonthsNode = vobjCommon.CreateNewElement("MONTHS", xmlComponent)
                '*-add the TERMINMONTHS attribute
                Call xmlCopyAttributeValue(xmlItem, xmlMonthsNode, "TERMINMONTHS", "TERMINMONTHS")
            End If
            'BC MAR1685 End
            
            '*-add the TERMINYEARS attribute
            Call xmlCopyAttributeValue(xmlItem, xmlComponent, "TERMINYEARS", "TERMINYEARS")
            '*-add the LOANCOMPONENTAMOUNT attribute
            Call xmlCopyAttributeValue(xmlItem, xmlComponent, "LOANAMOUNT", "LOANCOMPONENTAMOUNT")
            '*-add the PART attribute
            intPart = xmlGetAttributeAsInteger(xmlItem, "LOANCOMPONENTSEQUENCENUMBER") - intLowestSeqNum
            ' PB 05/06/2006 EP651/MAR1590
            'Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intPart))
            ' PB EP651/MAR1590 End
            
            '*-add the PRODUCT element
            Set xmlProduct = vobjCommon.CreateNewElement("PRODUCT", xmlComponent)
            Call AddProductNameAttribute(xmlItem, xmlProduct, vobjCommon)
            '*-add the first mortgage rate period element
            ' PB 24/05/2006 EP603/MAR1736
            Call xmlSetAttributeValue(xmlItem, "LOANCOMPONENTSEQUENCENUMBER", intPart) 'BC MAR1736
            ' EP603/MAR1736 End
            Call AddRatePeriodElements(vobjCommon, xmlItem, xmlProduct, True, False, _
                    blnHasMoreRateTypes)
            '*-we now need to add the MORTGAGETYPE attribute to the FULLTERM and
            '*-REMAININGTERM elements
            Call AddMortgageTypeAttribute(vobjCommon, xmlProduct.selectSingleNode("RATEPERIOD/FULLTERM"))
            Call AddMortgageTypeAttribute(vobjCommon, xmlProduct.selectSingleNode("RATEPERIOD/REMAININGTERM"))
            
            '*-add the subsequent rate periods to the FT element if any
            If blnHasMoreRateTypes Then
                'MAR88 - BC 06Oct Cretion of FT element movied inside 'If' statment
                Set xmlFT = vobjCommon.CreateNewElement("FT", xmlMainComp)
                Call AddRatePeriodElements(vobjCommon, xmlItem, xmlFT, False, True)
                '*-we now need to add the MORTGAGETYPE attribute to the FULLTERM and
                '*-REMAININGTERM elements
                Call AddMortgageTypeAttribute(vobjCommon, xmlFT.selectSingleNode("RATEPERIOD/FULLTERM"))
                Call AddMortgageTypeAttribute(vobjCommon, xmlFT.selectSingleNode("RATEPERIOD/REMAININGTERM"))
                '*-and add it to the FT element as well
                Call AddMortgageTypeAttribute(vobjCommon, xmlFT)
            End If
            
            '*-add any Insurance elements
            Call AddTiedInsuranceElements(vobjCommon, xmlTemp)
            
            '*-get the mortgage product for this loan component and update the
            '*-maximum percentage if it is larger
            Set xmlMortProd = xmlItem.selectSingleNode("MORTGAGEPRODUCT")
            dblPercent = xmlGetAttributeAsDouble(xmlMortProd, "MAXIMUMLTV")
            If dblMaxPercent < dblPercent Then
                dblMaxPercent = dblPercent
            End If
            
            'PB 05/06/2006 EP561/MAR1590 Begin
            If (InStr(1, xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD"), "3", vbTextCompare)) Then
                Set xmlComp2 = xmlComponent.cloneNode(True)
                xmlMainComp.appendChild xmlComp2
                Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intPart) + "(a)")
                Call xmlSetAttributeValue(xmlComponent, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlItem, "CAPITALANDINTERESTELEMENT"))
                'PB 12/06/2006 EP730/MAR1831 begin
                Call xmlSetAttributeValue(xmlComponent, "COMPONENTREPAYMENTMETHOD", "Repayment (i.e. capital and interest)")
                'PB EP730/MAR1831 End
                Call xmlSetAttributeValue(xmlComp2, "PART", CStr(intPart) + "(b)")
                Call xmlSetAttributeValue(xmlComp2, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlItem, "INTERESTONLYELEMENT"))
                'PB 12/06/2006 EP730/MAR1831 Begin
            Else
                Call xmlSetAttributeValue(xmlComponent, "PART", CStr(intPart))
            End If
            'PB 05/06/2006 EP561/MAR1590 End
        Next xmlItem
    End If
    
    '*-now add any INSURANCE elements to the document
    For Each xmlItem In xmlTemp.childNodes
        If xmlItem.baseName = "INSURANCE" Then
            Call vxmlNode.appendChild(xmlItem)
            '*-add the MORTGAGETYPE attribute to it
            Call AddMortgageTypeAttribute(vobjCommon, xmlItem)
        End If
    Next xmlItem
    
    'CORE82 These should be added to the RESTRICTIONS Node
'    '*-add the LTV element
    Set xmlLTV = vobjCommon.CreateNewElement("LTV", vxmlNode)
    '*-add the MORTGAGETYPE attribute to it
    Call AddMortgageTypeAttribute(vobjCommon, xmlLTV)
    '*-add the MAXPERCENT attribute
    'EP983 - reverse MAR1071
    'MAR1071 Use the Global Parameter for Maximum LTV percentage
'    Set xmlItem = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "MaximumLTVallowed" & Chr$(34) & "]")
'    strMaxPercent = xmlGetAttributeText(xmlItem, "PERCENTAGE")
'    Call xmlSetAttributeValue(xmlLTV, "MAXPERCENT", strMaxPercent)
    Call xmlSetAttributeValue(xmlLTV, "MAXPERCENT", CStr(dblMaxPercent))
    If vobjCommon.LoanPurposeText = "PURPOSEPURCHASE" Then
        '*-add the PURPOSEPURCHASE element
        Set xmlPurpose = vobjCommon.AddLoanPurposeElement(xmlLTV)
    End If
    
    '*-add the RESTRICTIONS element
    Call AddRestrictionsElements(vobjCommon, xmlItems, intLowestSeqNum, vxmlNode, dblMaxPercent)
    '*-add the MORTGAGETYPE attribute to the FINANCIALSTATUS node
    Set xmlFinStatus = vxmlNode.selectSingleNode("RESTRICTIONS/FINANCIALSTATUS")
    Call AddMortgageTypeAttribute(vobjCommon, xmlFinStatus)
    
    '*-add the CATSTANDARD element
    If IsCatStandardMortgage(vobjCommon.Data) Then
        Set xmlCat = vobjCommon.CreateNewElement("CATSTANDARD", vxmlNode)
        Call AddMortgageTypeAttribute(vobjCommon, xmlCat)
    End If

    'BBG1637*-add the SPECIALDEALS elements
    'ASSUMPTION THAT THERE IS ONLY A SINGLE LOANCOMPONENT
    If IsSpecialDeals(vobjCommon) Then
        Set xmlSpecialDeals = vobjCommon.CreateNewElement("SPECIALDEALS", vxmlNode)
        AddSpecialDealElements vobjCommon, vobjCommon.SingleLoanComponent, xmlSpecialDeals
    End If

    Set xmlItems = Nothing
    Set xmlItem = Nothing
    Set xmlMainComp = Nothing
    Set xmlComponent = Nothing
    Set xmlProduct = Nothing
    Set xmlFT = Nothing
    Set xmlTemp = Nothing
    Set xmlRate = Nothing
    Set xmlLTV = Nothing
    Set xmlPurpose = Nothing
    Set xmlFinStatus = Nothing
    Set xmlCat = Nothing
    Set xmlMortProd = Nothing
    Set xmlLC = Nothing
    Set xmlSpecialDeals = Nothing
    ' PB 05/06/2006 EP561/MAR1590
    Set xmlComp2 = Nothing
    ' PB EP561/MAR1590 End
Exit Sub
ErrHandler:
    Set xmlItems = Nothing
    Set xmlItem = Nothing
    Set xmlMainComp = Nothing
    Set xmlComponent = Nothing
    Set xmlProduct = Nothing
    Set xmlFT = Nothing
    Set xmlTemp = Nothing
    Set xmlRate = Nothing
    Set xmlLTV = Nothing
    Set xmlPurpose = Nothing
    Set xmlFinStatus = Nothing
    Set xmlCat = Nothing
    Set xmlMortProd = Nothing
    Set xmlLC = Nothing
    Set xmlSpecialDeals = Nothing
    ' PB 05/06/2006 EP561/MAR1590
    Set xmlComp2 = Nothing
    ' PB EP561/MAR1590 End
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection6
'** Created by:     Andy Maggs
'** Date:           27/05/2004
'** Description:    Sets the elements and attributes for the compulsory section6
'**                 (What you will need to pay each month) element. NB this is
'**                 only common to standard mortgage offers and transfer of equity
'**                 offer documents.  Lifetime offers have their own method.
'** Parameters:     vxmlNode - the section6 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection6(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection6"
    Dim xmlComponentNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    ' PB 05/06/2006 EP561/MAR1590 Begin
    Dim xmlNewNode2 As IXMLDOMNode
    ' PB EP561/MAR1590 End
    ' PB EP730/MAR1831 begin
    Dim xmlLoanComponentPaySched As IXMLDOMNode 'BC MAR1831
    ' PB EP730/MAR1831 End
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlLCList As IXMLDOMNodeList
    Dim xmlList As IXMLDOMNodeList
    Dim xmlRate As IXMLDOMNode
    Dim strRepayMethod As String
    Dim strRepayChar As String
    Dim blnInterestOnly As Boolean
    Dim blnPartandPart As Boolean '** MAR88
    Dim lngIOAmount As Long
    Dim xmlTemp As IXMLDOMNode
    Dim intIndex As Integer
    ' PB 05/06/2006 EP561/MAR1590 Begin
    Dim intPart As Integer
    ' PB EP561/MAR1590 End
    Dim xmlMDay As IXMLDOMNode
    Dim dteMDay As Date
    Dim intPostMday As Integer
    ' BBG1641
    Dim xmlLoanPurpNode As IXMLDOMNode
    
    Dim strTerminMonths As String 'BC MAR1510
    Dim intTerminMonths As Integer 'BC MAR1510
    Dim intYears As Integer  'BC MAR1510
    Dim intMonths As Integer  'BC MAR1510
    Dim dtInterestRateEndDate As Date  'BC MAR1510
    Dim strAssumedStartDate As String   ' PB EP529 / MAR1731
    ' PB EP730/MAR1831 Begin
    Dim strIntOnlyValue As String
    Dim strCapIntValue As String
    ' PB EP730/MAR1831 End
    ' PM EP824 23/06/2006 Begin
    Dim strRepayVehicle As String
    Dim xmlRepayVehicle As IXMLDOMNode
    Dim xmlPlan As IXMLDOMNode
    ' PM EP824 23/06/2006 End
    Dim strRepayVehicleType As String 'Peter Edney - EP1099
    'EP2_2042
    Dim repayVehicleMnthlyPayment As Double
    Dim planText As String
    Dim planType As String  ' ik_EP2_2458
    
    On Error GoTo ErrHandler

    blnInterestOnly = False
    blnPartandPart = False '** MAR88
    lngIOAmount = 0
    ' PM EP824 23/06/2006 Begin
    strRepayVehicle = ""
    ' PM EP824 23/06/2006 End
    Set xmlComponentNode = vobjCommon.AddComponentsTypeElement(vxmlNode)
    
    If xmlComponentNode.baseName = "SINGLECOMPONENT" Then
        '*-get the single loan component
        Set xmlLoanComponent = vobjCommon.SingleLoanComponent
        If xmlLoanComponent Is Nothing Then
            '*-if there is no loan component we can't really do a lot here!
            Exit Sub
        End If
        
        '*-get the repayment method
        strRepayMethod = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD")
        strRepayChar = Mid(strRepayMethod, 1, 1)
            
        If strRepayChar = "1" Then
            blnInterestOnly = True
            lngIOAmount = lngIOAmount + xmlGetAttributeAsLong(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT") 'SR EP2_2341-use TOTALLOANCOMPONENTAMOUNT (not LOANAMOUNT)
        End If
            
        If strRepayChar = "3" Then
            blnPartandPart = True
            lngIOAmount = lngIOAmount + xmlGetAttributeAsLong(xmlLoanComponent, "INTERESTONLYELEMENT")
        End If
        
        '*-add the mandatory PAYMENT attribute
        Call AddPaymentAttribute(xmlLoanComponent, xmlComponentNode, "PAYMENT")
        
        'BBG1596 Are we Pre or Post MDay?
        'EP2_1401 We no longer care
'        Set xmlMDay = vobjCommon.Data.selectSingleNode("//GLOBALDATAITEM[@NAME=" & Chr$(34) & "KFIMDay" & Chr$(34) & "]")
'        dteMDay = xmlGetAttributeAsDate(xmlMDay, "STRING")
'        intPostMday = DateDiff("d", Now, dteMDay)
'
'        If intPostMday <= 0 Then
            '*-add the loan component interest rate data
            Call AddPostMDayLoanComponentInterestRateData(vobjCommon, xmlLoanComponent, _
                    xmlComponentNode, False, blnInterestOnly)
'        Else
'            '*-add the loan component interest rate data
'            Call AddLoanComponentInterestRateData(vobjCommon, xmlLoanComponent, _
'                    xmlComponentNode, False)
'        End If
    
        If blnInterestOnly Or blnPartandPart Then
            ' PM EP824 23/06/2006 Begin
            '*-add the INTERESTONLY element
'            Call vobjCommon.CreateNewElement("INTERESTONLY", xmlComponentNode)
            '*-get the repayment vehicle
            strRepayVehicle = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_TEXT")
            ' PM EP824 23/06/2006 End
            
            'Peter Edney - EP1099
            strRepayVehicleType = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_VALIDID")
            'EP2_2042 get repayment cost
            repayVehicleMnthlyPayment = xmlGetAttributeAsLong(xmlLoanComponent, "REPAYMENTVEHICLEMONTHLYCOST")
            
            '*-add the INTERESTONLY element
            Set xmlNewNode = vobjCommon.CreateNewElement("INTERESTONLY", vxmlNode)
            '*-add the mandatory INTERESTONLYTOTAL attribute
            'BBG1596 Need to add any Fees Added To the Loan
            Call xmlSetAttributeValue(xmlNewNode, "INTERESTONLYTOTAL", _
                CStr(lngIOAmount))  'SR EP2_2341 - DO NOT add any fee as it has already been added to lngIOAmount
                
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, xmlNewNode)
            ' BBG1641
            Set xmlLoanPurpNode = vobjCommon.AddLoanPurposeElement(xmlNewNode)
            
            ' PM EP824 23/06/2006 Begin
            If strRepayVehicle = "" Then
                Call vobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
            Else
            
                'Peter Edney - EP1099
                Select Case True
                Case CheckForValidationType(strRepayVehicleType, "U")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
                Case CheckForValidationType(strRepayVehicleType, "M")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES_MIXED", xmlNewNode)
                Case CheckForValidationType(strRepayVehicleType, "E")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
                    planText = "an endowment policy/policies"
                    planType = "E" ' ik_EP2_2458
                Case CheckForValidationType(strRepayVehicleType, "I")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
                    planText = "an ISA/ISAs"
                    planType = "I" ' ik_EP2_2458
                Case CheckForValidationType(strRepayVehicleType, "P")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
                    planText = "a Pension/Pensions"
                    planType = "P" ' ik_EP2_2458
                Case CheckForValidationType(strRepayVehicleType, "SRP")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES2", xmlNewNode)
                    planText = "by selling your home"
                Case CheckForValidationType(strRepayVehicleType, "INH")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES2", xmlNewNode)
                    planText = "from an inheritance"
                Case CheckForValidationType(strRepayVehicleType, "SOB")
                    Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES2", xmlNewNode)
                    planText = "by selling your investment property or business"
                End Select
                
                If Not xmlRepayVehicle Is Nothing Then
                    Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", vobjCommon.Provider)
                    If Len(planText) > 0 Then
                        Call xmlSetAttributeValue(xmlRepayVehicle, "PLANTEXT", planText)
                    End If

                    '*-add the PLAN element to REPAYMENTVEHICLES element
                    ' ik_EP2_2458
                    Set xmlPlan = vobjCommon.CreateNewElement("PLANTYPE" & planType, xmlRepayVehicle)
                    Call xmlSetAttributeValue(xmlPlan, "PRODUCTNAME", strRepayVehicle)
                    If repayVehicleMnthlyPayment > 0 Then
                        Call xmlSetAttributeValue(xmlPlan, "PLANCOST", CStr(set2DP(repayVehicleMnthlyPayment)))
                    End If
                End If
            End If
            ' PM EP824 23/06/2006 End
        End If
    
    Else
        ' PB 16/05/2006 EP529 / MAR1731
        ' PB 24/05/2006 EP603/MAR1777
        'strAssumedStartDate = Format$(Now, "dd/mm/yyyy")
        'strAssumedStartDate = "01" & Mid(strAssumedStartDate, 3, 8)
        'strAssumedStartDate = Format$(DateAdd("m", 1, strAssumedStartDate), "dd/mm/yyyy")
        strAssumedStartDate = Format$(CalcExpectedCompletionDate(Date), "dd/mm/yyyy")
        ' EP603/MAR1777 End
        
        xmlSetAttributeValue xmlComponentNode, "ASSUMEDSTARTDATE", strAssumedStartDate
        ' EP529 / MAR1731 End
        
        '*-add the appropriate Fees and/or Premiums element
        Set xmlTemp = AddFeesPremiumsElement(vobjCommon, xmlComponentNode)
        If xmlTemp.baseName <> "NOFEESORPREMIUMS" Then
            '*-add the MORTGAGETYPE attribute
            Call AddMortgageTypeAttribute(vobjCommon, xmlTemp)
        End If
                
        '*-get the list of LoanComponents
        Set xmlLCList = vobjCommon.LoanComponents
        '*-add a COMPONENT element for each LoanComponent
        For Each xmlLoanComponent In xmlLCList
        
            '*-get the repayment vehicle
            strRepayVehicle = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_TEXT")
            ' PM EP824 23/06/2006 End
            
            'Peter Edney - EP1099
            strRepayVehicleType = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTVEHICLE_VALIDID")
            'EP2_2042 get repayment cost
            repayVehicleMnthlyPayment = xmlGetAttributeAsLong(xmlLoanComponent, "REPAYMENTVEHICLEMONTHLYCOST")

        
        
        
            strRepayMethod = xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD")
            strRepayChar = Mid(strRepayMethod, 1, 1)
            
            If strRepayChar = "1" Then
                blnInterestOnly = True
                lngIOAmount = lngIOAmount + xmlGetAttributeAsLong(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT") 'SR EP2_2341 - use TOTALLOANCOMPONENTAMOUNT (not LOANAMOUNT)
            End If
            
            If strRepayChar = "3" Then
                blnPartandPart = True
                lngIOAmount = lngIOAmount + xmlGetAttributeAsLong(xmlLoanComponent, "INTERESTONLYELEMENT")
            End If
            
            '*-add the COMPONENT element
            Set xmlNewNode = vobjCommon.CreateNewElement("COMPONENT", xmlComponentNode)
            '*-add the mandatory INITIALPAYMENT attribute
            Call AddPaymentAttribute(xmlLoanComponent, xmlNewNode, "INITIALPAYMENT")
            
            '*-and get the first interestratetype record
            Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlLoanComponent)
            
            '*-add the mandatory RATETYPE attribute
            Call AddInterestRateTypeAttribute(xmlRate, xmlNewNode)
            '*-add the mandatory INITIALRATE attribute
            Call AddRateAttribute(vobjCommon, xmlLoanComponent, _
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
                    ' PB 24/05/2006 EP603/MAR1777
                    'intTerminMonths = DateDiff("m", vobjCommon.ExpectedCompletionDate, dtInterestRateEndDate)
                    intTerminMonths = MonthDiff(vobjCommon.ExpectedCompletionDate, dtInterestRateEndDate)   'MAR1777 GHun
                    ' EP603/MAR1777 End
                    Call xmlSetAttributeValue(xmlNewNode, "TERMINMONTHS", CStr(intTerminMonths))
                End If
            End If
    
            '*-add the mandatory LOANCOMPONENTAMOUNT attribute
            'BC MAR1510 End
            
            Call AddLoanComponentAmountAttribute(xmlLoanComponent, xmlNewNode)
            '*-add the mandatory PART attribute
            Call AddComponentPartAttribute(xmlLoanComponent, xmlNewNode)
            
            'PB 05/06/2006 EP561/MAR1590
            If (InStr(1, xmlGetAttributeText(xmlLoanComponent, "REPAYMENTMETHOD"), "3", vbTextCompare)) Then
                
                'PB 12/06/2006 EP730/MAR1831 Begin
                Set xmlLoanComponentPaySched = xmlLoanComponent.selectSingleNode("LOANCOMPONENTPAYMENTSCHEDULE")
                strIntOnlyValue = xmlGetAttributeText(xmlLoanComponentPaySched, "INTONLYMONTHLYCOST")
                strCapIntValue = xmlGetAttributeText(xmlLoanComponentPaySched, "CAPINTMONTHLYCOST")
                'PB EP730/MAR1831 End

                Set xmlNewNode2 = xmlNewNode.cloneNode(True)
                xmlComponentNode.appendChild xmlNewNode2
                intPart = xmlGetAttributeText(xmlNewNode, "PART")
                Call xmlSetAttributeValue(xmlNewNode, "PART", CStr(intPart) + "(a)")
                Call xmlSetAttributeValue(xmlNewNode, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlLoanComponent, "CAPITALANDINTERESTELEMENT"))
                'PB 12/06/2006 EP730/MAR1831 begin
                'Call xmlSetAttributeValue(xmlNewNode, "REPAYMENTMETHOD", "Repayment (i.e. capital and interest)") 'SR EP2_1938
                
                'Call xmlSetAttributeValue(xmlNewNode, "INITIALPAYMENT", "0")
                Call xmlSetAttributeValue(xmlNewNode, "INITIALPAYMENT", strCapIntValue)
                'PB EP730/MAR1831 End
                Call xmlSetAttributeValue(xmlNewNode2, "PART", CStr(intPart) + "(b)")
                Call xmlSetAttributeValue(xmlNewNode2, "LOANCOMPONENTAMOUNT", _
                                    xmlGetAttributeAsLong(xmlLoanComponent, "INTERESTONLYELEMENT"))
                'PB 12/06/2006 EP730/MAR1831
                'Call xmlSetAttributeValue(xmlNewNode2, "REPAYMENTMETHOD", "Interest Only") 'SR EP2_1938
                Call xmlSetAttributeValue(xmlNewNode2, "INITIALPAYMENT", strIntOnlyValue)
                'PB EP730/MAR1831 End
            End If
            'PB EP561/MAR1590 - End
            
            
            If blnInterestOnly Or blnPartandPart Then
                '*-add the INTERESTONLY element if it doesn't already exist
                Set xmlNewNode = vxmlNode.selectSingleNode("INTERESTONLY")
                If xmlNewNode Is Nothing Then
                    Set xmlNewNode = vobjCommon.CreateNewElement("INTERESTONLY", vxmlNode)
                End If
                '*-add the mandatory INTERESTONLYTOTAL attribute
                'BBG1596 Need to add any Fees Added To the Loan
                Call xmlSetAttributeValue(xmlNewNode, "INTERESTONLYTOTAL", _
                   CStr(lngIOAmount)) 'SR EP2_2341 - Do not add any fee as they have already been added to lngIOAmount
                '*-add the MORTGAGETYPE attribute
                Call AddMortgageTypeAttribute(vobjCommon, xmlNewNode)
                ' BBG1641
                Set xmlLoanPurpNode = vobjCommon.AddLoanPurposeElement(xmlNewNode)
                
                ' PM EP824 23/06/2006 Begin
                If strRepayVehicle = "" Then
                    Call vobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
                Else
                
                    'Peter Edney - EP1099
                    Select Case True
                    Case CheckForValidationType(strRepayVehicleType, "U")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
                    Case CheckForValidationType(strRepayVehicleType, "M")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES_MIXED", xmlNewNode)
                    Case CheckForValidationType(strRepayVehicleType, "E")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
                        planText = "an endowment policy/policies"
                        planType = "E" ' ik_EP2_2458
                    Case CheckForValidationType(strRepayVehicleType, "I")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
                        planText = "an ISA/ISAs"
                        planType = "I" ' ik_EP2_2458
                    Case CheckForValidationType(strRepayVehicleType, "P")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
                        planText = "a Pension/Pensions"
                        planType = "P" ' ik_EP2_2458
                    Case CheckForValidationType(strRepayVehicleType, "SRP")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES2", xmlNewNode)
                        planText = "by selling your home"
                    Case CheckForValidationType(strRepayVehicleType, "INH")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES2", xmlNewNode)
                        planText = "from an inheritance"
                    Case CheckForValidationType(strRepayVehicleType, "SOB")
                        Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES2", xmlNewNode)
                        planText = "by selling your investment property or business"
                    End Select
                    If Not xmlRepayVehicle Is Nothing Then
                        Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", vobjCommon.Provider)
                        If Len(planText) > 0 Then
                            Call xmlSetAttributeValue(xmlRepayVehicle, "PLANTEXT", planText)
                        End If
    
                        '*-add the PLAN element to REPAYMENTVEHICLES element
                        ' ik_EP2_2458
                        Set xmlPlan = vobjCommon.CreateNewElement("PLANTYPE" & planType, xmlRepayVehicle)
                        Call xmlSetAttributeValue(xmlPlan, "PRODUCTNAME", strRepayVehicle)
                        If repayVehicleMnthlyPayment > 0 Then
                            Call xmlSetAttributeValue(xmlPlan, "PLANCOST", CStr(set2DP(repayVehicleMnthlyPayment)))
                        End If
                    End If
                End If
                ' PM EP824 23/06/2006 End
            End If
            
            
            
            
        Next xmlLoanComponent
        
        ' PM EP824 23/06/2006 Begin
'        If blnInterestOnly Or blnPartandPart Then
'            '*-add the INTERESTONLY element
'            Call vobjCommon.CreateNewElement("INTERESTONLY", xmlComponentNode)
'        End If
        ' PM EP824 23/06/2006 End
    
    End If
    
'    If blnInterestOnly Or blnPartandPart Then
'        '*-add the INTERESTONLY element
'        Set xmlNewNode = vobjCommon.CreateNewElement("INTERESTONLY", vxmlNode)
'        '*-add the mandatory INTERESTONLYTOTAL attribute
'        'BBG1596 Need to add any Fees Added To the Loan
'        Call xmlSetAttributeValue(xmlNewNode, "INTERESTONLYTOTAL", _
'            CStr(lngIOAmount + vobjCommon.FeesAddedToLoanAmount))
'
'        '*-add the MORTGAGETYPE attribute
'        Call AddMortgageTypeAttribute(vobjCommon, xmlNewNode)
'        ' BBG1641
'        Set xmlLoanPurpNode = vobjCommon.AddLoanPurposeElement(xmlNewNode)
'
'        ' PM EP824 23/06/2006 Begin
'        If strRepayVehicle = "" Then
'            Call vobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
'        Else
'
'            'Peter Edney - EP1099
'            Select Case True
'            Case CheckForValidationType(strRepayVehicleType, "U")
'                Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES_UNKNOWN", xmlNewNode)
'            Case CheckForValidationType(strRepayVehicleType, "M")
'                Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES_MIXED", xmlNewNode)
'            Case Else
'                Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
'            End Select
'
''            Set xmlRepayVehicle = vobjCommon.CreateNewElement("REPAYMENTVEHICLES", xmlNewNode)
'            Call xmlSetAttributeValue(xmlRepayVehicle, "PROVIDER", vobjCommon.Provider)
'            '*-add the PLAN element to REPAYMENTVEHICLES element
'            Set xmlPlan = vobjCommon.CreateNewElement("PLAN", xmlRepayVehicle)
'            Call xmlSetAttributeValue(xmlPlan, "PRODUCTNAME", strRepayVehicle)
'        End If
'        ' PM EP824 23/06/2006 End
'    End If

    Set xmlComponentNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlNewNode2 = Nothing 'EP730/MAR1831
    Set xmlLoanComponent = Nothing
    Set xmlLCList = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlTemp = Nothing
    Set xmlMDay = Nothing
    Set xmlLoanPurpNode = Nothing
    Set xmlLoanComponentPaySched = Nothing 'EP730/MAR1831
    ' PB 05/06/2006 EP561/MAR1590
    Set xmlNewNode2 = Nothing
    ' PB EP561/MAR1590
    ' PM EP824 23/06/2006 Begin
    Set xmlRepayVehicle = Nothing
    Set xmlPlan = Nothing
    ' PM EP824 23/06/2006 End
Exit Sub
ErrHandler:
    Set xmlComponentNode = Nothing
    Set xmlNewNode = Nothing
    Set xmlNewNode2 = Nothing 'EP730/MAR1831
    Set xmlLoanComponent = Nothing
    Set xmlLCList = Nothing
    Set xmlList = Nothing
    Set xmlRate = Nothing
    Set xmlTemp = Nothing
    Set xmlMDay = Nothing
    Set xmlLoanPurpNode = Nothing
    Set xmlLoanComponentPaySched = Nothing 'EP730/MAR1831
    ' PB 05/06/2006 EP561/MAR1590
    Set xmlNewNode2 = Nothing
    ' PB EP561/MAR1590
    ' PM EP824 23/06/2006 Begin
    Set xmlRepayVehicle = Nothing
    Set xmlPlan = Nothing
    ' PM EP824 23/06/2006 End
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection6A
'** Created by:     Andy Maggs
'** Date:
'** Description:    Sets the elements and attributes for the compulsory section6A
'**                 (What you will need to pay in the future) element.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section6A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection6A(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection6A"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlBalance As IXMLDOMNode
    Dim xmlRate As IXMLDOMNode
    'SR 23/03/2007 : EP2_1938
    Dim xmlPaymentScheduleDoc As FreeThreadedDOMDocument40, xslDoc As FreeThreadedDOMDocument40
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlListNode As IXMLDOMElement, xmlTemp As IXMLDOMNode
    Dim xmlPaymentSchedule As IXMLDOMNode
    Dim xmlPaymentScheduleList As IXMLDOMNodeList
    Dim strXslPattern As String, strSortedPaymentSchedule As String
    Dim strStartDate As String, strDateTemp As String, strCondition As String, strParts As String
    Dim strMortgageTypeDesc As String
    Dim strCurrStartDate As String
    Dim intLoanCompSeqNo As Integer, intNoOfComponents As Integer, iCount As Integer, intMonthsFromStart As String
    Dim dtStartDate As Date, dtCurrDate As Date
    Dim dblTotalPaymentAmount As Double
    
    Set xmlPaymentScheduleDoc = New FreeThreadedDOMDocument40
    
    'SR 23/03/2007 : EP2_1938 - End
    
    On Error GoTo ErrHandler

    Set xmlList = vobjCommon.LoanComponents
'SR 23/03/2007 : EP2_1938
    '----
'    If xmlList.length > 1 Then
'        For Each xmlItem In xmlList
'            '*-get the LOANCOMPONENTPAYMENTSCHEDULE element for this loan component
'            Set xmlBalance = xmlItem.selectSingleNode(".//LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=2]")
'            If Not xmlBalance Is Nothing Then
'
'                'MAR88 - BC 06Oct - Move "Set xmlComponent" and
'                '                        "Set xmlRate" and
'                '                        "Call AddRatePeriodAttribute" inside "'If"'
'                '*-add the COMPONENT element
'                Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", vxmlNode)
'
'                '*-and get the first interestratetyperecord
'                Set xmlRate = vobjCommon.GetLoanComponentFirstInterestRate(xmlItem)
'
'                'MAR54
'                '*-add the RATEPERIOD attribute
'                Call AddRatePeriodAttribute(vobjCommon, xmlRate, xmlComponent)
'
'                '*-add the NEWAMOUNT attribute
'                Call AddNewAmountAttribute(xmlBalance, xmlComponent)
'                '*-add the PART attribute
'                Call AddComponentPartAttribute(xmlBalance, xmlComponent)
'            End If
'        Next xmlItem
'    End If
    '---
  
    'Get all the records on LoanComponentSchedule records and sort them on startdate
    intNoOfComponents = xmlList.length
    Set xmlListNode = xmlPaymentScheduleDoc.createElement("LOANCOMPONENTPAYMENTSCHEDULELIST")
    xmlPaymentScheduleDoc.appendChild xmlListNode
    For Each xmlItem In xmlList
        Set xmlPaymentScheduleList = xmlItem.selectNodes("LOANCOMPONENTPAYMENTSCHEDULE")
        For Each xmlPaymentSchedule In xmlPaymentScheduleList
            Set xmlTemp = xmlPaymentSchedule.cloneNode(True)
            strStartDate = xmlGetAttributeText(xmlTemp, "STARTDATE")
            strStartDate = Right(strStartDate, 4) & Mid(strStartDate, 4, 2) + Left(strStartDate, 2)
            xmlSetAttributeValue xmlTemp, "STARTDATESTRING", strStartDate
            xmlListNode.appendChild xmlTemp
        Next xmlPaymentSchedule
    Next xmlItem
    
    'Now sort the records based on STARTDATESTRING
    strXslPattern = "<?xml version='1.0'?>" & _
                        "<xsl:stylesheet version='1.0' " & _
                        "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>" & _
                        "<xsl:template match='/'>" & _
                            "<LOANCOMPONENTPAYMENTSCHEDULELIST>" & _
                                "<xsl:for-each select='//LOANCOMPONENTPAYMENTSCHEDULE'> " & _
                                    "<xsl:sort order='ascending' select='@STARTDATESTRING'/>" & _
                                    "<xsl:copy-of select='.'/>" & _
                                "</xsl:for-each>" & _
                            "</LOANCOMPONENTPAYMENTSCHEDULELIST>" & _
                        "</xsl:template>" & _
                        "</xsl:stylesheet>"
                        
    Set xslDoc = xmlLoad(strXslPattern, "BuildCommonOfferSection6A")
    strSortedPaymentSchedule = xmlPaymentScheduleDoc.transformNode(xslDoc)
    Set xmlPaymentScheduleDoc = xmlLoad(strSortedPaymentSchedule, "BuildCommonOfferSection6A")
    Set xmlTempDoc = xmlLoad(strSortedPaymentSchedule, "BuildCommonOfferSection6A")
    
    Set xmlPaymentScheduleList = xmlPaymentScheduleDoc.selectNodes("//LOANCOMPONENTPAYMENTSCHEDULE")
    Set xmlPaymentSchedule = xmlPaymentScheduleDoc.documentElement.selectSingleNode("//LOANCOMPONENTPAYMENTSCHEDULE")
    
    strCurrStartDate = xmlGetAttributeText(xmlPaymentSchedule, "STARTDATESTRING")
    dtStartDate = xmlGetAttributeAsDate(xmlPaymentSchedule, "STARTDATE")
    
    For Each xmlPaymentSchedule In xmlPaymentScheduleList
        strDateTemp = xmlGetAttributeText(xmlPaymentSchedule, "STARTDATESTRING")
        If (strDateTemp > strCurrStartDate) Then
            intLoanCompSeqNo = xmlGetAttributeAsInteger(xmlPaymentSchedule, "LOANCOMPONENTSEQUENCENUMBER")
            dtCurrDate = xmlGetAttributeAsDate(xmlPaymentSchedule, "STARTDATE")
            
            'Check whether any other components starts on this date. Fetch all those parts
            strCondition = "[@LOANCOMPONENTSEQUENCENUMBER !=" & CStr(intLoanCompSeqNo) & " and " & "@STARTDATESTRING='" & strDateTemp & "']"
            Set xmlList = xmlTempDoc.selectNodes("//LOANCOMPONENTPAYMENTSCHEDULE" & strCondition)
            Set xmlItem = vobjCommon.Data.selectSingleNode(gcstrLOANCOMPONENT_PATH & "[@LOANCOMPONENTSEQUENCENUMBER =" & CStr(intLoanCompSeqNo) & "]")
            strParts = GetProductRateType(xmlItem)
            strParts = strParts & " on part " & CStr(intLoanCompSeqNo)
            strMortgageTypeDesc = strParts
            
            If xmlList.length > 0 Then
                For Each xmlItem In xmlList
                    intLoanCompSeqNo = xmlGetAttributeAsInteger(xmlItem, "LOANCOMPONENTSEQUENCENUMBER")
                    Set xmlItem = vobjCommon.Data.selectSingleNode(gcstrLOANCOMPONENT_PATH & "[@LOANCOMPONENTSEQUENCENUMBER =" & CStr(intLoanCompSeqNo) & "]")
                    strParts = GetProductRateType(xmlItem)
                    strParts = strParts & " on part " & CStr(intLoanCompSeqNo)
                    strMortgageTypeDesc = strMortgageTypeDesc & " and " & strParts
                Next xmlItem
            End If
            
            'get Number of Months from start date
            intMonthsFromStart = DateDiff("M", dtStartDate, dtCurrDate)
            
            'Get the total amount paid after this date (this period)
            dblTotalPaymentAmount = 0
            For iCount = 1 To intNoOfComponents
                ' Sum of Payment in all the components that have highest startdate that are less than the date in focus
                strCondition = "[@LOANCOMPONENTSEQUENCENUMBER =" & CStr(iCount) & " and " & "@STARTDATESTRING <='" & strDateTemp & "']"

                Set xmlList = xmlTempDoc.selectNodes("//LOANCOMPONENTPAYMENTSCHEDULE" & strCondition)
                If xmlList.length > 0 Then
                    'Choose the one that comes first and retrieve MONTHLYCOST
                    dblTotalPaymentAmount = dblTotalPaymentAmount + _
                                        xmlGetAttributeAsDouble(xmlList.Item(xmlList.length - 1), "MONTHLYCOST")
                End If
            Next iCount
            
            'Add all the values to to XML as COMPONENT node
            Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", vxmlNode)
            'EP2_2478 Set to 2DP
            xmlSetAttributeValue xmlComponent, "NEWAMOUNT", CStr(set2DP(dblTotalPaymentAmount))
            xmlSetAttributeValue xmlComponent, "MORTGAGETYPEDESC", strMortgageTypeDesc
            Set xmlTemp = vobjCommon.CreateNewElement("RATEPERIOD", xmlComponent)
            xmlSetAttributeValue xmlTemp, "RATEPERIOD", intMonthsFromStart
            
            strCurrStartDate = strDateTemp
        End If
    Next xmlPaymentSchedule
'SR 23/03/2007 : EP2_1938

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlComponent = Nothing
    Set xmlBalance = Nothing
    Set xmlRate = Nothing
    Set xmlPaymentScheduleDoc = Nothing
    Set xslDoc = Nothing
    Set xmlTempDoc = Nothing
    Set xmlListNode = Nothing
    Set xmlTemp = Nothing
    Set xmlPaymentSchedule = Nothing
    Set xmlPaymentScheduleList = Nothing
    
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlComponent = Nothing
    Set xmlBalance = Nothing
    Set xmlRate = Nothing
    Set xmlPaymentScheduleDoc = Nothing
    Set xslDoc = Nothing
    Set xmlTempDoc = Nothing
    Set xmlListNode = Nothing
    Set xmlTemp = Nothing
    Set xmlPaymentSchedule = Nothing
    Set xmlPaymentScheduleList = Nothing
    
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection7
'** Created by:     Andy Maggs
'** Date:           27/05/2004
'** Description:    Sets the elements and attributes for the compulsory section7
'**                 (Are you comfortable with the risks?) element. NB this is
'**                 only common to standard mortgage offers and transfer of equity
'**                 offer documents.  Lifetime offers have their own method.
'** Parameters:     vobjCommon - the common data helper.
'**                 vxmlNode - the section7 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection7(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection7"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlComponent As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode

    On Error GoTo ErrHandler

    Call BuildSection7Common(vobjCommon, vxmlNode)
        
    Set xmlList = vxmlNode.selectNodes("//COMPONENT")
    For Each xmlComponent In xmlList
        '*-add the provider attribute to the appropriate nodes if present
        Set xmlTemp = xmlComponent.selectSingleNode("RATEPERIOD/EXAMPLEONEPERCENTINCREASE")
        If Not xmlTemp Is Nothing Then
            Call xmlSetAttributeValue(xmlTemp, "PROVIDER", vobjCommon.Provider)
        End If
        Set xmlTemp = xmlComponent.selectSingleNode("RATEPERIOD/ONEPERCENTINCREASEAFTER")
        If Not xmlTemp Is Nothing Then
            Call xmlSetAttributeValue(xmlTemp, "PROVIDER", vobjCommon.Provider)
        End If
    Next xmlComponent

    Set xmlList = Nothing
    Set xmlComponent = Nothing
    Set xmlTemp = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlComponent = Nothing
    Set xmlTemp = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'********************************************************************************
'** Function:       BuildCommonOfferSection7A
'** Created by:     Andy Maggs
'** Date:           21/06/2004
'** Description:    Sets the elements and attributes for the compulsory section7A
'**                 (Total Borrowing) element.
'** Parameters:     vxmlNode - the section7A element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection7A(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection7A"

    On Error GoTo ErrHandler

    Call BuildSection7ACommon(vobjCommon, vxmlNode)
    '*-add the PROVIDER attribute
    Call xmlSetAttributeValue(vxmlNode, "PROVIDER", vobjCommon.Provider)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'BuildCommonOfferSection4Multiple
'[Added PB 12/06/2006]
'********************************************************************************
'** Function:       BuildCommonOfferSection4Single
'** Created by:     Paul Buck (originally Pat Morse)
'** Date:           08/06/2006
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element. This processes for
'**                 mortgages with multiple loan components
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferSection4Multiple(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferSection4Multiple"
    
    Dim xmlItem                 As IXMLDOMNode
    'input xml
    Dim xmlLC                   As IXMLDOMNode
    Dim xmlLCs                  As IXMLDOMNodeList
    Dim xmlInitialIntRate       As IXMLDOMNode
    Dim xmlSubsequentIntRate    As IXMLDOMNode
    
    'output xml
    Dim xmlMainComp             As IXMLDOMNode
    Dim xmlProduct              As IXMLDOMNode
    Dim xmlMortProd             As IXMLDOMNode
    Dim xmlFixedNode            As IXMLDOMNode
    Dim xmlDiscountNode         As IXMLDOMNode

    Dim eType As MortgageInterestRateType
    Dim intYears As Integer
    Dim intMths As Integer
    Dim intPrefPeriod As Integer
    Dim intCurrentRateType As Integer
    Dim intBaseType As Integer
    Dim dblCurrentBaseInterestRate As Double
    Dim dblResolvedRate As Double
    Dim dblFixedRate As Double
    'PM EP769 23/06/2006 Start
'    Dim dblBaseRateLoad As Double
'    Dim dblReverBaseSet As Double
    Dim dblInitialRateDiff As Double
    Dim dblSubseqentRateDiff As Double
    Dim dblMaxPercent As Double
    'PM EP769 23/06/2006 End
    Dim dblBaseRate As Double

    Dim dblInitialRate As Double
    Dim dblSubsequentRate As Double
    Dim strNatureOfLoan As String
    Dim strPrefPeriod As String
    Dim strBaseType As String
    Dim strDirection As String
    Dim strMaxLTV As String
    Dim strSpecialGroup As String
    Dim strTradingNameDescription As String 'LH 13/10/2006  EP1226/CC132
    
    Dim xmlMultiComponent As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlComponent2 As IXMLDOMNode
    Dim intComponentCount As Integer
    'Dim xmlProduct As IXMLDOMNode
    Dim xmlNewProduct As IXMLDOMNode
    Dim xmlDiscount As IXMLDOMNode
    Dim xmlInterestType As IXMLDOMNode
    Dim dblLoad As Double
    Dim xmlBaseRate As IXMLDOMNode
    Dim xmlNextBaseRate As IXMLDOMNode
    Dim xmlInterestRateType As IXMLDOMNode
    Dim dblRate As Double
    Dim xmlRateBands As IXMLDOMNodeList
    Dim xmlRateBand As IXMLDOMNode
    Dim dblRateAdjustment As Double
    Dim intLoop As Integer
    Dim xmlInterestRateTypes As IXMLDOMNodeList
    Dim dblDifference As Double
    Dim strDifference As String
    Dim xmlFT As IXMLDOMNode
    Dim xmlReferenceRate As IXMLDOMNode
    Dim strStep4ReferenceRateStep As String
    Dim strBBR As String
    Dim intRateType As Integer
    Dim dblDiscountAmount As Double
    Dim dblDiscountedRate As Double
    Dim dblUnDiscountedRate As Double
    Dim strProductName As String
    'EP2_2448
    Dim specialGroupType As String
    Dim specialGroupTypeValidation As String
    Dim notAllPrime As Boolean
    notAllPrime = False
    
    On Error GoTo ErrHandler
    
    'LH 13/10/2006  EP1226/CC132: start
    'If vobjCommon.ProviderCode = "001" Then 'db mortgage
    If vobjCommon.Provider = "db mortgages" Then
        strTradingNameDescription = ", a trading name of DB UK Bank Ltd."
    End If
    'LH 13/10/2006  EP1226/CC132: end
    
    Set xmlMultiComponent = vobjCommon.CreateNewElement("MULTICOMPONENT", vxmlNode)
    xmlSetAttributeValue xmlMultiComponent, "MORTGAGETYPE", vobjCommon.MortgageTypeText
    
    'EP2_2395
    'Name of lender
    strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN_VALIDID")
    If Not Len(strNatureOfLoan) > 0 Then
        strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN")
        strNatureOfLoan = GetValidationTypeForValueID("NatureOfLoan", strNatureOfLoan)
    End If
    
    If CheckForValidationType(strNatureOfLoan, "BI") Or CheckForValidationType(strNatureOfLoan, "BR") Then    'Buy To Let
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This Buy To Let mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    ElseIf CheckForValidationType(strNatureOfLoan, "LT") Then                          'Let To Buy
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This Let To Buy mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    ElseIf CheckForValidationType(strNatureOfLoan, "RS") Then                          'Residential
        Call xmlSetAttributeValue(vxmlNode, "PROVIDER", "This mortgage is provided by " + vobjCommon.Provider + strTradingNameDescription)
    End If
    'get a handle on the multiple loan components
    Set xmlLCs = vobjCommon.LoanComponents 'SingleLoanComponent
    
    'loop thru individual components
    For Each xmlLC In xmlLCs
    
        intComponentCount = intComponentCount + 1
        Set xmlComponent = vobjCommon.CreateNewElement("COMPONENT", xmlMultiComponent)
        Set xmlComponent2 = vobjCommon.CreateNewElement("COMPONENT2", xmlMultiComponent)
        
        xmlSetAttributeValue xmlComponent, "PART", CStr(intComponentCount)
         'SR EP2_2341 - use TOTALLOANCOMPONENTAMOUNT instead of LOANAMOUNT
        xmlSetAttributeValue xmlComponent, "LOANCOMPONENTAMOUNT", xmlGetAttributeAsDouble(xmlLC, "TOTALLOANCOMPONENTAMOUNT")
        xmlSetAttributeValue xmlComponent, "TERMINYEARS", xmlGetAttributeAsDouble(xmlLC, "TERMINYEARS")
        xmlSetAttributeValue xmlComponent, "TERMINMONTHS", xmlGetAttributeAsDouble(xmlLC, "TERMINMONTHS")
        'SR EP2_1938
        'xmlSetAttributeValue xmlComponent, "COMPONENTREPAYMENTMETHOD", xmlGetAttributeText(xmlLC, "REPAYMENTMETHOD_TEXT")
        Call AddRepaymentMethodAttribute(xmlLC, xmlComponent, "COMPONENTREPAYMENTMETHOD")
        'SR EP2_1938 - End
        Set xmlProduct = xmlLC.selectSingleNode(".//MORTGAGEPRODUCT")
        '
        dblMaxPercent = xmlGetAttributeAsDouble(xmlProduct, "MAXIMUMLTV")
        If dblMaxPercent > mdblMaxLTV Then
            mdblMaxLTV = dblMaxPercent
        End If
        '
        Set xmlNewProduct = vobjCommon.CreateNewElement("PRODUCT", xmlComponent)
        'EP2_2448 if coming from pre-dip KFI, don't have the specialgroup stuff
        specialGroupType = xmlGetAttributeText(xmlProduct.selectSingleNode(".//SPECIALGROUP"), "GROUPTYPE")
        If Not Len(specialGroupType) > 0 Then
            specialGroupType = vobjCommon.GetSpecialSchemeGroupType
        End If
        specialGroupTypeValidation = xmlGetAttributeText(xmlProduct.selectSingleNode(".//SPECIALGROUP"), "GROUPTYPESEQUENCENUMBER_VALIDID")
        If Not Len(specialGroupTypeValidation) > 0 Then
            specialGroupTypeValidation = vobjCommon.GetSpecialSchemeValidation
        End If
        ' Build product name string
        strProductName = xmlGetAttributeText(xmlProduct.selectSingleNode(".//MORTGAGEPRODUCTLANGUAGE"), "PRODUCTTEXTDETAILS") & ", " & _
                        specialGroupType & _
                        " (product code " & xmlGetAttributeText(xmlProduct, "MORTGAGEPRODUCTCODE") & ")"
        xmlSetAttributeValue xmlNewProduct, "PRODUCTNAME", strProductName
        
        If Not CheckForValidationType(specialGroupTypeValidation, "ST") Then
            notAllPrime = True
        End If
        '
        If vobjCommon.GetMainMortgageTypeGroup = "F" Then
            'Additional borrowing (further advance)
            xmlSetAttributeValue xmlNewProduct, "ADDBORROWAVAIL", "This product may not be available for any additional borrowing."
        End If
        
        'XML for multipart mortgages appears as follows:
        '
        '<MULTICOMPONENT MORTGAGETYPE='...'>
        '   <COMPONENT PART='' LOANCOMPONENTAMOUNT='' TERMINYEARS='' TERMINMONTHS='' COMPONENTREPAYMENTMETHOD=''>
        '       <PRODUCT PRODUCTNAME='' INITIALRATE='' ADDBORROWAVAIL=''>
        '           <DISCOUNT BASERATELOAD='' DIRECTION='' BASETYPE='' BASERATE='' DISCOUNT='' PREFPERIOD='' RATE1=''/>
        '       ...or...
        '           <FIXED FIXEDRATE='' PREFPERIOD=''/>
        '       </PRODUCT>
        '   </COMPONENT>
        '   <COMPONENT2 COMPONENTNO='' PREFPERIOD='' REVERBASESET='' DIRECTION='' BASETYPE='' BASERATE='' RATE2=''/>
        '</MULTICOMPONENT>
        '
        'Check if Discount or Fixed
        Set xmlInterestRateType = xmlLC.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]") ' Get the first one (might be multiple step)
        eType = GetInterestRateType(xmlInterestRateType)
        'GetInterestRates vobjCommon, xmlLC, xmlInterestRateType, eType, dblRate 'SR EP2_1938 - this rate is not used anywhere in the function
        'dblLoad = xmlGetAttributeAsDouble(xmlInterestType, "RATE") 'SR EP2_1938
        '
        'Adjust if any rate differences for LTV
        Set xmlRateBands = xmlInterestRateType.selectNodes("./BASERATESETDATA/BASERATEBAND")
        'Loop thru and get the best applicable rate adjustment
        For intLoop = xmlRateBands.length To 1 Step -1
            Set xmlRateBand = xmlRateBands(intLoop - 1)
            If vobjCommon.LTV <= xmlGetAttributeAsDouble(xmlRateBand, "MAXIMUMLTV") Then
                dblRateAdjustment = xmlGetAttributeAsDouble(xmlRateBand, "RATEDIFFERENCE")
            Else
                Exit For
            End If
        Next
        '
        intPrefPeriod = xmlGetAttributeAsInteger(xmlInterestRateType, "INTERESTRATEPERIOD")
        '
        Select Case xmlGetAttributeText(xmlInterestRateType, "RATETYPE")
            Case "D"
                'Discount
                Set xmlBaseRate = xmlInterestRateType.selectSingleNode(".//BASERATE")
                dblBaseRate = xmlGetAttributeAsDouble(xmlBaseRate, "BASEINTERESTRATE")
                
                ' Get the discounted rate
                dblDiscountedRate = xmlGetAttributeAsDouble(xmlLC.selectSingleNode("LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=1]"), "INTERESTRATE")
                ' Get the undiscounted rate
                dblUnDiscountedRate = xmlGetAttributeAsDouble(xmlLC.selectSingleNode("LOANCOMPONENTPAYMENTSCHEDULE[@INTERESTRATETYPESEQUENCENUMBER=2]"), "INTERESTRATE")
                ' Now we can get the amount of rate discount
                dblDiscountAmount = dblUnDiscountedRate - dblDiscountedRate
                
                ' And the rate adjustment
                dblLoad = dblUnDiscountedRate - dblBaseRate
                
                Set xmlDiscount = vobjCommon.CreateNewElement("DISCOUNT", xmlNewProduct)
                
                xmlSetAttributeValue xmlDiscount, "BASERATELOAD", Abs(dblLoad)
                If dblRateAdjustment >= 0 Then
                    xmlSetAttributeValue xmlDiscount, "DIRECTION", "above"
                Else
                    xmlSetAttributeValue xmlDiscount, "DIRECTION", "below"
                End If
                
                                
                xmlSetAttributeValue xmlDiscount, "BASETYPE", xmlGetAttributeText(xmlBaseRate, "RATEDESCRIPTION")
                xmlSetAttributeValue xmlDiscount, "BASERATE", dblBaseRate
                xmlSetAttributeValue xmlDiscount, "DISCOUNT", dblDiscountAmount
                xmlSetAttributeValue xmlDiscount, "PREFPERIOD", intPrefPeriod
                xmlSetAttributeValue xmlDiscount, "RATE1", dblDiscountedRate
                'EP2_2478 set to 2DP
                xmlSetAttributeValue xmlComponent, "INITIALRATE", CStr(set2DP(dblDiscountedRate))
            Case "B"
                'Base
                Set xmlDiscount = vobjCommon.CreateNewElement("BASE", xmlNewProduct)
                Set xmlBaseRate = xmlInterestRateType.selectSingleNode(".//BASERATE") 'SR EP2_2341
                xmlSetAttributeValue xmlDiscount, "FIXEDRATE", xmlGetAttributeText(xmlBaseRate, "BASEINTERESTRATE")
                xmlSetAttributeValue xmlDiscount, "PREFPERIOD", intPrefPeriod
                dblLoad = xmlGetAttributeAsDouble(xmlInterestRateType, "RATE") 'SR EP2_1938
                'EP2_1382
                'EP2_2478 set to 2DP
                xmlSetAttributeValue xmlComponent, "INITIALRATE", CStr(set2DP(CSafeDbl(xmlGetAttributeText(xmlBaseRate, "BASEINTERESTRATE")) + dblRateAdjustment + dblLoad))
            Case "F"
                'Fixed
                Set xmlDiscount = vobjCommon.CreateNewElement("FIXED", xmlNewProduct)
                Set xmlBaseRate = xmlInterestRateType.selectSingleNode(".//BASERATE")  'SR EP2_2341
                xmlSetAttributeValue xmlDiscount, "FIXEDRATE", xmlGetAttributeText(xmlBaseRate, "BASEINTERESTRATE")
                xmlSetAttributeValue xmlDiscount, "PREFPERIOD", intPrefPeriod
                dblLoad = xmlGetAttributeAsDouble(xmlInterestRateType, "RATE") 'SR EP2_1938
                'EP2_1382
                'SR EP2_1938
                Dim dblTemp As Double
                dblTemp = CSafeDbl(xmlGetAttributeText(xmlBaseRate, "BASEINTERESTRATE")) + dblRateAdjustment + dblLoad
                'EP2_2478 set to 2DP
                xmlSetAttributeValue xmlDiscount, "INITIALRATE", CStr(set2DP(dblTemp))
                xmlSetAttributeValue xmlComponent, "INITIALRATE", CStr(set2DP(dblTemp))
                'SR EP2_1938
            Case Else
                'Unknown rate type
                
        End Select
        '
    Next
    intComponentCount = 0
    '
    '*-add MaxLTV to XML
    xmlSetAttributeValue vxmlNode, "MAXLTV", mdblMaxLTV
    '
    ' Footer section
    'EP2_1983/1984 FT section is required for single component as well
    Set xmlFT = vobjCommon.CreateNewElement("FT", vxmlNode)
    For Each xmlLC In xmlLCs
        intComponentCount = intComponentCount + 1
        'Get info on rate steps
        Set xmlInterestRateTypes = xmlLC.selectNodes(".//INTERESTRATETYPE")
        If xmlInterestRateTypes.length > 1 Then

            'EP2_2478
            Set xmlInterestRateType = xmlLC.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=1]") ' Get the first one
            intPrefPeriod = xmlGetAttributeAsInteger(xmlInterestRateType, "INTERESTRATEPERIOD")
            'Steps
            Set xmlInterestRateType = xmlLC.selectSingleNode(".//INTERESTRATETYPE[@INTERESTRATETYPESEQUENCENUMBER=2]") ' Get the second one
            dblBaseRate = xmlGetAttributeAsDouble(xmlInterestRateType.selectSingleNode(".//BASERATE"), "BASEINTERESTRATE")
            dblDifference = xmlGetAttributeAsDouble(xmlInterestRateType.selectSingleNode(".//BASERATEBAND"), "RATEDIFFERENCE")
            'EP2_2395
            strBaseType = xmlGetAttributeText(xmlInterestRateType.selectSingleNode(".//BASERATE"), "RATEDESCRIPTION")
            If dblDifference > 0 Then
                strDifference = " above "
            ElseIf dblDifference < 0 Then
                strDifference = " below "
            Else
                strDifference = ""
            End If
            'EP2_2478 to 2DP
            strStep4ReferenceRateStep = "On part " & intComponentCount & ", after " & intPrefPeriod & _
                " months, a variable tracker rate which is " & CStr(set2DP(Abs(dblDifference))) & "% " & strDifference & " the " & _
                strBaseType & " rate currently " & CStr(set2DP(dblBaseRate)) & "% will apply for the remaining term " & _
                "of the mortgage to give a current rate payable of " & CStr(set2DP((dblBaseRate + dblDifference))) & "%."
            Set xmlReferenceRate = vobjCommon.CreateNewElement("ReferenceRate1", xmlFT)
            xmlSetAttributeValue xmlReferenceRate, "ReferenceRateStep", strStep4ReferenceRateStep
        End If
    Next
    '
    If xmlFT Is Nothing Then
        'EP2_1983/1984 FT section is required for single component as well
        Set xmlFT = vobjCommon.CreateNewElement("FT", vxmlNode)
    End If
    
        
    If notAllPrime Then
        Call vobjCommon.CreateNewElement("FINANCIALSTATUS", vxmlNode)
    End If
    
    '
    'Check for BBR rate
    'Will check combo validation when one exists!
    ' SR EP2_2159 - Check for first INTERESTRATETYPE record which is not fixed. On a mulit-component case, you can only have
    '               all the products based on either BOE or LIBOR
    intRateType = 0
    For Each xmlLC In xmlLCs
        For Each xmlInterestRateType In xmlLC.selectNodes(".//INTERESTRATETYPE")
            If xmlGetAttributeText(xmlInterestRateType, "RATETYPE") <> "F" Then
                intRateType = xmlGetAttributeAsInteger(xmlInterestRateType.selectSingleNode(".//BASERATE"), "RATEID")
                Exit For
            End If
        Next xmlInterestRateType
        If intRateType <> 0 Then
            Exit For
        End If
    Next xmlLC
    ' SR EP2_2159 - End
    If intRateType = 1 Then
        ' BBR (Bank of England Base Rate)
        strBBR = "Following a change in the Bank of England base rate, the mortgage rate applicable " & _
                "will be adjusted from the following business day."
     ElseIf intRateType <> 0 Then ' SR EP2_2159
        strBBR = "The LIBOR rate is reviewed on the 15th March, June, September and December each year. " & _
                "Following any changes, the mortgage rate applicable will be adjusted from the following " & _
                "business day."
    End If
    '
    Set xmlReferenceRate = vobjCommon.CreateNewElement("ReferenceRate2", xmlFT)
    xmlSetAttributeValue xmlReferenceRate, "ReferenceRateChange", strBBR
    '
    Set xmlItem = Nothing
    Set xmlLC = Nothing
    Set xmlInitialIntRate = Nothing
    Set xmlSubsequentIntRate = Nothing
    Set xmlMainComp = Nothing
    Set xmlProduct = Nothing
    Set xmlFixedNode = Nothing
    Set xmlDiscountNode = Nothing

Exit Sub
ErrHandler:
    Set xmlItem = Nothing
    Set xmlLC = Nothing
    Set xmlInitialIntRate = Nothing
    Set xmlSubsequentIntRate = Nothing
    Set xmlMainComp = Nothing
    Set xmlProduct = Nothing
    Set xmlFixedNode = Nothing
    Set xmlDiscountNode = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

'EP2_139 New Function
'********************************************************************************
'** Function:       BuildCommonOfferRiskSection
'** Created by:     Ian Ross
'** Date:           26/05/2004
'** Description:    Sets the elements and attributes for the compulsory RiskWarning
'**                 element.
'** Parameters:     vxmlNode - the element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub BuildCommonOfferRiskSection(ByVal vobjCommon As CommonDataHelper, _
        ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildCommonOfferRiskSection"
    
    Dim xmlRisk As IXMLDOMNode
    Dim strNatureOfLoan As String
   
    On Error GoTo ErrHandler
    'EP2_2395 CheckForValidationType(strNatureOfLoan, "BI")
    strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN_VALIDID")
    If Not Len(strNatureOfLoan) > 0 Then
        strNatureOfLoan = xmlGetAttributeText(vobjCommon.Data.selectSingleNode(gcstrAPPLICATIONFACTFIND_PATH), "NATUREOFLOAN")
        strNatureOfLoan = GetValidationTypeForValueID("NatureOfLoan", strNatureOfLoan)
    End If
    
    If CheckForValidationType(strNatureOfLoan, "BI") Or CheckForValidationType(strNatureOfLoan, "BR") Then    'Buy To Let
        Set xmlRisk = vobjCommon.CreateNewElement("RISKWARNINGBTL", vxmlNode)
    Else
        Set xmlRisk = vobjCommon.CreateNewElement("RISKWARNING", vxmlNode)
    End If

    Set xmlRisk = Nothing
Exit Sub
ErrHandler:
    Set xmlRisk = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Sub

Private Sub AddMortgageType1(vobjCommon As CommonDataHelper, vxmlNode As IXMLDOMNode)
    '
    If vobjCommon.GetMainMortgageTypeGroup = "F" Then
        'Additional borrowing
        xmlSetAttributeValue vxmlNode, "MORTGAGETYPE1", "additional borrowing"
        xmlSetAttributeValue vxmlNode, "MORTGAGETYPE3", "mortgage for the additional borrowing"
    Else
        xmlSetAttributeValue vxmlNode, "MORTGAGETYPE1", "mortgage"
        xmlSetAttributeValue vxmlNode, "MORTGAGETYPE3", "mortgage"
    End If
    '
End Sub
