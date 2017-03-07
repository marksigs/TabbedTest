VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransferOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         TransferOfer
'** Created by:     Andy Maggs
'** Date:           14/04/2004
'** Description:    Genderate ToEOffer documents
'********************************************************************************
'Prog   Date        Description
'BC     16/12/2005  MAR907 - Remortgages, ToE and Product Switch
'BC     18/05/2006  MAR1284  Retrieve Names Added and Removed from Memopad
'********************************************************************************
'PB     22/11/2006  EP2_139  Modified 'Names Added and Removed' to work with current method.
'PB     12/12/2006  EP2_422  KFI changes for phase 2
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
'** Date:           21/06/2004
'** Description:    The method responsible for formatting the data into the XML
'**                 format required for the transfer of equity mortgage offer
'**                 document template.
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
'    Set xmlNode = mobjCommon.CreateNewElement("COPYTO", xmlTemplate)
'    Call BuildCopyTo(xmlNode)
    
    '*-create the mandatory AUTONUMBERING element and set the attributes
    Set xmlNode = mobjCommon.CreateNewElement("AUTONUMBERING", xmlTemplate)
    Call AddAutonumberingAttributes(xmlNode)

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
    'Call AddMortgageCommencementDateAttribute(xmlNode)

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
    Set xmlNode = mobjCommon.CreateNewElement("SECTION13", xmlTemplate)
    Call BuildSection13(xmlNode)

    '*-create the mandatory SECTION14 (Where can you get more information on mortgages?) element
    Call mobjCommon.CreateNewElement("SECTION14", xmlTemplate)
    
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
'** Date:           21/06/2004
'** Description:    Builds the template data section of the document.
'** Parameters:     vxmlNode - the node to add the template data to.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildTemplateData(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildTemplateData"

    On Error GoTo ErrHandler

    Call BuildCommonOfferTemplateData(mobjCommon, vxmlNode)
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText) 'BC MAR907
    xmlSetAttributeValue vxmlNode, "PROVIDER", mobjCommon.Provider 'BC MAR907
    xmlSetAttributeValue vxmlNode, "TOTALLOANAMOUNT", mobjCommon.TotalLoanAmount 'BC MAR1284


Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
' BC 05/Jan/2004 - E2EM00003121 - Remove COPYTO element
'********************************************************************************
'** Function:       BuildCopyTo
'** Created by:     Andy Maggs
'** Date:           21/06/2004
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
'** Date:           21/06/2004
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
'** Date:           06/07/2004
'** Description:    Sets the attributes for the compulsory section2 (Which service
'**                 are we providing you with?) element.
'** Parameters:     vxmlNode - the section1 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection2(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection2"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim iAdded As Integer
    Dim iRemoved As Integer
    Dim blnCustomerAdded As Boolean
    Dim xmlAdded As IXMLDOMNodeList 'BC MAR1284
    Dim xmlRemoved As IXMLDOMNodeList 'BC MAR1284
    Dim xmlNamesAdded As IXMLDOMNode 'BC MAR1284
    Dim xmlNamesRemoved As IXMLDOMNode 'BC MAR1284
    Dim strNamesAdded As String 'BC MAR1284
    Dim strNamesRemoved As String 'BC MAR1284

    
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
    

    'BC MAR1284 Begin
    Set xmlAdded = mobjCommon.Data.selectNodes(gcstrMEMOPAD_PATH & "[contains(@ENTRYTYPE_VALIDID,'TOEA')]")

    If Not (xmlAdded Is Nothing) Then
        iAdded = xmlAdded.length
    End If
    
    Set xmlRemoved = mobjCommon.Data.selectNodes(gcstrMEMOPAD_PATH & "[contains(@ENTRYTYPE_VALIDID,'TOER')]")

    If Not (xmlRemoved Is Nothing) Then
        iRemoved = xmlRemoved.length
    End If
    'BC MAR1284 END
    
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
    
    Set xmlList = Nothing
    Set xmlRemoved = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlRemoved = Nothing
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
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim blnIsValuation As Boolean
    Dim xmlAddedList As IXMLDOMNodeList 'BC MAR1284
    Dim xmlRemovedList As IXMLDOMNodeList 'BC MAR1284
    Dim xmlAdded As IXMLDOMNode 'BC MAR1284
    Dim xmlRemoved As IXMLDOMNode 'BC MAR1284
    Dim xmlNamesAdded As IXMLDOMNode 'BC MAR1284
    Dim xmlNamesRemoved As IXMLDOMNode 'BC MAR1284
    Dim strNamesAdded As String 'BC MAR1284
    Dim strNamesRemoved As String 'BC MAR1284
    Dim intNamesAddedRemoved As Integer 'BC MAR1284
    
    On Error GoTo ErrHandler

    '*-add the ESTIMATEDVALUE attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, False, blnValuationUsed)
    '*-add TERMYEARS and TERMMONTHS attributes
    Call AddMortgageTermAttributes(mobjCommon, vxmlNode, "TERMYEARS", "TERMMONTHS")
    '*-add the TOTALLOANAMOUNT attribute
    Call xmlSetAttributeValue(vxmlNode, "TOTALLOANAMOUNT", CStr(mobjCommon.LoanAmount))
    '*-add the MORTGAGETYPE attribute
    Call xmlSetAttributeValue(vxmlNode, "MORTGAGETYPE", mobjCommon.MortgageTypeText)
    '*-add the LTV attribute
    Call xmlSetAttributeValue(vxmlNode, "LTV", CStr(mobjCommon.LTV))
    '*-add the PROPERTYVALUATIONAMOUNT attribute
    Call AddPurchPriceEstValueAttribute(mobjCommon.Data, vxmlNode, _
            Not mobjCommon.IsAdditionalBorrowing, blnIsValuation, _
            "PROPERTYVALUATIONAMOUNT")
    
    If mobjCommon.IsAdditionalBorrowing Then
        '*-now add the ADDITIONALBORROWING element to the PURPOSE node
        Call mobjCommon.CreateNewElement("ADDITIONALBORROWING", vxmlNode)
    End If
    
    '*-add the appropriate (1-16) mandatory FEESANDPREMIUMS element (exactly one)
    Call AddFeesAndPremiumsElement(mobjCommon, vxmlNode)
    
    '******************************
    '*-NB ADDNAMES AND REMOVENAMES WILL NOT WORK CORRECTLY UNTIL THE NEW
    '*-CUSTOMERADDED AND CUSTOMERREMOVED TABLES ARE ADDED TO THE DATABASE AND THE
    '*-APPROPRIATE SCHEMA IS ADDED TO omSchema.xml (SEE SECTION 2 FOR DETAILS)
    
    '*-add the ADDNAMES element and NAMESTOADD attribute if applicable
    
'    Call AddCustomerAddedRemovedData((gcstrCUSTOMERVERSION_PATH & "[contains(@NEWTOECUSTOMERIND,'1')]"), "ADDNAMES", _
            "NAMESTOADD", vxmlNode)
    
    '*-add the REMOVENAMES element and NAMESTOREMOVE attribute if applicable
'    Call AddCustomerAddedRemovedData(gcstrREMOVEDTOECUSTOMER_PATH, "REMOVENAMES", _
            "NAMESTOREMOVE", vxmlNode)
    '******************************
    
'PB 22/11/2006 EP2_139
    'This entire section of code commented-out because the added/removed list of customers does not work
    'in the way implied here. There are no tables called CUSTOMERADDED or CUSTOMERREMOVED, only a table
    'called REMOVEDTOECUSTOMER which contains both. To determine which type a customer is, check the
    'TYPE column: A=added, R=removed.
    '
    'BC MAR1284 Begin
    'Set xmlAddedList = mobjCommon.Data.selectNodes(gcstrMEMOPAD_PATH & "[contains(@ENTRYTYPE_VALIDID,'TOEA')]")
    '
    'intNamesAddedRemoved = 1
    'If Not (xmlAddedList.length = 0) Then
    '    For Each xmlAdded In xmlAddedList
    '        If intNamesAddedRemoved = 1 Then
    '            strNamesAdded = xmlGetAttributeText(xmlAdded, "MEMOENTRY")
    '        Else
    '            If intNamesAddedRemoved < xmlAddedList.length Then
    '                strNamesAdded = strNamesAdded & ", " & xmlGetAttributeText(xmlAdded, "MEMOENTRY")
    '            Else
    '                strNamesAdded = strNamesAdded & " and " & xmlGetAttributeText(xmlAdded, "MEMOENTRY")
    '            End If
    '        End If
    '
    '        intNamesAddedRemoved = intNamesAddedRemoved + 1
    '    Next xmlAdded
    '
    '    Set xmlNamesAdded = mobjCommon.CreateNewElement("ADDNAMES", vxmlNode)
    '    Call xmlSetAttributeValue(xmlNamesAdded, "NAMESTOADD", strNamesAdded)
    '
    'End If
    '
    'Set xmlRemovedList = mobjCommon.Data.selectNodes(gcstrMEMOPAD_PATH & "[contains(@ENTRYTYPE_VALIDID,'TOER')]")
    '
    'intNamesAddedRemoved = 1
    'If Not (xmlRemovedList.length = 0) Then
    '    For Each xmlRemoved In xmlRemovedList
    '    If intNamesAddedRemoved = 1 Then
    '        strNamesRemoved = xmlGetAttributeText(xmlRemoved, "MEMOENTRY")
    '    Else
    '       If intNamesAddedRemoved < xmlRemovedList.length Then
    '            strNamesRemoved = strNamesRemoved & ", " & xmlGetAttributeText(xmlRemoved, "MEMOENTRY")
    '        Else
    '            strNamesRemoved = strNamesRemoved & " and " & xmlGetAttributeText(xmlRemoved, "MEMOENTRY")
    '        End If
    '    End If
    '
    '    intNamesAddedRemoved = intNamesAddedRemoved + 1
    '    Next xmlRemoved
    '
    '    Set xmlNamesRemoved = mobjCommon.CreateNewElement("REMOVENAMES", vxmlNode)
    '    Call xmlSetAttributeValue(xmlNamesRemoved, "NAMESTOREMOVE", strNamesRemoved)
    '
    'End If
    'BC MAR1284 End
    '
'PB 22/11/2006 EP2_139 Begin
    'New code added here:
    Set xmlAddedList = mobjCommon.Data.selectNodes("//APPLICATION/REMOVEDTOECUSTOMER[@TYPE='A']")
    Set xmlRemovedList = mobjCommon.Data.selectNodes("//APPLICATION/REMOVEDTOECUSTOMER[@TYPE='R']")
    'Build lists of names
    For Each xmlAdded In xmlAddedList
        strNamesAdded = strNamesAdded & xmlAdded.Attributes("TITLE") & " " & xmlAdded.Attributes("FIRSTFORENAME") & " " & xmlAdded.Attributes("SURNAME") & ", "
    Next
    If strNamesAdded <> "" Then
        strNamesAdded = Left(strNamesAdded, Len(strNamesAdded) - 2)
        Set xmlNamesAdded = mobjCommon.CreateNewElement("ADDNAMES", vxmlNode)
        xmlSetAttributeValue xmlNamesAdded, "NAMESTOADD", strNamesAdded
    End If
    '
    For Each xmlRemoved In xmlRemovedList
        strNamesRemoved = strNamesRemoved & xmlRemoved.Attributes("TITLE") & " " & xmlRemoved.Attributes("FIRSTFORENAME") & " " & xmlRemoved.Attributes("SURNAME") & ", "
    Next
    If strNamesRemoved <> "" Then
        strNamesRemoved = Left(strNamesRemoved, Len(strNamesRemoved) - 2)
        Set xmlNamesRemoved = mobjCommon.CreateNewElement("REMOVENAMES", vxmlNode)
        xmlSetAttributeValue xmlNamesRemoved, "NAMESTOREMOVE", strNamesRemoved
    End If
'PB 22/11/2006 EP2_139 End
    '
    '
    '
    '*-add the type of mortgage mandatory REPAYMENT, INTERESTONLY or
    '*-PARTANDPART elements
    Call mobjCommon.AddMortgageRepaymentTypeElement(vxmlNode)
    
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
Private Sub AddCustomerAddedRemovedData(ByVal vstrXPath As String, _
        ByVal vstrNodeName As String, ByVal vstrAttribName As String, _
        ByVal vxmlNode As IXMLDOMNode)
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim strNamesList As String
    Dim xmlNames As IXMLDOMNode

    Set xmlList = mobjCommon.Data.selectNodes(vstrXPath)
    If Not xmlList Is Nothing Then
        For Each xmlItem In xmlList
            If Len(strNamesList) > 0 Then
                strNamesList = strNamesList & vbCrLf
            End If
            strNamesList = strNamesList & BuildName(xmlGetAttributeText(xmlItem, "TITLE_TEXT"), _
                    xmlGetAttributeText(xmlItem, "FORENAME"), _
                    xmlGetAttributeText(xmlItem, "SURNAME"))
        Next xmlItem

        If Len(strNamesList) > 0 Then
            '*-add the ADDNAMES element
            Set xmlNames = mobjCommon.CreateNewElement(vstrNodeName, vxmlNode)
            '*-add the NAMESTOADD attribute
            Call xmlSetAttributeValue(xmlNames, vstrAttribName, strNamesList)
        End If
    End If

    Set xmlList = Nothing
    Set xmlItem = Nothing
    Set xmlNames = Nothing
End Sub

'********************************************************************************
'** Function:       BuildSection4
'** Created by:     Andy Maggs
'** Date:           21/06/2004
'** Description:    Sets the elements and attributes for the compulsory section4
'**                 (Description of this mortgage) element.
'** Parameters:     vxmlNode - the section4 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection4(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection4"
    
    On Error GoTo ErrHandler
    
    '*-strictly the XML produced by this 'common' method contains two extra
    '*-occurrences of the MORTGAGETYPE attribute that are not required by the
    '*-Transfer of Equity document, however, the template processor does not
    '*-care about additional attributes, so we are not going to waste time
    '*-removing them
    Call BuildCommonOfferSection4(mobjCommon, vxmlNode)
    
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection5
'** Created by:     Andy Maggs
'** Date:           06/07/2004
'** Description:    Sets the elements and attributes for the compulsory section5
'**                 (Overall Costs) element.
'** Parameters:     vxmlNode - the section5 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection5(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection5"
    Dim xmlComment As IXMLDOMNode
    Dim xmlChild As IXMLDOMNode
    Dim xmlRepay As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    '*-strictly the XML produced by this 'common' method contains two extra
    '*-occurrences of the MORTGAGETYPE attribute that are not required by the
    '*-Transfer of Equity document, however, the template processor does not
    '*-care about additional attributes, so we are not going to waste time
    '*-removing them
    Call BuildOverallCostsSection(mobjCommon, vxmlNode, True)
    
'   BC MAR907
'    '*-get the PAYMENTCOMMENT2 element
'    Set xmlComment = vxmlNode.selectSingleNode("PAYMENTCOMMENT2")
'    '*-get its child element which will be either ADDITIONALBORROWING or
'    '*-MORTGAGE as appropriate
'    Set xmlChild = xmlComment.childNodes.Item(0)
'    '*-get the repayment type from the child
'    Set xmlRepay = xmlChild.childNodes.Item(0)
'
'    '*-remove the child node
'    Call xmlComment.removeChild(xmlChild)
'    '*-and add the repayment type node instead
'    Call xmlComment.appendChild(xmlRepay)
    
    Set xmlComment = Nothing
    Set xmlChild = Nothing
    Set xmlRepay = Nothing
Exit Sub
ErrHandler:
    Set xmlComment = Nothing
    Set xmlChild = Nothing
    Set xmlRepay = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6
'** Created by:     Andy Maggs
'** Date:           21/06/2004
'** Description:    Sets the elements and attributes for the compulsory section6
'**                 (What you will need to pay each month) element.
'** Parameters:     vxmlNode - the section6 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection6(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection6"
    
    On Error GoTo ErrHandler
    
    '*-strictly the XML produced by this 'common' method contains two extra
    '*-occurrences of the MORTGAGETYPE attribute that are not required by the
    '*-Transfer of Equity document, however, the template processor does not
    '*-care about additional attributes, so we are not going to waste time
    '*-removing them
    Call BuildCommonOfferSection6(mobjCommon, vxmlNode)

Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection6A
'** Created by:     Andy Maggs
'** Date:           21/06/2004
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
'** Date:           21/06/2004
'** Description:    Sets the elements and attributes for the compulsory section7
'**                 (Are you comfortable with the risks?) element.
'** Parameters:     vxmlNode - the section7 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection7(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection7"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlComponent As IXMLDOMNode
    Dim xmlPart As IXMLDOMAttribute
    Dim xmlMulti As IXMLDOMNode
    
    On Error GoTo ErrHandler

    Call BuildCommonOfferSection7(mobjCommon, vxmlNode)
        
    '*-remove the MORTGAGETYPE attribute from the section
    Call vxmlNode.Attributes.removeNamedItem("MORTGAGETYPE")
    
    '*-get all the components
    Set xmlList = vxmlNode.selectNodes("//COMPONENT")
'    For Each xmlComponent In xmlList
'        '*-get and remove the  PART attribute
'        Set xmlPart = xmlComponent.Attributes.removeNamedItem("PART")
'
'        '*-create a new MULTICOMPONENT element
'        Set xmlMulti = mobjCommon.CreateNewElement("MULTICOMPONENT", xmlComponent)
'        '*-add the PART attribute to it
'        Call xmlMulti.Attributes.setNamedItem(xmlPart)
'    Next xmlComponent
        
    Set xmlList = Nothing
    Set xmlComponent = Nothing
    Set xmlPart = Nothing
    Set xmlMulti = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlComponent = Nothing
    Set xmlPart = Nothing
    Set xmlMulti = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection7A
'** Created by:     Andy Maggs
'** Date:           21/06/2004
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

    '*-strictly the XML produced by this 'common' method contains extra
    '*-occurrences of the MORTGAGETYPE attribute on the NOTADDEDTOLOAN element
    '*-that are not required by the Transfer of Equity document, however, the
    '*-template processor does not care about additional attributes, so we are
    '*-not going to waste time removing them
    Call BuildWhatFeesYouMustPaySection(mobjCommon, vxmlNode, True)
    
Exit Sub
ErrHandler:
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection9
'** Created by:     Andy Maggs
'** Date:           21/06/2004
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
'** Date:           21/06/2004
'** Description:    Sets the elements and attributes for the compulsory section10
'**                 (What happens if you do not want this mortgage any more?) element.
'** Parameters:     vxmlNode - the section10 element to set the attributes on.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub BuildSection10(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "BuildSection10"
    Dim xmlSummary As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlPort As IXMLDOMNode
    
    On Error GoTo ErrHandler

    Call BuildMortgageNoMoreSection(mobjCommon, vxmlNode, True, False)
    
'   BC MAR907
'    Set xmlSummary = vxmlNode.selectSingleNode("//SUMMARY")
'    If Not xmlSummary Is Nothing Then
'        '*-remove the SINGLECOMPONENT/MULTICOMPONENT element from the SUMMARY element
'        Set xmlComponent = xmlSummary.selectSingleNode("SINGLECOMPONENT")
'        If xmlComponent Is Nothing Then
'            Set xmlComponent = xmlSummary.selectSingleNode("MULTICOMPONENT")
'        End If
'
'        If Not xmlComponent Is Nothing Then
'            '*-remove it
'            Call xmlSummary.removeChild(xmlComponent)
'        End If
'
'        '*-add the MORTGAGETYPE attribute to the SUMMARY element
'        Call AddMortgageTypeAttribute(mobjCommon, xmlSummary)
'
'    End If
'
'    '*-remove the MORTGAGETYPE attribute from the PORTABLE/NOTPORTABLE element
'    Set xmlPort = vxmlNode.selectSingleNode("//PORTABLE")
'    If xmlPort Is Nothing Then
'        Set xmlPort = vxmlNode.selectSingleNode("//NOTPORTABLE")
'    End If
'
'    If Not xmlPort Is Nothing Then
'        '*-remove the MORTGAGETYPE attribute from it
'        Call xmlPort.Attributes.removeNamedItem("MORTGAGETYPE")
'    End If
    
    Set xmlSummary = Nothing
    Set xmlComponent = Nothing
    Set xmlPort = Nothing
Exit Sub
ErrHandler:
    Set xmlSummary = Nothing
    Set xmlComponent = Nothing
    Set xmlPort = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection11
'** Created by:     Andy Maggs
'** Date:           21/06/2004
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
    Dim xmlList As IXMLDOMNodeList
    Dim xmlIncentive As IXMLDOMNode
    Dim xmlValue As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Call BuildAdditionalFeaturesSection(mobjCommon, vxmlNode, True, False)
        
    '*-NB if the xml above contains SINGLECOMPONENT elements then each one
    '*-contains a MORTGAGETYPE attribute that is surplus to requirements for
    '*-this document type, however, it does not cause the template processor
    '*-any problems, so we have not removed them
    
    '*-modify the INCENTIVES element
    
'   BC MAR907
'    Set xmlList = vxmlNode.selectNodes("//INCENTIVES")
'    For Each xmlIncentive In xmlList
'        '*-move the INCENTIVEAMOUNT attribute from the child MONETARYVALUE element to
'        '*-the INCENTIVES element
'        Set xmlValue = xmlIncentive.selectSingleNode("MONETARYVALUE")
'        If Not xmlValue Is Nothing Then
'            Call xmlCopyAttribute(xmlValue, xmlIncentive, "INCENTIVEAMOUNT")
'            '*-remove the MONETARYVALUE element
'            Call xmlIncentive.removeChild(xmlValue)
'        End If
'        '*-remove the MORTGAGETYPE attribute
'        Call xmlIncentive.Attributes.removeNamedItem("MORTGAGETYPE")
'    Next xmlIncentive
    
    Set xmlList = Nothing
    Set xmlIncentive = Nothing
    Set xmlValue = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlIncentive = Nothing
    Set xmlValue = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'********************************************************************************
'** Function:       BuildSection13
'** Created by:     Andy Maggs
'** Date:           21/06/2004
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
'** Date:           21/06/2004
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
'** Date:           21/06/2004
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

