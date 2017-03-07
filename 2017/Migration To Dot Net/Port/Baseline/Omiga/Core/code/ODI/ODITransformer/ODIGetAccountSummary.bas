Attribute VB_Name = "ODIGetAccountSummary"
'Workfile:      ODIGetAccountSummary.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     04/09/01    Expanded class created by LD.
'DS     23/03/02    Loads of fixes to make it work. SYS4306.
'DS     17/04/02    Telephonnumberlist was not being output correctly.
'DS     30/04/02    Updates to fix problems with PRODUCTANDLOANS section
'DS     30/04/02    PAYMENT details corrected. SYS4468.
'DS     1/05/02     OUTSTANDINGBALANCE and REPAYENTINSTALLMENT fixed. SYS4523.
'DS     1/05/02     INTERESTRATE AND LTV made into percentages. SYS4524.
'JLD    21/05/02    SYS4665 changed VALUATIONAMOUNT to equal just the APPRAISEDTOTALVALUE
'DS     19/06/02    SYS4865 PaidInAdvance should now have correct value
'DS     35/06/02    SYS4702 Added an extra test to PaidInAdvance being populated
'STB    11/07/02    SYS4610 Prefix the area code with a zero. Optimus strips it.
'SG     28/03/03    SYS6150 Included MORTGAGELOANLIST/MORTGAGELOAN in response
'AS     02/07/03    MSMS178 OMIP110 Fixes to arrears
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub GetAccountSummary( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountSummaryExit
    
    Const strFunctionName As String = "GetAccountSummary"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeMortgageAccount As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    Dim strAccountNumber As String
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodePCKey As IXMLDOMNode
    Dim nodelistPrimaryCustomerKeys As IXMLDOMNodeList
    Dim nodeOmigaCustomerList As IXMLDOMNode
    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodelistCustomerObjects As IXMLDOMNodeList
    Dim nodeCustomer As IXMLDOMNode

    '------------------------------------------------------------------------------------------
    ' find which account number is required
    '------------------------------------------------------------------------------------------
    
    Set nodeTemp = xmlGetMandatoryNode(vxmlRequestNode, "MORTGAGEACCOUNT")
    strAccountNumber = xmlGetMandatoryAttributeText(nodeTemp, "ACCOUNTNUMBER")
    
    '-----------------------------------------------------------------------------------------
    ' create MORTGAGEACCOUNT response element
    '------------------------------------------------------------------------------------------
    
    Set nodeMortgageAccount = _
            vxmlResponseNode.ownerDocument.createElement("ACCOUNTSUMMARY")
        
    '------------------------------------------------------------------------------------------
    ' Add Mortgage details
    '------------------------------------------------------------------------------------------
    
    ' this method both finds the Mortgage and handles the response
    AddMortgageDetails vobjODITransformerState, strAccountNumber, nodeMortgageAccount
            
    '------------------------------------------------------------------------------------------
    ' get list of PrimaryCustomers
    '------------------------------------------------------------------------------------------
        
    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_CustomerListPattern( _
            vobjODITransformerState, strAccountNumber)
            
    CheckConverterResponse nodeConverterResponse, True
    
    Set nodelistCustomerObjects = nodeConverterResponse.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT")

    ' customer list node
    Set nodeOmigaCustomerList = _
        vxmlResponseNode.ownerDocument.createElement("CUSTOMERLIST")
    nodeMortgageAccount.appendChild nodeOmigaCustomerList
    
    '------------------------------------------------------------------------------------------
    ' add PrimaryCustomer details
    '------------------------------------------------------------------------------------------
    
    For Each nodeCustomer In nodelistCustomerObjects
        AddCustomerDetails vobjODITransformerState, nodeCustomer, nodeOmigaCustomerList
    Next nodeCustomer
    
    vxmlResponseNode.appendChild nodeMortgageAccount

    AddExceptionsToResponse vxmlResponseNode, vxmlResponseNode

GetAccountSummaryExit:

    Set nodeMortgageAccount = Nothing
    Set nodeTemp = Nothing
    Set nodeConverterResponse = Nothing
    Set nodePCKey = Nothing
    Set nodelistPrimaryCustomerKeys = Nothing
    Set nodeOmigaCustomerList = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub

Private Sub AddMortgageDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vstrAccountNumber As String, _
    ByVal vnodeTarget As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Get the details for the mortgage having the a key of
'   vstrAccountNumber and add them to vnodeTarget.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddMortgageDetailsExit

    Const strFunctionName As String = "AddMortgageDetails"
    
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodeMortgageKey As IXMLDOMNode
    Dim nodeMortgageAccount As IXMLDOMNode
    Dim nodeOptimusCharge As IXMLDOMNode
    Dim elemProperty As IXMLDOMElement
    Dim elemPropertyAddress As IXMLDOMElement
    Dim elemMortgageLoanList As IXMLDOMElement
    Dim elemArrearsHistory As IXMLDOMElement
    Dim nodeTemp As IXMLDOMNode
    Dim strTemp As String
    Dim strTempValue1 As String
    Dim strTempValue2 As String
    Dim elemLTV As IXMLDOMElement
    Dim elemSuspenseAccount As IXMLDOMElement
    Dim elemAdvanceAccount As IXMLDOMElement
    Dim elemArrearsDetails As IXMLDOMElement
    Dim elemPaymentDetails As IXMLDOMElement
    Dim elemAdvance As IXMLDOMElement
    Dim elemBuildingsContentsInsurances As IXMLDOMElement
    Dim elemPaymentProtection As IXMLDOMElement
    Dim elemProductAndLoanDetails As IXMLDOMElement
    Dim nodeOptimusRealEstate As IXMLDOMNode
    Dim nodeComponent As IXMLDOMNode
    Dim nodelistOptimusComponents As IXMLDOMNodeList
    Dim nodelistOptimusPaymentSources As IXMLDOMNodeList
    Dim nodeChargeImpl As IXMLDOMNode
    Dim strPremium As String
    Dim num As Double

    
    '------------------------------------------------------------------------------------------
    ' Call PlexusHome.findByPrimaryKey supplying a MortgageKey
    '------------------------------------------------------------------------------------------
    
    Set nodeMortgageKey = CreateMortgageKey(vstrAccountNumber)
    Set nodeConverterResponse = PlexusHomeImpl_findByPrimaryKey_MortgageKey( _
            vobjODITransformerState, nodeMortgageKey)
            
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    ' get some useful pointers into nodeConverterResponse
    
    Set nodeOptimusRealEstate = xmlGetMandatoryNode(nodeConverterResponse, "ARGUMENTS/OBJECT/MORTGAGEIMPL/REALESTATE/REALESTATEIMPL")
    
    Set nodeOptimusCharge = xmlGetMandatoryNode( _
        nodeConverterResponse, "ARGUMENTS/OBJECT/MORTGAGEIMPL/CHARGE/CHARGEIMPL")
    
    Set nodeChargeImpl = xmlGetNode(nodeOptimusCharge, "COMPONENTS/COMPONENTIMPL")
    If Not nodeChargeImpl Is Nothing Then
        Set nodelistOptimusComponents = nodeChargeImpl.childNodes
    End If
    
        
    '------------------------------------------------------------------------------------------
    ' MORTGAGEACCOUNT element
    '------------------------------------------------------------------------------------------
        
    Dim xmlCollateralNumber As IXMLDOMElement
    Set xmlCollateralNumber = xmlGetNode(nodeConverterResponse, "ARGUMENTS/OBJECT/MORTGAGEIMPL/PRIMARYKEY/MORTGAGEKEY/REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER")
        
    'xmlSetAttributeValue nodeMortgageAccount, "ACCOUNTNUMBER", xmlGetAttributeText(xmlCollateralNumber, "DATA")
    
    '------------------------------------------------------------------------------------------
    ' PROPERTY element
    '------------------------------------------------------------------------------------------
    
    Set elemProperty = vnodeTarget.ownerDocument.createElement("PROPERTY")
    vnodeTarget.appendChild elemProperty

    strTemp = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISALDATE/@DATA")
    elemProperty.setAttribute "VALUATIONDATE", OptimusDateToOmiga(strTemp)

    strTempValue1 = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISEDTOTALVALUE/@DATA")
    'strTempValue2 = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISEDLANDVALUE/@DATA")
    'If strTempValue1 = "" And strTempValue2 = "" Then
    '    strTemp = ""
    'Else
    '    strTemp = CStr(CSafeLng(strTempValue1) + CSafeLng(strTempValue2))
    'End If
    elemProperty.setAttribute "VALUATIONAMOUNT", strTempValue1    'JLD SYS4665
    
    '------------------------------------------------------------------------------------------
    ' PROPERTYADDRESS element
    '------------------------------------------------------------------------------------------

    Set elemPropertyAddress = vnodeTarget.ownerDocument.createElement("PROPERTYADDRESS")
    
    Set nodeTemp = xmlGetNode(nodeOptimusRealEstate, "ADDRESSDETAIL/ADDRESSDETAIL")
    
    If Not nodeTemp Is Nothing Then
        OptimusAddressToOmiga elemPropertyAddress, nodeTemp, True, True
        elemProperty.appendChild elemPropertyAddress
    End If
    
    '------------------------------------------------------------------------------------------
    ' SUSPENSEACCOUNT
    '------------------------------------------------------------------------------------------

    Set elemSuspenseAccount = vnodeTarget.ownerDocument.createElement("SUSPENSEACCOUNT")
    vnodeTarget.appendChild elemSuspenseAccount
    
    Dim strSuspenseBalance As String
    strSuspenseBalance = xmlGetNodeText(nodeOptimusCharge, "BALANCEINSUSPENSE/@DATA")
    If strSuspenseBalance = "" Then
        strSuspenseBalance = "0"
    End If
    
    elemSuspenseAccount.setAttribute "BALANCE", strSuspenseBalance

    '------------------------------------------------------------------------------------------
    ' ARREARSDETAILS
    '------------------------------------------------------------------------------------------

    Set elemArrearsDetails = vnodeTarget.ownerDocument.createElement("ARREARSDETAILS")
    vnodeTarget.appendChild elemArrearsDetails
    
    'AS MSMS178 02/07/03
    strTemp = xmlGetNodeText(nodeOptimusCharge, "ORIGINALFIRSTPAYMENTDATE/@DATA")
    Dim dateOriginalFirstPaymentDate As Date
    dateOriginalFirstPaymentDate = CSafeDate(OptimusDateToOmiga(strTemp))
    If dateOriginalFirstPaymentDate <= Date Then
        strTemp = xmlGetNodeText(nodeOptimusCharge, "ARREARSPAYMENTCOUNTER/@DATA")
        If strTemp = "" Then
            strTemp = "0"
        End If
    Else
        strTemp = "0"
    End If
    'END AS MSMS178 02/07/03
    
    strTemp = strTemp & " PAYMENTS"
    elemArrearsDetails.setAttribute "ARREARSSTATUS", strTemp
 
    '------------------------------------------------------------------------------------------
    ' ADVANCE
    '------------------------------------------------------------------------------------------

    Set elemAdvance = vnodeTarget.ownerDocument.createElement("ADVANCE")
    vnodeTarget.appendChild elemAdvance
    
    strTemp = xmlGetNodeText(nodeOptimusCharge, "FIRSTADVANCEDATE/@DATA")
    elemAdvance.setAttribute "FIRSTDATE", OptimusDateToOmiga(strTemp)

    'SG 28/03/03 SYS6150 Start
    '------------------------------------------------------------------------------------------
    ' MORTGAGELOANLIST element
    '------------------------------------------------------------------------------------------
    
    Dim elemMortgageLoan As IXMLDOMElement
    Dim strOriginalLoanAmount As String
    Dim strAccountNumber As String
    Dim nodeMortgageImpl As IXMLDOMNode
    
    Set nodeMortgageImpl = xmlGetMandatoryNode(nodeConverterResponse, "ARGUMENTS/OBJECT/MORTGAGEIMPL")
    
    strAccountNumber = xmlGetMandatoryNodeText(nodeMortgageImpl, "PRIMARYKEY/MORTGAGEKEY/REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER/@DATA") & _
                        "." & _
                        xmlGetMandatoryNodeText(nodeMortgageImpl, "PRIMARYKEY/MORTGAGEKEY/CHARGETYPE/@DATA")
    
    
    Set elemMortgageLoanList = vnodeTarget.ownerDocument.createElement("MORTGAGELOANLIST")
    vnodeTarget.appendChild elemMortgageLoanList
    'SG 28/03/03 SYS6150 End

    '------------------------------------------------------------------------------------------
    ' PRODUCTANDLOANDETAILS, BC PREMIUM and PP PREMIUM
    '------------------------------------------------------------------------------------------

    Dim dblTotalOutstanding As Double
    Dim dblAdvanceAccount As Double
    Dim dblPaymentAmount As Double
    Dim strLedger As String
    Dim strSubLedger As String
    Dim node As IXMLDOMElement
    
    dblTotalOutstanding = 0
    dblAdvanceAccount = 0
    dblPaymentAmount = 0

    If Not nodelistOptimusComponents Is Nothing Then
        For Each nodeComponent In nodelistOptimusComponents
                
            strLedger = xmlGetNodeText(nodeComponent, ".//LEDGER/@DATA")
            strSubLedger = xmlGetNodeText(nodeComponent, ".//SUBLEDGER/@DATA")
            
            If strLedger = "A" Then
                
                Set elemProductAndLoanDetails = _
                    vnodeTarget.ownerDocument.createElement("PRODUCTANDLOANDETAILS")
                
                ' fixme - need text value, not a code
                strTemp = xmlGetNodeText(nodeComponent, "MORTGAGEPRODUCTCODE/@DATA")
                elemProductAndLoanDetails.setAttribute "PRODUCTDESCRIPTION", _
                    lookupMortgageProductCode(strTemp, cdOptimusToOmiga)
                
                elemProductAndLoanDetails.setAttribute "PRODUCTTYPE", _
                    xmlGetNodeText(nodeComponent, "PAYMENTTYPE/@DATA")
                    
                Dim strCurrentRate As String
                strCurrentRate = xmlGetNodeText(nodeComponent, "CURRENTRATE/@DATA")
                If strCurrentRate = "" Then
                    strCurrentRate = "0"
                End If
                
                num = CDbl(strCurrentRate) * 100
                elemProductAndLoanDetails.setAttribute "INTERESTRATE", num
                
                strTemp = CStr( _
                    CSafeDbl(xmlGetNodeText(nodeComponent, "PRINCIPALBALANCE/@DATA")) + _
                    CSafeDbl(xmlGetNodeText(nodeComponent, "INTERESTINARREARS/@DATA")) + _
                    CSafeDbl(xmlGetNodeText(nodeComponent, "PAIDINADVANCE/@DATA")))
                elemProductAndLoanDetails.setAttribute "OUTSTANDINGBALANCE", strTemp
                
                ' PRODUCTENDDATE and MATURITYDATE are both PIComponent.maturityDate
                strTemp = xmlGetNodeText(nodeComponent, "MATURITYDATE/@DATA")
                strTemp = OptimusDateToOmiga(strTemp)
                elemProductAndLoanDetails.setAttribute "PRODUCTENDDATE", strTemp
                elemProductAndLoanDetails.setAttribute "LOANENDDATE", strTemp
                
                strTemp = xmlGetNodeText(nodeComponent, "PRINCIPALADVANCED/@DATA")
                If CSafeDbl(strTemp) > 0 Then
                    strTemp = xmlGetNodeText(nodeComponent, "FIXEDPAYMENTAMOUNT/@DATA")
                    elemProductAndLoanDetails.setAttribute "REPAYMENTINSTALLMENT", strTemp
                    dblPaymentAmount = dblPaymentAmount + CSafeDbl(strTemp)
                Else
                    elemProductAndLoanDetails.setAttribute "REPAYMENTINSTALLMENT", "0"
                End If
                            
                elemProductAndLoanDetails.setAttribute "REGULAREXTRAPAYMENT", _
                    xmlGetNodeText(nodeComponent, "SECONDARYPAYMENTAMOUNT/@DATA")
                
                'AS MSMS178 02/07/03
                Dim dAmountOfArrears As Double
                dAmountOfArrears = 0
                If dateOriginalFirstPaymentDate <= Date Then
                    strTempValue1 = xmlGetNodeText(nodeComponent, "PRINCIPALINARREARS/@DATA")
                    strTempValue2 = xmlGetNodeText(nodeComponent, "INTERESTINARREARS/@DATA")
                    dAmountOfArrears = CSafeDbl(strTempValue1) + CSafeDbl(strTempValue2)
                End If
                elemProductAndLoanDetails.setAttribute "AMOUNTOFARREARS", dAmountOfArrears
                'END AS MSMS178 02/07/03
                
                vnodeTarget.appendChild elemProductAndLoanDetails
                
                ' dblAdvanceAccount (used to set ADVANCEACCOUNT) is the sum of
                ' Component.paidInAdvance for each component where
                ' ComponentKey.Ledger="A"
                If strLedger = "A" Then
    '                dblAdvanceAccount = dblAdvanceAccount + _
    '                    CSafeDbl(xmlGetAttributeText(nodeComponent, "PAIDINADVANCE/DATA"))
                    Set node = nodeComponent.selectSingleNode("PAIDINADVANCE")
                    If IsObject(node) And Not (node Is Nothing) Then
                        dblAdvanceAccount = dblAdvanceAccount + CSafeDbl(node.getAttribute("DATA"))
                    End If
                    
                End If
                
                ' dblTotalOutstanding (used in LTV calculation) is the sum of
                ' (PIComponent.principalBalance + PIComponent.interestInArrears -
                ' PIComponent.paidInAdvance) for each component
                Set node = nodeComponent.selectSingleNode("PRINCIPALBALANCE")
                If Not node Is Nothing Then
                    dblTotalOutstanding = dblTotalOutstanding + CSafeDbl(node.getAttribute("DATA"))
                End If
                Set node = nodeComponent.selectSingleNode("INTERESTINARREARS")
                If Not node Is Nothing Then
                    dblTotalOutstanding = dblTotalOutstanding + CSafeDbl(node.getAttribute("DATA"))
                End If
                Set node = nodeComponent.selectSingleNode("PAIDINADVANCE")
                If Not node Is Nothing Then
                    dblTotalOutstanding = dblTotalOutstanding + CSafeDbl(node.getAttribute("DATA"))
                End If
            
                'SG 28/03/03 SYS6150 Start
                '----------------------------------------------------------------------------
                ' MORTGAGELOAN
                '----------------------------------------------------------------------------
                Dim strtemp2 As String
       
                Set elemMortgageLoan = vnodeTarget.ownerDocument.createElement("MORTGAGELOAN")
                elemMortgageLoanList.appendChild elemMortgageLoan
                
                elemMortgageLoan.setAttribute "ACCOUNTNUMBER", strAccountNumber
                
                strtemp2 = xmlGetNodeText(nodeComponent, ".//SUBLEDGER/@DATA")
                elemMortgageLoan.setAttribute "LOANACCOUNTNUMBER", _
                    strAccountNumber & "." & strtemp2
        
                elemMortgageLoan.setAttribute "ORIGINALLOANAMOUNT", xmlGetNodeText(nodeComponent, ".//MAXIMUMAPPROVEDLOANAMOUNT/@DATA")
                    
                'SG 28/03/03 SYS6150 End
            
            End If
            
            '------------------------------------------------------------------------------------------
            ' BUILDINGSCONTENTSINSURANCES and PAYMENTPROTECTION
            '------------------------------------------------------------------------------------------
            
            If strLedger = "C" Then
                If strSubLedger = "02" Or strSubLedger = "03" Or strSubLedger = "04" Then
                    strPremium = xmlGetNodeText(nodeComponent, "FIXEDPAYMENTAMOUNT/@DATA")
                    If strPremium <> "" Then
                        Set elemBuildingsContentsInsurances = _
                            vnodeTarget.ownerDocument.createElement("BUILDINGSCONTENTSINSURANCES")
                        vnodeTarget.appendChild elemBuildingsContentsInsurances
                        elemBuildingsContentsInsurances.setAttribute "PREMIUM", strPremium
                    End If
                Else
                    If strSubLedger = "05" Or strSubLedger = "06" Or strSubLedger = "07" Then
                        strPremium = xmlGetNodeText(nodeComponent, "FIXEDPAYMENTAMOUNT/@DATA")
                        If strPremium <> "" Then
                            Set elemPaymentProtection = _
                                vnodeTarget.ownerDocument.createElement("PAYMENTPROTECTION")
                            vnodeTarget.appendChild elemPaymentProtection
                            
                            elemPaymentProtection.setAttribute "PREMIUM", strPremium
                        End If
                    End If
                End If
            End If
        
        Next nodeComponent
    End If

    '------------------------------------------------------------------------------------------
    ' LTV
    '------------------------------------------------------------------------------------------

    Set elemLTV = vnodeTarget.ownerDocument.createElement("LTV")
    vnodeTarget.appendChild elemLTV
    
    strTempValue1 = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISEDTOTALVALUE/@DATA")
    If CSafeDbl(strTempValue1) > 0 Then
        strTemp = CStr(dblTotalOutstanding * 100 / CSafeDbl(strTempValue1))
    Else
        strTemp = ""
    End If
    elemLTV.setAttribute "RATIOVALUE", strTemp

    '------------------------------------------------------------------------------------------
    ' ADVANCEACCOUNT
    '------------------------------------------------------------------------------------------

    Set elemAdvanceAccount = vnodeTarget.ownerDocument.createElement("ADVANCEACCOUNT")
    vnodeTarget.appendChild elemAdvanceAccount
    
    elemAdvanceAccount.setAttribute "BALANCE", CStr(dblAdvanceAccount)

    '------------------------------------------------------------------------------------------
    ' PAYMENTDETAILS
    '------------------------------------------------------------------------------------------

    Set elemPaymentDetails = vnodeTarget.ownerDocument.createElement("PAYMENTDETAILS")
    vnodeTarget.appendChild elemPaymentDetails
    
    strTemp = xmlGetNodeText(nodeOptimusCharge, "NEXTPAYMENTDATE/@DATA")
    elemPaymentDetails.setAttribute "PAYMENTDATE", OptimusDateToOmiga(strTemp)
    
    ' PAYMENTMETHOD:
    ' lookup PaymentSourceKey.paymentMethod where
    ' PaymentSourceKey.sequenceNumber is the greatest
       
    Set nodelistOptimusPaymentSources = _
        nodeOptimusCharge.selectNodes("PAYMENTSOURCES")
        
    Dim strThisSeqNo As String
    Dim intMaxSeqNo As Integer
    Dim strPaymentMethod As String
        
    For Each nodeTemp In nodelistOptimusPaymentSources
        strThisSeqNo = xmlGetMandatoryNodeText(nodeTemp, _
            ".//PAYMENTSOURCEKEY/SEQUENCENUMBER/@DATA")
        
        If CSafeInt(strThisSeqNo) > intMaxSeqNo Then
            intMaxSeqNo = CSafeInt(strThisSeqNo)
            strPaymentMethod = xmlGetNodeText(nodeTemp, _
                ".//PAYMENTSOURCEKEY/PAYMENTMETHOD/@DATA")
        End If
                        
    Next nodeTemp
    
    elemPaymentDetails.setAttribute "PAYMENTMETHOD", strPaymentMethod
    
    elemPaymentDetails.setAttribute "PAYMENTAMOUNT", CStr(dblPaymentAmount)
    
   
    
AddMortgageDetailsExit:

    'SG 28/03/03 SYS6150 Start
    Set elemMortgageLoan = Nothing
    Set nodeMortgageImpl = Nothing
    'SG 28/03/03 SYS6150 End

    Set nodeConverterResponse = Nothing
    Set nodeMortgageAccount = Nothing
    Set elemProperty = Nothing
    Set elemPropertyAddress = Nothing
    Set elemMortgageLoanList = Nothing
    Set elemArrearsHistory = Nothing
    Set nodeTemp = Nothing
    Set nodeMortgageKey = Nothing
    Set elemLTV = Nothing
    Set elemSuspenseAccount = Nothing
    Set elemAdvanceAccount = Nothing
    Set elemArrearsDetails = Nothing
    Set elemPaymentDetails = Nothing
    Set elemAdvance = Nothing
    Set elemBuildingsContentsInsurances = Nothing
    Set elemPaymentProtection = Nothing
    Set elemProductAndLoanDetails = Nothing
    Set nodeOptimusRealEstate = Nothing
    Set nodeComponent = Nothing
    
    errCheckError strFunctionName

End Sub

Private Sub AddCustomerDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeCustomerObject As IXMLDOMNode, _
    ByVal vnodeTarget As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Add to vnodeTarget the details for the customers with a key of vnodePrimaryCustomerKey.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerDetailsExit

    Const strFunctionName As String = "AddCustomerDetails"

    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim strCustomerNumber As String
    Dim nodePrimaryCustomer As IXMLDOMNode
    Dim strTemp As String
    Dim nodeTemp, nodeCustomerKey As IXMLDOMNode

    Set nodeCustomerKey = xmlGetMandatoryNode(vnodeCustomerObject, "TARGETKEY/PRIMARYCUSTOMERKEY")
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey
    '------------------------------------------------------------------------------------------
    
    Set nodeConverterResponse = PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey( _
        vobjODITransformerState, nodeCustomerKey)
    
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    Set nodePrimaryCustomer = xmlGetMandatoryNode( _
        nodeConverterResponse, _
        "//RESPONSE/ARGUMENTS/OBJECT/PRIMARYCUSTOMERIMPL", _
        oeRecordNotFound)
        
    ' customer response node
    Set nodeOmigaCustomer = vnodeTarget.ownerDocument.createElement("CUSTOMER")
    
    xmlSetAttributeValue nodeOmigaCustomer, "OTHERSYSTEMCUSTOMERNUMBER", _
        xmlGetNodeText(nodePrimaryCustomer, ".//PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA")
    
    OptimusNameToOmiga _
        nodeOmigaCustomer, _
        nodePrimaryCustomer.selectSingleNode( _
            ".//CONTACTDETAIL/CONTACTDETAIL/NAMEDETAIL/NAMEDETAIL"), _
        vblnIncludeSalutation:=False, _
        vblnIncludeTitle:=False
        
    '------------------------------------------------------------------------------------------
    ' handle home telephone number
    '------------------------------------------------------------------------------------------
    
    Dim nodeTelephoneNumberList As IXMLDOMNode
    Dim cmPreferredMethod As CONTACTMETHOD
    Dim strPreferredMethod As String
    Dim strPreferredTime As String
            
    Set nodeTelephoneNumberList = _
        vnodeTarget.ownerDocument.createElement("TELEPHONENUMBERLIST")
    nodeOmigaCustomer.appendChild nodeTelephoneNumberList
    
    strPreferredMethod = xmlGetNodeText( _
        nodePrimaryCustomer, ".//PREFERREDCOMMUNICATIONMETHOD/@DATA")
    cmPreferredMethod = lookupPreferredContactMethod(strPreferredMethod)
    
    strPreferredTime = xmlGetNodeText( _
        nodePrimaryCustomer, ".//PREFERREDCOMMUNICATIONTIME/@DATA")
    
    ' home phone
    
    strTemp = xmlGetNodeText(nodePrimaryCustomer, ".//HOMEPHONENUMBER/@DATA")
    
    If Len(strTemp) > 0 Then
        Set nodeTemp = vnodeTarget.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
        xmlSetAttributeValue nodeTemp, "USAGE", GetFirstComboValueId("TelephoneUsage", "H")
        xmlSetAttributeValue nodeTemp, "COUNTRYCODE", _
            xmlGetNodeText(nodePrimaryCustomer, ".//HOMEPHONECOUNTRYCODE/@DATA")
        
        'SYS4610 - Prefix the area code with a zero. Optimus strips it.
        xmlSetAttributeValue nodeTemp, "AREACODE", "0" & xmlGetNodeText(nodePrimaryCustomer, ".//HOMEPHONEAREACODE/@DATA")
        
        xmlSetAttributeValue nodeTemp, "TELEPHONENUMBER", strTemp
        xmlSetAttributeValue nodeTemp, "EXTENSIONNUMBER", _
            xmlGetNodeText(nodePrimaryCustomer, ".//HOMEPHONEEXTNUMBER/@DATA")
        
        If cmPreferredMethod = cmHome Then
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", strPreferredTime
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "1"
        Else
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", ""
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "0"
        End If
        nodeTelephoneNumberList.appendChild nodeTemp
        
    End If
   
    vnodeTarget.appendChild nodeOmigaCustomer
    
AddCustomerDetailsExit:

    Set nodeOmigaCustomer = Nothing
    Set nodeConverterResponse = Nothing
    Set nodeTemp = Nothing
    Set nodeTelephoneNumberList = Nothing
    Set nodePrimaryCustomer = Nothing

End Sub


