Attribute VB_Name = "ODIFindAccountDetails"
'Workfile:      ODIFindAccountDetails.cls
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
'RF     27/08/01    Expanded class created by LD.
'DS     23/01/02    Untold number of fixes to make it work. SYS4306.
'DS     28/03/02    Fix for CUSTOMERROLETYPE and CUSTOMERORDER
'DS     05/04/02    LENDERCODE is now populated from the database.
'DS     25/04/02    PURPOSEOFLOAN maps to ChargeImpl.ProductLoadDetails
'DS     30/04/02    Use FreeThreadedDOMDocument40.
'DS     16/05/02    Set DESCRIPTIONOFLOAN to 1
'JLD    21/05/02    SYS4665 changed VALUATIONAMOUNT to equal just the APPRAISEDTOTALVALUE
'DS     26/05/02    SYS4703 If there were no arrears, don't create a histroy.
'STB    29/05/02    SYS4637 Altered element names returned to directly match Omiga field names.
'DS     31/05/02    SYS4792 Mortgage lookup should be on description not code.
'DS     25/06/02    SYS4702 Now use ACTUALAMORTIZATION NOT CONTRACTAMORTIZATION for term of loan
'DS     27/06/02    SYS4974 Change the property type and description back to how it was (!)
'DS     01/07/02    SYS4984 Various changes detailed in AQR.
'CL     02/09/02    SYS5394 Change to AddMortgageDetails.
'SG     17/12/02    SYS5650 Retrieve Joint Borrower Customer Number
'SG     04/02/03    SYS6034 Arrears calculation error.
'SG     05/02/03    SYS6035 Amendments to Arrears block
'AS     02/07/03    MSMS178 OMIP110 Fixes to arrears
'DRC    16/10/03    MSMS206 OMIP146 Add in field FIRSTPAYMENTDATE to Add mortgage details
'PSC    17/01/2007  EP2_855 Changes to bring into line with current dtd
'PSC    19/01/2007  EP2_928 Additional data mapping
'PSC    26/01/2007  EP2_1033 Additional mapping
'PSC    30/01/2007  EP2_1129 Format Interest Rate and Sort Code
'PSC    09/02/2007  EP2_1271 Set up PRODUCTSTARTDATE on mortgage loan
'PSC    15/02/2007  EP2_1415 Change where PRODUCTSTARTDATE is populated from
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub FindAccountDetails( _
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
'   oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindAccountDetailsError
    
    Const strFunctionName As String = "FindAccountDetails"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodeCustomer As IXMLDOMNode
    Dim nodelistShortcuts As IXMLDOMNodeList
    Dim strCustomerNumber As String
    Dim nodeOmigaMortgageAccountList As IXMLDOMNode
    Dim nodeMortgageKey As IXMLDOMNode
    Dim nodeMortgageAccount As IXMLDOMNode
    Dim nodeShortcut As IXMLDOMNode
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern
    ' to get all the account numbers we are interested in
    '------------------------------------------------------------------------------------------
        
    Set nodeCustomer = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    strCustomerNumber = xmlGetMandatoryAttributeText(nodeCustomer, "CUSTOMERNUMBER")
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern( _
            vobjODITransformerState, strCustomerNumber)
    
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' from the response get objectshortcuts
    '------------------------------------------------------------------------------------------
    
    Set nodelistShortcuts = _
        nodeConverterResponse.selectNodes( _
            "ARGUMENTS/OBJECT/OBJECTSHORTCUT")
            
    If nodelistShortcuts Is Nothing Then
        errThrowError strFunctionName, oerecordnotfound
    End If
        
    Set nodeOmigaMortgageAccountList = _
        vxmlResponseNode.ownerDocument.createElement("MORTGAGEACCOUNTLIST")
    vxmlResponseNode.appendChild nodeOmigaMortgageAccountList
    
    '------------------------------------------------------------------------------------------
    ' Handle each Account number returned
    '------------------------------------------------------------------------------------------
    
    For Each nodeShortcut In nodelistShortcuts
    
        Set nodeMortgageKey = nodeShortcut.selectSingleNode("TARGETKEY/MORTGAGEKEY")
        Set nodeMortgageAccount = _
            vxmlResponseNode.ownerDocument.createElement("MORTGAGEACCOUNT")
        nodeOmigaMortgageAccountList.appendChild nodeMortgageAccount
        
        AddMortgageDetails vobjODITransformerState, _
            nodeMortgageKey, nodeMortgageAccount
        
'DS - remove this when done
'        AddCustomerDetails vobjODITransformerState, _
'           nodeMortgageKey, nodeMortgageAccount
    
    Next nodeShortcut
      
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode

FindAccountDetailsExit:

    Set nodeCustomer = Nothing
    Set nodeOmigaMortgageAccountList = Nothing
    Set nodeConverterResponse = Nothing
    Set nodelistShortcuts = Nothing
    Set nodeMortgageKey = Nothing
    Set nodeMortgageAccount = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    Exit Sub
FindAccountDetailsError:
    errCheckError strFunctionName

End Sub

Private Sub AddMortgageDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeMortgageKey As IXMLDOMNode, _
    ByVal vnodeTarget As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Add to vnodeTarget the details for the mortgage having the a key of
'   vnodeMortgageKey.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddMortgageDetailsExit

    Const strFunctionName As String = "AddMortgageDetails"
    
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodeMortgageAccount As IXMLDOMNode
    Dim nodeMortgageImpl, nodeChargeImpl, nodePIComponentImpl As IXMLDOMNode
    Dim elemProperty As IXMLDOMElement
    Dim elemPropertyAddress As IXMLDOMElement
    Dim elemMortgageLoanList As IXMLDOMElement
    Dim elemMortgageLoan As IXMLDOMElement
    Dim elemArrearsHistory As IXMLDOMElement
    Dim nodeTemp As IXMLDOMNode
    Dim strTemp As String
    Dim strTempValue1 As String
    Dim strTempValue2 As String
    Dim strTermInMonths As String
    Dim strLineOfBusiness As String
    Dim nodelistComponents As IXMLDOMNodeList
    Dim nodeComponent As IXMLDOMNode
    Dim dblArrearsHistoryMaximumBalance As Double
    Dim nodelistPeripheralSecurities As IXMLDOMNodeList
    Dim nodelistBankingDetails As IXMLDOMNodeList
    Dim nodeOptimusRealEstate As IXMLDOMNode
    Dim strBankerCIFNumber, strBankerCIFSuffix, strLenderCode As String
    'DRC 17/04/2003 - check for the number of optimus customers found
    Dim vIntCustomerCount As Integer
    
    'SG 05/02/03 SYS6035
    Dim dblFixedPaymentAmount As Double
    Dim strTempValue3 As String
    
    ' PSC 19/01/2007 EP2_928 - Start
    Dim dblMonthlyRentalIncome As Double
    Dim dblAvailableDisbursement As Double
    Dim strAppraiserCIF As String
    
    Dim nodeAppraiserKey As IXMLDOMNode
    Dim nodeAppraiserResponse As IXMLDOMNode
    Dim nodeAppraiserDetails As IXMLDOMNode
    ' PSC 19/01/2007 EP2_928 - End
    
    '------------------------------------------------------------------------------------------
    ' Call PlexusHome.findByPrimaryKey supplying a MortgageKey
    '------------------------------------------------------------------------------------------
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_findByPrimaryKey_MortgageKey( _
            vobjODITransformerState, vnodeMortgageKey)
            
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response -
    '
    ' MORTGAGEACCOUNT element
    '------------------------------------------------------------------------------------------
        
    ' get useful pointer into nodeConverterResponse
    
    Set nodeMortgageImpl = xmlGetMandatoryNode(nodeConverterResponse, "ARGUMENTS/OBJECT/MORTGAGEIMPL")
    Set nodeOptimusRealEstate = xmlGetMandatoryNode(nodeMortgageImpl, "REALESTATE/REALESTATEIMPL")
    Set nodeMortgageImpl = xmlGetMandatoryNode(nodeConverterResponse, "ARGUMENTS/OBJECT/MORTGAGEIMPL")
    Set nodeChargeImpl = xmlGetMandatoryNode(nodeMortgageImpl, "CHARGE/CHARGEIMPL")
    Set nodePIComponentImpl = xmlGetMandatoryNode(nodeChargeImpl, "COMPONENTS/COMPONENTIMPL/PICOMPONENTIMPL")
    
    strBankerCIFNumber = xmlGetNodeText(nodeChargeImpl, ".//BANKERCIFNUMBER/@DATA")
    strBankerCIFSuffix = xmlGetNodeText(nodeChargeImpl, ".//BANKERCIFSUFFIX/@DATA")
    strLineOfBusiness = xmlGetNodeText(nodeChargeImpl, ".//PRODUCTLINEOFBUSINESS/@DATA")
            
    Dim strAccountNumber As String
    
    ' MortgageKey.RealEstateKey.collateralNumber + "." + MortgageKey.chargeType
    strAccountNumber = xmlGetMandatoryNodeText(nodeMortgageImpl, "PRIMARYKEY/MORTGAGEKEY/REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER/@DATA") & _
                        "." & _
                        xmlGetMandatoryNodeText(nodeMortgageImpl, "PRIMARYKEY/MORTGAGEKEY/CHARGETYPE/@DATA")
    
    xmlSetAttributeValue vnodeTarget, "ACCOUNTNUMBER", strAccountNumber
    
    xmlSetAttributeValue vnodeTarget, "SECONDCHARGEINDICATOR", ""
    
    'strTemp = xmlGetNodeText(nodeChargeImpl, "MORTGAGEINSURANCETYPE/@DATA")
    'strTemp = lookupIndemnityCompanyName(strTemp, cdOptimusToOmiga)
    xmlSetAttributeValue vnodeTarget, "INDEMNITYCOMPANYNAME", ""
    
    ' PSC 17/01/2007 EP2_855 - Start
    xmlSetAttributeValue vnodeTarget, "INDEMNITYMORTGAGEAMOUNT", "0"
    xmlSetAttributeValue vnodeTarget, "INDEMNITYAMOUNT", "0"
    ' PSC 17/01/2007 EP2_855 - End
    
    strTemp = xmlGetNodeText(nodeChargeImpl, "FIRSTADVANCEDATE/@DATA")
    xmlSetAttributeValue vnodeTarget, "CREATIONDATE", OptimusDateToOmiga(strTemp)
    
    ' PSC 17/01/2007 EP2_855 - Start
    xmlSetAttributeValue vnodeTarget, "REPOSSESSIONFLAG", ""
    
    ' PSC 19/01/2007 EP2_928 - Start
    strTemp = xmlGetNodeText(nodeOptimusRealEstate, "ESTIMATEDINCOME/@DATA")
    If Len(strTemp) > 0 Then
        dblMonthlyRentalIncome = CSafeDbl(strTemp) / 12#
        xmlSetAttributeValue vnodeTarget, "MONTHLYRENTALINCOME", CStr(dblMonthlyRentalIncome)
    Else
        xmlSetAttributeValue vnodeTarget, "MONTHLYRENTALINCOME", ""
    End If
    ' PSC 19/01/2007 EP2_928 - End
    
    xmlSetAttributeValue vnodeTarget, "COLLATERALID", ""
    xmlSetAttributeValue vnodeTarget, "INTERESTFREQUENCY", ""
    xmlSetAttributeValue vnodeTarget, "DSSFLAG", ""
    xmlSetAttributeValue vnodeTarget, "BUSINESSCHANNEL", ""
    ' PSC 17/01/2007 EP2_855 - End

    ' PSC 19/01/2007 EP2_928 - Start
    AddBankingDetailsDetails vobjODITransformerState, nodeChargeImpl, vnodeTarget
    
    xmlSetAttributeValue vnodeTarget, "REGULATEDMORTGAGEINDICATOR", ""
    
    strTemp = xmlGetNodeText(nodeChargeImpl, ".//NEXTPAYMENTDATE/@DATA")
    xmlSetAttributeValue vnodeTarget, "PAYMENTDUEDATE", Left$(OptimusDateToOmiga(strTemp), 2)
    
    strTemp = xmlGetNodeText(nodeChargeImpl, "STATUS/@DATA")
    xmlSetAttributeValue vnodeTarget, "REDEMPTIONSTATUS", lookupRedemptionStatus(strTemp, cdOptimusToOmiga)
    
    strTemp = xmlGetNodeText(nodeChargeImpl, ".//USERDEFINEDDATE4/@DATA")
    xmlSetAttributeValue vnodeTarget, "PREEMPTIONENDDATE", OptimusDateToOmiga(strTemp)
    
    xmlSetAttributeValue vnodeTarget, "OTHERPROPERTYMORTGAGE", ""
    xmlSetAttributeValue vnodeTarget, "ORIGINALNATUREOFLOAN", strLineOfBusiness
    ' PSC 19/01/2007 EP2_928 - End
   
      
    '------------------------------------------------------------------------------------------
    ' CUSTOMERLIST element
    '------------------------------------------------------------------------------------------
    
    AddCustomerDetails vobjODITransformerState, vnodeMortgageKey, vnodeTarget
    
    vIntCustomerCount = vnodeTarget.selectNodes(".//CUSTOMER").length
    
    '  DRC 17/04/2003 MSMS38 - only get joint rec number if the returned number of customers is > 1
    If vIntCustomerCount > 1 Then
      'SG 17/12/02 SYS5650
        xmlSetAttributeValue vnodeTarget, "OTHERSYSTEMJOINTRECNUMBER", xmlGetNodeText(nodeChargeImpl, ".//BORROWERCIFNUMBER/@DATA")
    Else
        xmlSetAttributeValue vnodeTarget, "OTHERSYSTEMJOINTRECNUMBER", ""
    End If
    
    xmlSetAttributeValue vnodeTarget, "LENDERCODE", GetGlobalParamString("AccountDownloadLenderCode")
    
    '------------------------------------------------------------------------------------------
    ' PROPERTY element
    '------------------------------------------------------------------------------------------
    
    Set elemProperty = vnodeTarget.ownerDocument.createElement("PROPERTY")
    vnodeTarget.appendChild elemProperty

    elemProperty.setAttribute "CREATIONDATE", ""

    elemProperty.setAttribute "PURCHASEPRICE", _
        xmlGetNodeText(nodeOptimusRealEstate, "PURCHASEPRICE/@DATA")

    strTemp = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISALDATE/@DATA")
    elemProperty.setAttribute "VALUATIONDATE", OptimusDateToOmiga(strTemp)

    strTempValue1 = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISEDTOTALVALUE/@DATA")
    'strTempValue2 = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISEDLANDVALUE/@DATA")
    'If strTempValue1 = "" And strTempValue2 = "" Then
    '    strTemp = ""
    'Else
    '   strTemp = CStr(CSafeLng(strTempValue1) + CSafeLng(strTempValue2))
    'End If
    elemProperty.setAttribute "VALUATIONAMOUNT", strTempValue1    'JLD SYS4665

    ' PSC 19/01/2007 EP2_928 - Start
    strAppraiserCIF = xmlGetNodeText(nodeOptimusRealEstate, "APPRAISERCIFNUMBER/@DATA")
    strTemp = strAppraiserCIF & xmlGetNodeText(nodeOptimusRealEstate, "APPRAISERCIFSUFFIX/@DATA")
    ' PSC 19/01/2007 EP2_928 - End
    elemProperty.setAttribute "VALUERID", strTemp
    
    'SYS4637, SYS4974 - DESCRIPTIONOFPROPERTY.
    ' PSC 26/01/2007 EP2_1033 - Start
    strTemp = xmlGetNodeText(nodeOptimusRealEstate, "COLLATERALIMPROVEMENTTYPE/@DATA")
    
    If Len(strTemp) = 0 Or strTemp = "99" Then
        strTempValue1 = xmlGetNodeText(nodeOptimusRealEstate, "COLLATERALUSAGE/@DATA")
        
        If Len(strTempValue1) > 0 Then
            strTemp = strTempValue1
        End If
    End If
    
    elemProperty.setAttribute "DESCRIPTIONOFPROPERTY", strTemp
    ' PSC 26/01/2007 EP2_1033 - End

    elemProperty.setAttribute "YEARBUILT", ""
    
    'SYS4637 - TENURETYPE.
    strTemp = xmlGetNodeText(nodeOptimusRealEstate, "FREEHOLDTITLE/@DATA")
    elemProperty.setAttribute "TENURETYPE", strTemp
    
    'SYS4637, SYS4974 - TYPEOFPROPERTY.
    ' PSC 26/01/2007 EP2_1033
    strTemp = xmlGetNodeText(nodeOptimusRealEstate, "COLLATERALOCCUPANCYSTATUS/@DATA")
    elemProperty.setAttribute "TYPEOFPROPERTY", strTemp
    
    Dim xmlPerSecParentNode As IXMLDOMNode
    Set xmlPerSecParentNode = nodeConverterResponse.selectSingleNode(".//PERIPHERALSECURITIES/PERIPHERALSECURITYIMPL")
    
    Set nodelistPeripheralSecurities = xmlPerSecParentNode.selectNodes( _
                "PERIPHERALSECURITYIMPL")
                
    Dim dblBldgsSumInsured As Double
    Dim blnHomeInsTypeFound As Boolean
    
    dblBldgsSumInsured = 0
    blnHomeInsTypeFound = False
    
    For Each nodeTemp In nodelistPeripheralSecurities
    
        strTemp = xmlGetMandatoryNodeText(nodeTemp, _
            "PRIMARYKEY/PERIPHERALSECURITYKEY/EXPIRYSUFFIX/@DATA")
            
        If strTemp = "10" Or strTemp = "11" Or strTemp = "12" Then
            dblBldgsSumInsured = dblBldgsSumInsured + CSafeDbl( _
                xmlGetNodeText(nodeTemp, "PERIPHERALVALUE/@DATA"))
        End If
        
        If blnHomeInsTypeFound = False Then
            If strTemp = "10" Or strTemp = "11" Or strTemp = "12" Or _
                strTemp = "13" Or strTemp = "15" Then
                
                elemProperty.setAttribute "HOMEINSURANCETYPE", _
                    lookupHomeInsuranceType(strTemp, cdOptimusToOmiga)
                
                blnHomeInsTypeFound = True
            End If
        End If
    Next nodeTemp
    
    If blnHomeInsTypeFound = False Then
        elemProperty.setAttribute "HOMEINSURANCETYPE", ""
    End If
    
    elemProperty.setAttribute "BUILDINGSSUMINSURED", CStr(dblBldgsSumInsured)

    elemProperty.setAttribute "REINSTATEMENTAMOUNT", ""
    
    ' PSC 19/01/2007 EP2_928 - Start
    If Len(strAppraiserCIF) > 0 Then
        
        Set nodeAppraiserKey = CreatePrimaryCustomerKey(strAppraiserCIF)
        
        Set nodeAppraiserResponse = PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey( _
            vobjODITransformerState, nodeAppraiserKey)
        
        CheckConverterResponse nodeAppraiserResponse, True
        
        Set nodeAppraiserDetails = xmlGetNode(nodeAppraiserResponse, _
            "ARGUMENTS/OBJECT/PRIMARYCUSTOMERIMPL")
            
    End If
    
    If Not nodeAppraiserDetails Is Nothing Then
        elemProperty.setAttribute "LASTVALUERNAME", xmlGetNodeText(nodeAppraiserDetails, "CONTACT/CONTACT/NAME/NAME/LINE1/@DATA")
    Else
        elemProperty.setAttribute "LASTVALUERNAME", ""  ' PSC 17/01/2007 EP2_855
    End If
    ' PSC 19/01/2007 EP2_928 - End

    '------------------------------------------------------------------------------------------
    ' PROPERTYADDRESS element
    '------------------------------------------------------------------------------------------

    Set elemPropertyAddress = vnodeTarget.ownerDocument.createElement("PROPERTYADDRESS")
    elemProperty.appendChild elemPropertyAddress
    
    
    'DS - THE SECURITY ADDRESS ISSUE NEEDS SORTING OUT!
    'See AQR SYS4291
    Set nodeTemp = xmlGetNode(nodeOptimusRealEstate, "ADDRESSDETAIL/ADDRESSDETAIL")
    OptimusAddressToOmiga elemPropertyAddress, nodeTemp, True, True
    
    '------------------------------------------------------------------------------------------
    ' MORTGAGELOANLIST element
    '------------------------------------------------------------------------------------------
    
    Set elemMortgageLoanList = vnodeTarget.ownerDocument.createElement("MORTGAGELOANLIST")
    vnodeTarget.appendChild elemMortgageLoanList

    dblArrearsHistoryMaximumBalance = 0
    
    Dim nodeComponentImpl As IXMLDOMNode
    Set nodeComponentImpl = nodeChargeImpl.selectSingleNode("COMPONENTS/COMPONENTIMPL")
    
    Set nodelistComponents = nodeComponentImpl.childNodes
    
    ' PSC 19/01/2007 EP2_928 - Start
    Dim blnFirstComponent As Boolean
    blnFirstComponent = True
    Dim dblTotalCollateralBalance As Double
    Dim dblOutstandingBalance As Double
    Dim dblTotalMonthlyCost As Double
    ' PSC 19/01/2007 EP2_928 - End
    
    ' PSC 30/01/2007 EP2_1129 - Start
    Dim dblInterestRate As Double
    Dim strInterestRateText As Double
    ' PSC 30/01/2007 EP2_1129 - End
    
    For Each nodeComponent In nodelistComponents
    
        If xmlGetMandatoryNodeText(nodeComponent, ".//LEDGER/@DATA") = "A" Then
        
            ' PSC 19/01/2007 EP2_928 - Start
            If blnFirstComponent Then
                xmlSetAttributeValue vnodeTarget, "ORIGINALCREDITSCHEME", xmlGetNodeText(nodeComponent, ".//USERDEFINEDCODE2/@DATA")
                blnFirstComponent = False
            End If
            ' PSC 19/01/2007 EP2_928 - End
            
            Set elemMortgageLoan = vnodeTarget.ownerDocument.createElement("MORTGAGELOAN")
            elemMortgageLoanList.appendChild elemMortgageLoan
            
            elemMortgageLoan.setAttribute "ACCOUNTNUMBER", strAccountNumber
            
            strTemp = xmlGetNodeText(nodeComponent, ".//SUBLEDGER/@DATA")
            elemMortgageLoan.setAttribute "LOANACCOUNTNUMBER", _
                strAccountNumber & "." & strTemp
            
            'DS SYS4792
            ' PSC 19/01/2007 EP2_928
            strTemp = xmlGetNodeText(nodeComponent, ".//USERDEFINEDAMOUNT1/@DATA")
            elemMortgageLoan.setAttribute "MORTGAGEPRODUCTCODE", CStr(CSafeLng(strTemp))
                
            
            'convert the code to a description
            
            elemMortgageLoan.setAttribute "MORTGAGEPRODUCTDESCRIPTION", lookupMortgageProductCode(strTemp, cdOptimusToOmiga)
            
            ' PSC 30/01/2007 EP2_1129 - Start
            strInterestRateText = xmlGetNodeText(nodeComponent, ".//CURRENTRATE/@DATA")
            
            If Len(strInterestRateText) > 0 Then
                dblInterestRate = CSafeDbl(strInterestRateText) * 100#
                strInterestRateText = Format$(CStr(dblInterestRate), "0.00")
            End If
            
            elemMortgageLoan.setAttribute "INTERESTRATE", strInterestRateText
            ' PSC 30/01/2007 EP2_1129 - End
  
                
            ' Call lookupInterestRateType to convert PIComponent.maximumRate and
            ' PIComponent.rateChangeFrequency to an Omiga combo id
            ' PSC 19/01/2007 EP2_928 - Start
            Dim strMaximumRate As String
            Dim strMinimumRate As String
            Dim strRateFrequency As String
            strMaximumRate = xmlGetNodeText(nodeComponent, ".//MAXIMUMRATE/@DATA")
            strMinimumRate = xmlGetNodeText(nodeComponent, ".//MINIMUMRATE/@DATA")
            strRateFrequency = xmlGetNodeText(nodeComponent, ".//RATECHANGEFREQUENCY/@DATA")
            elemMortgageLoan.setAttribute "INTERESTRATETYPE", _
                lookupInterestRateType(strRateFrequency, strMinimumRate, strMaximumRate, cdOptimusToOmiga)
            
            elemMortgageLoan.setAttribute "PURPOSEOFLOAN", ""
    
            dblOutstandingBalance = _
                CSafeDbl(xmlGetNodeText(nodeComponent, ".//PRINCIPALBALANCE/@DATA")) + _
                CSafeDbl(xmlGetNodeText(nodeComponent, ".//INTERESTINARREARS/@DATA")) + _
                CSafeDbl(xmlGetNodeText(nodeComponent, ".//PAIDINADVANCE/@DATA"))
            elemMortgageLoan.setAttribute "OUTSTANDINGBALANCE", CStr(dblOutstandingBalance)
            
            dblTotalCollateralBalance = dblTotalCollateralBalance + dblOutstandingBalance
            ' PSC 19/01/2007 EP2_928 - End
            
            strTemp = xmlGetNodeText(nodeComponent, ".//PAYMENTTYPE/@DATA")
            elemMortgageLoan.setAttribute "REPAYMENTTYPE", lookupRepaymentType(strTemp, cdOptimusToOmiga)
            
            elemMortgageLoan.setAttribute "ORIGINALLOANAMOUNT", xmlGetNodeText(nodeComponent, ".//MAXIMUMAPPROVEDLOANAMOUNT/@DATA")
        
            ' SYS4841
            ' PSC 19/01/2007 EP2_928 - Start
            strTemp = xmlGetNodeText(nodeComponent, ".//FIXEDPAYMENTAMOUNT/@DATA")
            elemMortgageLoan.setAttribute "MONTHLYREPAYMENT", strTemp
            
            dblTotalMonthlyCost = dblTotalMonthlyCost + CSafeDbl(strTemp)

            strTermInMonths = xmlGetNodeText(nodeComponent, ".//TERMINMONTHS/@DATA")
            ' PSC 19/01/2007 EP2_928 - End
            
            If strTermInMonths = "" Then
                elemMortgageLoan.setAttribute "ORIGINALTERMYEARS", ""
                elemMortgageLoan.setAttribute "ORIGINALTERMMONTHS", ""
            Else
                elemMortgageLoan.setAttribute "ORIGINALTERMYEARS", CStr(Int((CSafeLng(strTermInMonths) / 12)))
                elemMortgageLoan.setAttribute "ORIGINALTERMMONTHS", CInt(strTermInMonths) Mod 12
            End If
                        
            strTemp = xmlGetNodeText(nodeComponent, ".//FIRSTADVANCEDATE/@DATA")
            elemMortgageLoan.setAttribute "STARTDATE", OptimusDateToOmiga(strTemp)
            
            'MSMS206 - Add this in so that remaining term can be calculated in Omiga4
            
            strTemp = xmlGetNodeText(nodeComponent, ".//FIRSTPAYMENTDATE/@DATA")
            elemMortgageLoan.setAttribute "FIRSTPAYMENTDATE", OptimusDateToOmiga(strTemp)
            
            strTemp = xmlGetNodeText(nodeComponent, "STATUS/@DATA")
            elemMortgageLoan.setAttribute "REDEMPTIONSTATUS", lookupRedemptionStatus(strTemp, cdOptimusToOmiga)
            
            strTemp = xmlGetNodeText(nodeComponent, ".//PAYOUTDATE/@DATA")
            elemMortgageLoan.setAttribute "REDEMPTIONDATE", OptimusDateToOmiga(strTemp)
            
            'CL SYS5394 02/09/02
            ' PSC 19/01/2007 EP2_928 - Start
            strTemp = xmlGetNodeText(nodeComponent, ".//USERDEFINEDDATE3/@DATA")
            elemMortgageLoan.setAttribute "LOANPRODUCTSTARTDATE", ""
            ' PSC 19/01/2007 EP2_928 - End
            'END CL SYS5394 02/09/02
                      
            'AS MSMS178 02/07/03
            strTemp = xmlGetNodeText(nodeChargeImpl, ".//ORIGINALFIRSTPAYMENTDATE/@DATA")
            Dim dateOriginalFirstPaymentDate As Date
            dateOriginalFirstPaymentDate = CSafeDate(OptimusDateToOmiga(strTemp))
            If dateOriginalFirstPaymentDate <= Date Then
                ' dblArrearsHistoryMaximumBalance =
                ' sum of principalInArrears + interestInArrears
                ' for each component where ComponentKey.Ledger="A"
                strTempValue1 = xmlGetNodeText(nodeComponent, ".//PRINCIPALINARREARS/@DATA")
                strTempValue2 = xmlGetNodeText(nodeComponent, ".//INTERESTINARREARS/@DATA")
            
                'SG 04/02/03 SYS6034 - Arrears should be total of all components
                'dblArrearsHistoryMaximumBalance = CSafeDbl(strTempValue1) + CSafeDbl(strTempValue2)
                dblArrearsHistoryMaximumBalance = _
                    dblArrearsHistoryMaximumBalance + CSafeDbl(strTempValue1) + CSafeDbl(strTempValue2)
            End If
            'END AS MSMS178 02/07/03
            
            'SG 05/02/03 SYS6035 sum values
            strTempValue3 = xmlGetNodeText(nodeComponent, ".//FIXEDPAYMENTAMOUNT/@DATA")
            dblFixedPaymentAmount = dblFixedPaymentAmount + CSafeDbl(strTempValue3)
            
            ' PSC 17/01/2007 EP2_855 - Start
            elemMortgageLoan.setAttribute "DISBURSEDAMOUNT", ""
            elemMortgageLoan.setAttribute "CCAINDICATOR", ""
            elemMortgageLoan.setAttribute "CCIINDICATOR", ""
            ' PSC 19/01/2007 EP2_928
            elemMortgageLoan.setAttribute "CALCULATEDOUTSTANDINGTERM", xmlGetNodeText(nodeComponent, ".//REMAININGTERM/@DATA")
            elemMortgageLoan.setAttribute "PENALTYPLANCODE", ""
            elemMortgageLoan.setAttribute "PENALTYPLANDESCRIPTION", ""
            elemMortgageLoan.setAttribute "PENALTYPLANENDDATE", ""
            elemMortgageLoan.setAttribute "LOANCLASSTYPE", ""
            elemMortgageLoan.setAttribute "OVERPAYMENTS", ""
            elemMortgageLoan.setAttribute "LOANTYPE", ""
            elemMortgageLoan.setAttribute "EXISTINGRETENTIONS", ""
            
            ' PSC 19/01/2007 EP2_928 - Start
            dblAvailableDisbursement = _
                    xmlGetNodeAsDouble(nodeChargeImpl, ".//PIMAXIMUMAMOUNT/@DATA") - _
                    xmlGetNodeAsDouble(nodeChargeImpl, ".//PIPRINCIPAL/@DATA") - _
                    xmlGetNodeAsDouble(nodeChargeImpl, ".//BALANCEINHOLDBACKS/@DATA")
            
            elemMortgageLoan.setAttribute "AVAILABLEFORDISBURSEMENT", CStr(dblAvailableDisbursement)
            ' PSC 19/01/2007 EP2_928 - End
            
            elemMortgageLoan.setAttribute "REDEMPTIONPENALTY", ""
            elemMortgageLoan.setAttribute "INDEXCODE", ""
            elemMortgageLoan.setAttribute "VARIANCE", ""
            elemMortgageLoan.setAttribute "PRODUCTSTEP", ""
            elemMortgageLoan.setAttribute "REMAININGSTEPDURATION", ""
            elemMortgageLoan.setAttribute "REMAININGINTERESTONLYAMOUNT", ""
            elemMortgageLoan.setAttribute "REMAININGCAPITALINTERESTAMOUNT", ""
            elemMortgageLoan.setAttribute "ORIGINALPARTANDPARTINTONLYAMT", ""
            ' PSC 17/01/2007 EP2_855 - End

            ' PSC 19/01/2007 EP2_928 - Start
            elemMortgageLoan.setAttribute "ORIGINALLTV", xmlGetNodeText(nodeComponent, ".//USERDEFINEDAMOUNT2/@DATA")
            elemMortgageLoan.setAttribute "ORIGINALINCOMESTATUS", xmlGetNodeText(nodeChargeImpl, ".//PRODUCTLINE/@DATA")
            elemMortgageLoan.setAttribute "FLEXIBLEPRODUCTINDICATOR", ""
            ' PSC 19/01/2007 EP2_928 - End
            
            ' PSC 09/02/2007 EP2_1271 - Start
            strTemp = xmlGetNodeText(nodeComponent, ".//USERDEFINEDDATE2/@DATA")
            
            If Len(strTemp) = 0 Then
                ' PSC 15/02/2007 EP2_1415
                strTemp = xmlGetNodeText(nodeComponent, ".//FIRSTADVANCEDATE/@DATA")
            End If
            
            elemMortgageLoan.setAttribute "PRODUCTSTARTDATE", OptimusDateToOmiga(strTemp)
            ' PSC 09/02/2007 EP2_1271 - End

        End If
    
    Next nodeComponent
    
    ' PSC 19/01/2007 EP2_928 - Start
    xmlSetAttributeValue vnodeTarget, "TOTALMONTHLYCOST", CStr(dblTotalMonthlyCost)
    xmlSetAttributeValue vnodeTarget, "TOTALCOLLATERALBALANCE", CStr(dblTotalCollateralBalance)
    ' PSC 19/01/2007 EP2_928 - End
    
    '------------------------------------------------------------------------------------------
    ' ARREARSHISTORYLIST element
    '------------------------------------------------------------------------------------------
    
    'DS SYS4703
    If dblArrearsHistoryMaximumBalance > 0 Then
        Set elemArrearsHistory = vnodeTarget.ownerDocument.createElement("ARREARSHISTORY")
        vnodeTarget.appendChild elemArrearsHistory
    
        elemArrearsHistory.setAttribute "MAXIMUMBALANCE", CStr(dblArrearsHistoryMaximumBalance)
        elemArrearsHistory.setAttribute "MAXIMUMNUMBEROFMONTHS", ""
    
        'SG 05/02/03 SYS6035 corrected XSL pattern
        strTemp = xmlGetNodeText(nodeConverterResponse, ".//CHARGEIMPL/PAIDTODATE/@DATA")
        elemArrearsHistory.setAttribute "DATECLEARED", OptimusDateToOmiga(strTemp)
        
        elemArrearsHistory.setAttribute "DESCRIPTIONOFLOAN", "1"
        
        'SG 05/02/03 SYS6035 corrected XSL pattern
        elemArrearsHistory.setAttribute "INSTALLMENTSINARREARS", _
            xmlGetNodeText(nodeConverterResponse, ".//CHARGEIMPL/ARREARSPAYMENTCOUNTER/@DATA")
         
        ' PSC 19/01/2007 EP2_928
        elemArrearsHistory.setAttribute "CURRENTYEARSINARREARS", ""
   
        'SG 05/02/03 SYS6035 added attribute
        elemArrearsHistory.setAttribute "MONTHLYREPAYMENT", CStr(dblFixedPaymentAmount)
    
    End If
AddMortgageDetailsExit:

    Set nodeConverterResponse = Nothing
    Set nodeMortgageAccount = Nothing
    Set elemProperty = Nothing
    Set elemPropertyAddress = Nothing
    Set elemMortgageLoanList = Nothing
    Set elemMortgageLoan = Nothing
    Set elemArrearsHistory = Nothing
    Set nodeTemp = Nothing
    
    ' PSC 19/01/2007 EP2_928 - Start
    Set nodeAppraiserKey = Nothing
    Set nodeAppraiserResponse = Nothing
    Set nodeAppraiserDetails = Nothing
    ' PSC 19/01/2007 EP2_928 - End

    errCheckError strFunctionName

End Sub
        
Private Sub AddCustomerDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeMortgageKey As IXMLDOMNode, _
    ByVal vnodeTarget As IXMLDOMNode _
    )
' header ----------------------------------------------------------------------------------
' description:
'   Add to vnodeTarget the details for the customers attached to the mortgage with a
'   key of vnodeMortgageKey.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerDetailsExit

    Const strFunctionName As String = "AddCustomerDetails"

    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodelistOptimusCustomers As IXMLDOMNodeList
    Dim strSearchKey As String
    Dim nodePrimaryCustomerKey As IXMLDOMNode
    Dim nodeOmigaCustomerList As IXMLDOMNode
    Dim nodeOmigaCustomer, nodeOptimusCustomer As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim index As Integer
    
    
    '------------------------------------------------------------------------------------------
    'Call PlexusHomeImpl.searchByPattern, supplying a CustomerListPattern.
    '------------------------------------------------------------------------------------------

    strSearchKey = xmlGetNodeText(vnodeMortgageKey, "REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER/@DATA") _
                        & "." & xmlGetNodeText(vnodeMortgageKey, "CHARGETYPE/@DATA")

    Set xmlDoc = New FreeThreadedDOMDocument40
    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_CustomerListPattern( _
            vobjODITransformerState, strSearchKey)
            
    Set nodelistOptimusCustomers = nodeConverterResponse.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT")
        
    If nodelistOptimusCustomers Is Nothing Then
        errThrowError strFunctionName, oerecordnotfound
    End If
    
    ' customer list node
    Set nodeOmigaCustomerList = _
        vnodeTarget.ownerDocument.createElement("CUSTOMERLIST")
    vnodeTarget.appendChild nodeOmigaCustomerList

    index = 1
    For Each nodeOptimusCustomer In nodelistOptimusCustomers
    
        Set nodePrimaryCustomerKey = xmlGetNode(nodeOptimusCustomer, ".//PRIMARYCUSTOMERKEY")
        Set nodeOmigaCustomer = xmlDoc.createElement("CUSTOMER")
        xmlSetAttributeValue nodeOmigaCustomer, "CUSTOMERNUMBER", _
            xmlGetMandatoryNodeText(nodePrimaryCustomerKey, ".//ENTITYNUMBER/@DATA")
        xmlSetAttributeValue nodeOmigaCustomer, "CUSTOMERORDER", index
        xmlSetAttributeValue nodeOmigaCustomer, "CUSTOMERROLETYPE", xmlGetNodeText(nodeOptimusCustomer, ".//EXPIRYSUFFIX/@DATA")
        nodeOmigaCustomerList.appendChild nodeOmigaCustomer
        index = index + 1
    
    Next nodeOptimusCustomer
    Set xmlDoc = Nothing
        
AddCustomerDetailsExit:

    Set nodeConverterResponse = Nothing
    Set nodelistOptimusCustomers = Nothing
    Set nodePrimaryCustomerKey = Nothing
    Set nodeOmigaCustomerList = Nothing
    Set nodeOmigaCustomer = Nothing
    Set xmlDoc = Nothing

End Sub

' PSC 19/01/2007 EP2_928 - Start
Private Sub AddBankingDetailsDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeChargeImpl As IXMLDOMNode, _
    ByVal vnodeTarget As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Add to vnodeTarget the details for the customers attached to the mortgage with a
'   key of vnodeMortgageKey.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddBankingDetailsDetailsExit

    Const strFunctionName As String = "AddBankingDetailsDetails"

    Dim nodeCustomerKey As IXMLDOMNode
    Dim nodePaymentSource As IXMLDOMNode
    Dim nodeBankingDetails As IXMLDOMNode
    Dim nodeListPaymentSources As IXMLDOMNodeList
    
    Dim strPayorCIF As String
    Dim srrPayorCIFSuffix As String
 
    Dim nodeConverterResponse As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim index As Integer
    
    Dim strStartDate As String
    Dim strEndDate As String
    Dim dteStartDate As Date
    Dim dteEndDate As Date
    Dim intPayorCIFSuffix As Integer
    Dim intHighestCIFSuffix As Integer
    
    Dim strSortCode As String       ' PSC 30/01/2007 EP2_1129
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    
    index = 0
    intHighestCIFSuffix = 0
    
    Set nodeListPaymentSources = vnodeChargeImpl.selectNodes("PAYMENTSOURCES/PAYMENTSOURCEIMPL/PAYMENTSOURCEIMPL")

    Do While index < nodeListPaymentSources.length
        Set nodePaymentSource = nodeListPaymentSources(index)
        
        strStartDate = xmlGetNodeText(nodePaymentSource, "PAYMENTSTARTDATE/@DATA")
        strStartDate = OptimusDateToOmiga(strStartDate)
        strEndDate = xmlGetNodeText(nodePaymentSource, "EXPIRYDATE/@DATA")
        strEndDate = OptimusDateToOmiga(strEndDate)
        intPayorCIFSuffix = xmlGetNodeAsInteger(nodePaymentSource, "PAYORCIFSUFFIX/@DATA")
        
        If intPayorCIFSuffix > intHighestCIFSuffix Then
        
            dteEndDate = CSafeDate(strEndDate)
            dteStartDate = CSafeDate(strStartDate)
            
            If dteStartDate <= Date And dteEndDate > Date Then
                strPayorCIF = xmlGetNodeText(nodePaymentSource, "PAYORCIFNUMBER/@DATA")
                intHighestCIFSuffix = intPayorCIFSuffix
            End If
        End If
        
        index = index + 1
    Loop
    
    If Len(strPayorCIF) > 0 Then
        Set nodeCustomerKey = CreatePrimaryCustomerKey(strPayorCIF)
        
        Set nodeConverterResponse = PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey( _
            vobjODITransformerState, nodeCustomerKey)
        
        CheckConverterResponse nodeConverterResponse, True
        
        Set nodeBankingDetails = xmlGetNode(nodeConverterResponse, _
            "ARGUMENTS/OBJECT/PRIMARYCUSTOMERIMPL/EXTENSIONS/ENTITYEXTENSIONIMPL/BANKINGDETAILIMPL[PRIMARYKEY/BANKINGDETAILKEY[ENTITYNUMBER/@DATA='" & strPayorCIF & "' and ENTITYSUFFIX/@DATA='" & CStr(intHighestCIFSuffix) & "']]")
            
        If Not nodeBankingDetails Is Nothing Then
            
            ' PSC 30/01/2007 EP2_1129 - Start
            strSortCode = xmlGetNodeText(nodeBankingDetails, "TRANSITNUMBER/@DATA")
            
            If Len(strSortCode) > 0 Then
                strSortCode = Format$(strSortCode, "000000")
            End If
            xmlSetAttributeValue vnodeTarget, "BANKSORTCODE", strSortCode
            ' PSC 30/01/2007 EP2_1129 - End
            
            xmlSetAttributeValue vnodeTarget, "BANKACCOUNTNUMBER", xmlGetNodeText(nodeBankingDetails, "ACCOUNTNUMBER/@DATA")
            xmlSetAttributeValue vnodeTarget, "BANKACCOUNTNAME", xmlGetNodeText(nodeBankingDetails, "CONTACT/CONTACT/NAME/NAME/LINE1/@DATA")
            xmlSetAttributeValue vnodeTarget, "BANKNAME", xmlGetNodeText(nodeBankingDetails, "CONTACT/CONTACT/ADDRESS/ADDRESS/LINE1/@DATA")
        End If
        
    End If
    

AddBankingDetailsDetailsExit:

    Set nodeCustomerKey = Nothing
    Set nodePaymentSource = Nothing
    Set nodeBankingDetails = Nothing
    Set nodeListPaymentSources = Nothing

    errCheckError strFunctionName

End Sub
' PSC 19/01/2007 EP2_928 - End


