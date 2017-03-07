VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CalcsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CalcsBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Calculations Business Object
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'PSC   24/09/99 Created
'JLD   25/02/2000   Added mortgage product details to return from CalcMortCalcCosts()
'JLD   03/05/00     In CalculateMonthlyCosts the LOANAMOUNT needs to be called
'                   TOTALLOANCOMPONENTAMOUNT so that it matches the call from cost modelling/Quick quote
' MS    21/07/00    performance rework.
' DJP   13/09/00    SYS1546 - CalcMortcCalcCosts, pass the document element, not the document.
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'--- BMIDS specific history --------------------------------------------------
'Prog   Date        Description
'DPF    19/7/2002   CMWP3 - Added new method called CalculateRedemptionFees()
'GHun   21/04/2004  BMIDS736 Alpha+ integration
'-----------------------------------------------------------------------------
'Mars History
'Prog   Date        Description
'GHun   04/11/2005  MAR301 CreateKFI
'GHun   08/11/2005  MAR458 Use Location, TypeOfValuation and TypeOfApplication from the request,
'                   and minor tidy up
'GHun   11/11/2005  MAR490 Changed CalculateOneOffCosts to use TypeOfApplication from the request
'INR    30/11/2005  MAR717 Removed redundant code
'GHun   12/05/2006  MAR1779 Change CalcMortCalcCosts to insert OneOffCosts into the correct place in the XML
'-----------------------------------------------------------------------------
'Epsom History
'Prog   Date        Description
'PSC    07/11/2006  EP2_41 CalcMortCalcCosts to calculate as far as oneoffcosts and return them and also to
'                   allow them to be passed in as input
'PSC    29/11/2006  EP2_249 Add CalculateProcFees
'PSC    14/12/2006  EP2_507 Change CalculateMonthlyCosts to take cater for drawdown
'PSC    29/01/2007  EP2_1034 Change CalculateProcFees to use Basic Proc Fee Adjustment
'-----------------------------------------------------------------------------
Option Explicit

Implements ObjectControl

 '=============================================
 'Constant Declaration Section
 '=============================================
Private Const cstrTABLE_NAME As String = "CALCS"

 '=============================================
 'Variable Declaration Section
 '=============================================
Private m_objContext As ObjectContext

Public Function CalcMortCalcCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  calculates Mortgage Calculator Costs
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
'
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format:
'
'   <REQUEST>
'       <CALCS>
'           <MORTGAGEPRODUCTCODE>Value</MORTGAGEPRODUCTCODE>
'           <STARTDATE>Value</STARTDATE>
'           <TERMINYEARS>Value</TERMINYEARS>
'           <TERMINMONTHS>Value</TERMINMONTHS>
'           <AMOUNTREQUESTED>Value</AMOUNTREQUESTED>
'           <LTV>Value</LTV>                                            Optional
'           <PURCHASEPRICE>Value</PURCHASEPRICE>                        Optional
'           <INTERESTONLYELEMENT>Value</INTERESTONLYELEMENT>
'           <CAPITALANDINTERESTELEMENT>Value</CAPITALANDINTERESTELEMENT>
'       </CALCS>
'   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CalcMortCalcCostsVbErr
    
    Const cstrFunctionName As String = "CalcMortCalcCosts"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlOneOffCosts As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlMortProdElement As IXMLDOMElement
    Dim xmlMonthlyCosts As IXMLDOMElement
    Dim xmlCostsElement As IXMLDOMElement
    Dim xmlLTVElem As IXMLDOMElement
    Dim xmlPurchasePriceElem As IXMLDOMElement
    Dim objGlobalParamDO As GlobalParameterDO
    Dim xmlTempElement As IXMLDOMElement
    Dim blnCalcFeesOnly As Boolean  ' PSC 07/11/2006 EP2_41

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = xmlLoad(vstrXMLRequest, cstrFunctionName)
    
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    'Set xmlRequestNode = xmlIn.getElementsByTagName("REQUEST").Item(0)

    If xmlRequestNode Is Nothing Then
        errThrowError cstrFunctionName, oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    ' Set the LTV to the default if it has not been supplied
    Set xmlTempElement = xmlRequestNode.cloneNode(False)
    Set xmlLTVElem = xmlGetNode(xmlRequestNode, "//LTV")
    
    If xmlLTVElem Is Nothing Then
        Set xmlLTVElem = xmlIn.createElement("LTV")
    'MAR490 GHun
    Else
        Set xmlLTVElem = xmlLTVElem.cloneNode(True)
    End If
    
    xmlTempElement.appendChild xmlLTVElem
    'MAR490 End
    
    Dim dblLTV As Double
    
    If CSafeDbl(xmlLTVElem.Text) = 0 Then
        dblLTV = objGlobalParamDO.GetCurrentParameterByType("DefaultLTV", "PERCENTAGE")
        xmlLTVElem.Text = CStr(dblLTV)
    End If
    
    Dim blnPurchasePricePresent As Boolean
    Dim lngAmountRequested As Long
    Dim lngPurchasePrice As Double
        
    Set xmlElement = xmlIn.getElementsByTagName("AMOUNTREQUESTED").Item(0)
    lngAmountRequested = CSafeLng(xmlElement.Text)
    
    blnPurchasePricePresent = True
    
    Set xmlPurchasePriceElem = xmlIn.getElementsByTagName("PURCHASEPRICE").Item(0)
    
    If xmlPurchasePriceElem Is Nothing Then
        Set xmlPurchasePriceElem = xmlIn.createElement("PURCHASEPRICE")
        xmlTempElement.appendChild xmlPurchasePriceElem
    End If
    
    lngPurchasePrice = CSafeLng(xmlPurchasePriceElem.Text)
    
    ' If purchase price is 0 set it up based on the LTV and amount requested
    If lngPurchasePrice = 0 Then
        blnPurchasePricePresent = False
        dblLTV = CSafeDbl(xmlLTVElem.Text)
        lngPurchasePrice = Fix(CDbl(lngAmountRequested) / dblLTV * 100)
        xmlPurchasePriceElem.Text = CStr(lngPurchasePrice)
    End If
            
    ' Get the mortgage product and append it to the request for calculating
    ' the one off costs. Also apend to the output to return.
    ' DJP SYS1546 - pass the document element, not the document.
    
    GetMortgageProduct xmlIn.documentElement
    
    xmlResponseElem.appendChild xmlIn.selectSingleNode("/REQUEST/CALCS").cloneNode(True)
    
    ' If a purchase price is entered get the one off costs and append them
    ' to the request for calculating regular costs and to the response

'   MCS spec change 02/03/00
'
'    If blnPurchasePricePresent = True Then
    If blnPurchasePricePresent = False Then
        
        'Remove the tag
        xmlTempElement.removeChild xmlPurchasePriceElem
    
    End If
    
    ' PSC 07/11/2006 EP2_41 - Start
    Set xmlOneOffCosts = xmlGetNode(xmlRequestNode, "ONEOFFCOSTLIST")

    If xmlOneOffCosts Is Nothing Then
        Set xmlOneOffCosts = CalculateOneOffCosts(xmlRequestNode.xml)
    End If
    
    blnCalcFeesOnly = xmlGetNodeAsBoolean(xmlRequestNode, "CALCULATEFEESONLY")
   
    If blnCalcFeesOnly = False Then

        'APS 10/03/00 - SYS0275 On a warning we resume next and xmlOneOffCosts = nothing
        If Not xmlOneOffCosts Is Nothing Then
            'MAR1779 GHun AlphaPlusRequest.xsl expect OneOffCosts to be found under LoanComponent
            Set xmlNode = xmlRequestNode.selectSingleNode("//LOANCOMPONENTLIST/LOANCOMPONENT")
            If Not xmlNode Is Nothing Then
                xmlNode.appendChild xmlOneOffCosts.cloneNode(True)
            End If
            'MAR1779 End
                        
            xmlResponseElem.appendChild xmlOneOffCosts.cloneNode(True)
        End If
        
        If blnPurchasePricePresent = False Then
            
            'Add the tag back in
            xmlTempElement.appendChild xmlPurchasePriceElem
        
        End If
    
    
        ' Get the incentives and append it to the output to return.
        Set xmlMortProdElement = GetIncentives(xmlRequestNode.xml)
        For Each xmlNode In xmlMortProdElement.childNodes
            xmlResponseElem.appendChild xmlNode.cloneNode(True)
        Next
    
        Set xmlMonthlyCosts = CalculateMonthlyCosts(xmlRequestNode.xml)
        errCheckXMLResponseNode xmlMonthlyCosts, , True
    
        xmlChangeNodeName xmlMonthlyCosts, "RESPONSE", "ALPHAPLUS"
        xmlResponseElem.appendChild xmlMonthlyCosts.cloneNode(True)
    Else
        If Not xmlOneOffCosts Is Nothing Then
            xmlResponseElem.appendChild xmlOneOffCosts
        End If
    End If

            
    CalcMortCalcCosts = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    ' PSC 07/11/2006 EP2_41 - End

CalcMortCalcCostsVbErr:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set xmlOneOffCosts = Nothing
    Set xmlElement = Nothing
    Set xmlMortProdElement = Nothing
    Set xmlMonthlyCosts = Nothing
    Set xmlCostsElement = Nothing
    Set xmlLTVElem = Nothing
    Set xmlPurchasePriceElem = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlTempElement = Nothing
    
    errCheckError cstrFunctionName
End Function

Private Function CalculateOneOffCosts(ByVal vstrXMLRequest As String) As IXMLDOMElement

' header ----------------------------------------------------------------------------------
' description:  Gets the  one off costs
'
' pass:         vstrXmlRequest  xml Request data stream containing request
'
' return:       One Off costs element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateOneOffCostsVbErr
    
    Const strFunctionName As String = "CalculateOneOffCosts"
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objResponseDoc As FreeThreadedDOMDocument40
    Dim objChargesElem As IXMLDOMElement
    Dim objOneOffCostsElem As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objChargesBO As ChargesBO
    Dim objComboDO As ComboDO
       
    Dim strResponse As String
        
    Set objXmlIn = xmlLoad(vstrXMLRequest, strFunctionName)
            
    Set objChargesElem = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    xmlChangeNodeName objChargesElem, cstrTABLE_NAME, "CHARGES"
    
    If m_objContext Is Nothing Then
        Set objChargesBO = New ChargesBO
        Set objComboDO = New ComboDO
    Else
        Set objChargesBO = m_objContext.CreateInstance(App.Title & ".ChargesBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
                   
    Set objElement = objXmlIn.createElement("CONTEXT")
    objElement.Text = "Mortgage Calculator"
    objChargesElem.appendChild objElement
   
    Set objElement = objXmlIn.createElement("TOTALLOANAMOUNT")
    objElement.Text = xmlGetNodeText(objChargesElem, "//AMOUNTREQUESTED")
    objChargesElem.appendChild objElement
    
    Set objElement = objXmlIn.createElement("TYPEOFAPPLICATION")
    'MAR490 GHun Only default TypeOfApplication if it is not passed in
    strResponse = xmlGetNodeText(objXmlIn, "//TYPEOFAPPLICATION")
    If Len(strResponse) > 0 Then
        objElement.Text = strResponse
    Else
        objElement.Text = objComboDO.GetNewLoanValue()
    End If
    'MAR490 End
    objChargesElem.appendChild objElement
    
    Set objElement = objXmlIn.createElement("LOANSNOTREDEEMED")
    objElement.Text = "0"
    objChargesElem.appendChild objElement

    Set objElement = objXmlIn.getElementsByTagName("PURCHASEPRICE").Item(0)
    
    If objElement Is Nothing Then
        errThrowError strFunctionName, 167
    End If
    
    ' Calculate the charges
    strResponse = objChargesBO.CalcCharges(objXmlIn.xml)
    errCheckXMLResponse strResponse, True
    Set objResponseDoc = xmlLoad(strResponse, strFunctionName)
    Set objOneOffCostsElem = objResponseDoc.getElementsByTagName("ONEOFFCOSTLIST").Item(0)
    
    Set CalculateOneOffCosts = objOneOffCostsElem.cloneNode(True)
    
    Set objXmlIn = Nothing
    Set objResponseDoc = Nothing
    Set objChargesElem = Nothing
    Set objElement = Nothing
    Set objChargesBO = Nothing
    Set objComboDO = Nothing
    Set objOneOffCostsElem = Nothing
    
    Exit Function

CalculateOneOffCostsVbErr:

    Set objXmlIn = Nothing
    Set objResponseDoc = Nothing
    Set objChargesElem = Nothing
    Set objElement = Nothing
    Set objChargesBO = Nothing
    Set objComboDO = Nothing
    Set objOneOffCostsElem = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Sub GetMortgageProduct(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Gets the relevant mortgage product structure
'
' pass:         vstrXmlRequest       xml Request data stream
'
' return:       The mortgage product details element
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageProductVbErr
    
    Const cstrFunctionName      As String = "GetMortgageProduct"
    
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    Dim xmlResponseDoc          As FreeThreadedDOMDocument40
    Dim xmlMortProdReqElem      As IXMLDOMElement
    Dim xmlElement              As IXMLDOMElement
    Dim xmlElementCopy          As IXMLDOMElement
    Dim objIComboDO             As IComboDO
    Dim objComboDO              As ComboDO
    Dim objMortgageProductBO    As MortgageProductBO
    Dim xmlCopyRequest          As IXMLDOMElement
    Dim xmlLoanComponent        As IXMLDOMNode
    Dim strValue                As String
    
    ' DJP SYS1546 - the request is changed in here, so don't change the original because it's
    ' used in the calling method. Allowing this method to directly change the input XML
    ' breaks methods elsewhere
    Set xmlCopyRequest = vXMLRequest.cloneNode(True)
    'Set xmlMortProdReqElem = xmlCopyRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    'Set xmlMortProdReqElem = xmlCopyRequest.cloneNode(False)
    'xmlMortProdReqElem.appendChild xmlGetMandatoryNode(xmlCopyRequest, cstrTABLE_NAME).cloneNode(False)
    'xmlChangeNodeName xmlMortProdReqElem, cstrTABLE_NAME, "MORTGAGEPRODUCT"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlMortProdReqElem = xmlDoc.createElement("MORTGAGEPRODUCT")
    
    If m_objContext Is Nothing Then
        Set objIComboDO = New ComboDO
        Set objMortgageProductBO = New MortgageProductBO
    Else
        Set objMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
        Set objIComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
        
    ' Add the values to the request to get the mortgage product details
    Set xmlElement = xmlDoc.createElement("LOCATION")
    'MAR458 GHun Only default location if it is not passed in
    strValue = xmlGetNodeText(xmlCopyRequest, "//LOCATION")
    If Len(strValue) > 0 Then
        xmlElement.Text = strValue
    Else
        xmlElement.Text = objIComboDO.GetQuickQuoteLocationValueId()
    End If
    'MAR458 End
    xmlMortProdReqElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("TYPEOFVALUATION")
    'MAR458 GHun Only default TypeOfValuation if it is not passed in
    strValue = xmlGetNodeText(xmlCopyRequest, "//TYPEOFVALUATION")
    If Len(strValue) > 0 Then
        xmlElement.Text = strValue
    Else
        xmlElement.Text = objIComboDO.GetQuickQuoteValuationTypeValueId()
    End If
    'MAR458 End
    xmlMortProdReqElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("FEETYPE")
    xmlElement.Text = objIComboDO.GetDormantLegalFeeValueId()
    xmlMortProdReqElem.appendChild xmlElement
        
    Set xmlElement = xmlDoc.createElement("TYPEOFAPPLICATION")
    'MAR458 GHun Only default TypeOfApplication if it is not passed in
    strValue = xmlGetNodeText(xmlCopyRequest, "//TYPEOFAPPLICATION")
    If Len(strValue) > 0 Then
        xmlElement.Text = strValue
    Else
        Set objComboDO = objIComboDO
        xmlElement.Text = objComboDO.GetNewLoanValue()
    End If
    'MAR458 End
    xmlMortProdReqElem.appendChild xmlElement
    
    xmlMortProdReqElem.appendChild xmlGetMandatoryNode(xmlCopyRequest, "//AMOUNTREQUESTED").cloneNode(True)
    xmlMortProdReqElem.appendChild xmlGetMandatoryNode(xmlCopyRequest, "//PURCHASEPRICE").cloneNode(True)
    xmlMortProdReqElem.appendChild xmlGetMandatoryNode(xmlCopyRequest, "//LTV").cloneNode(True)
    
    ' Get mortgage product
    Dim strResponse As String
    
    ' SR 19/12/2002 : BM0209 - get data required for CostModelling only
    Set xmlElement = xmlDoc.createElement("CALCSDATAONLY")
    xmlElement.Text = "1"
    xmlMortProdReqElem.appendChild xmlElement
    
    xmlMortProdReqElem.appendChild xmlGetMandatoryNode(xmlCopyRequest, "//APPLICATIONDATE").cloneNode(True)
    
    Set xmlElementCopy = xmlMortProdReqElem.cloneNode(True)
    
    For Each xmlLoanComponent In vXMLRequest.selectNodes("//LOANCOMPONENT")
        Set xmlMortProdReqElem = xmlElementCopy.cloneNode(True)
        
        For Each xmlElement In xmlLoanComponent.childNodes
            xmlMortProdReqElem.appendChild xmlElement.cloneNode(True)
        Next
        
        strResponse = objMortgageProductBO.GetFullProductDetails(xmlMortProdReqElem.xml)
    
        ' Copy product and append copy to the request
        Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
        Set xmlElement = xmlGetNode(xmlResponseDoc, "/RESPONSE/MORTGAGEPRODUCTDETAILS")
        xmlLoanComponent.appendChild xmlElement
    Next
            
GetMortgageProductVbErr:

    Set xmlResponseDoc = Nothing
    Set xmlMortProdReqElem = Nothing
    Set xmlElement = Nothing
    Set objIComboDO = Nothing
    Set objMortgageProductBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Function GetIncentives(ByVal vstrXMLRequest As String) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Gets the incentives for this product
'
' pass:         vstrXmlRequest       xml Request data stream
'
' return:       The incentives element
'------------------------------------------------------------------------------------------
On Error GoTo GetIncentivesVbErr
    
    Const strFunctionName As String = "GetIncentives"
    
    Dim objResponseDoc As FreeThreadedDOMDocument40
    Dim objMortProdReqElem As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objMortgageProductBO As MortgageProductBO
    Dim objXmlIn As FreeThreadedDOMDocument40
    
    Set objXmlIn = xmlLoad(vstrXMLRequest, strFunctionName)
    Set objMortProdReqElem = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    xmlChangeNodeName objMortProdReqElem, cstrTABLE_NAME, "MORTGAGEPRODUCT"
    xmlChangeNodeName objMortProdReqElem, "AMOUNTREQUESTED", "LOANAMOUNT"
    
    If m_objContext Is Nothing Then
        Set objMortgageProductBO = New MortgageProductBO
    Else
        Set objMortgageProductBO = _
                    m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
    End If
        
    ' Get mortgage product incentives
    Dim strResponse As String
    strResponse = objMortgageProductBO.GetIncentivesForProduct(objXmlIn.xml)
    
    ' Copy incentives and append copy to the request
    Set objResponseDoc = xmlLoad(strResponse, strFunctionName)
        
    Set GetIncentives = objResponseDoc.documentElement.cloneNode(True)
        
    Set objResponseDoc = Nothing
    Set objMortProdReqElem = Nothing
    Set objElement = Nothing
    Set objMortgageProductBO = Nothing
    Set objXmlIn = Nothing
    
    Exit Function

GetIncentivesVbErr:

    Set objResponseDoc = Nothing
    Set objMortProdReqElem = Nothing
    Set objElement = Nothing
    Set objMortgageProductBO = Nothing
    Set objXmlIn = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function CalculateMonthlyCosts(ByVal vstrXMLRequest As String) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Calculates the monthly costs
'
' pass:         vstrXmlRequest       xml Request data stream
'
' return:       The monthly costs elements
'------------------------------------------------------------------------------------------
On Error GoTo CalculateMonthlyCostsVbErr

    Const strFunctionName As String = "CalculateMonthlyCosts"

    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    
    Dim xmlCostsElement As IXMLDOMElement
    Dim xmlCalcElem As IXMLDOMElement
    Dim xmlRepayTypeElem As IXMLDOMElement
    'Dim xmlCapAndIntAmtElem As IXMLDOMElement
    'Dim xmlIntOnlyAmtElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlMonthlyCost As IXMLDOMElement
    Dim xmlMortgageCostsElem As IXMLDOMElement
    Dim objMortgageProductBO As MortgageProductBO
    'Dim strIntOnlyAmount As String
    'Dim strCapAndIntAmount As String
    
    ' PSC 14/12/2006 EP2_507 - Start
    Dim xmlRequestCopy As IXMLDOMNode
    Dim xmlLoanAmount As IXMLDOMNode
    Dim xmlDrawdownCalcResponse As IXMLDOMNode
    Dim xmlDrawDownMonthlyCost As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlLoanComponentList As IXMLDOMNodeList

    Dim dblDrawdown As Double
    Dim dblLoanAmount As Double
    ' PSC 14/12/2006 EP2_507 - End

    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    
    Dim xmlResponseElem  As IXMLDOMElement
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlOutDoc.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim xmlLoanComponent As IXMLDOMNode
    Dim intSequence As Integer
    
    Set xmlInDoc = xmlLoad(vstrXMLRequest, strFunctionName)
    Set xmlMortgageCostsElem = xmlInDoc.getElementsByTagName(cstrTABLE_NAME).Item(0)
    xmlChangeNodeName xmlMortgageCostsElem, cstrTABLE_NAME, "MORTGAGECOSTS"
    'xmlChangeNodeName xmlMortgageCostsElem, "AMOUNTREQUESTED", "LOANAMOUNT"
    'The LOANAMOUNT needs to be called TOTALLOANCOMPONENTAMOUNT so that it matches the call from cost modelling/Quick quote
    xmlChangeNodeName xmlMortgageCostsElem, "AMOUNTREQUESTED", "TOTALLOANCOMPONENTAMOUNT"

'MAR717 Redundant Code
'    Set xmlIntOnlyAmtElem = xmlInDoc.getElementsByTagName("INTERESTONLYELEMENT").Item(0)
'
'    If Not xmlIntOnlyAmtElem Is Nothing Then
'        strIntOnlyAmount = xmlIntOnlyAmtElem.Text
'        xmlIntOnlyAmtElem.Text = vbNullString
'    End If
'
'    Set xmlCapAndIntAmtElem = xmlInDoc.getElementsByTagName("CAPITALANDINTERESTELEMENT").Item(0)
'
'    If Not xmlCapAndIntAmtElem Is Nothing Then
'        strCapAndIntAmount = xmlCapAndIntAmtElem.Text
'        xmlCapAndIntAmtElem.Text = vbNullString
'    End If

    If m_objContext Is Nothing Then
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
    Else
        Set objIMortgageSubQuoteBO = _
                    m_objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
    End If

    Set xmlCostsElement = xmlOutDoc.createElement("REGULARCOSTS")
    xmlResponseElem.appendChild xmlCostsElement
    
    intSequence = 1
    For Each xmlLoanComponent In xmlMortgageCostsElem.selectNodes("//LOANCOMPONENT")
        xmlChangeNodeName xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT", "LOANAMOUNT"
        Set xmlElement = xmlInDoc.createElement("LOANCOMPONENTSEQUENCENUMBER")
        xmlElement.Text = intSequence
        xmlLoanComponent.appendChild xmlElement
        intSequence = intSequence + 1
    Next
    ' Calculate repayments
    
    ' PSC 14/12/2006 EP2_507 - Start
    dblDrawdown = xmlGetNodeAsDouble(xmlMortgageCostsElem, ".//MORTGAGESUBQUOTE/DRAWDOWN")
    
    If dblDrawdown > 0 Then
       Set xmlRequestCopy = xmlMortgageCostsElem.cloneNode(True)
    End If
    ' PSC 14/12/2006 EP2_507 - End

    Set xmlCalcElem = objIMortgageSubQuoteBO.RunCompositeRateCalc(xmlMortgageCostsElem)
    errCheckXMLResponseNode xmlCalcElem, xmlResponseElem, True
    
    ' PSC 14/12/2006 EP2_507 - Start
    If dblDrawdown > 0 Then
        Set xmlLoanComponent = xmlGetMandatoryNode(xmlRequestCopy, ".//MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT[position()=last()]")
        dblLoanAmount = xmlGetNodeAsDouble(xmlLoanComponent, "LOANAMOUNT")
        dblLoanAmount = dblLoanAmount - dblDrawdown
        
        If dblLoanAmount = 0 Then
            xmlLoanComponent.parentNode.removeChild xmlLoanComponent
        Else
            Set xmlLoanAmount = xmlGetMandatoryNode(xmlLoanComponent, "LOANAMOUNT")
            xmlLoanAmount.Text = CStr(dblLoanAmount)
        End If
     
        Set xmlDrawdownCalcResponse = objIMortgageSubQuoteBO.RunCompositeRateCalc(xmlRequestCopy)
        errCheckXMLResponseNode xmlCalcElem, xmlResponseElem, True
        
        ' Add a MONTHLYCOSTLESSDRAWDOWN node onto each output loan component
        Set xmlLoanComponentList = xmlCalcElem.selectNodes("LOANCOMPONENTLIST/LOANCOMPONENT")
        
        For Each xmlLoanComponent In xmlLoanComponentList
            intSequence = xmlGetMandatoryNodeAsInteger(xmlLoanComponent, "LOANCOMPONENTSEQUENCENUMBER")

            Set xmlNode = xmlGetNode(xmlDrawdownCalcResponse, "LOANCOMPONENTLIST/LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER='" & CStr(intSequence) & "']/NETMONTHLYCOST")
            Set xmlDrawDownMonthlyCost = xmlOutDoc.createElement("MONTHLYCOSTLESSDRAWDOWN")
            
            If Not xmlNode Is Nothing Then
                xmlDrawDownMonthlyCost.Text = xmlNode.Text
            Else
                xmlDrawDownMonthlyCost.Text = 0#
            End If
            xmlLoanComponent.appendChild xmlDrawDownMonthlyCost
        Next
    
        ' Add a node at the top level to show total cost less drawdown
        Set xmlNode = xmlGetMandatoryNode(xmlDrawdownCalcResponse, "TOTALNETMONTHLYCOST")
        Set xmlDrawDownMonthlyCost = xmlOutDoc.createElement("MONTHLYCOSTLESSDRAWDOWN")
        xmlDrawDownMonthlyCost.Text = xmlNode.Text
        xmlCalcElem.appendChild xmlDrawDownMonthlyCost
    End If
    ' PSC 14/12/2006 EP2_507 - End

    Set CalculateMonthlyCosts = xmlCalcElem

CalculateMonthlyCostsVbErr:

    Set xmlOutDoc = Nothing
    Set xmlInDoc = Nothing
    Set xmlCostsElement = Nothing

    Set xmlCalcElem = Nothing
    Set xmlRepayTypeElem = Nothing
    'Set xmlCapAndIntAmtElem = Nothing
    'Set xmlIntOnlyAmtElem = Nothing
    Set xmlElement = Nothing
    Set objMortgageProductBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set xmlMonthlyCost = Nothing
    Set xmlMortgageCostsElem = Nothing
    
    ' PSC 14/12/2006 EP2_507 - Start
    Set xmlRequestCopy = Nothing
    Set xmlLoanAmount = Nothing
    Set xmlDrawdownCalcResponse = Nothing
    Set xmlDrawDownMonthlyCost = Nothing
    Set xmlNode = Nothing
    Set xmlLoanComponentList = Nothing
    ' PSC 14/12/2006 EP2_507 - End

    errCheckError strFunctionName
    
End Function


Public Function CalculateRedemptionFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  calculates Redemption Fees
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
'
' return:       xml Response data stream containing results of operation
'               either: TYPE="SUCCESS"
'               or: TYPE="SYSERR" and <ERROR> element
'
' Request:      xml Request data stream containing loan component details passed in from
'               the MortgageSubQuoteTxBo.Calc&SaveRedemptionFees
'------------------------------------------------------------------------------------------
On Error GoTo CalculateRedemptionFeesVbErr
    
    Const strFunctionName As String = "CalculateRedemptionFees"

    'Variable declaration
    
    'Assist modules and external objects
    Dim objCMRules As Object
    
    'incoming xml
    Dim objXmlLoanComponentIn As IXMLDOMElement
    
    'feeding xml
    Dim objXmlFeedCMRules As IXMLDOMNode
    Dim objXmlCMRulesRespAttribBased As FreeThreadedDOMDocument40
    Dim strCMRulesAttBasedResponse As String
    Dim objXmlCMRulesRespElemBased As FreeThreadedDOMDocument40
    
    'outgoing xml
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim objSeqNo As IXMLDOMElement
    Dim objAppNo As IXMLDOMElement
    Dim objAppFactFindNo As IXMLDOMElement
    Dim objSubQuoteNo As IXMLDOMElement
    Dim objStepNo As IXMLDOMElement
    Dim objFeePeriod As IXMLDOMElement
    Dim objEndDate As IXMLDOMElement
    Dim objFeeAmount As IXMLDOMElement
    Dim objXmlFeeListElem As IXMLDOMElement
    Dim objXmlFeeElem As IXMLDOMElement
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objFeeBandNode As IXMLDOMNode
    Dim objCMRulesBandNodeList As IXMLDOMNode
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Set objXmlCMRulesRespAttribBased = New FreeThreadedDOMDocument40
    objXmlCMRulesRespAttribBased.validateOnParse = False
    objXmlCMRulesRespAttribBased.setProperty "NewParser", True
    Set objXmlCMRulesRespElemBased = New FreeThreadedDOMDocument40
    objXmlCMRulesRespElemBased.validateOnParse = False
    objXmlCMRulesRespElemBased.setProperty "NewParser", True
    
    'set up CMRules object
    If m_objContext Is Nothing Then
        Set objCMRules = CreateObject("omCMRules.CalcsBO")
    Else
        Set objCMRules = m_objContext.CreateInstance("omCMRules.CalcsBO")
    End If
        
    'build response block
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    objXmlOut.appendChild objXmlResponseElem
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'copy request into empty Dom Document
    If vXMLRequest.nodeName = "LOANCOMPONENT" Then
        Set objXmlLoanComponentIn = vXMLRequest
    Else
        Set objXmlLoanComponentIn = vXMLRequest.getElementsByTagName("LOANCOMPONENT").Item(0)
    End If
            
    'Check we have some data in the response
    If objXmlLoanComponentIn Is Nothing Then
        errThrowError strFunctionName, oeMissingPrimaryTag, "LOANCOMPONENT tag not found"
    End If
    
    'convert request to attribute based as the component we are passing request into is Phase2
    Set objXmlFeedCMRules = xmlCreateAttributeBasedResponse(objXmlLoanComponentIn, True)
    
    'call cmRules.CalculateRedemptionFees and pull back redemption fee calculations
    strCMRulesAttBasedResponse = objCMRules.CalculateRedemptionFees(objXmlFeedCMRules.xml)
    Set objXmlCMRulesRespAttribBased = xmlLoad(strCMRulesAttBasedResponse, strFunctionName)
    
    'convert response back to element based
    Set objXmlCMRulesRespElemBased = xmlCreateElementRequestFromNode(objXmlCMRulesRespAttribBased.documentElement, "REDEMPTIONFEEAMOUNTLIST", True, "REDEMPTIONFEEAMOUNTLIST")
    
    'build XML response block
    Set objXmlFeeListElem = objXmlOut.createElement("LOANCOMPONENTREDEMPTIONFEELIST")
    Set objCMRulesBandNodeList = objXmlCMRulesRespElemBased.documentElement.selectSingleNode("//REDEMPTIONFEEAMOUNTLIST")
                     
    'Loop through Redemption Fee records passed in from CMRules object
    For Each objFeeBandNode In objCMRulesBandNodeList.selectNodes("REDEMPTIONFEEAMOUNT")
        If Not objFeeBandNode Is Nothing Then
            
            'create required elements for each record for the response
            Set objXmlFeeElem = objXmlOut.createElement("LOANCOMPONENTREDEMPTIONFEE")
            Set objSeqNo = objXmlOut.createElement("LOANCOMPONENTSEQUENCENUMBER")
            Set objAppNo = objXmlOut.createElement("APPLICATIONNUMBER")
            Set objAppFactFindNo = objXmlOut.createElement("APPLICATIONFACTFINDNUMBER")
            Set objSubQuoteNo = objXmlOut.createElement("MORTGAGESUBQUOTENUMBER")
            Set objStepNo = objXmlOut.createElement("REDEMPTIONFEESTEPNUMBER")
            Set objFeePeriod = objXmlOut.createElement("REDEMPTIONFEEPERIOD")
            Set objEndDate = objXmlOut.createElement("REDEMPTIONFEEPERIODENDDATE")
            Set objFeeAmount = objXmlOut.createElement("REDEMPTIONFEEAMOUNT")
            
            'assign values to the elements for the response
            objSeqNo.Text = objXmlLoanComponentIn.selectSingleNode("LOANCOMPONENTSEQUENCENUMBER").Text
            objAppNo.Text = objXmlLoanComponentIn.selectSingleNode("APPLICATIONNUMBER").Text
            objAppFactFindNo.Text = objXmlLoanComponentIn.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
            objSubQuoteNo.Text = objXmlLoanComponentIn.selectSingleNode("MORTGAGESUBQUOTENUMBER").Text
            objStepNo.Text = objFeeBandNode.selectSingleNode("STEP").Text
            If Not IsDate(objEndDate.Text) > 0 Then
                objFeePeriod.Text = objFeeBandNode.selectSingleNode("FEEPERIOD").Text
            Else
                objEndDate.Text = objFeeBandNode.selectSingleNode("FEEPERIODENDDATE").Text
            End If
            objFeeAmount.Text = objFeeBandNode.selectSingleNode("AMOUNT").Text
            
            'append these elements to the response
            objXmlFeeElem.appendChild objSeqNo
            objXmlFeeElem.appendChild objAppNo
            objXmlFeeElem.appendChild objAppFactFindNo
            objXmlFeeElem.appendChild objSubQuoteNo
            objXmlFeeElem.appendChild objStepNo
            objXmlFeeElem.appendChild objFeePeriod
            objXmlFeeElem.appendChild objEndDate
            objXmlFeeElem.appendChild objFeeAmount
            
            'add individual record to fees block.
            objXmlFeeListElem.appendChild objXmlFeeElem
            
        Else
            errThrowError strFunctionName, oeMissingPrimaryTag, "REDEMPTIONFEEBAND tag not found"
        End If
    Next
    
    'add fee details to the response block
    objXmlResponseElem.appendChild objXmlFeeListElem
    
    'send response back to calling function
    Set CalculateRedemptionFees = objXmlOut.documentElement

'**************************************
'Close down function

CalculateRedemptionFeesExit:

    Set objCMRules = Nothing
    
    'incoming xml
    Set objXmlLoanComponentIn = Nothing
    
    'feeding xml
    Set objXmlFeedCMRules = Nothing
    Set objXmlCMRulesRespAttribBased = Nothing
    Set objXmlCMRulesRespElemBased = Nothing
    
    'outgoing xml
    Set objXmlOut = Nothing
    Set objSeqNo = Nothing
    Set objAppNo = Nothing
    Set objSubQuoteNo = Nothing
    Set objStepNo = Nothing
    Set objFeePeriod = Nothing
    Set objEndDate = Nothing
    Set objFeeAmount = Nothing
    Set objXmlFeeListElem = Nothing
    Set objXmlFeeElem = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objFeeBandNode = Nothing
    Set objCMRulesBandNodeList = Nothing

    Exit Function

CalculateRedemptionFeesVbErr:
           
    CalculateRedemptionFees = errCreateErrorResponse()
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalculateRedemptionFeesExit

End Function

' PSC 29/11/2006 EP2_249 - Start
Public Function CalculateProcFees(ByVal vstrXMLRequest As String) As String
On Error GoTo CalculateProcFeesVbErr
    
    Const cstrFunctionName As String = "CalculateProcFees"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCrudResponse As FreeThreadedDOMDocument40
    Dim objCRUD As omCRUDBO
    Dim xmlRequest As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    Dim xmlCrudRequest As IXMLDOMNode
    Dim xmlOperation As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlProcFees As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlLoanComponent As IXMLDOMNode

    Dim strSubmissionRoute As String
    Dim strProductScheme As String
    Dim strProductCategory As String
    Dim strResponse As String

    Dim intTypeOfApplication As Integer

    Dim dblLTV As Double
    Dim dblLoanAmount As Double
    Dim dblBasicProcFeeRate As Double
    Dim dblLoanComponentAmount As Double
    Dim dblExclusiveLoading As Double
    
    Dim dblAssociationFeeRate As Double
    Dim dblAssociationFeeAmount As Double
    Dim dblProcFee As Double
    Dim dblAssociationFee As Double
    Dim dblPackagingFee As Double
    Dim dblBasicProcFeeAdj As Double        ' PSC 29/01/2007 EP2_1034
    
    Dim blnLoanAndLTVInd As Boolean
    Dim blnOnlineInd As Boolean
    
    dblProcFee = 0#
    dblAssociationFee = 0#
    dblPackagingFee = 0#
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponse, "TYPE", "SUCCESS"
    xmlOut.appendChild xmlResponse
    
    Set xmlIn = xmlLoad(vstrXMLRequest, cstrFunctionName)
    
    Set xmlRequest = xmlGetMandatoryNode(xmlIn, "REQUEST")
    Set xmlProcFees = xmlGetMandatoryNode(xmlRequest, "PROCFEES")
    
    strSubmissionRoute = xmlGetMandatoryNodeText(xmlProcFees, "SUBMISSIONROUTE")
    strProductScheme = xmlGetMandatoryNodeText(xmlProcFees, "PRODUCTSCHEME")
    strProductCategory = xmlGetMandatoryNodeText(xmlProcFees, "PRODUCTCATEGORY")
    dblLoanAmount = xmlGetMandatoryNodeAsDouble(xmlProcFees, "LOANAMOUNT")
    dblLTV = xmlGetMandatoryNodeAsDouble(xmlProcFees, "LTV")
    intTypeOfApplication = xmlGetMandatoryNodeAsInteger(xmlProcFees, "TYPEOFAPPLICATION")
    blnLoanAndLTVInd = xmlGetNodeAsBoolean(xmlProcFees, "LOANANDLTVINDICATOR")
    blnOnlineInd = xmlGetNodeAsBoolean(xmlProcFees, "ONLINEINDICATOR")
    dblAssociationFeeAmount = xmlGetNodeAsDouble(xmlProcFees, "ASSOCIATIONFEEAMOUNT")
    dblAssociationFeeRate = xmlGetNodeAsDouble(xmlProcFees, "ASSOCIATIONFEERATE")
    
    ' PSC 29/01/2007 EP2_1034
    dblBasicProcFeeAdj = xmlGetNodeAsDouble(xmlProcFees, "BASICPROCFEEADJ")
    
    If Not (IsValidationType("TypeOfMortgage", intTypeOfApplication, "ABO") Or _
            IsValidationType("TypeOfMortgage", intTypeOfApplication, "TOE") Or _
            IsValidationType("TypeOfMortgage", intTypeOfApplication, "PSW") Or _
            IsValidationType("TypeOfMortgage", intTypeOfApplication, "CLI") Or _
            IsValidationType("TypeOfMortgage", intTypeOfApplication, "NP")) Then
        
        If m_objContext Is Nothing Then
            Set objCRUD = New omCRUDBO
        Else
            Set objCRUD = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDBO")
        End If
        
        ' Set up request to get the appropriate calcs data
        Set xmlCrudRequest = xmlRequest.cloneNode(False)
        
        Set xmlOperation = xmlOut.createElement("OPERATION")
        xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
        xmlSetAttributeValue xmlOperation, "ENTITY_REF", "PROCFEEDEFAULT"
        xmlCrudRequest.appendChild xmlOperation
        
        Set xmlRoot = xmlOut.createElement("PROCFEEDEFAULT")
        xmlSetAttributeValue xmlRoot, "SUBMISSIONROUTE", strSubmissionRoute
        xmlSetAttributeValue xmlRoot, "PRODUCTSCHEME", strProductScheme
        xmlOperation.appendChild xmlRoot
        
        If blnLoanAndLTVInd Then
            Set xmlOperation = xmlOut.createElement("OPERATION")
            xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
            xmlSetAttributeValue xmlOperation, "ENTITY_REF", "PROCFEEADJBYLOAN"
            xmlSetAttributeValue xmlOperation, "SCHEMA_NAME", "omCM"
            xmlCrudRequest.appendChild xmlOperation
        
            Set xmlRoot = xmlOut.createElement("PROCFEEADJBYLOAN")
            xmlSetAttributeValue xmlRoot, "PRODUCTCATEGORY", strProductCategory
            xmlSetAttributeValue xmlRoot, "LOANAMOUNT", dblLoanAmount
            xmlOperation.appendChild xmlRoot
        
            Set xmlOperation = xmlOut.createElement("OPERATION")
            xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
            xmlSetAttributeValue xmlOperation, "ENTITY_REF", "PROCFEEADJBYLTV"
            xmlSetAttributeValue xmlOperation, "SCHEMA_NAME", "omCM"
            xmlCrudRequest.appendChild xmlOperation
        
            Set xmlRoot = xmlOut.createElement("PROCFEEADJBYLTV")
            xmlSetAttributeValue xmlRoot, "PRODUCTCATEGORY", strProductCategory
            xmlSetAttributeValue xmlRoot, "LTV", dblLTV
            xmlOperation.appendChild xmlRoot
        End If
        
        If blnOnlineInd Then
            Set xmlOperation = xmlOut.createElement("OPERATION")
            xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
            xmlSetAttributeValue xmlOperation, "ENTITY_REF", "GLOBALPARAMETER"
            xmlCrudRequest.appendChild xmlOperation
        
            Set xmlRoot = xmlOut.createElement("GLOBALPARAMETER")
            xmlSetAttributeValue xmlRoot, "NAME", "ProcFeeOnlineLoading"
            xmlOperation.appendChild xmlRoot
        End If
        
        If IsValidationType("ProcFeeSubmissionRoute", strSubmissionRoute, "PA") Or _
           IsValidationType("ProcFeeSubmissionRoute", strSubmissionRoute, "P") Then
        
            Set xmlOperation = xmlOut.createElement("OPERATION")
            xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
            xmlSetAttributeValue xmlOperation, "ENTITY_REF", "GLOBALPARAMETER"
            xmlCrudRequest.appendChild xmlOperation
        
            Set xmlRoot = xmlOut.createElement("GLOBALPARAMETER")
            xmlSetAttributeValue xmlRoot, "NAME", "PackagingFee"
            xmlOperation.appendChild xmlRoot
        
        End If
        
        strResponse = objCRUD.OmRequest(xmlCrudRequest.xml)
        Set xmlCrudResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlCrudResponse.documentElement, xmlResponse, True
        
        Set xmlTempNode = xmlGetMandatoryNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/PROCFEEDEFAULT")
        
        dblBasicProcFeeRate = xmlGetAttributeAsDouble(xmlTempNode, "FEERATE")
        
        ' PSC 29/01/2007 EP2_1034
        dblBasicProcFeeRate = dblBasicProcFeeRate + dblBasicProcFeeAdj
        
        Set xmlTempNode = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/PROCFEEADJBYLOAN")
        
        If Not xmlTempNode Is Nothing Then
            dblBasicProcFeeRate = dblBasicProcFeeRate + xmlGetAttributeAsDouble(xmlTempNode, "FEERATE")
        End If
        
        Set xmlTempNode = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/PROCFEEADJBYLTV")
        
        If Not xmlTempNode Is Nothing Then
            dblBasicProcFeeRate = dblBasicProcFeeRate + xmlGetAttributeAsDouble(xmlTempNode, "FEERATE")
        End If
    
        Set xmlTempNode = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='ProcFeeOnlineLoading']")
        
        If Not xmlTempNode Is Nothing Then
            dblBasicProcFeeRate = dblBasicProcFeeRate + xmlGetAttributeAsDouble(xmlTempNode, "PERCENTAGE")
        End If
    
        ' Calculate proc fee
        Set xmlLoanComponentList = xmlProcFees.selectNodes("LOANCOMPONENTLIST/LOANCOMPONENT")
        
        For Each xmlLoanComponent In xmlLoanComponentList
        
            dblLoanComponentAmount = xmlGetNodeAsDouble(xmlLoanComponent, "LOANAMOUNT")
            dblExclusiveLoading = xmlGetNodeAsDouble(xmlLoanComponent, "EXCLUSIVELOADING")
        
            dblProcFee = dblProcFee + (dblLoanComponentAmount * ((dblBasicProcFeeRate + dblExclusiveLoading) / 100#))
        Next
        
        ' Calculate Association Fee
        If IsValidationType("ProcFeeSubmissionRoute", strSubmissionRoute, "PA") Or _
           IsValidationType("ProcFeeSubmissionRoute", strSubmissionRoute, "MC") Then
           
            If dblAssociationFeeRate > 0 Then
                dblAssociationFee = dblLoanAmount * (dblAssociationFeeRate / 100#)
        
            ElseIf dblAssociationFeeAmount > 0 Then
                dblAssociationFee = dblAssociationFeeAmount
            End If
        End If
        
        ' Calculate Packaging Fee
        If IsValidationType("ProcFeeSubmissionRoute", strSubmissionRoute, "PA") Or _
           IsValidationType("ProcFeeSubmissionRoute", strSubmissionRoute, "P") Then
           
            Set xmlTempNode = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='PackagingFee']")
        
            If Not xmlTempNode Is Nothing Then
                dblPackagingFee = xmlGetAttributeAsDouble(xmlTempNode, "AMOUNT")
            End If
        End If
    End If
    
    Set xmlRoot = xmlOut.createElement("PROCFEES")
    xmlResponse.appendChild xmlRoot
    Set xmlTempNode = xmlOut.createElement("PROCFEE")
    xmlTempNode.Text = CLng(dblProcFee)
    xmlResponse.appendChild xmlTempNode
    Set xmlTempNode = xmlOut.createElement("ASSOCIATIONFEE")
    xmlTempNode.Text = CLng(dblAssociationFee)
    xmlResponse.appendChild xmlTempNode
    Set xmlTempNode = xmlOut.createElement("PACKAGINGFEE")
    xmlTempNode.Text = CLng(dblPackagingFee)
    xmlResponse.appendChild xmlTempNode
    
    CalculateProcFees = xmlOut.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateProcFeesExit:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlCrudResponse = Nothing
    Set objCRUD = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlOperation = Nothing
    Set xmlRoot = Nothing
    Set xmlTempNode = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlProcFees = Nothing
    
    Exit Function

CalculateProcFeesVbErr:
    
    CalculateProcFees = errCreateErrorResponse()
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalculateProcFeesExit

End Function
' PSC 29/11/2006 EP2_249 - End


Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub
