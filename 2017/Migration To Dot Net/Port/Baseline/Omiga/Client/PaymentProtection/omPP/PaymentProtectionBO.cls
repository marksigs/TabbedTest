VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PaymentProtectionBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProtectionBO.cls
'
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Payment Protection Business Object which 'supports transactions' only

'Dependencies:  PaymentProtectionDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    14/10/99    Created
'SR     08/12/99    Modified method - Calculate Costs
'DRC    05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
Private Const cstrCALC_NAME = "PAYMENTPROTECTION"


Public Function FindProductList() As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         n/a
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindProductListVbErr
    
    Const strFunctionName As String = "FindProductList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objPaymentProtectionDO As PaymentProtectionDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objPaymentProtectionDO = New PaymentProtectionDO
    Else
        Set objPaymentProtectionDO = _
                            objContext.CreateInstance(App.Title & ".PaymentProtectionDO")
    End If
            
    Dim strXMLData As String
        
    strXMLData = objPaymentProtectionDO.FindProductList()
    Set xmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
        
    FindProductList = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindProductListExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objPaymentProtectionDO = Nothing

    Exit Function

FindProductListVbErr:

    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindProductList = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindProductListExit

End Function

Public Function CalculateCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'
'
'   <PAYMENTPROTECTION>
'       <APPLICANT1GENDER>Value</APPLICANT1GENDER>
'       <APPLICANT1AGE>Value</APPLICANT1AGE>
'       <APPLICANT2GENDER>Value</APPLICANT2GENDER>
'       <APPLICANT2AGE>Value</APPLICANT2AGE>
'       <DISTRIBUTIONCHANNEL>Value</DISTRIBUTIONCHANNEL>
'       <COVERTYPE>Value</COVERTYPE>
'       <MORTGAGEAMOUNT>Value</MORTGAGEAMOUNT>
'       <MORTGAGECOVERFORAPPLICANT1>Value</MORTGAGECOVERFORAPPLICANT1>
'       <MORTGAGECOVERFORAPPLICANT2>Value</MORTGAGECOVERFORAPPLICANT2>
'       <RECOMMENDEDMORTGAGECOVERAMOUNT>Value</RECOMMENDEDMORTGAGECOVERAMOUNT>
'       <PAYMENTPROTECTIONPRODUCT>
'           ...
'           ...
'       </PAYMENTPROTECTIONPRODUCT>
'   </PAYMENTPROTECTION>
'------------------------------------------------------------------------------------------
On Error GoTo CalculateCostsVbErr
    
    Const strFunctionName As String = "CalculateCosts"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    Dim xmlPayProtProductElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objGlobalParamDO As GlobalParameterDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
 
    ValidateCalculateCostsInput vstrXMLRequest
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrCALC_NAME).Item(0)
    
    Set xmlPayProtProductElem = _
                    xmlRequestElem.getElementsByTagName("PAYMENTPROTECTIONPRODUCT").Item(0)
    
    Dim dblApp1ASURate As Double
    Dim dblApp1ASRate As Double
    Dim dblApp1URate As Double
    Dim dblApp2ASURate As Double
    Dim dblApp2ASRate As Double
    Dim dblApp2URate As Double
    Dim strDistributionChannel As String
    Dim strApp1Gender As String
    Dim strApp1Age As String
    Dim strApp2Gender As String
    Dim strApp2Age As String
    Dim strPPProdRateNo As String
    Dim strTagName As String
    Dim strTagValue As String
    
    strTagName = "PAYMENTPROTECTIONRATESNUMBER"
    strPPProdRateNo = objXmlAssist.GetTagValue(xmlPayProtProductElem, strTagName)
    strTagName = "DISTRIBUTIONCHANNEL"
    strDistributionChannel = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    strTagName = "APPLICANT1GENDER"
    strApp1Gender = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    strTagName = "APPLICANT1AGE"
    strApp1Age = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    strTagName = "APPLICANT2GENDER"
    strApp2Gender = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    strTagName = "APPLICANT2AGE"
    strApp2Age = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    
    ' Get the rates for each applicant
    If Len(strApp1Gender) > 0 And Len(strApp1Age) > 0 Then
        GetRates strPPProdRateNo, strDistributionChannel, strApp1Gender, strApp1Age, _
                 dblApp1ASURate, dblApp1ASRate, dblApp1URate
    End If
    
    If Len(strApp2Gender) > 0 And Len(strApp2Age) > 0 Then
        GetRates strPPProdRateNo, strDistributionChannel, strApp2Gender, strApp2Age, _
                 dblApp2ASURate, dblApp2ASRate, dblApp2URate
    End If
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
                        objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    Dim dblRoundFactor As Double
    Dim dblMortgageCoverAmt As Double
    Dim dblMaxCoverAmt As Double
    Dim dblApp1MortgageCover As Double
    Dim dblApp2MortgageCover As Double
    
    strTagName = "MORTGAGEAMOUNT"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMortgageCoverAmt = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "MORTGAGECOVERFORAPPLICANT1"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblApp1MortgageCover = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "MORTGAGECOVERFORAPPLICANT2"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblApp2MortgageCover = objConvAssist.CSafeDbl(strTagValue)
    
    dblRoundFactor = objGlobalParamDO.GetCurrentParameterByType("PPRoundingFactor", "AMOUNT")
    
    If (dblMortgageCoverAmt Mod dblRoundFactor) <> 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 229, , dblRoundFactor
    End If
    
    dblMaxCoverAmt = objGlobalParamDO.GetCurrentParameterByType("PPMaximumCover", "AMOUNT")
    
    ' Adjust the cover amounts if they exceed the maximum permissible
    If dblMortgageCoverAmt > dblMaxCoverAmt Then
        
        ' Raise a warning
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 230
        
        dblMortgageCoverAmt = dblMaxCoverAmt
    
        If dblApp1MortgageCover > 0 And dblApp2MortgageCover > 0 Then
            dblApp2MortgageCover = dblMortgageCoverAmt / 2
            dblApp1MortgageCover = dblMortgageCoverAmt - dblApp2MortgageCover
        ElseIf dblApp1MortgageCover > 0 Then
            dblApp1MortgageCover = dblMortgageCoverAmt
        ElseIf dblApp2MortgageCover > 0 Then
            dblApp2MortgageCover = dblMortgageCoverAmt
        End If
        
    End If
    
    Dim strCoverType As String
    Dim dblApp1Rate As Double
    Dim dblApp2Rate As Double
    
    strTagName = "COVERTYPE"
    strCoverType = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)

    Select Case strCoverType
        Case "A"
            dblApp1Rate = dblApp1ASURate
            dblApp2Rate = dblApp2ASURate
        Case "S"
            dblApp1Rate = dblApp1ASRate
            dblApp2Rate = dblApp2ASRate
        Case "U"
            dblApp1Rate = dblApp1URate
            dblApp2Rate = dblApp2URate
        Case Else
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeInvalidParameter, _
                                    "Invalid Cover Type: " & strCoverType
    End Select
    
    
    Dim dblInsPremiumTaxRate As Double
    Dim dblApp1Premium As Double
    Dim dblApp1IPT As Double
    Dim dblApp2Premium As Double
    Dim dblApp2IPT As Double
   
    dblInsPremiumTaxRate = _
                objGlobalParamDO.GetCurrentParameterByType("PPInsPremTaxRate", "PERCENTAGE")
    
    If dblApp1MortgageCover > 0 Then
        dblApp1Premium = dblApp1MortgageCover * dblApp1Rate / 100
        dblApp1IPT = dblApp1Premium * dblInsPremiumTaxRate / 100
    End If
    
    If dblApp2MortgageCover > 0 Then
        dblApp2Premium = dblApp2MortgageCover * dblApp2Rate / 100
        dblApp2IPT = dblApp2Premium * dblInsPremiumTaxRate / 100
    End If
    
    strTagName = "APPLICANT1PREMIUM"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblApp1Premium, "0.00")
    xmlResponseElem.appendChild xmlElement
    
    strTagName = "APPLICANT1IPT"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblApp1IPT, "0.00")
    xmlResponseElem.appendChild xmlElement
     
    strTagName = "APPLICANT2PREMIUM"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblApp2Premium, "0.00")
    xmlResponseElem.appendChild xmlElement
    
    strTagName = "APPLICANT2IPT"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblApp2IPT, "0.00")
    xmlResponseElem.appendChild xmlElement
    
    strTagName = "MORTGAGEAMOUNT"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblMortgageCoverAmt, "0.00")
    xmlResponseElem.appendChild xmlElement

    strTagName = "MORTGAGECOVERFORAPPLICANT1"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblApp1MortgageCover, "0.00")
    xmlResponseElem.appendChild xmlElement
        
    strTagName = "MORTGAGECOVERFORAPPLICANT2"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(dblApp2MortgageCover, "0.00")
    xmlResponseElem.appendChild xmlElement
           
    CalculateCosts = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalculateCostsExit:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlResponseNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlPayProtProductElem = Nothing
    Set xmlElement = Nothing
    Set objGlobalParamDO = Nothing

    Exit Function

CalculateCostsVbErr:
    
    If objErrAssist.IsWarning = True Then
        
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateCosts = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalculateCostsExit

End Function

Private Sub GetRates(ByVal vstrPPProdRateNo As String, _
                     ByVal vstrChannelIndicator As String, _
                     ByVal vstrGender As String, _
                     ByVal vstrAge As String, _
                     ByRef rdblASURate, _
                     ByRef rdblASRate, _
                     ByRef rdblURate)
' header ----------------------------------------------------------------------------------
' description:  Gets the payment protection rates
'
' pass:         vstrChannelIndicator Channel indicator
'               vstrGender           Gender Combo ValueId for applicant
'               vstrAge              Applicant age
'               rdblASURate          Accident/Sickness/Unemployment Rate
'               rdblASRate           Accident/Sickness Rate
'               rdblURate            Unemployment rate
'
' return n/a
'------------------------------------------------------------------------------------------
On Error GoTo GetRatesVbErr
                     
    Const strFunctionName As String = "GetRates"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlRatesDoc As FreeThreadedDOMDocument40
    Dim xmlTableElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    Dim objPaymentProtectionDO As PaymentProtectionDO
    
    Dim strTagName As String
    Dim strTagValue As String
    
    If objContext Is Nothing Then
        Set objPaymentProtectionDO = New PaymentProtectionDO
    Else
        Set objPaymentProtectionDO = objContext.CreateInstance(App.Title & ".PaymentProtectionDO")
    End If
    
    Set xmlTableElement = xmlRequestDoc.createElement("PAYMENTPROTECTIONRATES")
    xmlRequestDoc.appendChild xmlTableElement
    
    strTagName = "PAYMENTPROTECTIONRATESNUMBER"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrPPProdRateNo
    xmlTableElement.appendChild xmlElement
    
    strTagName = "CHANNEL"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrChannelIndicator
    xmlTableElement.appendChild xmlElement
                     
    strTagName = "APPLICANTSGENDER"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrGender
    xmlTableElement.appendChild xmlElement
                     
    strTagName = "APPLICANTSAGE"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrAge
    xmlTableElement.appendChild xmlElement
                     
    Dim strResponse As String
    strResponse = objPaymentProtectionDO.GetRates(xmlRequestDoc.xml)
    Set xmlRatesDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    strTagName = "ASURATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRatesDoc.documentElement, strTagName)
    rdblASURate = objConvAssist.CSafeDbl(strTagValue)
                     
    strTagName = "ASRATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRatesDoc.documentElement, strTagName)
    rdblASRate = objConvAssist.CSafeDbl(strTagValue)
                     
    strTagName = "URATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRatesDoc.documentElement, strTagName)
    rdblURate = objConvAssist.CSafeDbl(strTagValue)
                                   
    Set xmlRequestDoc = Nothing
    Set xmlRatesDoc = Nothing
    Set xmlTableElement = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objPaymentProtectionDO = Nothing
                     
    Exit Sub
                     
GetRatesVbErr:
    
    Set xmlRequestDoc = Nothing
    Set xmlRatesDoc = Nothing
    Set xmlTableElement = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objPaymentProtectionDO = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
                     
End Sub

Private Sub ValidateCalculateCostsInput(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Validates that the correct parameters are present to calculate the
'               payment protection costs
'
' pass:         vstrXMLRequest CalcCosts Input xml Request
'
' return n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCalculateCostsInputVbErr
                     
    Const strFunctionName As String = "ValidateCalculateCostsInput"

    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlRequestElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    Dim xmlPayProtProductElem As IXMLDOMElement
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrCALC_NAME).Item(0)
        
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
        
    Dim strCoverType As String
    Dim intGenderApp1 As Integer
    Dim intAgeApp1 As Integer
    Dim intGenderApp2 As Integer
    Dim intAgeApp2 As Integer
    Dim dblMortgageCoverAmt As Double
    Dim dblMortgageCoverApp1 As Double
    Dim dblMortgageCoverApp2 As Double
    Dim strDistributionChannel As String
    
    Dim strTagName As String
    Dim strTagValue As String
    
    strTagName = "COVERTYPE"
    strCoverType = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    
    strTagName = "APPLICANT1GENDER"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    intGenderApp1 = objConvAssist.CSafeInt(strTagValue)
    
    strTagName = "APPLICANT1AGE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    intAgeApp1 = objConvAssist.CSafeInt(strTagValue)
    
    strTagName = "APPLICANT2GENDER"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    intGenderApp2 = objConvAssist.CSafeInt(strTagValue)
    
    strTagName = "APPLICANT2AGE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    intAgeApp2 = objConvAssist.CSafeInt(strTagValue)
    
    strTagName = "MORTGAGEAMOUNT"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMortgageCoverAmt = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "MORTGAGECOVERFORAPPLICANT1"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMortgageCoverApp1 = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "MORTGAGECOVERFORAPPLICANT2"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMortgageCoverApp2 = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "DISTRIBUTIONCHANNEL"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    strDistributionChannel = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    
    Set xmlPayProtProductElem = _
                    xmlRequestElem.getElementsByTagName("PAYMENTPROTECTIONPRODUCT").Item(0)
    
    Dim blnApplDetailsPresent As Boolean
    blnApplDetailsPresent = False
    
    If (intGenderApp1 <> 0 And intAgeApp1 <> 0) Or _
       (intGenderApp2 <> 0 And intAgeApp2 <> 0) Then
        blnApplDetailsPresent = True
    End If
    
    ' Check all parameters have been entered
    If xmlPayProtProductElem Is Nothing Or Len(strCoverType) = 0 Or _
       dblMortgageCoverAmt = 0 Or blnApplDetailsPresent = False Or _
       (dblMortgageCoverApp1 = 0 And dblMortgageCoverApp2 = 0) Or _
                                   Len(strDistributionChannel) = 0 Then
        
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 228
    
    End If
    
    If (dblMortgageCoverApp1 + dblMortgageCoverApp2) <> dblMortgageCoverAmt Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 231
    End If
                     
    Set xmlIn = Nothing
    Set xmlRequestElem = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlPayProtProductElem = Nothing

    Exit Sub
                     
ValidateCalculateCostsInputVbErr:
    
    Set xmlIn = Nothing
    Set xmlRequestElem = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlPayProtProductElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
                     
End Sub




