VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ChargesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ChargesBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Charges Business Object

'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    17/09/99    Created
'MCS    01/03/00    Not all one off costs appearing in GUI
'MCS    03/03/00    SYS0301
'APS    06/06/00    SYS0786 - Removed excess validation on ValidateCalcChargesInput
'PSC    07/07/00    SYS1153 - Correct CalculateArrangementFee to take into account the
'                             fee amount from the lender
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'BMIDS History
'------------------------------------------------------------------------------------------
'MO     12/6/02     CMWP2:BM022, BM023, BM024
'MO    2/7/2002     Modified for BMIDS00089, CalculateSpecificCosts
'MV     04/07/2002  BMIDS00156 - Modified CalculateArrangementFee()
'MO     16/07/2002  Modified for BMIDS00216, There was a bug in CalculateArrangementFee
'MO     13/08/2002  Made change for BMIDS00326, and commented out debug code.
'GD     28/05/2003  BM0198 Changes made to remove TT Fee calculations.
'DRC    28/06/2004  BMIDS763 -   Added in ProductSwitchFee
'HMA    30/09/2004  BMIDS902 Change CalculateArrangementFee to exclude Loan components that are to be ported.
'JD     09/11/2004  BMIDS943 Don't calc arrangement fee for any additional borrowing type applications
'------------------------------------------------------------------------------------------
'MARS History
'HMA    18/08/2005  MAR28 Changes to Arrangement fee and Valuation fee.
'                         Add Refund Amount to MortgageOneOffCosts
'GHun   12/10/2005  MAR46 Changed CalculateArrangementFee
' PJO   29/11/2005  MAR627 Missing valuation fee on multicomponent loans
' INR   01/12/2005  MAR717 Picking up wrong amount for ValuationFee
' JD    15/12/2005  MAR846 calculate arrangement fee for disposableKFI's
'GHun   22/12/2005  MAR846 changed CalculateArrangementFee
'------------------------------------------------------------------------------------------
'
'EPSOM History
'
'AW     09/05/2006  EP514 Allow for manual override of broker and valuation fees
'IK     16/06/2006  EP776 back out BM0198 - TT Fees required
'GHun   27/10/2006  EP2_9 Add additionalBorrowing and CreditLimitIncrease fees
'GHun   13/11/2006  EP2_9 Changed ONEOFFCOSTTYPES and GetIdentifierType to private to avoid future binary compatibility issues
'GHun   23/11/2006  EP2_9 Changed CalculateOneOffCosts to get NatureOfLoan from request XML and handle it being missing
'GHun   24/11/2006  EP2_19 Changed Calculate*SpecificCosts to handle FreeLegalFees
'GHun   24/11/2006  EP2_56 Add TransferOfEquity fee
'PSC    05/02/2007  EP2_1113 Change CalculateOneOffCosts to get LTV correctly for TOE
'PEdney 19/02/2007  EP2_1281 Do not calculate a valuation fee in one-off costs if no valuation is required.
'SR     01/03/2007  EP2_1272 Modified methods 'CalculateOneOffCosts',  'CalculateSpecificCosts', 'AppendCharges'
'PSC    02/03/2007  EP2_1347 Fee Refund processing
'SR     21/03/2007  EP2_1753 - modified AppendCharges - Add REBATEAMOUNT to MortgageOneOffCost
'SR     25/03/2007  EP2_1519 - modified CalculateArramgementFee
'SR     26/03/2007  EP2_1777 - modified AppendCharges - Append TransferOfEquityFee
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================
Private Const cstrCALC_NAME = "CHARGES"

 '=============================================
 'Variable Declaration Section
 '=============================================
Private Enum ONEOFFCOSTTYPES    'EP2_9 GHun changed to private
    octArrangementFee
    octPortingFee
    octValuationFee
    octReinspectionFee
    octAdministrationFee
    octStampDuty
    octSealingFee
    octDeedsReleaseFee

    octTTFee

    octLegalFee
    octMIGPremium
    octTotalIndemnity
    octRevaluation 'MO - CMWP2
    octProductSwitchFee 'DRC - BMIDS763
    octBrokerFee              'EP514
    octThirdPartyValuationFee 'EP514
    octAdditionalBorrowingFee   'EP2_9 GHun
    octCreditLimitIncreaseFee   'EP2_9 GHun
    octTransferOfEquityFee      'EP2_56 GHun
    octOther
End Enum

Private Type ONEOFFCOSTS
    lngArrangementFee As Long
    lngPortingFee As Long
    lngValuationFee As Long
    lngReinspectionFee As Long
    lngAdministrationFee As Long
    lngStampDuty As Long
    lngSealingFee As Long
    lngDeedsReleaseFee As Long
    'GD BM0198 START
    'IK EP776
    lngTTFee As Long
    'GD BM0198 END
    lngLegalFee As Long
    lngMIGPremium As Long
    lngTotalIndemnity As Long
    'MO - CMWP2 - Start
    'dblMtgArrangementFee As Double
    'dblMtgArrangeFeePercent As Double
    'dblMtgArrangeFeePercentMax As Double
    'MO - CMWP2 - End
    dblNewLoanAmount As Double
    'MO - CMWP2 - Start
    lngRevaluationFee As Long
    lngProductSwitchFee As Long 'DRC - BMIDS763
    dblValuationFeeAmount As Double
    dblValuationFeePC As Double
    dblValuationFeePCMax As Double
    dblValuationFeePCMin As Double
    'MO - CMWP2 - End
    lngRefundAmount As Long         ' MAR28
    lngBrokerFeeAmount As Long      ' EP514
    lngThirdPartyValuationFee As Long ' EP514
    lngAdditionalBorrowingFee As Long   'EP2_9 GHun
    lngCreditLimitIncreaseFee As Long   'EP2_9 GHun
    lngTransferOfEquityFee As Long      'EP2_56 GHun
End Type

Public Function CalcCharges(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the one-off costs to be applied to an illustration
'               or quotation for a loan. The Loan Components need not be present for
'               mortgage calculator
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Request Format:
'      <CHARGES>
'          <CONTEXT>Value</CONTEXT>
'          <AMOUNTREQUESTED>Value</AMOUNTREQUESTED>
'          <PURCHASEPRICE>Value</PURCHASEPRICE>
'          <LTV>Value</LTV>
'          <TYPEOFAPPLICATION>Value</TYPEOFAPPLICATION>
'          <TOTALLOANAMOUNT>Value</TOTALLOANAMOUNT>
'          <LOANSNOTREDEEMED>Value</LOANSNOTREDEEMED>
'          <LOANCOMPONENTLIST>
'              <LOANCOMPONENT>
'                <MORTGAGEPRODUCTDETAILS>
'                ...
'                </MORTGAGEPRODUCTDETAILS>
'                  ...
'              </LOANCOMPONENT>
'              ...
'          <LOANCOMPONENTLIST>
'      </CHARGES>
'
' Change history ---------------------------------------------------------------------------
' MO    6/6/2002    Modified for CMWP2, requirements BM022, BM023 and BM024
' MO    13/08/2002  Modified for BMIDS00326, commented out debug code
'------------------------------------------------------------------------------------------
On Error GoTo CalcChargesVbErr
        
    Const strFunctionName As String = "CalcCharges"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlIn As FreeThreadedDOMDocument40
    'AW 10/05/2006 EP514
    Dim xmlFactFindDoc As FreeThreadedDOMDocument40

    Dim objXmlAssist As New XMLAssist
    
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    'AW 10/05/2006 EP514 - Start
    Dim oocFees As ONEOFFCOSTS
    'Dim strCharges As String
    
    ValidateCalcChargesInput objXmlIn
    
    Set xmlFactFindDoc = New FreeThreadedDOMDocument40
    xmlFactFindDoc.async = False
    xmlFactFindDoc.loadXML GetOverrideFees(objXmlIn.documentElement)
    
    CalculateOneOffCosts objXmlIn, xmlFactFindDoc, oocFees
    'AW 10/05/2006 EP514 - End
    
    AppendCharges objXmlIn, objXmlResponseElem, oocFees

    CalcCharges = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalcChargesExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objErrAssist = Nothing
    Set objXmlNode = Nothing
    Set objXmlIn = Nothing
    Set xmlFactFindDoc = Nothing
    Set objXmlAssist = Nothing

    Exit Function

CalcChargesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CalcCharges = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalcChargesExit

End Function


Private Sub ValidateCalcChargesInput(ByVal vobjXmlRequest As FreeThreadedDOMDocument40)

' header ----------------------------------------------------------------------------------
' description:  Validates the input request for CalcCharges
'
' pass:         vobjXmlRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCalcChargesInputVbErr
    
    Const strFunctionName As String = "ValidateCalcChargesInput"
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlRequestElem As IXMLDOMElement
    
    Dim strContext As String
    Dim strAmountRequested As String
    Dim strPurchasePrice As String
    'GD BM0198 START
    'Dim strLTV As String
    'GD BM0198 END
    Dim strTypeOfApplication As String
    Dim strTotalLoanAmount As String
    Dim strLoansNotRedeemed As String
    Dim strTagName As String
    Dim strDetails As String
        
    Set xmlRequestElem = vobjXmlRequest.getElementsByTagName(cstrCALC_NAME).Item(0)
        
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
   
    strTagName = "CONTEXT"
    strContext = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    
    strTagName = "AMOUNTREQUESTED"
    strAmountRequested = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    
    strTagName = "PURCHASEPRICE"
    strPurchasePrice = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    'GD BM0198 START
    'strTagName = "LTV"
    'strLTV = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    'GD BM0198 END
    strTagName = "TYPEOFAPPLICATION"
    strTypeOfApplication = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    
    strTagName = "TOTALLOANAMOUNT"
    strTotalLoanAmount = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    
    strTagName = "LOANSNOTREDEEMED"
    strLoansNotRedeemed = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, strTagName)
    'GD BM0198 START
    'If Len(strContext) = 0 Or Len(strAmountRequested) = 0 Or Len(strLTV) = 0 Or _
       Len(strTypeOfApplication) = 0 Or Len(strTotalLoanAmount) = 0 Or _
       Len(strLoansNotRedeemed) = 0 Then
     If Len(strContext) = 0 Or Len(strAmountRequested) = 0 Or _
       Len(strTypeOfApplication) = 0 Or Len(strTotalLoanAmount) = 0 Or _
       Len(strLoansNotRedeemed) = 0 Then
    'GD BM0198 END
       strDetails = "CONTEXT = " & strContext
       strDetails = strDetails & "; AMOUNTREQUESTED = " & strAmountRequested
       strDetails = strDetails & "; PURCHASEPRICE = " & strPurchasePrice
       'GD BM0198 START
       'strDetails = strDetails & "; LTV = " & strLTV
       'GD BM0198 END
       strDetails = strDetails & "; TYPEOFAPPLICATION = " & strTypeOfApplication
       strDetails = strDetails & "; TOTALLOANAMOUNT = " & strTotalLoanAmount
       strDetails = strDetails & "; LOANSNOTREDEEMED = " & strLoansNotRedeemed
       
       objErrAssist.ThrowError TypeName(Me), _
                               strFunctionName, _
                               oeInvalidParameter, _
                               strDetails
    End If
    
    'APS SYS0786 06/06/00 - Removed excess validation on ValidateCalcChargesInput
    'If vobjXmlRequest.getElementsByTagName("MORTGAGEPRODUCTDETAILS").length = 0 Then
    '       objErrAssist.ThrowError TypeName(Me), _
    '                               strFunctionName, _
    '                               oeMissingPrimaryTag, _
    '                               "MORTGAGEPRODUCTDETAILS tag is missing"
    'End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Sub

ValidateCalcChargesInputVbErr:
    
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing

    Err.Raise Err.number, Err.Source, Err.Description

End Sub


Private Sub CalculateOneOffCosts(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                                 ByVal vobjXmlFactFind As FreeThreadedDOMDocument40, _
                                 ByRef roocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Calculates the one off costs
'
' pass:         vobjXmlRequest  xml Document containing the request
'               roocFees        fees structure
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CalculateOneOffCostsVbErr

    Const strFunctionName As String = "CalculateOneOffCosts"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    'AW 10/05/2006  EP514
    Dim xmlNode As IXMLDOMNode
    Dim blnUsingValuationOverride As Boolean
    'AW 10/05/2006  EP514 - End
    Dim objMtgLenderElem As IXMLDOMElement
    Dim objComboDO As ComboDO
    
    Dim strContext As String
    Dim blnPortingFeeNeeded As Boolean
    
    'EP2_9 GHun
    Dim blnIsAdditionalBorrowing    As Boolean
    Dim strNatureOfLoan             As String
    Dim blnChargeValFee             As Boolean
    Dim lngMaxMonths                As Long
    Dim strLastValuationDate        As String
    Dim dteLastValuationDate        As Date
    'EP2_9 End
    
     'EP2_56 GHun
    Dim blnIsTransferOfEquity As Boolean
    Dim objGlobalParamDO As GlobalParameterDO
     'EP2_56 End
    Dim xmlIntroducerFeeListNode As IXMLDOMNode  'SR 28/02/2007 : EP2_1272
    Dim strTPVFeeType As String 'SR 28/02/2007 : EP2_1272
 
  
    strContext = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, "CONTEXT")
    
    If StrComp(strContext, "Mortgage Calculator", vbTextCompare) = 0 Then
        CalculateMortCalcSpecificCosts vobjXmlRequest, roocFees
    Else
        CalculateSpecificCosts vobjXmlRequest, roocFees, blnPortingFeeNeeded
    End If
            
    Dim strApplicationType As String
    
    strApplicationType = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, _
                                                  "TYPEOFAPPLICATION")
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' Get the stamp duty if this is a new loan
    If objComboDO.IsNewLoan(strApplicationType) = True Then
        Dim strPurchasePrice As String
        strPurchasePrice = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, _
                                                    "PURCHASEPRICE")
        GetStampDuty strPurchasePrice, roocFees
    End If
        
    ' Get sealing fee and deeds release fee
    Dim strTagValue As String
    
    Set objMtgLenderElem = vobjXmlRequest.getElementsByTagName("MORTGAGELENDER").Item(0)
    
    If Not objMtgLenderElem Is Nothing Then
        strTagValue = objXmlAssist.GetTagValue(objMtgLenderElem, "SEALINGFEE")
        roocFees.lngSealingFee = objConvAssist.CSafeLng(strTagValue)
        strTagValue = objXmlAssist.GetTagValue(objMtgLenderElem, "DEEDSRELEASEFEE")
        roocFees.lngDeedsReleaseFee = objConvAssist.CSafeLng(strTagValue)
    End If
    
    'MAR28 Initialise the Porting and Arrangement fees to 0
    roocFees.lngPortingFee = 0
    roocFees.lngArrangementFee = 0
    
    ' If we don't need to charge a porting fee then set the fee to zero and
    ' calculate the arrangement fee
    If blnPortingFeeNeeded = False Then
        roocFees.lngPortingFee = 0
        roocFees.lngArrangementFee = 0   ' MAR28

        ' MO - 13/06/2002 - CMWP2:BM022 - Mortgage calculator is to be removed
        ' from the BMIDS Omiga4, no requirement has been made for the new
        ' CalculateArrangementFee to calculate a mortgage calculator fee, so it
        ' doesnt! I have inserted this line of code so that if a mortgage calculator
        ' quote is requested (although it shouldnt be cause it isnt going to be available)
        ' it wont generate an error, it just wont calculate an arrangement fee.
        'JD BMIDS943 Don't calc arrangement fee for any additional borrowing type applications
        
        'MAR28 Calculate Arrangement Fee for Additional Borrowing
            
'        If StrComp(strContext, "Mortgage Calculator", vbTextCompare) <> 0 And _
'           objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "F") = False And _
'           objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "M") = False _
'        Then
'            CalculateArrangementFee vobjXmlRequest, roocFees
'        End If
        'MAR846 calculate arrangement fee for disposableKFI's
        Dim strKFI As String
        strKFI = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, _
                                                    "DISPOSABLEKFI")
        If ((StrComp(strContext, "Mortgage Calculator", vbTextCompare) <> 0) Or _
             strKFI = "1") Then
            CalculateArrangementFee vobjXmlRequest, roocFees
        End If
    End If
    
    blnIsAdditionalBorrowing = objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "ABO")   'EP2_9 GHun
    blnIsTransferOfEquity = objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "TOE")      'EP2_56 GHun
    
    'MAR28 Initialise the Valuation fee to 0
    roocFees.lngValuationFee = 0

    'MAR28 Do not calculate Valuation fee for Product Switch mortgages
    If (objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "PSW") = False) Then
    
        'AW 10/05/2006  EP514
        Set xmlNode = vobjXmlFactFind.selectSingleNode("//APPLICATIONFACTFIND")
        If Not xmlNode Is Nothing Then
        
            If Not xmlNode.Attributes.getNamedItem("PACKAGEDVALUATIONFEE") Is Nothing Then
            
                roocFees.lngThirdPartyValuationFee = objConvAssist.CSafeLng(xmlNode.Attributes.getNamedItem("PACKAGEDVALUATIONFEE").nodeValue)
                
                If roocFees.lngThirdPartyValuationFee > 0 Then
                    blnUsingValuationOverride = True
                End If
            End If
        End If
    
        If Not blnUsingValuationOverride Then
            'EP2_9 GHun
            blnChargeValFee = True
            If blnIsAdditionalBorrowing Then
                strNatureOfLoan = xmlGetNodeText(vobjXmlRequest.documentElement, "NATUREOFLOAN")
                If Len(strNatureOfLoan) > 0 Then
                    On Error Resume Next
                    lngMaxMonths = CSafeLng(objComboDO.GetFirstComboValidation("AddBorrowChargeValFeeCategory", strNatureOfLoan))
                    On Error GoTo CalculateOneOffCostsVbErr
                    If Err.number <> 0 Then
                        Err.Clear
                        blnChargeValFee = False
                    Else
                        
                        If lngMaxMonths = 0 Then
                            blnChargeValFee = True
                        ElseIf lngMaxMonths > 0 Then
                            strLastValuationDate = xmlGetNodeText(vobjXmlRequest.documentElement, "LASTVALUATIONDATE")
                            If Len(strLastValuationDate) > 0 Then
                                dteLastValuationDate = CSafeDate(strLastValuationDate)
                                If MonthDiff(dteLastValuationDate, Now()) > lngMaxMonths Then
                                    blnChargeValFee = True
                                Else
                                    blnChargeValFee = False
                                End If
                            Else
                                blnChargeValFee = True
                            End If
                            
                        End If
                    End If
                End If
            'EP2_56 GHun
            ElseIf blnIsTransferOfEquity Then
                If objContext Is Nothing Then
                    Set objGlobalParamDO = New GlobalParameterDO
                Else
                    Set objGlobalParamDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
                End If
                Dim dblLTV As Double
                ' PSC 05/02/2007 EP2_1113
                dblLTV = xmlGetMandatoryNodeAsDouble(vobjXmlRequest, "CHARGES/LTV")
                If dblLTV > objGlobalParamDO.GetCurrentParameterByType("TOEValFeeRequiredMaxLTV", "PERCENTAGE") Then
                    blnChargeValFee = True
                Else
                    blnChargeValFee = False
                End If
            'EP2_56 End
            Else  'SR EP2_1272 - if PackagedValuationFee is passed in, do not calculate valuation fee
                strTPVFeeType = objComboDO.GetFirstComboValueId("OneOffCost", "TPV")
                Set xmlIntroducerFeeListNode = vobjXmlRequest.selectSingleNode("//INTRODUCERFEELIST")
                If Not xmlIntroducerFeeListNode Is Nothing Then
                    If xmlIntroducerFeeListNode.selectNodes("INTRODUCERFEE[FEETYPE='" & strTPVFeeType & "']").length > 0 Then
                        blnChargeValFee = False
                    End If
                End If
            End If 'SR EP2_1272 - end
            
            If blnChargeValFee Then
                CalculateValuationFee vobjXmlRequest, roocFees
            End If
            'EP2_9 End
        End If
        'AW 10/05/2006  EP514 - End
    End If
    
    'EP2_9 GHun
    If blnIsAdditionalBorrowing Then
        CalculateAdditionalBorrowingFee vobjXmlRequest, roocFees
    End If
    
    If (objComboDO.IsItemInValidation("TypeOfMortgage", strApplicationType, "CLI") = True) Then
        CalculateCreditLimitIncreaseFee vobjXmlRequest, roocFees
    End If
    'EP2_9 End
    
    'EP2_56 GHun
    If blnIsTransferOfEquity Then
        GetTransferOfEquityFee vobjXmlRequest, roocFees
    End If
    'EP2_56 End
    
    CalculateMortgageIndemnity vobjXmlRequest, roocFees
        
    'AW 10/05/2006  EP514
    Set xmlNode = vobjXmlFactFind.selectSingleNode("//APPLICATIONFACTFIND")
    If Not xmlNode Is Nothing Then
        If Not xmlNode.Attributes.getNamedItem("ADDITIONALBROKERFEE") Is Nothing Then
            roocFees.lngBrokerFeeAmount = objConvAssist.CSafeLng(xmlNode.Attributes.getNamedItem("ADDITIONALBROKERFEE").nodeValue)
        End If
    End If
    'AW 10/05/2006  EP514 - End
     
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objComboDO = Nothing
    Set objMtgLenderElem = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlIntroducerFeeListNode = Nothing
    
    Exit Sub

CalculateOneOffCostsVbErr:
    
    Set objXmlAssist = Nothing
    Set objComboDO = Nothing
    Set objMtgLenderElem = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlIntroducerFeeListNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub CalculateValuationFee(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                                    ByRef roocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Calculates the Valuation/Revaluation/Reinspection Fees
'
' pass:         vobjXmlRequest  xml Document containing the request
'               roocFees        fees structure
'
' return:       n/a
'------------------------------------------------------------------------------------------
'MO 06/06/2002 CMWP2 - Start - New sub to calculate valuation fees
On Error GoTo CalculateValuationFeeVbErr

    Const strFunctionName As String = "CalculateValuationFee"
    Dim strValuationType As String
    Dim dblPurchasePrice As String
    Dim dblTempValuationFee As Double
       
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objConvAssist As New ConvertAssist
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objValuationElem As IXMLDOMElement
    Dim objComboDO As ComboDO
    Dim objValuationList As IXMLDOMNodeList   ' MAR28
    Dim iNumberOfValuations As Integer        ' MAR28
    Dim lngRefundAmount As Long               ' MAR28
    Dim bCalculateValuationRefund As Boolean  ' MAR28
    Dim objGlobalParamDO As GlobalParameterDO ' MAR28
    Dim bNonSC1Exists As Boolean           ' PJO 29/11/2005 MAR627
    
    ' PSC 02/03/2007 EP2_1347 - Start
    Dim xmlProductList As IXMLDOMNodeList
    Dim blnIsRefundOfValuation As Boolean
    ' PSC 02/03/2007 EP2_1347 - End

    Dim strTagValue As String
    
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objGlobalParamDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
          
    Set objValuationList = vobjXmlRequest.getElementsByTagName("VALUATIONFEE")
    
    'MAR28 If there are 2 Valuation fees, the one with type SC1 will be used to get the Refund Amount
    iNumberOfValuations = objValuationList.length
    
    lngRefundAmount = 0
    
    bNonSC1Exists = False    ' PJO 29/11/2005 MAR627
    
    'MAR28 Get the Refund Amount from the Valuation Fee with type SC1
    'PJO 29/11/2005 MAR627 - At the same time check for the existence of a non-SC1
    ' valuation type. We need to know this because an SC1 type may be for refund only. If
    ' we have a multicomponent loan where all the valuations are type SC1 they will be
    ' ignored if the software belives that SC1 types are for a refund
    
    For Each objValuationElem In objValuationList
        'Get the valuation type
        strTagValue = objXmlAssist.GetTagValue(objValuationElem, "TYPEOFVALUATION")
        strValuationType = objConvAssist.CSafeLng(strTagValue)

        If (objComboDO.IsItemInValidation("ValuationType", strValuationType, "SC1") = True) Then
            strTagValue = objXmlAssist.GetTagValue(objValuationElem, "AMOUNT")
            lngRefundAmount = objConvAssist.CSafeLng(strTagValue)
        Else
            bNonSC1Exists = True        ' PJO 29/11/2005 MAR627
        End If
    Next
        
    
    If (iNumberOfValuations > 0) Then

        For Each objValuationElem In objValuationList
    
            'Get the valuation type
            strTagValue = objXmlAssist.GetTagValue(objValuationElem, "TYPEOFVALUATION")
            strValuationType = objConvAssist.CSafeLng(strTagValue)
    
            ' Do not use the one with type SC1 if there is more than one Valuation present unles
            ' (PJO 29/11/2005 MAR627) they are all SC1 types
            If ((iNumberOfValuations = 1) Or _
                 (bNonSC1Exists = False) Or _
                 (objComboDO.IsItemInValidation("ValuationType", strValuationType, "SC1") <> True)) Then
         
                dblPurchasePrice = objConvAssist.CSafeDbl(objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, "PURCHASEPRICE"))
        
                If roocFees.dblValuationFeePC <> 0 Then
                    dblTempValuationFee = dblPurchasePrice * roocFees.dblValuationFeePC / 100
            
                    If roocFees.dblValuationFeePCMax > 0 Then
                        If dblTempValuationFee > roocFees.dblValuationFeePCMax Then
                            dblTempValuationFee = roocFees.dblValuationFeePCMax
                        End If
                    End If
            
                    If dblTempValuationFee < roocFees.dblValuationFeePCMin Then
                        dblTempValuationFee = roocFees.dblValuationFeePCMin
                    End If
            
                End If
        
                If roocFees.dblValuationFeeAmount > dblTempValuationFee Then
                    dblTempValuationFee = roocFees.dblValuationFeeAmount
                End If
        
                'get the valuation type
                strTagValue = objXmlAssist.GetTagValue(objValuationElem, "TYPEOFVALUATION")
                strValuationType = objConvAssist.CSafeLng(strTagValue)
        
                'check the valuation fee type and output the correct fee type
                'Is this a reinspection?
                If objComboDO.IsItemInValidation("ValuationType", strValuationType, "R") = True Then
                    roocFees.lngReinspectionFee = Fix(dblTempValuationFee)
                Else
                    'Is this a revalution
                    If objComboDO.IsItemInValidation("ValuationType", strValuationType, "REV") = True Then
                        roocFees.lngRevaluationFee = Fix(dblTempValuationFee)
                    Else
                        'No its a normal valuation
                        roocFees.lngValuationFee = Fix(dblTempValuationFee)
                    End If
                End If
            End If
        Next
        
        'Set up the Refund Amount
        bCalculateValuationRefund = objGlobalParamDO.GetCurrentParameterByType("CalculateValuationRefund", "BOOLEAN")
    
        If (bCalculateValuationRefund = True) Then
            roocFees.lngRefundAmount = lngRefundAmount
        Else
            ' PSC 02/03/2007 EP2_1347 - Start
            Set xmlProductList = vobjXmlRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCT[REFUNDOFVALUATION='1']")
            
            blnIsRefundOfValuation = xmlProductList.length > 0
            
            If blnIsRefundOfValuation Then
                roocFees.lngRefundAmount = roocFees.lngValuationFee
            Else
                roocFees.lngRefundAmount = 0
            End If
            ' PSC 02/03/2007 EP2_1347 - End
        End If
        
    End If
    
    Set objConvAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objValuationElem = Nothing
    Set xmlProductList = Nothing   ' PSC 02/03/2007 EP2_1347
    
    Exit Sub
    
CalculateValuationFeeVbErr:
    
    Set objConvAssist = Nothing
    Set objXmlAssist = Nothing
    Set objValuationElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description
'MO 06/06/2002 CMWP2 - End
End Sub

Private Sub CalculateArrangementFee(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                                    ByRef roocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Calculates the Arrangement Fee
'
' pass:         vobjXmlRequest  xml Document containing the request
'               roocFees        fees structure
'
' return:       n/a
' MO - 12/06/2002 : CMWP2, BM022 - Complete rewrite to cope with new arrangement fee
'                   percentage calculations and summed per product.
'------------------------------------------------------------------------------------------
On Error GoTo CalculateArrangementFeeVbErr

    Const strFunctionName As String = "CalculateArrangementFee"
    
    Dim objConvAssist As ConvertAssist
    Set objConvAssist = New ConvertAssist
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim xmlLoanComponentList As IXMLDOMNode
    Dim xmlMortgageProductCodeList As IXMLDOMNode
    Dim xmlMortgageProductCode As IXMLDOMNode
    Dim xmlMortgageProductDOM As FreeThreadedDOMDocument40
    Dim xmlLoanComponentsForMortgageProduct As IXMLDOMNodeList
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlMortgageProductDetails As IXMLDOMNode
    Dim xmlMortgageLenderDetails As IXMLDOMNode
    
    Dim dblProductArrangementFee As Double
    
    'MO - 12/06/2002 - CMWP2, BM022 - removed
    'Dim objProductElem As IXMLDOMElement
    'Dim objMtgLenderElem As IXMLDOMElement
    
    Dim dblProductArrangementFeeAmount As Double
    Dim dblProductArrangementFeePercent As Double
    Dim dblProductArrangementFeePercentMax As Double
    Dim dblProductArrangementFeePercentMin As Double
    Dim blnProductArrangementFeePercentInd As Boolean 'MO     16/07/2002  BMIDS00216
    
    Dim dblLenderArrangementFeeAmount As Double
    Dim dblLenderArrangementFeePercent As Double
    Dim dblLenderArrangementFeePercentMax As Double
    Dim dblLenderArrangementFeePercentMin As Double
    Dim blnLenderArrangementFeePercentInd As Boolean 'MO     16/07/2002  BMIDS00216
    
    'Dim strTagValue As String
    Dim strSearch As String                            'BMIDS902
    'SR EP2_1519 - Start
    Dim strTypeOfApplication As String
    Dim blnIsABO As Boolean, blnIsTOE As Boolean, blnIsCLI As Boolean, blnIsPSW As Boolean, blnIsNP As Boolean
    Dim blnGPChargeArFeeForPort As Boolean
    'SR EP2_1519 - End
    
    'SR EP2_1519 - Start
    strTypeOfApplication = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, "TYPEOFAPPLICATION")
    blnIsABO = IsValidationType("TypeOfMortgage", strTypeOfApplication, "ABO")
    If Not blnIsABO Then
        blnIsTOE = IsValidationType("TypeOfMortgage", strTypeOfApplication, "TOE")
    End If
    If (Not blnIsABO) And (Not blnIsTOE) Then
        blnIsCLI = IsValidationType("TypeOfMortgage", strTypeOfApplication, "CLI")
    End If
    If (Not blnIsABO) And (Not blnIsTOE) And (Not blnIsCLI) Then
        blnIsPSW = IsValidationType("TypeOfMortgage", strTypeOfApplication, "PSW")
    End If
    If (Not blnIsABO) And (Not blnIsTOE) And (Not blnIsCLI) And (Not blnIsPSW) Then
        blnIsNP = IsValidationType("TypeOfMortgage", strTypeOfApplication, "NP")
    End If
    
    blnGPChargeArFeeForPort = GetMandatoryGlobalParamBoolean("CMChargeARRFeeForPort")
    'SR EP2_1519 - End
    
    If (Not blnIsTOE) And (Not blnIsABO) And (Not blnIsCLI) And _
       (Not blnIsNP Or (blnIsNP And blnGPChargeArFeeForPort)) Then 'SR EP2_1519 - Start
        'get the loan components list
        Set xmlLoanComponentList = objXmlAssist.GetNode(vobjXmlRequest.documentElement, ".//LOANCOMPONENTLIST", True)
        
        Set xmlMortgageProductDOM = New FreeThreadedDOMDocument40
        xmlMortgageProductDOM.validateOnParse = False
        xmlMortgageProductDOM.setProperty "NewParser", True
        
        'create an element which will be the parent node for my mortgageproductcodes
        Set xmlMortgageProductCodeList = xmlMortgageProductDOM.createElement("MORTGAGEPRODUCTCODELIST")
        
        'get the mortgage product codes which are in this calculation
        For Each xmlMortgageProductCode In xmlLoanComponentList.selectNodes("LOANCOMPONENT/MORTGAGEPRODUCTCODE")
            'append the mortgageproduct codes to the mortgageproductcodelistfactfind
            xmlMortgageProductCodeList.appendChild xmlMortgageProductCode.cloneNode(True)
        Next
        
        If xmlMortgageProductCodeList.hasChildNodes = False Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "MORTGAGEPRODUCTCODE tag missing"
        End If
        
        
        'remove the duplicate mortgage product codes from the the list of mortgageproducts
        objXmlAssist.RemoveDuplicates xmlMortgageProductCodeList
                
        'loop through the mortgage product codes
        For Each xmlMortgageProductCode In xmlMortgageProductCodeList.childNodes
                
            dblProductArrangementFee = 0
            'MO     16/07/2002  BMIDS00216 - Start
            blnProductArrangementFeePercentInd = False
            blnLenderArrangementFeePercentInd = False
            'MO     16/07/2002  BMIDS00216 - End
            
            'BMIDS902 Get all the loan components that are for this mortgage product code and are NOT ported
            'MAR28    - and where the ProductSwitchRetained indicator = false
            'MAR46 GHun
            'MAR846 GHun handle missing attributes
            'SR 25/03/2007: EP2_1519
            strSearch = "LOANCOMPONENT[MORTGAGEPRODUCTCODE='" & xmlMortgageProductCode.Text & "'"
            If blnIsPSW Then
                 strSearch = strSearch & "and (not(PRODUCTSWITCHRETAINPRODUCTIND) or PRODUCTSWITCHRETAINPRODUCTIND!='1')]"
            Else
                 strSearch = strSearch & "]"
            End If
            'SR 25/03/2007: EP2_1519 - End
                                       
            Set xmlLoanComponentsForMortgageProduct = xmlLoanComponentList.selectNodes(strSearch)
            
            ' loop through the loan components
            For Each xmlLoanComponent In xmlLoanComponentsForMortgageProduct
                
                dblProductArrangementFeeAmount = 0
                dblProductArrangementFeePercent = 0
                dblLenderArrangementFeeAmount = 0
                dblLenderArrangementFeePercent = 0
                
                'get the mortgage product details out of the loan component
                Set xmlMortgageProductDetails = objXmlAssist.GetNode(xmlLoanComponent, "MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCT", True)
                
                dblProductArrangementFeeAmount = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageProductDetails, "MPARRANGEMENTFEEAMOUNT", True).Text)
                
                'has this product got a fixed arrangement fee?  If it has exit the loop as we dont want to calculate any more
                If dblProductArrangementFeeAmount > 0 Then
                    dblProductArrangementFee = dblProductArrangementFeeAmount
                    Exit For
                End If
                
                dblProductArrangementFeePercent = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageProductDetails, "MPARRANGEMENTFEEPERCENT", True).Text)
                
                'has this product got a percentage based arrangement fee?
                If dblProductArrangementFeePercent > 0 Then
                    
                    blnProductArrangementFeePercentInd = True
                    
                    dblProductArrangementFeePercentMin = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageProductDetails, "MPARRANGEMENTFEEPERCENTMIN", True).Text)
                    dblProductArrangementFeePercentMax = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageProductDetails, "MPARRANGEMENTFEEPERCENTMAX", True).Text)
                    
                    'calculate the fee and add it to the total for this product
                    dblProductArrangementFee = dblProductArrangementFee + objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlLoanComponent, "LOANAMOUNT", True).Text) * dblProductArrangementFeePercent / 100
                                  
                End If
                
                Set xmlMortgageProductDetails = Nothing
                
                'if there isnt a product arrangement fee, check if there is a lender arrangement fee
                If dblProductArrangementFeeAmount = 0 And dblProductArrangementFeePercent = 0 Then
                    
                    'get the lender product details out of the loan component
                    Set xmlMortgageLenderDetails = objXmlAssist.GetNode(xmlLoanComponent, "MORTGAGEPRODUCTDETAILS/MORTGAGELENDER", True)
                        
                    dblLenderArrangementFeeAmount = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageLenderDetails, "MLARRANGEMENTFEEAMOUNT", True).Text)
                        
                    'has this lender got a fixed arrangement fee?  If it has exit the loop as we dont want to calculate any more
                    If dblLenderArrangementFeeAmount > 0 Then
                        dblProductArrangementFee = dblLenderArrangementFeeAmount
                        Exit For
                    End If
                    
                    dblLenderArrangementFeePercent = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageLenderDetails, "MLARRANGEMENTFEEPERCENT", True).Text)
                    
                    'has this lender got a percentage based arrangement fee?
                    If dblLenderArrangementFeePercent > 0 Then
                        
                        blnLenderArrangementFeePercentInd = True
                        
                        dblLenderArrangementFeePercentMin = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageLenderDetails, "MLARRANGEMENTFEEPERCENTMIN", True).Text)
                        dblLenderArrangementFeePercentMax = objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlMortgageLenderDetails, "MLARRANGEMENTFEEPERCENTMAX", True).Text)
                        
                        'calculate the fee and add it to the total for this product
                        dblProductArrangementFee = dblProductArrangementFee + objConvAssist.CSafeDbl(objXmlAssist.GetNode(xmlLoanComponent, "LOANAMOUNT", True).Text) * dblLenderArrangementFeePercent / 100
                        
                    End If
                    
                    Set xmlMortgageLenderDetails = Nothing
                    
                End If
            
            Next
            
            'MO     16/07/2002  BMIDS00216 - Start
            'Arrangement fees calculate as a percentage are based per product, not loan component,
            '   so minimum and maximum values are calculated here
            'Has this been calculated as a product based percentage arrangement fee?
            If blnProductArrangementFeePercentInd = True Then
                'MO 12/06/2002 AQR-BMIDS000333 - Changed so if a percentage max isnt specified the fee isnt set to 0 - Start
                If dblProductArrangementFeePercentMax > 0 Then
                    If dblProductArrangementFee > dblProductArrangementFeePercentMax Then
                        dblProductArrangementFee = dblProductArrangementFeePercentMax
                    End If
                End If
                
                If dblProductArrangementFee < dblProductArrangementFeePercentMin Then
                    dblProductArrangementFee = dblProductArrangementFeePercentMin
                End If
            Else
                'Has this been calculated as a lender based percentage arrangement fee?
                If blnLenderArrangementFeePercentInd = True Then
                    'MO 12/06/2002 AQR-BMIDS000333 - Changed so if a percentage max isnt specified the fee isnt set to 0 - Start
                    If dblLenderArrangementFeePercentMax > 0 Then
                        If dblProductArrangementFee > dblLenderArrangementFeePercentMax Then
                            dblProductArrangementFee = dblLenderArrangementFeePercentMax
                        End If
                    End If
                    
                    If dblProductArrangementFee < dblLenderArrangementFeePercentMin Then
                        dblProductArrangementFee = dblLenderArrangementFeePercentMin
                    End If
                End If
            End If
            'MO     16/07/2002  BMIDS00216 - End
            
            Set xmlLoanComponentsForMortgageProduct = Nothing
            
            'MAR28. Set the product arrangement fee to be the largest of the arrangements fees calculated.
            If (Fix(dblProductArrangementFee) > roocFees.lngArrangementFee) Then
                roocFees.lngArrangementFee = Fix(dblProductArrangementFee)
            End If
          
        Next
        'SR 25/03/2007  : EP2_1519 - Removed comments . See V11 for deleted lines.
    End If ' SR EP2_1519 - End : Application is not ABO, TOE, CLI
    
    Set objConvAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlMortgageProductCodeList = Nothing
    Set xmlMortgageProductCode = Nothing
    Set xmlMortgageProductDOM = Nothing
    Set xmlLoanComponentsForMortgageProduct = Nothing
    Set xmlMortgageProductDetails = Nothing
    Set xmlMortgageLenderDetails = Nothing
    
    'MO - 12/06/2002 - CMWP2, BM022 - removed
    'Set objProductElem = Nothing
    'Set objMtgLenderElem = Nothing

    Exit Sub
    
CalculateArrangementFeeVbErr:
    
    Set objConvAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlMortgageProductCodeList = Nothing
    Set xmlMortgageProductCode = Nothing
    Set xmlMortgageProductDOM = Nothing
    Set xmlLoanComponentsForMortgageProduct = Nothing
    
    'MO - 12/06/2002 - CMWP2, BM022 - removed
    'Set objProductElem = Nothing
    'Set objMtgLenderElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub CalculateMortgageIndemnity(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                                       ByRef roocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Calculates the Mortgage Indemnity Fee
'
' pass:         vobjXmlRequest  xml Document containing the request
'               roocFees        fees structure
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CalculateMortgageIndemnityVbErr
    
    Const strFunctionName As String = "CalculateMortgageIndemnity"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objMigBO As MIGBO
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objXmlMIGDoc As FreeThreadedDOMDocument40
    Dim objXmlRequestElem As IXMLDOMElement
    Dim objXmlMIGNode As IXMLDOMNode
    Dim objMIGRequestDoc As New FreeThreadedDOMDocument40
    objMIGRequestDoc.validateOnParse = False
    objMIGRequestDoc.setProperty "NewParser", True
    
    If objContext Is Nothing Then
        Set objMigBO = New MIGBO
    Else
        Set objMigBO = objContext.CreateInstance(App.Title & ".MIGBO")
    End If
    
    Set objXmlRequestElem = vobjXmlRequest.getElementsByTagName(cstrCALC_NAME).Item(0)
    Set objXmlMIGNode = objXmlRequestElem.cloneNode(True)
    objMIGRequestDoc.appendChild objXmlMIGNode
    objXmlAssist.ChangeNodeName objXmlMIGNode, cstrCALC_NAME, "MIG"
    
    Dim strMIGResponse As String
    Dim strTagValue As String
    
    strMIGResponse = objMigBO.CalcMIG(objMIGRequestDoc.xml)
    objErrAssist.CheckXMLResponse strMIGResponse, True
    Set objXmlMIGDoc = objXmlAssist.load(strMIGResponse, TypeName(Me), strFunctionName)
    
    strTagValue = objXmlAssist.GetTagValue(objXmlMIGDoc.documentElement, "MIGPREMIUM")
    roocFees.lngMIGPremium = objConvAssist.CSafeLng(strTagValue)
    strTagValue = objXmlAssist.GetTagValue(objXmlMIGDoc.documentElement, "TOTALINDEMNITY")
    roocFees.lngTotalIndemnity = objConvAssist.CSafeLng(strTagValue)
    
    Set objMigBO = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objXmlMIGDoc = Nothing
    Set objXmlRequestElem = Nothing
    Set objXmlMIGNode = Nothing
    Set objMIGRequestDoc = Nothing
    
  
    Exit Sub
    
CalculateMortgageIndemnityVbErr:

    Set objMigBO = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objXmlMIGDoc = Nothing
    Set objXmlRequestElem = Nothing
    Set objXmlMIGNode = Nothing
    Set objMIGRequestDoc = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description
    
End Sub

Private Sub GetStampDuty(ByVal vstrPurchasePrice As String, _
                         ByRef roocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Gets the stamp duty
'
' pass:         vstrPurchasePrice   Purchase Price
'               roocFees            Fees structure
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo GetStampDutyVbErr
    
    Const strFunctionName As String = "GetStampDuty"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objGlobalBandedParamDO As GlobalBandedParameterDO
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objXmlResponseDoc As FreeThreadedDOMDocument40
    
    If objContext Is Nothing Then
        Set objGlobalBandedParamDO = New GlobalBandedParameterDO
    Else
        Set objGlobalBandedParamDO = _
                    objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalBandedParameterDO")
    End If
    
    Dim strResponse As String
    Dim strTagValue As String
    Dim dblStampDutyPercent As Double
    Dim dblPurchasePrice As Double
   
    strResponse = objGlobalBandedParamDO.GetCurrentParameter("StampDuty", vstrPurchasePrice)
    Set objXmlResponseDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    strTagValue = objXmlAssist.GetTagValue(objXmlResponseDoc.documentElement, _
                                           "PERCENTAGE")
    
    dblStampDutyPercent = objConvAssist.CSafeDbl(strTagValue)
    dblPurchasePrice = objConvAssist.CSafeDbl(vstrPurchasePrice)
    
    roocFees.lngStampDuty = Fix(dblPurchasePrice * dblStampDutyPercent / 100)
        
    Set objGlobalBandedParamDO = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlResponseDoc = Nothing
  
    Exit Sub
    
GetStampDutyVbErr:

    Set objGlobalBandedParamDO = Nothing
    Set objXmlAssist = Nothing
    Set objXmlResponseDoc = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description
    
End Sub

Private Sub AppendCharges(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                          ByVal vobjXmlOutElement As IXMLDOMElement, _
                          ByRef voocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Appends the charges to the element passed in
'
' pass:         vobjXmlRequest      xml request document
'               vobjXmlOutElement   element to add charges to
'               roocFees            fees structure
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo AppendChargesVbErr

    Const strFunctionName As String = "AppendCharges"
    
    Dim objFeeElement As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objCostsElem As IXMLDOMElement
    Dim objProductList As IXMLDOMNodeList
    Dim objProductElem As IXMLDOMElement
    Dim objOtherFeeList As IXMLDOMNodeList
    Dim objOtherFeeElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim blnValReqd As Boolean
    Dim xmlValuationType As IXMLDOMNode
    Dim xmlNoValnReqd As IXMLDOMNode
    Dim objComboDO As ComboDO
    Dim xmlCombo As FreeThreadedDOMDocument40  'SR EP2_1272
    'SR 28/02/2007 : EP2_1272
    Dim xmlListEntry As IXMLDOMElement, xmlValidationType As IXMLDOMNode
    Dim strValidationType As String, strIdentifier As String, strFeeType As String
    Dim intMaxInstanceOfFeeType As Integer, intCount As Integer, intFeesAlreadyUsedFromInput As Integer, intFeesAlreadyInOneOffCostXml As Integer
    'Dim intFeesAddedToOutput As Integer
    Dim xmlIntroduceFeeListNode As IXMLDOMNode, xmlIntroducerFeeList As IXMLDOMNodeList, xmlFeesAlreadyAddedList As IXMLDOMNodeList
    Dim xmlFeeNode As IXMLDOMNode, xmlFeeAmountElem As IXMLDOMNode, xmlRefundElem As IXMLDOMNode, xmlIntroducerFeeNode As IXMLDOMNode
    Dim intIntroducerFeePassesIn As Integer
    Dim objConversionAssist As New ConvertAssist
    'SR 28/02/2007 : EP2_1272 - End
    
    ' PSC 02/03/2007 EP2_1347 - Start
    Dim xmlValRefundProdList As IXMLDOMNodeList
    
    Dim blnIsRefundOfValuation As Boolean
    Dim lngFeeAmount As Long
    Dim lngRefundAmount As Long
    Dim strTPVFeeType As String
    ' PSC 02/03/2007 EP2_1347 - End
    Dim lngRebateAmount As Long 'SR EP2_1753
    Dim xmlRebateElem As IXMLDOMNode 'SR EP2_1753
    
    Dim strCostTagName As String
    strCostTagName = "ONEOFFCOST"
  
    Dim objContext As ObjectContext  'SR EP2_1272
    Set objContext = GetObjectContext()  'SR EP2_1272
  
    Set xmlCombo = New FreeThreadedDOMDocument40  'SR EP2_1272
    xmlCombo.validateOnParse = False  'SR EP2_1272
    xmlCombo.setProperty "NewParser", True  'SR EP2_1272
    
     'SR EP2_1272
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    'SR EP2_1272 - End
  
    'EP2_1281
    blnValReqd = True
    Set xmlValuationType = vobjXmlRequest.selectSingleNode("//VALUATIONTYPE")
    If Not xmlValuationType Is Nothing Then
        If IsNumeric(xmlValuationType.Text) Then
            'Set xmlCombo = New DomDocument  'SR EP2_1272 - End
            If xmlCombo.loadXML(objComboDO.GetComboValueId("ValuationType", "NONE")) Then
                Set xmlNoValnReqd = xmlCombo.selectSingleNode("VALUEID")
                If Not xmlNoValnReqd Is Nothing Then
                    If IsNumeric(xmlNoValnReqd.Text) Then
                        blnValReqd = CInt(xmlNoValnReqd.Text) <> CInt(xmlValuationType.Text)
                    End If
                End If
            End If
        End If
    End If
    
    ' PSC 02/03/2007 EP2_1347 - Start
    Set xmlValRefundProdList = vobjXmlRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCT[REFUNDOFVALUATION='1']")
    blnIsRefundOfValuation = xmlValRefundProdList.length > 0
    ' PSC 02/03/2007 EP2_1347 - End
    
    Set objCostsElem = vobjXmlOutElement.ownerDocument.createElement("ONEOFFCOSTLIST")
    vobjXmlOutElement.appendChild objCostsElem
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octArrangementFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngArrangementFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octPortingFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngPortingFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    strTPVFeeType = GetFirstComboValueId("OneOffCost", "TPV")
    Set xmlIntroducerFeeNode = xmlGetNode(vobjXmlRequest, ".//INTRODUCERFEELIST/INTRODUCERFEE[FEETYPE='" & strTPVFeeType & "']")
    ' PSC 02/03/2007 EP2_1347 - End

    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octValuationFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    If blnValReqd Then objElement.Text = CStr(voocFees.lngValuationFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    ' Don't create refund element if we have a TPV sent in
    If xmlIntroducerFeeNode Is Nothing Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")   ' MAR28
        If blnValReqd Then objElement.Text = CStr(voocFees.lngRefundAmount)              ' MAR28
        objFeeElement.appendChild objElement                                             ' MAR28
    End If
    ' PSC 02/03/2007 EP2_1347 - Start
    
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = xmlGetNode(objFeeElement, "REFUNDAMOUNT")
        
        If objElement Is Nothing Then
            Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
            objElement.Text = "0"
            objFeeElement.appendChild objElement
        End If
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octReinspectionFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    If blnValReqd Then objElement.Text = CStr(voocFees.lngReinspectionFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    'MO 10/7/2002 Modified for BMIDS00196, append the revaluation fee to the returned XML
    'NOTE: This was originally changed for BMIDS CMWP2, but the INHERIT tags were not removed,
    '       and the change was lost
    'Start
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octRevaluation)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    If blnValReqd Then objElement.Text = CStr(voocFees.lngRevaluationFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
        
    objCostsElem.appendChild objFeeElement
    'End
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octAdministrationFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngAdministrationFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octStampDuty)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngStampDuty)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octSealingFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(Fix(voocFees.lngSealingFee))
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octDeedsReleaseFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(Fix(voocFees.lngDeedsReleaseFee))
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    'GD BM0198 START
    'IK EP776 starts
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octTTFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngTTFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    'IK EP776 ends
    'GD BM0198 END
    
    If voocFees.lngLegalFee > 0 Then 'SR 28/02/2007 : EP2_1272 - Add LegalFee only when it was calculated earlier
        Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
        objElement.Text = GetFeeIdentifier(octLegalFee)
        strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
        objFeeElement.appendChild objElement
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
        objElement.Text = CStr(voocFees.lngLegalFee)
        objFeeElement.appendChild objElement
    
        ' PSC 02/03/2007 EP2_1347 - Start
        If IsFeeRefundable(strIdentifier) Then
            Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
            objElement.Text = "0"
            objFeeElement.appendChild objElement
        End If
        ' PSC 02/03/2007 EP2_1347 - End
    
        objCostsElem.appendChild objFeeElement
    End If
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octMIGPremium)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngMIGPremium)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    ' BMIDS763 DRC - Add in a dummy place holder for Product Switch Fee
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octProductSwitchFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = "0" ' must be set to 0
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    ' BMIDS763 DRC - End
    
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octTotalIndemnity)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngTotalIndemnity)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    'AW 10/05/2006 EP514 - Start
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octBrokerFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngBrokerFeeAmount)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octThirdPartyValuationFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngThirdPartyValuationFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If blnIsRefundOfValuation Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    'AW 10/05/2006 EP514 - End
    
    'EP2_9 GHun
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octAdditionalBorrowingFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngAdditionalBorrowingFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octCreditLimitIncreaseFee)
    strIdentifier = objElement.Text ' PSC 02/03/2007 EP2_1347
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngCreditLimitIncreaseFee)
    objFeeElement.appendChild objElement
    
    ' PSC 02/03/2007 EP2_1347 - Start
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    ' PSC 02/03/2007 EP2_1347 - End
    
    objCostsElem.appendChild objFeeElement
    'EP2_9 End
    'SR 26/03/2007: EP2_1777
    Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
    objElement.Text = GetFeeIdentifier(octTransferOfEquityFee)
    strIdentifier = objElement.Text
    objFeeElement.appendChild objElement
    Set objElement = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
    objElement.Text = CStr(voocFees.lngTransferOfEquityFee)
    objFeeElement.appendChild objElement
    If IsFeeRefundable(strIdentifier) Then
        Set objElement = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
        objElement.Text = "0"
        objFeeElement.appendChild objElement
    End If
    objCostsElem.appendChild objFeeElement
    'SR 26/03/2007: EP2_1777 - End
    
    Dim strTagName As String
    
    Set objProductList = vobjXmlRequest.getElementsByTagName("MORTGAGEPRODUCTDETAILS")
    
    For Each objProductElem In objProductList
        Set objOtherFeeList = objProductElem.getElementsByTagName("OTHERFEE")
        
        For Each objOtherFeeElem In objOtherFeeList
        
            Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
            Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
            objElement.Text = GetFeeIdentifier(octOther)
            objFeeElement.appendChild objElement
            
            strTagName = "AMOUNT"
            Set objElement = vobjXmlOutElement.ownerDocument.createElement(strTagName)
            objElement.Text = objXmlAssist.GetTagValue(objOtherFeeElem, "OTHERFEEAMOUNT") 'MCS     03/03/00    SYS0301
            objFeeElement.appendChild objElement
            
            strTagName = "NAME"
            Set objElement = vobjXmlOutElement.ownerDocument.createElement(strTagName)
            objElement.Text = objXmlAssist.GetTagValue(objOtherFeeElem, strTagName)
            objFeeElement.appendChild objElement
           
            objCostsElem.appendChild objFeeElement
        Next
    Next
    
    'SR EP2_1272 - begin
    'Get all the Introducer Fees - OneOffCosts combo with validationType='I'
    xmlCombo.loadXML objComboDO.GetComboList("<REQUEST><LIST><LISTNAME>OneOffCost</LISTNAME></LIST></REQUEST>")
    
    ' PSC 02/03/2007 EP2_1347 - Start
    For Each xmlListEntry In xmlCombo.selectNodes("//LISTENTRY[not(VALIDATIONTYPELIST/VALIDATIONTYPE= 'I')]")
        xmlListEntry.parentNode.removeChild xmlListEntry
    Next xmlListEntry
    ' PSC 02/03/2007 EP2_1347 - End

    Set xmlIntroduceFeeListNode = vobjXmlRequest.selectSingleNode("//INTRODUCERFEELIST")
            
    'Check for the instances of each FeeType allowed. Update the OneOffCost records already added, with FeeAmount as ZERO,
    ' with values from IntroduceFeeXML passed in
    For Each xmlListEntry In xmlCombo.selectNodes("//LISTENTRY")
        intMaxInstanceOfFeeType = 0
        strIdentifier = ""
        strFeeType = ""
        
        ' PSC 02/03/2007 EP2_1347 - Start
        strValidationType = xmlGetNodeText(xmlListEntry, "VALIDATIONTYPELIST/VALIDATIONTYPE[starts-with(.,'@')]")
        
        If Len(strValidationType) > 1 Then
            intMaxInstanceOfFeeType = CSafeInt(Right(strValidationType, Len(strValidationType) - 1))
        Else
            intMaxInstanceOfFeeType = 1
        End If
        
        strValidationType = xmlGetNodeText(xmlListEntry, "VALIDATIONTYPELIST/VALIDATIONTYPE[starts-with(.,'#')]")
        
        If Len(strValidationType) > 1 Then
            strIdentifier = Right(strValidationType, Len(strValidationType) - 1)
        End If
        
        strFeeType = xmlGetNodeText(xmlListEntry, "VALUEID")
        ' PSC 02/03/2007 EP2_1347 - End
              
        
        If Not xmlIntroduceFeeListNode Is Nothing Then
            Set xmlIntroducerFeeList = xmlIntroduceFeeListNode.selectNodes("INTRODUCERFEE[FEETYPE='" & strFeeType & "']")
            intIntroducerFeePassesIn = xmlIntroducerFeeList.length
        Else
            intIntroducerFeePassesIn = 0
        End If
        
        'Create OneOffCost records so that we have got the maximum number of records allowed.
        'If we have not used up IntroduceFeeXML nodes for this FeeType, use these values to update FeeAmount and RefundAmount
        'else set them to zero
        Set xmlFeesAlreadyAddedList = objCostsElem.selectNodes("//" & strCostTagName & "[IDENTIFIER='" & strIdentifier & "']")
        intFeesAlreadyInOneOffCostXml = xmlFeesAlreadyAddedList.length
        intFeesAlreadyUsedFromInput = 0
        If intFeesAlreadyInOneOffCostXml > 0 And intIntroducerFeePassesIn > 0 Then
             For intCount = 0 To intFeesAlreadyInOneOffCostXml - 1
                If intIntroducerFeePassesIn >= intCount + 1 Then
                    Set xmlIntroducerFeeNode = xmlIntroducerFeeList.Item(intCount)
                    Set xmlFeeNode = xmlFeesAlreadyAddedList.Item(intFeesAlreadyUsedFromInput)
                    Set xmlFeeAmountElem = xmlFeeNode.selectSingleNode("AMOUNT")
                    Set xmlRefundElem = xmlFeeNode.selectSingleNode("REFUNDAMOUNT")
                    Set xmlRebateElem = xmlFeeNode.selectSingleNode("REBATEAMOUNT")
                    
                    ' PSC 02/03/2007 EP2_1347 - Start
                    lngFeeAmount = xmlGetNodeAsLong(xmlIntroducerFeeNode, "FEEAMOUNT")
                    lngRefundAmount = xmlGetNodeAsLong(xmlIntroducerFeeNode, "REFUNDAMOUNT")
                    ' PSC 02/03/2007 EP2_1347 - End
                    lngRebateAmount = xmlGetNodeAsLong(xmlIntroducerFeeNode, "REBATEAMOUNT") 'SR 21/03/2007  Ep2_1753
                    
                    If xmlFeeAmountElem Is Nothing Then
                        Set xmlFeeAmountElem = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
                        xmlFeeNode.appendChild xmlFeeAmountElem
                    End If
                    If xmlRefundElem Is Nothing Then
                        Set xmlRefundElem = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
                        xmlFeeNode.appendChild xmlRefundElem
                    End If
                    If xmlRebateElem Is Nothing Then
                        Set xmlRebateElem = vobjXmlOutElement.ownerDocument.createElement("REBATEAMOUNT")
                        xmlFeeNode.appendChild xmlRebateElem
                    End If
                    
                    ' PSC 02/03/2007 EP2_1347 - Start
                    xmlFeeAmountElem.Text = CStr(lngFeeAmount)
                    
                    If strIdentifier = "TPV" Or strIdentifier = "TPVA" Then
                        If blnIsRefundOfValuation Then
                            xmlRefundElem.Text = CStr(lngFeeAmount)
                        ElseIf blnIsRefundOfValuation = False And lngRefundAmount > 0 Then
                            xmlRefundElem.Text = CStr(lngRefundAmount)
                        ElseIf IsValidationType("OneOffCost", CSafeInt(strFeeType), "R") Then
                            xmlRefundElem.Text = "0"
                        Else
                            ' Remove the refund element as it is not applicable and an empty
                            ' node cannot be processed by create KFI
                            xmlRefundElem.parentNode.removeChild xmlRefundElem
                        End If
                    Else
                        
                        If lngRefundAmount > 0 Then
                            xmlRefundElem.Text = CStr(lngRefundAmount)
                        ElseIf IsValidationType("OneOffCost", CSafeInt(strFeeType), "R") Then
                            xmlRefundElem.Text = "0"
                        Else
                            ' Remove the refund element as it is not applicable and an empty
                            ' node cannot be processed by create KFI
                            xmlRefundElem.parentNode.removeChild xmlRefundElem
                        End If
                    End If
                    ' PSC 02/03/2007 EP2_1347 - End
                    'SR 21/03/2007 : EP2_1753
                    If lngRebateAmount > 0 Then
                        xmlRebateElem.Text = CStr(lngRebateAmount)
                    ElseIf IsValidationType("OneOffCost", CSafeInt(strFeeType), "REB") Then
                        xmlRebateElem.Text = "0"
                    Else
                        xmlRebateElem.parentNode.removeChild xmlRebateElem
                    End If
                    'SR 21/03/2007 : EP2_1753 - End
                    intFeesAlreadyUsedFromInput = intFeesAlreadyUsedFromInput + 1
                End If
             Next intCount
        End If
        
       ' Set xmlFeesAlreadyAddedList = objCostsElem.selectNodes("//" & strCostTagName & "[IDENTIFIER='" & strIdentifier & "']")
       ' intFeesAddedToOutput = xmlFeesAlreadyAddedList.length
        
        ' Create OneOffCost records with values from IntroducerFeeXML passed in or FeeAmount values as ZERO
        For intCount = intFeesAlreadyInOneOffCostXml To intMaxInstanceOfFeeType - 1 Step 1
            Set objFeeElement = vobjXmlOutElement.ownerDocument.createElement(strCostTagName)
            Set objElement = vobjXmlOutElement.ownerDocument.createElement("IDENTIFIER")
            objElement.Text = strIdentifier
            objFeeElement.appendChild objElement
            
            Set xmlRefundElem = vobjXmlOutElement.ownerDocument.createElement("REFUNDAMOUNT")
            Set xmlFeeAmountElem = vobjXmlOutElement.ownerDocument.createElement("AMOUNT")
            ' PSC 02/03/2007 EP2_1347
            If intIntroducerFeePassesIn > 0 And _
                intIntroducerFeePassesIn - intFeesAlreadyUsedFromInput > 0 Then
                'All Records not used up from the input xml IntroducerFeeList
                Set xmlIntroducerFeeNode = xmlIntroducerFeeList.Item(intFeesAlreadyUsedFromInput)
                
                ' PSC 02/03/2007 EP2_1347 - Start
                lngFeeAmount = xmlGetNodeAsLong(xmlIntroducerFeeNode, "FEEAMOUNT")
                lngRefundAmount = xmlGetNodeAsLong(xmlIntroducerFeeNode, "REFUNDAMOUNT")
            
                xmlFeeAmountElem.Text = CStr(lngFeeAmount)
                
                If strIdentifier = "TPV" Or strIdentifier = "TPVA" Then
                    If blnIsRefundOfValuation Then
                        xmlRefundElem.Text = CStr(lngFeeAmount)
                    ElseIf blnIsRefundOfValuation = False And lngRefundAmount > 0 Then
                        xmlRefundElem.Text = CStr(lngRefundAmount)
                    ElseIf IsValidationType("OneOffCost", CSafeInt(strFeeType), "R") Then
                        xmlRefundElem.Text = "0"
                    Else
                        xmlRefundElem.Text = ""
                    End If
                Else
                    
                    If lngRefundAmount > 0 Then
                        xmlRefundElem.Text = CStr(lngRefundAmount)
                    ElseIf IsValidationType("OneOffCost", CSafeInt(strFeeType), "R") Then
                        xmlRefundElem.Text = "0"
                    Else
                        xmlRefundElem.Text = ""
                    End If
                End If
                ' PSC 02/03/2007 EP2_1347 - End
                
                intFeesAlreadyUsedFromInput = intFeesAlreadyUsedFromInput + 1
            Else 'Records used up from the input IntroducerFeeList xml / not got any IntroducerFeeList xml in the Input
                xmlFeeAmountElem.Text = 0
                
                ' PSC 02/03/2007 EP2_1347 - Start
                If (strIdentifier = "TPV" Or strIdentifier = "TPVA") And blnIsRefundOfValuation Then
                    xmlRefundElem.Text = "0"
                ElseIf IsValidationType("OneOffCost", CSafeInt(strFeeType), "R") Then
                    xmlRefundElem.Text = 0
                Else
                    xmlRefundElem.Text = ""
                End If
                ' PSC 02/03/2007 EP2_1347 - End

            End If
            objFeeElement.appendChild objElement
            objFeeElement.appendChild xmlFeeAmountElem
            
            ' PSC 02/03/2007 EP2_1347 - Start
            If Len(xmlRefundElem.Text) > 0 Then
                objFeeElement.appendChild xmlRefundElem
            End If
            ' PSC 02/03/2007 EP2_1347 - End
            
            objCostsElem.appendChild objFeeElement
        Next intCount
    Next xmlListEntry
    'SR EP2_1272 - End
    
    Set objFeeElement = Nothing
    Set objElement = Nothing
    Set objProductList = Nothing
    Set objProductElem = Nothing
    Set objOtherFeeList = Nothing
    Set objOtherFeeElem = Nothing
    Set objCostsElem = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    'SR 28/02/2007 : EP2_1272
    Set xmlListEntry = Nothing
    Set xmlValidationType = Nothing
    Set xmlIntroduceFeeListNode = Nothing
    Set xmlIntroducerFeeList = Nothing
    Set xmlRefundElem = Nothing
    Set xmlFeeAmountElem = Nothing
    Set xmlIntroducerFeeNode = Nothing
    Set xmlFeesAlreadyAddedList = Nothing
    Set xmlFeeNode = Nothing
    Set objConversionAssist = Nothing
    'SR 28/02/2007 : EP2_1272 - End
    Set xmlValRefundProdList = Nothing  ' PSC 02/03/2007 EP2_1347
    Exit Sub
    
AppendChargesVbErr:

    Set objFeeElement = Nothing
    Set objElement = Nothing
    Set objProductList = Nothing
    Set objProductElem = Nothing
    Set objOtherFeeList = Nothing
    Set objOtherFeeElem = Nothing
    Set objCostsElem = Nothing
    Set objXmlAssist = Nothing
    'SR 28/02/2007 : EP2_1272
    Set xmlListEntry = Nothing
    Set xmlValidationType = Nothing
    Set xmlIntroduceFeeListNode = Nothing
    Set xmlIntroducerFeeList = Nothing
    Set xmlRefundElem = Nothing
    Set xmlFeeAmountElem = Nothing
    Set xmlIntroducerFeeNode = Nothing
    Set objConversionAssist = Nothing
    'SR 28/02/2007 : EP2_1272 - End
    Set xmlRebateElem = Nothing 'SR EP2_1753
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description
    
End Sub

Private Function GetFeeIdentifier(ByVal voctCostType As ONEOFFCOSTTYPES) As String  'EP2_9 GHun changed to private

' header ----------------------------------------------------------------------------------
' description:  Identifier for a particular fee type
'
' pass:         voctCostType      Cost type to get the identifier for
'
' return:       GetFeeIdentifier The string that identifies the cost type
'------------------------------------------------------------------------------------------
On Error GoTo GetFeeIdentifierVbErr

    Const strFunctionName As String = "GetFeeIdentifier"

    
    Dim strIdentifier As String
    Dim objErrAssist As New ErrAssist
   
    Select Case voctCostType
        Case octArrangementFee
            strIdentifier = "ARR"
        Case octPortingFee
            strIdentifier = "POR"
        Case octValuationFee
             strIdentifier = "VAL"
        Case octReinspectionFee
             strIdentifier = "REI"
        Case octAdministrationFee
            strIdentifier = "ADM"
        Case octStampDuty
            strIdentifier = "STA"
        Case octSealingFee
            strIdentifier = "SEA"
        Case octDeedsReleaseFee
            strIdentifier = "DEE"
        'GD BM0198 START
        'IK EP776 starts
        Case octTTFee
            strIdentifier = "TTF"
        'IK EP776 ends
        'GD BM0198 END
        Case octLegalFee
            strIdentifier = "LEG"
        Case octMIGPremium
            strIdentifier = "MIG"
        Case octTotalIndemnity
            strIdentifier = "TID"
        Case octOther
            strIdentifier = "OTH"
        'MO 06/06/2002 CMWP2 - A new identifer for revaluation - Start
        Case octRevaluation
            strIdentifier = "REV"
        'MO 06/06/2002 CMWP2 - End
        'DRC 28/06/2004 - New Identifier for Product Switching Fee - Start
        Case octProductSwitchFee
            strIdentifier = "PSF"
        'DRC 28/06/2004 - End
        'AW 10/05/2006 EP514 - Start
        Case octBrokerFee
            strIdentifier = "BRK"
        Case octThirdPartyValuationFee
            strIdentifier = "TPV"
        'AW 10/05/2006 EP514 - End
        'EP2_9 GHun
        Case octAdditionalBorrowingFee
            strIdentifier = "AB"
        Case octCreditLimitIncreaseFee
            strIdentifier = "CLI"
        'EP2_9 End
        Case octTransferOfEquityFee 'SR 26/03/2007 : EP2_1777
            strIdentifier = "TOE"
        Case Else
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeInvalidParameter, _
                                    "Invalid One off cost type"
    End Select

    GetFeeIdentifier = strIdentifier
    
    Set objErrAssist = Nothing
    
    Exit Function
    
GetFeeIdentifierVbErr:

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub CalculateMortCalcSpecificCosts(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                                           ByRef roocFees As ONEOFFCOSTS)
' header ----------------------------------------------------------------------------------
' description:  Gets the specific one off costs as used for Mortgage Calculator
'
' pass:         vobjXmlRequest  xml Document containing the request
'               roocFees        fees structure
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CalculateMortCalcSpecificCostsVbErr

    Const strFunctionName As String = "CalculateMortCalcSpecificCosts"
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    
    Dim objProductElem As IXMLDOMElement
    Dim objValFeeElem As IXMLDOMElement
    Dim objAdminFeeElem As IXMLDOMElement
    Dim objLegalFeeElem As IXMLDOMElement
    Dim objMtgLenderElem As IXMLDOMElement
    
    Dim strTagValue As String
        
    Set objProductElem = vobjXmlRequest.getElementsByTagName("MORTGAGEPRODUCTDETAILS").Item(0)
    
    If objProductElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "MORTGAGEPRODUCTDETAILS tag missing"
    End If
    
    'MO - 12/06/2002 - CMWP2, BM022 - removed
    ' Get arrangement fee values
    'strTagValue = objXmlAssist.GetTagValue(objProductElem, "MPARRANGEMENTFEEAMOUNT") 'MCS 01/03/00
    'roocFees.dblMtgArrangementFee = objConvAssist.CSafeDbl(strTagValue)
    'strTagValue = objXmlAssist.GetTagValue(objProductElem, "MPARRANGEMENTFEEPERCENT") 'MCS 01/03/00
    'roocFees.dblMtgArrangeFeePercent = objConvAssist.CSafeDbl(strTagValue)
    'strTagValue = objXmlAssist.GetTagValue(objProductElem, "MPARRANGEMENTFEEPERCENTMAX") 'MCS 01/03/00
    'roocFees.dblMtgArrangeFeePercentMax = objConvAssist.CSafeDbl(strTagValue)
    
    strTagValue = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, _
                                           "AMOUNTREQUESTED")
    roocFees.dblNewLoanAmount = objConvAssist.CSafeDbl(strTagValue)
    
    Set objValFeeElem = objProductElem.getElementsByTagName("VALUATIONFEE").Item(0)
    
    If Not objValFeeElem Is Nothing Then
        strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "AMOUNT")
        'MAR717 INR Picking up wrong amount for ValuationFee
        roocFees.dblValuationFeeAmount = objConvAssist.CSafeLng(strTagValue)
        
        'MO - 12/06/2002 - CMWP2 - Start
        strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "FEEPERCENTAGE")
        roocFees.dblValuationFeePC = objConvAssist.CSafeDbl(strTagValue)

        strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "MAXIMUMFEEVALUE")
        roocFees.dblValuationFeePCMax = objConvAssist.CSafeDbl(strTagValue)
        
        strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "MINIMUMFEEVALUE")
        roocFees.dblValuationFeePCMin = objConvAssist.CSafeDbl(strTagValue)
        'MO - 12/06/2002 - CMWP2 - End
    End If
    
    Set objAdminFeeElem = _
        objProductElem.getElementsByTagName("ADMINISTRATIONFEE").Item(0)

    If Not objAdminFeeElem Is Nothing Then
        strTagValue = objXmlAssist.GetTagValue(objAdminFeeElem, "AMOUNT")
        roocFees.lngAdministrationFee = objConvAssist.CSafeLng(strTagValue)
    End If

    Set objMtgLenderElem = vobjXmlRequest.getElementsByTagName("MORTGAGELENDER").Item(0)
    
    'GD BM0198 START
    'IK EP776 starts
    If Not objMtgLenderElem Is Nothing Then
        strTagValue = objXmlAssist.GetTagValue(objMtgLenderElem, "TTFEE")
        roocFees.lngTTFee = objConvAssist.CSafeLng(strTagValue)
    End If
    'IK EP776 ends
    'GD BM0198 END
    
    Set objLegalFeeElem = objProductElem.getElementsByTagName("LEGALFEE").Item(0)

    If Not objLegalFeeElem Is Nothing Then
       'EP2_19 GHun
        If objXmlAssist.GetTagValue(objProductElem, "FREELEGALFEES") = "1" Then
            roocFees.lngLegalFee = 0
        Else
        'EP2_19 End
            strTagValue = objXmlAssist.GetTagValue(objLegalFeeElem, "AMOUNT")
            roocFees.lngLegalFee = objConvAssist.CSafeLng(strTagValue)
        End If 'EP2_19 GHun
    End If
    
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objProductElem = Nothing
    Set objValFeeElem = Nothing
    Set objAdminFeeElem = Nothing
    Set objLegalFeeElem = Nothing
    Set objMtgLenderElem = Nothing
    
    Exit Sub

CalculateMortCalcSpecificCostsVbErr:
    
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objProductElem = Nothing
    Set objValFeeElem = Nothing
    Set objAdminFeeElem = Nothing
    Set objLegalFeeElem = Nothing
    Set objMtgLenderElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub


Private Sub CalculateSpecificCosts(ByVal vobjXmlRequest As FreeThreadedDOMDocument40, _
                                   ByRef roocFees As ONEOFFCOSTS, _
                                   ByRef rblnPortingFeeNeeded As Boolean)
' header ----------------------------------------------------------------------------------
' description:  Gets the specific one off costs
'
' pass:         vobjXmlRequest  xml Document containing the request
'               roocFees        fees structure
'
' return:       n/a
' Change history --------------------------------------------------------------------------
' MO    2/7/2002    Modified for BMIDS00131, changed calculations for valuation fee
'------------------------------------------------------------------------------------------

On Error GoTo CalculateSpecificCostsVbErr

    Const strFunctionName As String = "CalculateSpecificCosts"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    
    Dim objProductElem As IXMLDOMElement
    Dim objValFeeElem As IXMLDOMElement
    Dim objAdminFeeElem As IXMLDOMElement
    Dim objLegalFeeElem As IXMLDOMElement
    Dim objMtgLenderElem As IXMLDOMElement
    Dim objLoanCompNodeList As IXMLDOMNodeList
    Dim objLoanCompElem As IXMLDOMElement
    Dim objComboDO As ComboDO
         
    Dim blnIsPortedLoan As Boolean
    'Dim lngPortedLoanAmt As Long
    'Dim lngNewLoanAmount As Long
    Dim lngCurrPortingFee As Long
    Dim lngCurrAdminFee As Long
    'Dim lngCurrValuationFee As Long

    'MO 12/06/2002 CMWP2
    'Dim dblCurrMtgArrangementFee As Double
    'Dim dblCurrMtgArrangeFeePercent As Double
    'Dim dblCurrMtgArrangeFeePercentMax As Double

    'MO 06/06/2002 CMWP2 - Start
    Dim dblCurrValFeeAmount As Double
    Dim dblCurrValFeePC As Double
    Dim dblCurrValFeePCMax As Double
    Dim dblCurrValFeePCMin As Double
    'MO 06/06/2002 CMWP2 - End
    
    'MO 02/07/2002 BMIDS00089 - Start
    Dim blnFirstLoanComponent As Boolean
    'MO 02/07/2002 BMIDS00089 - End
    'SR 28/02/2007 : EP2_1272
    Dim xmlIntroducerFeeListNode As IXMLDOMNode
    Dim strLegalFeeType As String
    Dim blnLegalPassedIn As Boolean
    'SR 28/02/2007 : EP2_1272 - End
    
    'Dim strValuationType As String
           
    rblnPortingFeeNeeded = False

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
   
    Dim strTagValue As String
    
    Set objLoanCompNodeList = vobjXmlRequest.getElementsByTagName("LOANCOMPONENT")
    
    'MO 02/07/2002 BMIDS00089
    blnFirstLoanComponent = True
    ' Go round each loan component to get the maximum value of the fees
    For Each objLoanCompElem In objLoanCompNodeList
        strTagValue = objXmlAssist.GetTagValue(objLoanCompElem, "PORTEDLOAN")
        blnIsPortedLoan = objConvAssist.CSafeBool(strTagValue)
   
        ' Only get the fees if this loan component is not to be ported
        If blnIsPortedLoan = False Then
        
            Set objProductElem = _
                    objLoanCompElem.getElementsByTagName("MORTGAGEPRODUCTDETAILS").Item(0)
                    
            If objProductElem Is Nothing Then
                objErrAssist.ThrowError TypeName(Me), _
                                        strFunctionName, _
                                        oeMissingPrimaryTag, _
                                        "MORTGAGEPRODUCTDETAILS tag missing"
            End If
            
            strTagValue = objXmlAssist.GetTagValue(objLoanCompElem, "LOANAMOUNT")
            roocFees.dblNewLoanAmount = _
                            roocFees.dblNewLoanAmount + objConvAssist.CSafeDbl(strTagValue)
            
            'SR EP2_1272 : removed commented code. See V8 for the code removed
                        
            ' Get porting fee
            strTagValue = objXmlAssist.GetTagValue(objProductElem, "PORTINGFEE")
            lngCurrPortingFee = objConvAssist.CSafeLng(strTagValue)
                        
            If lngCurrPortingFee > roocFees.lngPortingFee Then
                roocFees.lngPortingFee = lngCurrPortingFee
            End If
            
            'MO 02/07/2002 BMIDS00089
            'Only use the valuation fee from the first component
            If blnFirstLoanComponent = True Then
            
            ' Get valuation/reinspection fees
                Set objValFeeElem = objProductElem.getElementsByTagName("VALUATIONFEE").Item(0)
                
                If Not objValFeeElem Is Nothing Then
                    
                    'MO 02/07/2002 BMIDS00089 - Start
                    strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "AMOUNT")
                    dblCurrValFeeAmount = objConvAssist.CSafeDbl(strTagValue)
                    roocFees.dblValuationFeeAmount = dblCurrValFeeAmount
                    
                    strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "FEEPERCENTAGE")
                    dblCurrValFeePC = objConvAssist.CSafeDbl(strTagValue)
                    roocFees.dblValuationFeePC = dblCurrValFeePC
                    
                    strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "MAXIMUMFEEVALUE")
                    dblCurrValFeePCMax = objConvAssist.CSafeDbl(strTagValue)
                    roocFees.dblValuationFeePCMax = dblCurrValFeePCMax
                    
                    strTagValue = objXmlAssist.GetTagValue(objValFeeElem, "MINIMUMFEEVALUE")
                    dblCurrValFeePCMin = objConvAssist.CSafeDbl(strTagValue)
                    roocFees.dblValuationFeePCMin = dblCurrValFeePCMin
                    'MO 02/07/2002 BMIDS00089 - End
                    
                    'SR EP2_1272 : removed commented code. See V8 for the code removed
                    
                End If
                        
            End If
            'MO 02/07/2002 BMIDS00089 - End
                  
            ' get admin fee
            Set objAdminFeeElem = _
                objProductElem.getElementsByTagName("ADMINISTRATIONFEE").Item(0)
        
            If Not objAdminFeeElem Is Nothing Then
                strTagValue = objXmlAssist.GetTagValue(objAdminFeeElem, "AMOUNT")
                lngCurrAdminFee = objConvAssist.CSafeLng(strTagValue)
                
                If lngCurrAdminFee > roocFees.lngAdministrationFee Then
                    roocFees.lngAdministrationFee = lngCurrAdminFee
                End If
            End If
        Else
            rblnPortingFeeNeeded = True
        End If
        
        'MO 02/07/2002 BMIDS00089
        blnFirstLoanComponent = False
        
    Next
        
    Set objMtgLenderElem = vobjXmlRequest.getElementsByTagName("MORTGAGELENDER").Item(0)
    
    Dim strApplicationType As String
    Dim blnIsFurtherAdvance As Boolean
    
    strApplicationType = objXmlAssist.GetTagValue(vobjXmlRequest.documentElement, _
                                                  "TYPEOFAPPLICATION")
    blnIsFurtherAdvance = objComboDO.IsFurtherAdvance(strApplicationType)
    
    ' Set up fees that don't relate to further advances
    If blnIsFurtherAdvance = False Then
        'BM0198 START
        'IK EP776 starts
        If Not objMtgLenderElem Is Nothing Then
            strTagValue = objXmlAssist.GetTagValue(objMtgLenderElem, "TTFEE")
            roocFees.lngTTFee = objConvAssist.CSafeLng(strTagValue)
        End If
        'IK EP776 ends
        'BM0198 END
        'SR 28/02/3007 : EP2_1272 - if LegalFee passed in (via IntroducerFeeXML) do not calculate Legal fee
        strLegalFeeType = objComboDO.GetFirstComboValueId("OneOffCost", "LEG")
        
        Set xmlIntroducerFeeListNode = vobjXmlRequest.selectSingleNode("//INTRODUCERFEELIST")
        If Not xmlIntroducerFeeListNode Is Nothing Then
            If xmlIntroducerFeeListNode.selectNodes("INTRODUCERFEE[FEETYPE='" & strLegalFeeType & "']").length > 0 Then
                blnLegalPassedIn = True
            End If
        End If
        'SR 28/02/3007 : EP2_1272 - End
        If Not blnLegalPassedIn Then   'SR 28/02/3007 : EP2_1272
            Set objLegalFeeElem = _
                     objProductElem.getElementsByTagName("LEGALFEE").Item(0)
        
            If Not objLegalFeeElem Is Nothing Then
                'EP2_19 GHun
                If objXmlAssist.GetTagValue(objProductElem, "FREELEGALFEES") = "1" Then
                    roocFees.lngLegalFee = 0
                Else
                'EP2_19 End
                    strTagValue = objXmlAssist.GetTagValue(objLegalFeeElem, "AMOUNT")
                    roocFees.lngLegalFee = objConvAssist.CSafeLng(strTagValue)
                End If  'EP2_19 GHun
            End If
        End If   'SR 28/02/3007 : EP2_1272
    End If
            
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objProductElem = Nothing
    Set objValFeeElem = Nothing
    Set objAdminFeeElem = Nothing
    Set objLegalFeeElem = Nothing
    Set objMtgLenderElem = Nothing
    Set objLoanCompNodeList = Nothing
    Set objLoanCompElem = Nothing
    Set objComboDO = Nothing
    Set xmlIntroducerFeeListNode = Nothing  'SR 28/02/3007 : EP2_1272
    
    Exit Sub

CalculateSpecificCostsVbErr:
    
    Set objConvAssist = Nothing
    Set objProductElem = Nothing
    Set objValFeeElem = Nothing
    Set objAdminFeeElem = Nothing
    Set objLegalFeeElem = Nothing
    Set objMtgLenderElem = Nothing
    Set objLoanCompNodeList = Nothing
    Set objLoanCompElem = Nothing
    Set objComboDO = Nothing
    Set xmlIntroducerFeeListNode = Nothing  'SR 28/02/3007 : EP2_1272
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub
'AW 10/05/2006 EP514 - Start
Private Function GetOverrideFees(ByVal vXMLRequest As IXMLDOMElement) As String
    
    Const cstrFunctionName As String = "GetOverrideFees"
    On Error GoTo GetOverrideFeesExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim objCRUD As Object
    
    If Not vXMLRequest.selectSingleNode("//CHARGES/APPLICATIONNUMBER") Is Nothing Then
    
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
            
        Set xmlElem = xmlDoc.createElement("REQUEST")
        xmlElem.setAttribute "CRUD_OP", "READ"
        Set xmlNode = xmlDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlDoc.createElement("APPLICATIONFACTFIND")
        xmlElem.setAttribute "APPLICATIONNUMBER", vXMLRequest.selectSingleNode("//CHARGES/APPLICATIONNUMBER").Text
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", vXMLRequest.selectSingleNode("//CHARGES/APPLICATIONFACTFINDNUMBER").Text
        xmlNode.appendChild xmlElem
    
        Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        xmlDoc.loadXML objCRUD.OmRequest(xmlDoc.xml)
        Set objCRUD = Nothing
    
        If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving APPLICATIONFACTFIND details"
        End If
        
        GetOverrideFees = xmlDoc.xml
    
    End If
    
GetOverrideFeesExit:
        
    Set objCRUD = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'AW 10/05/2006 EP514 - end

'EP2_9 GHun
Private Sub CalculateAdditionalBorrowingFee(ByVal vXMLRequest As FreeThreadedDOMDocument40, _
                                    ByRef roocFees As ONEOFFCOSTS)

Const cstrFunctionName  As String = "CalculateAdditionalBorrowingFee"
    Dim objXmlAssist        As XMLAssist
    Dim objErrAssist        As ErrAssist
    Dim lngFeeMax           As Long
    Dim lngFee              As Long
    Dim xmlABFee            As IXMLDOMElement
    Dim strFeeAmount        As String
    Dim lngFeePercent       As Long
    Dim lngMaxFeeValue      As Long
    Dim lngMinFeeValue      As Long
    Dim lngAmountRequested  As Long
    
On Error GoTo CalculateAdditionalBorrowingFeeErr
    
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    
    lngAmountRequested = xmlGetMandatoryNodeAsLong(vXMLRequest, "CHARGES/AMOUNTREQUESTED")
    lngFeeMax = 0
    For Each xmlABFee In vXMLRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/ADDITIONALBORROWINGFEE")
        strFeeAmount = xmlGetNodeText(xmlABFee, "AMOUNT")
        
        If Len(strFeeAmount) > 0 Then
            lngFee = CSafeLng(strFeeAmount)
        Else
            lngFeePercent = xmlGetNodeAsLong(xmlABFee, "FEEPERCENTAGE")
            lngMinFeeValue = xmlGetNodeAsLong(xmlABFee, "MINFEEVALUE")
            lngMaxFeeValue = xmlGetNodeAsLong(xmlABFee, "MAXFEEVALUE")
            
            lngFee = CLng(lngFeePercent * lngAmountRequested / 100)
            If (lngMinFeeValue > 0) And (lngFee < lngMinFeeValue) Then
                lngFee = lngMinFeeValue
            ElseIf (lngMaxFeeValue > 0) And (lngFee > lngMaxFeeValue) Then
                lngFee = lngMaxFeeValue
            End If
        End If
        
        If lngFee > lngFeeMax Then
            lngFeeMax = lngFee
        End If
    Next
    
    roocFees.lngAdditionalBorrowingFee = lngFeeMax
    
    Set objXmlAssist = Nothing
    Set xmlABFee = Nothing
    Set objErrAssist = Nothing
    Exit Sub
    
CalculateAdditionalBorrowingFeeErr:
    
    Set objXmlAssist = Nothing
    Set xmlABFee = Nothing
    objErrAssist.AddToErrSource (cstrFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub CalculateCreditLimitIncreaseFee(ByVal vXMLRequest As FreeThreadedDOMDocument40, _
                                    ByRef roocFees As ONEOFFCOSTS)
    
    Const cstrFunctionName  As String = "CalculateCreditLimitIncreaseFee"
    Dim objXmlAssist        As XMLAssist
    Dim objErrAssist        As ErrAssist
    Dim lngFeeMax           As Long
    Dim lngFee              As Long
    Dim xmlCLIFee           As IXMLDOMElement
    Dim strFeeAmount        As String
    Dim lngFeePercent       As Long
    Dim lngMaxFeeValue      As Long
    Dim lngMinFeeValue      As Long
    Dim lngAmountRequested  As Long
    
On Error GoTo CalculateCreditLimitIncreaseFeeErr
    
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist

    lngAmountRequested = xmlGetMandatoryNodeAsLong(vXMLRequest, "CHARGES/AMOUNTREQUESTED")
    lngFeeMax = 0
    For Each xmlCLIFee In vXMLRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/CREDITLIMITINCREASEFEE")
        strFeeAmount = xmlGetNodeText(xmlCLIFee, "AMOUNT")
        
        If Len(strFeeAmount) > 0 Then
            lngFee = CSafeLng(strFeeAmount)
        Else
            lngFeePercent = xmlGetNodeAsLong(xmlCLIFee, "FEEPERCENTAGE")
            lngMinFeeValue = xmlGetNodeAsLong(xmlCLIFee, "MINFEEVALUE")
            lngMaxFeeValue = xmlGetNodeAsLong(xmlCLIFee, "MAXFEEVALUE")
            
            lngFee = CLng(lngFeePercent * lngAmountRequested / 100)
            If (lngMinFeeValue > 0) And (lngFee < lngMinFeeValue) Then
                lngFee = lngMinFeeValue
            ElseIf (lngMaxFeeValue > 0) And (lngFee > lngMaxFeeValue) Then
                lngFee = lngMaxFeeValue
            End If
        End If
        
        If lngFee > lngFeeMax Then
            lngFeeMax = lngFee
        End If
    Next
    
    roocFees.lngCreditLimitIncreaseFee = lngFeeMax
    
    Set objXmlAssist = Nothing
    Set xmlCLIFee = Nothing
    Set objErrAssist = Nothing
    Exit Sub
    
CalculateCreditLimitIncreaseFeeErr:
    
    Set objXmlAssist = Nothing
    Set xmlCLIFee = Nothing
    objErrAssist.AddToErrSource (cstrFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'EP2_9 End

'EP2_9 GHun Copied from MortgageSubQuoteBO
'Calculate the number of whole months difference between 2 dates (i.e. rounding down)
'The VB DataDiff function ignores the day of the month in the dates, so it sometimes
'rounds this value up, and other times rounds it down.
Private Function MonthDiff(ByVal vdteDate1 As Date, ByVal vdteDate2 As Date) As Long
    Dim lngMonths As Long
    
    lngMonths = DateDiff("m", vdteDate1, vdteDate2)
    'If the difference has been rounded up then decrement the number of months
    If DateDiff("d", DateAdd("m", lngMonths, vdteDate1), vdteDate2) < 0 Then
        lngMonths = lngMonths - 1
    End If
    
    MonthDiff = lngMonths
End Function
'EP2_9 GHun

'EP2_56 GHun
Private Sub GetTransferOfEquityFee(ByVal vXMLRequest As FreeThreadedDOMDocument40, _
                                    ByRef roocFees As ONEOFFCOSTS)
    
    Const cstrFunctionName  As String = "GetTransferOfEquityFee"
    Dim objErrAssist        As ErrAssist
    Dim lngFeeMax           As Long
    Dim lngFee              As Long
    Dim xmlTOEFee           As IXMLDOMElement
    
On Error GoTo GetTransferOfEquityFeeErr
    
    Set objErrAssist = New ErrAssist

    lngFeeMax = 0
    For Each xmlTOEFee In vXMLRequest.selectNodes(".//LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/TRANSFEROFEQUITYFEE")
        lngFee = xmlGetNodeAsLong(xmlTOEFee, "AMOUNT")
        
        If lngFee > lngFeeMax Then
            lngFeeMax = lngFee
        End If
    Next
    
    roocFees.lngTransferOfEquityFee = lngFeeMax
    
    Set xmlTOEFee = Nothing
    Set objErrAssist = Nothing
    Exit Sub
    
GetTransferOfEquityFeeErr:

    Set xmlTOEFee = Nothing
    objErrAssist.AddToErrSource (cstrFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'EP2_56 End
    
' PSC 02/03/2007 EP2_1347 - Start
Private Function IsFeeRefundable(ByVal strCostIdentifier) As Boolean
On Error GoTo IsFeeRefundableExit

    Const cstrFunctionName  As String = "IsFeeRefundable"
    
    Dim strValueId As String
    
    strValueId = GetFirstComboValueId("OneOffCost", strCostIdentifier)
    
    If IsValidationType("OneOffCost", CSafeInt(strValueId), "I") And _
       IsValidationType("OneOffCost", CSafeInt(strValueId), "R") Then
        IsFeeRefundable = True
    Else
        IsFeeRefundable = False
    End If
    
IsFeeRefundableExit:

    errCheckError cstrFunctionName, TypeName(Me)

End Function
' PSC 02/03/2007 EP2_1347 - End





