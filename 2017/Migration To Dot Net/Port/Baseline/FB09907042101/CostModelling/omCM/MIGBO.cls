VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MIGBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MIGBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   MIG Business Object
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    17/09/99    Created
'APS    02/12/99    Set bounding MIG tag on the RESPONSE xml for CalcMIG method
'MS     21/07/00    performance rework.
'MS     03/03/00    SYS1386 Pass Rateset to use in calculation
'MS     03/08/00    SYS1384 usew mig threshold in calculation
'DJP    29/09/00    SYS1467 - Add rounding to CalcMIG
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
Implements ObjectControl
 '=============================================
 'Constant Declaration Section
 '=============================================
Private Const cstrCALC_NAME = "MIG"
Private m_objGeneralAssist As GeneralAssist
 
 '=============================================
 'Variable Declaration Section
 '=============================================



'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQr         Description
'MV     04/09/2002  BMIDS00353  Modified CalcMIG() , GetMIGRate()
'MV     16/09/2002  BMIDS00455  Modified CalcMIG()
'MV     02/10/2002  BMIDS00536,
'                   BMIDS00523  Modified CalcMIG()
'MV     11/10/2002  BMIDS00590  Amended CalcMIG()
'MV     18/10/2002  BMIDS00666  Amended CalcMIG()
'MV     23/10/2002  BMIDS00430  Typo Error for "ORIGINALMIGAMOUNT"
'BS     20/03/2003  BM0151  Amended  CalcMIG() to allow for StaffMIGDiscount being a percentage
'MC     16/03/2004  BMIDS763    To find the right MIG rate for the case, application date is used.
'------------------------------------------------------------------------------------------

Public Function CalcMIG(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  calculates the MIG
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       CalcMIG         xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'
'<REQUEST>
'   <MIG>
'       <ORGANISATIONID>Value</ORGANISATIONID>
'       <MIGSTARTLTV>Value</MIGSTARTLTV>
'       <IGNOREMIGPREMIUMAMOUNT>Value</IGNOREMIGPREMIUMAMOUNT>
'       <MINIMUMMIGPREMIUM>Value</MINIMUMMIGPREMIUM>
'       <ADDIPTTOMIGPREMIUM>Value</ADDIPTTOMIGPREMIUM>
'       <TOTALLOANAMOUNT>Value</TOTALLOANAMOUNT>
'       <PURCHASEPRICE>Value</PURCHASEPRICE>
'       <LTV>Value</LTV>
'       <TYPEOFAPPLICATION>Value</TYPEOFAPPLICATION>
'       <LOANSNOTREDEEMED>Value</LOANSNOTREDEEMED>
'   </MIG>
' </REQUEST>
'
'<RESPONSE TYPE=SUCCESS>
'   <MIG>
'       <MIGPREMIUM>Value</MIGPREMIUM>
'       <TOTALINDEMNITY>Value</TOTALINDEMNITY>
'       <IPT>Value</IPT>
'   </MIG>
' </RESPONSE>

'------------------------------------------------------------------------------------------
On Error GoTo CalcMIGVbErr
    
    Const strFunctionName As String = "CalcMIG"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlMIGElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlMIGRateNode As IXMLDOMNode
    
    Dim objComboDO As ComboDO
    Dim objGlobalParameter  As GlobalParameterBO
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    
    Dim dblLTV As Double
    Dim dblMIGStartLTV As Double
    Dim dblMIGThreshold As Double
    Dim dblEligibleAmount As Double
    Dim dblLoansNotRedeemed As Double
    Dim dblPurchasePrice  As Double
    Dim dblIgnoreMigPremAmt As Double
    Dim dblMinMigPremium As Double
    Dim dblIndemnityRate As Double
    Dim dblInsPremiumTaxRate As Double
    Dim dblIndemnityAmount As Double
    Dim dblDiscount  As Double
    Dim dblMIGPremium  As Double
    Dim dblMinimumPremium  As Double
    Dim blnMemberOfStaff  As Boolean
    Dim dblSharedOwnerShipPercentage  As Double
    Dim dblOriginalMIGAmt As Double
    Dim dblStaffMIGThreashold As Double
    Dim dblActualMIGPremium As Double
    
    Dim blnFurtherAdvance  As Boolean
    Dim blnAddIPTToMIG As Boolean
            
    Dim strTagName As String
    Dim strTagValue As String
    Dim strTypeOfApplication As String
    Dim strRateSet As String
    '*=[MC]BMIDS763 APPLICATIONDATE
    Dim strApplicationDate As String
    '*=section end
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                 
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrCALC_NAME).Item(0)
        
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
    
    strTagName = "LTV"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblLTV = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "MPMIGSTARTLTV"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMIGStartLTV = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "MIGTHRESHOLD"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMIGThreshold = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "LOANSNOTREDEEMED"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblLoansNotRedeemed = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "TOTALLOANAMOUNT"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblEligibleAmount = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "PURCHASEPRICE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblPurchasePrice = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "IGNOREMIGPREMIUMAMOUNT"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblIgnoreMigPremAmt = objConvAssist.CSafeDbl(strTagValue)
       
    strTagName = "MINIMUMMIGPREMIUM"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblMinMigPremium = objConvAssist.CSafeDbl(strTagValue)
  
    strTagName = "TYPEOFAPPLICATION"
    strTypeOfApplication = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
     
    strTagName = "ADDIPTTOMIGPREMIUM"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    blnAddIPTToMIG = objConvAssist.CSafeBool(strTagValue)
    
    strTagName = "MPMIGRATESET"
    strRateSet = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    
    strTagName = "MEMBEROFSTAFF"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    blnMemberOfStaff = objConvAssist.CSafeBool(strTagValue)
    
    strTagName = "SHAREDPERCENTAGE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    If strTagValue = "" Or strTagValue = "0" Or strTagValue = Null Then
        strTagValue = "100"
    End If
    dblSharedOwnerShipPercentage = objConvAssist.CSafeDbl(strTagValue) / 100
    
    strTagName = "TOTALORIGINALMIGAMOUNT"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblOriginalMIGAmt = objConvAssist.CSafeDbl(strTagValue)
    '*=[MC]BMIDS763 - APPLICATIONDATE
    strTagName = "APPLICATIONDATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    strApplicationDate = strTagValue
    'SECTION END
'stage1:

    'Calculate The Mortgage Idemnity Amount
    If dblLTV > dblMIGStartLTV Then
        
        If objContext Is Nothing Then
            Set objComboDO = New ComboDO
            Set objGlobalParameter = New GlobalParameterBO
        Else
            Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
            Set objGlobalParameter = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        End If
    
        If objComboDO.IsFurtherAdvance(strTypeOfApplication) = True Then
            blnFurtherAdvance = True
        End If
        
        If blnMemberOfStaff Then
        
            Dim strParamResponse As String
            Dim objGlobalParamDoc  As New FreeThreadedDOMDocument40
            objGlobalParamDoc.validateOnParse = False
            objGlobalParamDoc.setProperty "NewParser", True
            
            strParamResponse = objGlobalParameter.GetCurrentParameter("StaffMIGThreshold")
            Set objGlobalParamDoc = objXmlAssist.load(strParamResponse, TypeName(Me), strFunctionName)
            strTagValue = objXmlAssist.GetTagValue(objGlobalParamDoc.documentElement, "PERCENTAGE")
    
            dblStaffMIGThreashold = objConvAssist.CSafeDbl(strTagValue)

            Set objGlobalParamDoc = Nothing

        Else
            If dblMIGStartLTV <> 0 Then
                dblStaffMIGThreashold = dblMIGStartLTV
            Else
                dblStaffMIGThreashold = dblMIGThreshold
            End If
            
        End If
        dblIndemnityAmount = dblEligibleAmount - ((dblStaffMIGThreashold / 100) * dblPurchasePrice * dblSharedOwnerShipPercentage)
        
        If blnFurtherAdvance Then
            dblIndemnityAmount = dblIndemnityAmount + dblLoansNotRedeemed - dblOriginalMIGAmt
        End If
            
'Stage2:
    'Calculate Indemnity Premium
        '*=[MC]BMIDS763 - GetMIGRate Member Function Extended to pass ApplicationDate
        Set xmlMIGRateNode = GetMIGRate(strRateSet, dblEligibleAmount, dblLTV, strApplicationDate)
        'SECTION END
        objErrAssist.CheckResponse xmlMIGRateNode, xmlResponseElem, True
        
        dblIndemnityRate = objConvAssist.CSafeDbl(xmlMIGRateNode.selectSingleNode(".//INDEMNITYRATE").Text)
        dblMinimumPremium = objConvAssist.CSafeDbl(xmlMIGRateNode.selectSingleNode(".//MINIMUMPREMIUM").Text)
        
        dblMIGPremium = (dblIndemnityAmount * dblIndemnityRate) / 100
        
        If blnMemberOfStaff Then

            strParamResponse = objGlobalParameter.GetCurrentParameter("StaffMIGDiscount")
            Set objGlobalParamDoc = objXmlAssist.load(strParamResponse, TypeName(Me), strFunctionName)

            strTagValue = objXmlAssist.GetTagValue(objGlobalParamDoc.documentElement, "AMOUNT")

            'BS BM0151 20/02/03
            If strTagValue <> "" Then
                'Staff discount is a set amount
                dblDiscount = objConvAssist.CSafeDbl(strTagValue)
            Else
                'Staff discount is a percentage
                strTagValue = objXmlAssist.GetTagValue(objGlobalParamDoc.documentElement, "PERCENTAGE")
                dblDiscount = (dblMIGPremium * objConvAssist.CSafeDbl(strTagValue)) / 100
            End If
            'BS BM0151 End 20/02/03

            dblMIGPremium = dblMIGPremium - dblDiscount
        End If
        
        If dblMIGPremium > 0 Then
        
            If dblMinimumPremium <> 0 Then
            
                If dblMIGPremium < dblMinimumPremium Then
                   dblMIGPremium = dblMinimumPremium
                End If
                
            End If
            
            dblActualMIGPremium = Int(dblMIGPremium + 0.5)
        Else
            dblActualMIGPremium = 0
        End If
        
    Else
    
        dblActualMIGPremium = 0
        dblIndemnityAmount = 0
    
    End If
    
'Stage3:

    Set xmlMIGElem = xmlOut.createElement("MIG")
    xmlResponseElem.appendChild xmlMIGElem
    
    Set xmlElement = xmlOut.createElement("MIGPREMIUM")
    xmlElement.Text = CStr(dblActualMIGPremium)
    xmlMIGElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("TOTALINDEMNITY")
    xmlElement.Text = CStr(dblIndemnityAmount)
    xmlMIGElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("IPT")
    xmlElement.Text = "0"
    xmlMIGElem.appendChild xmlElement
        
    CalcMIG = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalcMIGExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlMIGElem = Nothing
    Set xmlNode = Nothing
    Set xmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objComboDO = Nothing
    Set xmlElement = Nothing
    Set xmlRequestElem = Nothing

    Exit Function

CalcMIGVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlRequestElem
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    CalcMIG = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalcMIGExit

End Function


'*=[MC]Application Date Input parameter added to the method signature
Private Function GetMIGRate(ByVal vstrRateSet As String, _
                            ByVal vdblTotalLoanAmount As Double, _
                            ByVal vdblLTV, _
                            ByVal vstrApplicationDate) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the MIG Rate
' pass:         vstrRateSet           Rateset
'               vdblTotalLoanAmount     Total Loan Amount
'               vdblLTV                 LTV
'
' return:       GetMIGRate              The MIG Rate
'
' Raise Errors:
'------------------------------------------------------------------------------------------
    On Error GoTo GetMIGRateVbErr
            
    Const strFunctionName As String = "GetMIGRate"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objMortgageProductDO As MortgageProductDO
    Dim objConvAssist As New ConvertAssist
    
    Dim objMIGDoc As FreeThreadedDOMDocument40
    Dim xmlMIGResponse As IXMLDOMElement
    Dim strTagValue As String
    Dim xmlResponseElem  As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTempResponse As IXMLDOMNode
    Dim objRequestDoc As New FreeThreadedDOMDocument40
    objRequestDoc.validateOnParse = False
    objRequestDoc.setProperty "NewParser", True
    Dim objElement As IXMLDOMElement
    Dim objMIGElement As IXMLDOMElement
    Dim objTableElement As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Set up the MIGRateDO request
    Set objMIGElement = objRequestDoc.createElement("MIGRATE")
    objRequestDoc.appendChild objMIGElement
    
    Set objElement = objRequestDoc.createElement("RATESET")
    objElement.Text = vstrRateSet
    objMIGElement.appendChild objElement
        
    Set objElement = objRequestDoc.createElement("TOTALLOANAMOUNT")
    objElement.Text = CStr(vdblTotalLoanAmount)
    objMIGElement.appendChild objElement
    
    Set objElement = objRequestDoc.createElement("LTV")
    objElement.Text = CStr(vdblLTV)
    objMIGElement.appendChild objElement
    '*=[MC]BMIDS763
    Set objElement = objRequestDoc.createElement("APPLICATIONDATE")
    objElement.Text = CStr(vstrApplicationDate)
    objMIGElement.appendChild objElement
    'SECTION END

    If objContext Is Nothing Then
        Set objMortgageProductDO = New MortgageProductDO
    Else
        Set objMortgageProductDO = objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductDO")
    End If
    
    Set xmlTempResponse = objMortgageProductDO.GetMIGRate(objRequestDoc.documentElement)
    xmlResponseElem.appendChild xmlTempResponse
    
    Set GetMIGRate = xmlResponseElem
    
    Set objRequestDoc = Nothing
    Set objElement = Nothing
    Set objMIGElement = Nothing
    Set objTableElement = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objMortgageProductDO = Nothing
    Set objConvAssist = Nothing

    Exit Function
    
GetMIGRateVbErr:

    Set objRequestDoc = Nothing
    Set objElement = Nothing
    Set objMIGElement = Nothing
    Set objTableElement = Nothing
    Set objXmlAssist = Nothing
    Set objMortgageProductDO = Nothing
    Set objConvAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function GetInsurancePremiumTaxRate() As Double
' header ----------------------------------------------------------------------------------
' description:  Gets the Insurance Premium Tax Rate from the Global Parameters
' pass:         n/a
' return:       GetInsurancePremiumTaxRate  The Insurance Premium Tax Rate
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetInsurancePremiumTaxRateVbErr
                      
    Const strFunctionName As String = "GetInsurancePremiumTaxRate"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objGlobalParamDoc As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objGlobalParamDO As GlobalParameterDO
        
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    Dim strParamResponse As String
    
    strParamResponse = objGlobalParamDO.GetCurrentParameter("MIGInsPremTaxRate")
    Set objGlobalParamDoc = objXmlAssist.load(strParamResponse, _
                                              TypeName(Me), _
                                              strFunctionName)
    
    Dim strTagValue As String
    strTagValue = objXmlAssist.GetTagValue(objGlobalParamDoc.documentElement, "PERCENTAGE")
    
    GetInsurancePremiumTaxRate = objConvAssist.CSafeDbl(strTagValue)

    Set objGlobalParamDoc = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set objConvAssist = Nothing

    Exit Function
    
GetInsurancePremiumTaxRateVbErr:

    Set objGlobalParamDoc = Nothing
    Set objXmlAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set objConvAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Function


Public Function RoundNumber(number As Double) As Integer
    If Int(number + 0.5) > Int(number) Then
        RoundNumber = Int(number) + 1
    Else
        RoundNumber = Int(number)
    End If
End Function

Private Sub ObjectControl_Activate()
    Set m_objGeneralAssist = New GeneralAssist
End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function


Private Sub ObjectControl_Deactivate()
    Set m_objGeneralAssist = Nothing
End Sub
