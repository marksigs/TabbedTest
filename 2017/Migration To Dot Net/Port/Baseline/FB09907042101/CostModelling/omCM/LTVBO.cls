VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LTVBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LTVBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   LTV Business Object

'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' PSC   17/09/99 Created
' DRC    3/10/01 SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
' Epsom-Specific History:
'
' Prog  Date     Description
' SAB   22/06/06 EP837 Format LTV to two significant numbers.  NOTE: This overrides MAR1648 which
'                  formats the value to 1 decimal place before storing it in MORTGAGESUBQUOTE
' GHun  17/11/06 EP2_19 Change CalcCostModelLTV for Right-to-Buy
' PSC   05/02/07 EP2_1113 Change CalcCostModelLTV for TOE
' PSC   15/02/07 EP2_1239 Change CalcCostModelLTV to take into account discount correctly
'------------------------------------------------------------------------------------------



Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================
Private Const cstrCALC_NAME = "LTV"

 '=============================================
 'Variable Declaration Section
 '=============================================


'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     01/10/2002  BMIDS00534  Amended CalcCostModelLTV
'MV     11/10/2002  BMIDS00590  Modified CalcCostModelLTV
'------------------------------------------------------------------------------------------

Public Function CalcMortCalcLTV(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  calculates LTV for mortgage calculator
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Request Format:
'
'       <LTV>
'           <AMOUNTREQUESTED>Value</AMOUNTREQUESTED>
'           <PURCHASEPRICE>Value</PURCHASEPRICE>
'       </LTV>
'------------------------------------------------------------------------------------------
On Error GoTo CalcMortCalcLTVVbErr
    
    Const strFunctionName As String = "CalcMortCalcLTV"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlNode As IXMLDOMNode
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objElement As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrCALC_NAME).Item(0)
    
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
    
    Dim strTagValue As String
    Dim lngAmountRequested As Long
    Dim lngPurchasePrice As Long
    
    strTagValue = objXmlAssist.GetTagValue(xmlIn.documentElement, "AMOUNTREQUESTED")
    lngAmountRequested = objConvAssist.CSafeLng(strTagValue)
    strTagValue = objXmlAssist.GetTagValue(xmlIn.documentElement, "PURCHASEPRICE")
    lngPurchasePrice = objConvAssist.CSafeLng(strTagValue)
   
    If lngAmountRequested = 0 Or lngPurchasePrice = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 163
    End If
    
    Dim dblLTV As Double
    
    ' Format does the same as adding 0.005
    dblLTV = (lngAmountRequested / lngPurchasePrice) * 100
    
    Set objElement = xmlOut.createElement("LTV")
    objElement.Text = Format(dblLTV, "0.00")
    
    xmlResponseElem.appendChild objElement
        
    CalcMortCalcLTV = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalcMortCalcLTVExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlRequestElem = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing

    Exit Function

CalcMortCalcLTVVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    CalcMortCalcLTV = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalcMortCalcLTVExit

End Function

Public Function CalcCostModelLTV(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  calculates LTV for cost modelling
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Request Format:
'
'       <LTV>
'           <TYPEOFAPPLICATION>Value</TYPEOFAPPLICATION>
'           <PURCHASEPRICE>Value<PURCHASEPRICE>
'           <AMOUNTREQUESTED>Value</AMOUNTREQUESTED>
'           <LOANSNOTREDEEMED>Value</LOANSNOTREDEEMED>
'           <SHAREDOWNERSHIPTYPE>Value</SHAREDOWNERSHIPTYPE>
'           <SHAREDOWNERSHIPAMOUNT>Value</SHAREDOWNERSHIPAMOUNT>
'           <DISCOUNTAMOUNT>Value</DISCOUNTAMOUNT>
'       </LTV>
'------------------------------------------------------------------------------------------
On Error GoTo CalcCostModelLTVVbErr
    
    Const strFunctionName As String = "CalcCostModelLTV"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlNode As IXMLDOMNode
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objElement As IXMLDOMElement
    Dim objComboDO As ComboDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlRequestElem = xmlIn.getElementsByTagName(cstrCALC_NAME).Item(0)
    
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
    
    Dim strTagValue As String
    Dim strTagName As String
    Dim strTypeOfApp As String
    
    Dim dblPurchasePrice As Double
    Dim dblAmountRequested As Double
    Dim dblLoansNotRedeemed As Double
    Dim dblSharedOwnershipAmt As Double
    Dim dblDiscountAmount As Double
    Dim dblPurchaseAmount As Double
    Dim dblLTV As Double
    
    Dim intSharedOwnershipType As Integer
    
    Dim blnIsFurtherAdvance As Boolean
    Dim blnIsNewLoan As Boolean
    Dim blnIsRemortgage As Boolean
    
    Dim blnIsTOE As Boolean         ' PSC 05/02/2007 EP2_1113
    
    strTagName = "TYPEOFAPPLICATION"
    strTypeOfApp = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    
    strTagName = "PURCHASEPRICE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblPurchasePrice = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "AMOUNTREQUESTED"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblAmountRequested = objConvAssist.CSafeDbl(strTagValue)
       
    strTagName = "LOANSNOTREDEEMED"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblLoansNotRedeemed = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "SHAREDOWNERSHIPTYPE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    intSharedOwnershipType = objConvAssist.CSafeInt(strTagValue)
    
    'strTagName = "SHAREDOWNERSHIPAMOUNT"
    'MV - 01/10/2002 - BMISS00536
    'strTagName = "SHAREDAMOUNT"
    strTagName = "SHAREDPERCENTAGE"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblSharedOwnershipAmt = (objConvAssist.CSafeDbl(strTagValue) / 100) * dblPurchasePrice
    
    strTagName = "DISCOUNTAMOUNT"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblDiscountAmount = objConvAssist.CSafeDbl(strTagValue)
    
    ' Must have a type of application. purchase price and amount requested
    If Len(strTypeOfApp) = 0 Or dblPurchasePrice = 0 Or dblAmountRequested = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "TYPEOFAPPLICATION = " & strTypeOfApp & _
                                "; PURCHASEPRICE = " & CStr(dblPurchasePrice) & _
                                "; AMOUNTREQUESTED = " & CStr(dblAmountRequested)
    End If
    
    ' If we have a Shared Ownership Type then must have a Shared Ownership Amount
    If intSharedOwnershipType > 0 And dblSharedOwnershipAmt = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "SHAREDOWNERSHIPTYPE = " & CStr(intSharedOwnershipType) & _
                                "; SHAREDPERCENTAGE = " & CStr(dblSharedOwnershipAmt)
    End If
    
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
     
    If intSharedOwnershipType > 0 Then
        dblPurchaseAmount = dblSharedOwnershipAmt
    Else
        dblPurchaseAmount = dblPurchasePrice
    End If
    
    blnIsFurtherAdvance = objComboDO.IsFurtherAdvance(strTypeOfApp)
    
    If blnIsFurtherAdvance = True Then
        dblLTV = ((dblAmountRequested + dblLoansNotRedeemed) * 100) / dblPurchaseAmount
    Else
        blnIsNewLoan = objComboDO.IsNewLoan(strTypeOfApp)
        blnIsRemortgage = objComboDO.IsRemortgage(strTypeOfApp)
        ' PSC 05/02/2007 EP2_1113
        blnIsTOE = IsValidationType("TypeOfMortgage", CSafeInt(strTypeOfApp), "T")
    
        ' PSC 05/02/2007 EP2_1113
        If blnIsNewLoan = True Or blnIsRemortgage = True Or blnIsTOE Then
            If GetGlobalParamBoolean("CMUsesDiscountPurchasePrice") = "1" Then  'EP2_19 GHun
                ' PSC 15/02/2007 EP2_1239
                dblLTV = (dblAmountRequested * 100) / (dblPurchaseAmount - dblDiscountAmount)
            'EP2_19 GHun
            Else
                dblLTV = (dblAmountRequested * 100) / dblPurchaseAmount
            End If
            'EP2_19 End
        End If
    End If
    
    Set objElement = xmlOut.createElement("LTV")
    objElement.Text = Format(dblLTV, "0.00")
    
    xmlResponseElem.appendChild objElement
        
    CalcCostModelLTV = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalcCostModelLTVExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set objErrAssist = Nothing
    Set xmlNode = Nothing
    Set xmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objElement = Nothing
    Set objComboDO = Nothing

    Exit Function

CalcCostModelLTVVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    CalcCostModelLTV = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalcCostModelLTVExit

End Function


