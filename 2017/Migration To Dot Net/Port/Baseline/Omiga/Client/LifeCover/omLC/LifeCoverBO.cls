VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LifeCoverBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LifeCoverBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Life Cover Business Object which 'supports' transactions
'Dependencies:  LifeCoverDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    25/10/99    Created
'SR     10-12-99    Added function 'GetComboValue'
'                   Modified function 'CalculatePremiumsForCoverType'
'SR     17-02-00    Modified method - ValidateCalcuatePremiumsInput - AQR - SYS0175
'JLD    17/05/00    SYS0725 change tag names of returned xml in CalculatePremiumsForCovertype
'MS     19/07/00    Modified for performance enhancements
'DRC    04/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
Implements ILifeCoverBO
'MS 19/07/00 Modified for performance enhancements
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrCALC_NAME = "LIFECOVER"
'MS 19/07/00 Modified for performance enhancements

Private Function ILifeCoverBO_CalculatePremiums(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Calculates the life cover premiums for each applicant and joint premiums
'               if required
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'
'
'   Request Format:
'
'                   <LIFECOVER>
'                       <ENDOWMENTNEED>Value</ENDOWMENTNEED>
'                       <TERMASSURANCENEED>Value</TERMASSURANCENEED>
'                       <INTERESTONLYTERM>Value</INTERESTONLYTERM>
'                       <CAPITALANDINTERESTTERM>Value</CAPITALANDINTERESTTERM>
'                       <RELATIONSHIP>Value</RELATIONSHIP>
'                       <APPLICANT1SMOKERSTATUS>Value</APPLICANT1SMOKERSTATUS>
'                       <APPLICANT1GOODHEALTH>Value</APPLICANT1GOODHEALTH>
'                       <APPLICANT1GENDER>Value</APPLICANT1GENDER>
'                       <APPLICANT1AGE>Value</APPLICANT1AGE>
'                       <APPLICANT2SMOKERSTATUS>Value</APPLICANT2SMOKERSTATUS>
'                       <APPLICANT2GOODHEALTH>Value</APPLICANT2GOODHEALTH>
'                       <APPLICANT2GENDER>Value</APPLICANT2GENDER>
'                       <APPLICANT2AGE>Value</APPLICANT2AGE>
'                   </LIFECOVER>
'
'------------------------------------------------------------------------------------------
On Error GoTo ILifeCoverBO_CalculatePremiumsVbErr
    
    Const strFunctionName As String = "ILifeCoverBO_CalculatePremiums"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    'MS 19/07/00 Modified for performance enhancements
'    Dim xmlIn As New DOMDocument
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlApplicant1Elem As IXMLDOMElement
    Dim xmlJointElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim objConvAssist As ConvertAssist
    Set objConvAssist = New ConvertAssist
    
    Dim objGlobalParamDO As GlobalParameterDO
'MS 19/07/00 Modified for performance enhancements
'    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlRequestElem = vxmlRequest.getElementsByTagName(cstrCALC_NAME).Item(0)
        
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
    
    ValidateCalculatePremiumsInput vxmlRequest
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
                    objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
                
    Dim dblLifeCoverIPTRate As Double
    
    dblLifeCoverIPTRate = _
            objGlobalParamDO.GetCurrentParameterByType("LifeCoverIPTRate", "PERCENTAGE")
    
    Dim strTagValue As String
    Dim strTagName As String
    Dim strRelationship As String
    Dim strInterestOnlyTerm As String
    Dim strCapAndIntTerm As String
    Dim strApp1Gender As String
    Dim strApp2Gender As String
    Dim blnApp1IsSmoker As Boolean
    Dim blnApp2IsSmoker As Boolean
    Dim blnApp1IsInGoodHealth As Boolean
    Dim blnApp2IsInGoodHealth As Boolean
    Dim blnIsSmoker As Boolean
    Dim blnIsInGoodHealth As Boolean
    Dim dblEndowmentNeed As Double
    Dim dblTermAssuranceNeed As Double
    Dim intApp1Age As Integer
    Dim intApp2Age As Integer
    Dim dblEndowmentPremium As Double
    Dim dblTermAssurancePremium As Double
    
    'Set xmlApplicant1Elem = xmlOut.createElement("APPLICANT1")
    'Set xmlJointElem = xmlOut.createElement("JOINT")
    Set xmlApplicant1Elem = xmlOut.createElement("LIFEBENEFITS")
    Set xmlJointElem = xmlOut.createElement("LIFEBENEFITS")
    
    strTagName = "ENDOWMENTNEED"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblEndowmentNeed = objConvAssist.CSafeDbl(strTagValue)
        
    strTagName = "TERMASSURANCENEED"
    strTagValue = objXmlAssist.GetTagValue(xmlRequestElem, strTagName)
    dblTermAssuranceNeed = objConvAssist.CSafeDbl(strTagValue)
    
    ' Calculate endowmnent premiums
    If dblEndowmentNeed > 0 Then
        dblEndowmentPremium = _
                    CalculatePremiumsForCoverType("E", vxmlRequest, _
                                                  dblEndowmentNeed, dblLifeCoverIPTRate, _
                                                  xmlJointElem, xmlApplicant1Elem)
    End If
    
     ' Calculate Term Assurance premiums
    If dblTermAssuranceNeed > 0 Then
        dblTermAssurancePremium = _
                    CalculatePremiumsForCoverType("T", vxmlRequest, _
                                                   dblTermAssuranceNeed, dblLifeCoverIPTRate, _
                                                   xmlJointElem, xmlApplicant1Elem)
    End If
    
    If xmlJointElem.hasChildNodes = True Then
        xmlResponseElem.appendChild xmlJointElem
    End If
    
    If xmlApplicant1Elem.hasChildNodes = True Then
        xmlResponseElem.appendChild xmlApplicant1Elem
    End If
    
    Set xmlElement = xmlOut.createElement("TOTALLIFEMONTHLYCOST")
    xmlElement.Text = Format(dblTermAssurancePremium + dblEndowmentPremium, "0.00")
    xmlResponseElem.appendChild xmlElement

    Set ILifeCoverBO_CalculatePremiums = xmlOut.documentElement
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

ILifeCoverBO_CalculatePremiumsExit:
    
'    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlApplicant1Elem = Nothing
    Set xmlJointElem = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlElement = Nothing

    Exit Function

ILifeCoverBO_CalculatePremiumsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ILifeCoverBO_CalculatePremiums = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume ILifeCoverBO_CalculatePremiumsExit

End Function

'MS 19/07/00 Modified for performance enhancements
'Private Sub ValidateCalculatePremiumsInput(ByVal vstrXMLRequest As String)
Private Sub ValidateCalculatePremiumsInput(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Validates that the correct parameters are present to calculate the
'               payment protection costs
'
' pass:         vstrXMLRequest CalcCosts Input xml Request
'
' return n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCalculatePremiumsInputVbErr
                     
    Const strFunctionName As String = "ValidateCalculatePremiumsInput"

'    Dim xmlIn As New DOMDocument
    Dim xmlRequestElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    
    'MS 19/07/00 Modified for performance enhancements
'    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    Set xmlRequestElem = vxmlRequest.getElementsByTagName(cstrCALC_NAME).Item(0)
    
    Dim strXMLData As String
    
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrCALC_NAME & " tag not found"
    End If
    
    Dim strEndowmentNeed As String
    Dim strTermAssuranceNeed As String
    Dim strInterestOnlyTerm As String
    Dim strCapAndIntTerm As String
    Dim strRelationship As String
    Dim strApp1Smoker As String
    Dim strApp2Smoker As String
    Dim strApp1GoodHealth As String
    Dim strApp2GoodHealth As String
    Dim strApp1Gender As String
    Dim strApp2Gender As String
    Dim strApp1Age As String
    Dim strApp2Age As String
    
    strEndowmentNeed = objXmlAssist.GetTagValue(xmlRequestElem, "ENDOWMENTNEED")
    strTermAssuranceNeed = objXmlAssist.GetTagValue(xmlRequestElem, "TERMASSURANCENEED")
    strInterestOnlyTerm = objXmlAssist.GetTagValue(xmlRequestElem, "INTERESTONLYTERM")
    strCapAndIntTerm = objXmlAssist.GetTagValue(xmlRequestElem, "CAPITALANDINTERESTTERM")
    strRelationship = objXmlAssist.GetTagValue(xmlRequestElem, "RELATIONSHIP")
    strApp1Smoker = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT1SMOKERSTATUS")
    strApp1GoodHealth = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT1GOODHEALTH")
    strApp1Gender = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT1GENDER")
    strApp1Age = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT1AGE")
    strApp2Smoker = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT2SMOKERSTATUS")
    strApp2GoodHealth = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT2GOODHEALTH")
    strApp2Gender = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT2GENDER")
    strApp2Age = objXmlAssist.GetTagValue(xmlRequestElem, "APPLICANT2AGE")

    ' Must have an endowment or term assurance need
    If objConvAssist.CSafeDbl(strEndowmentNeed) = 0 And _
       objConvAssist.CSafeDbl(strTermAssuranceNeed) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "ENDOWMENTNEED = " & strEndowmentNeed & _
                                ";TERMASSURANCENEED = " & strTermAssuranceNeed
    End If
    
    ' Must have an Interest only term if there is an endowment need
    If objConvAssist.CSafeDbl(strEndowmentNeed) > 0 And _
       objConvAssist.CSafeInt(strInterestOnlyTerm) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "ENDOWMENTNEED = " & strEndowmentNeed & _
                                ";INTERESTONLYTERM = " & strInterestOnlyTerm
    End If
    
    ' Must have a Capital and Interest term if there is a Term Assurance need
    If objConvAssist.CSafeDbl(strTermAssuranceNeed) > 0 And _
       objConvAssist.CSafeInt(strCapAndIntTerm) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "TERMASSURANCENEED = " & strTermAssuranceNeed & _
                                ";CAPITALANDINTERESTTERM = " & strCapAndIntTerm
    End If
    
    ' Must have details for applicant 1
    If Len(strApp1Gender) = 0 Or objConvAssist.CSafeInt(strApp1Age) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICANT1GENDER = " & strApp1Gender & _
                                ";APPLICANT1AGE = " & strApp1Age
    End If
    
    ' Must have all applicant 2 details if they are to be used
    If (Len(strApp2Gender) > 0 And objConvAssist.CSafeInt(strApp2Age) = 0) Or _
       (Len(strApp2Gender) = 0 And objConvAssist.CSafeInt(strApp2Age) > 0) Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICANT2GENDER = " & strApp2Gender & _
                                ";APPLICANT2AGE = " & strApp2Age
    End If
      
    Exit Sub
    
ValidateCalculatePremiumsInputVbErr:
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description


End Sub

Private Sub GetRates(ByVal vstrCoverType As String, _
                     ByVal vstrGender As String, _
                     ByVal vstrAge As String, _
                     ByVal vstrTerm As String, _
                     ByRef rdblSmokerRate As Double, _
                     ByRef rdblAnnualRate As Double, _
                     ByRef rdblBadHealthRate As Double)
' header ----------------------------------------------------------------------------------
' description:  Gets the Life Cover rates
'
' pass:         vstrCoverType        Type of Cover required
'               vstrGender           Gender Combo ValueId for applicant
'               vstrAge              Applicant age
'               vstrTerm             Term of life cover
'               rdblSmokerRate       Additional rate for smoker
'               rdblAnnualRate       Annual life cover rate
'               rdblBadHealthRate    Additional rate for bad health
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
'    Dim xmlRatesDoc As DOMDocument
    Dim xmlRates As IXMLDOMElement
    Dim xmlTableElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objConvAssist As New ConvertAssist
    Dim objILifeCoverDO As ILifeCoverDO
    
    Dim strTagName As String
    Dim strTagValue As String
    
    If objContext Is Nothing Then
        Set objILifeCoverDO = New LifeCoverDO
    Else
        Set objILifeCoverDO = objContext.CreateInstance(App.Title & ".LifeCoverDO")
    End If
    
    Set xmlTableElement = xmlRequestDoc.createElement("LIFECOVERRATES")
    xmlRequestDoc.appendChild xmlTableElement
    
    strTagName = "COVERTYPE"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrCoverType
    xmlTableElement.appendChild xmlElement
                         
    strTagName = "APPLICANTSGENDER"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrGender
    xmlTableElement.appendChild xmlElement
                     
    strTagName = "APPLICANTSAGE"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrAge
    xmlTableElement.appendChild xmlElement
    
    strTagName = "TERM"
    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    xmlElement.Text = vstrTerm
    xmlTableElement.appendChild xmlElement
                    
'    Dim strResponse As String
'    strResponse = objLifeCoverDO.GetRates(xmlRequestDoc.xml)
'    Set xmlRatesDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
'MS 19/07/00 Modified for performance enhancements
    Set xmlRates = objILifeCoverDO.GetRates(xmlRequestDoc.documentElement)
    
    strTagName = "ADDITIONALSMOKERRATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRates, strTagName)
    rdblSmokerRate = objConvAssist.CSafeDbl(strTagValue)
                     
    strTagName = "ANNUALRATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRates, strTagName)
    rdblAnnualRate = objConvAssist.CSafeDbl(strTagValue)
                     
    strTagName = "NOTGOODHEALTHRATE"
    strTagValue = objXmlAssist.GetTagValue(xmlRates, strTagName)
    rdblBadHealthRate = objConvAssist.CSafeDbl(strTagValue)
                                   
    Set xmlRequestDoc = Nothing
    Set xmlRates = Nothing
    Set xmlTableElement = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objILifeCoverDO = Nothing
                     
    Exit Sub
                     
GetRatesVbErr:
    
    Set xmlRequestDoc = Nothing
    Set xmlRates = Nothing
    Set xmlTableElement = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objILifeCoverDO = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
                     
End Sub

Private Sub CalculatePremiumAndIPT(ByVal vstrCoverType As String, _
                                   ByVal vstrGender As String, _
                                   ByVal vstrAge As String, _
                                   ByVal vstrTerm As String, _
                                   ByVal vblnIsSmoker As Boolean, _
                                   ByVal vblnIsInGoodHealth As Boolean, _
                                   ByVal vdblLifeCoverIPTRate As Double, _
                                   ByVal vdblAmount As Double, _
                                   ByRef rdblPremium As Double, _
                                   ByRef rdblInsurancePremiumTax As Double)
' header ----------------------------------------------------------------------------------
' description:
'   Calculates the Life Cover Premium and Insurance Premium Tax
'
' pass: vstrCoverType           The value id of the type of cover required
'       vstrGender              Gender of applicant
'       vstrAge                 Age of applicant
'       vstrTerm                Term of cover
'       vblnIsSmoker            Whether the applicant smokes
'       vblnIsInGoodHealth      Whether the applicant is in good health
'       vdblLifeCoverIPTRate    The Insurance Premium Tax Rate
'       vdblAmount              The amount of cover required
'       rdblPremium             Calcultated Premium
'       rdblInsurancePremiumTax Calculated Insurance Premium Tax
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CalculatePremiumAndIPTVbErr

    Const strFunctionName As String = "CalculatePremiumAndIPT"

    Dim dblAnnualRate As Double
    Dim dblSmokerRate As Double
    Dim dblBadHealthRate As Double
    
    Dim objErrAssist As New ErrAssist
    
    rdblPremium = 0
    rdblInsurancePremiumTax = 0

    GetRates vstrCoverType, vstrGender, vstrAge, vstrTerm, _
             dblSmokerRate, dblAnnualRate, dblBadHealthRate

    If vblnIsSmoker = True Then
        dblAnnualRate = dblAnnualRate + dblSmokerRate
    End If
    
    If vblnIsInGoodHealth = False Then
        dblAnnualRate = dblAnnualRate + dblBadHealthRate
    End If
    
    rdblPremium = dblAnnualRate * vdblAmount / (1000 * 12)
    rdblInsurancePremiumTax = rdblPremium * vdblLifeCoverIPTRate / 100

    Set objErrAssist = Nothing
     
    Exit Sub

CalculatePremiumAndIPTVbErr:

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'Private Function CalculatePremiumsForCoverType(ByVal vstrCoverType As String, _
'                                          ByVal vstrXMLRequest As String, _
'                                          ByVal vdblAmount As Double, _
'                                          ByVal vdblLifeCoverIPTRate, _
'                                          ByVal vxmlJointElem As IXMLDOMElement, _
'                                          ByVal vxmlApplicant1Elem As IXMLDOMElement) As Double
'MS 19/07/00 Modified for performance enhancements
Private Function CalculatePremiumsForCoverType(ByVal vstrCoverType As String, _
                                          ByVal vxmlRequest As IXMLDOMElement, _
                                          ByVal vdblAmount As Double, _
                                          ByVal vdblLifeCoverIPTRate, _
                                          ByVal vxmlJointElem As IXMLDOMElement, _
                                          ByVal vxmlApplicant1Elem As IXMLDOMElement) As Double

' header ----------------------------------------------------------------------------------
' description:
'   Calculates the Life Cover Premium and Insurance Premium Tax for the particular cover
'   type and add the results to the appropriate xml element
'
' pass: vstrCoverType           The cover type required for the calculations
'       vstrXMLRequest          xml Request
'       vdblAmount              Amount of Cover Required
'       vdblLifeCoverIPTRate    Life cover insurance premium tax rate
'       vxmlJointElem           xml element to which joint life premiums are to be added
'       vxmlApplicant1Elem      xml element to which applicant1 premiums are to be added
'
' return:       n/a
'------------------------------------------------------------------------------------------
                                                
On Error GoTo CalculatePremiumsForCoverTypeVbErr
    
    Const strFunctionName As String = "CalculatePremiumsForCoverType"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
'    Dim xmlIn As New DOMDocument
    Dim xmlElement As IXMLDOMElement
    Dim xmlBenefitElem As IXMLDOMElement
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim objConvAssist As ConvertAssist
    Set objConvAssist = New ConvertAssist
    
    Dim objComboDO As ComboDO
        
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    
    Dim strTagValue As String
    Dim strTagName As String
    Dim strRelationship As String
    Dim strInterestOnlyTerm As String
    Dim strCapAndIntTerm As String
    Dim strApp1Gender As String
    Dim strApp2Gender As String
    Dim blnApp1IsSmoker As Boolean
    Dim blnApp2IsSmoker As Boolean
    Dim blnApp1IsInGoodHealth As Boolean
    Dim blnApp2IsInGoodHealth As Boolean
    Dim blnIsSmoker As Boolean
    Dim blnIsInGoodHealth As Boolean
    Dim intApp1Age As Integer
    Dim intApp2Age As Integer
 'MS 19/07/00 Modified for performance enhancements
'    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    
    strTagName = "APPLICANT1SMOKERSTATUS"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    blnApp1IsSmoker = objConvAssist.CSafeBool(strTagValue)
    
    strTagName = "APPLICANT2SMOKERSTATUS"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    blnApp2IsSmoker = objConvAssist.CSafeBool(strTagValue)
    
    strTagName = "APPLICANT1GOODHEALTH"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    blnApp1IsInGoodHealth = objConvAssist.CSafeBool(strTagValue)
    
    strTagName = "APPLICANT2GOODHEALTH"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    blnApp2IsInGoodHealth = objConvAssist.CSafeBool(strTagValue)
    
    strTagName = "APPLICANT1GENDER"
    strApp1Gender = objXmlAssist.GetTagValue(vxmlRequest, strTagName)

    strTagName = "APPLICANT2GENDER"
    strApp2Gender = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    
    strTagName = "APPLICANT1AGE"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    intApp1Age = objConvAssist.CSafeInt(strTagValue)
    
    strTagName = "APPLICANT2AGE"
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    intApp2Age = objConvAssist.CSafeInt(strTagValue)
    
    strTagName = "INTERESTONLYTERM"
    strInterestOnlyTerm = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    
    strTagName = "CAPITALANDINTERESTTERM"
    strCapAndIntTerm = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
        
    strTagName = "RELATIONSHIP"
    strRelationship = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    
    blnIsSmoker = False
    blnIsInGoodHealth = True
    
    ' Set smoker to true if either applicant smokes
    If blnApp1IsSmoker = True Or blnApp2IsSmoker = True Then
        blnIsSmoker = True
    End If
    
    'Set Good Health to false if Applicant 1 is false or Applicant 2 details have been entered
    ' and they are false
    If blnApp1IsInGoodHealth = False Or (intApp2Age > 0 And blnApp2IsInGoodHealth = False) Then
        blnIsInGoodHealth = False
    End If
    
    Dim strCoverTypeValue As String
    Dim strTerm As String
    
    ' Set up variables depending on whether we are calculating for term assurance
    ' or endowment
    
    If StrComp(vstrCoverType, "E", vbTextCompare) = 0 Then
'MS 19/07/00 Modified for performance enhancements
'        strCoverTypeValue = GetEndowmentCoverValueId()
'        strCoverTypeValue = objComboDO.GetFirstComboValueId("LifeCoverType", "E")
        strTerm = strInterestOnlyTerm
    ElseIf StrComp(vstrCoverType, "T", vbTextCompare) = 0 Then
'MS 19/07/00 Modified for performance enhancements
'        strCoverTypeValue = GetTermAssuranceCoverValueId()
'        strCoverTypeValue = objComboDO.GetFirstComboValueId("LifeCoverType", "T")
        strTerm = strCapAndIntTerm
    Else
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "Cover Type = " & vstrCoverType
    End If
    
    strCoverTypeValue = objComboDO.GetFirstComboValueId("LifeCoverType", vstrCoverType)
    
    Dim dblPremium As Double
    Dim dblIPT As Double
    Dim dblApp1Premium As Double
    Dim dblApp2Premium As Double
    Dim dblApp1IPT As Double
    Dim dblApp2IPT As Double
        
    ' Joint relationship
    If StrComp(strRelationship, "Joint", vbTextCompare) = 0 Then
        Dim intAverageAge As Integer, intGender As Integer
        
        ' Take the average age and round up
        intAverageAge = ((intApp1Age + intApp2Age) / 2) + 0.5
    
        ' Calculate premium and IPT for joint endowment
        
        ' SR 10-12-99. Find the ComoboId corresponding to validation 'Joint' and pass it
'MS 19/07/00 Modified for performance enhancements
'        intGender = GetComboValue(strRelationship, "LifeCoverGender")
        intGender = CInt(objComboDO.GetFirstComboValueId("LifeCoverGender", "J"))
        
        CalculatePremiumAndIPT strCoverTypeValue, intGender, CStr(intAverageAge), _
                               strTerm, blnIsSmoker, blnIsInGoodHealth, _
                               vdblLifeCoverIPTRate, vdblAmount, dblPremium, dblIPT
                               
        Set xmlBenefitElem = vxmlJointElem.ownerDocument.createElement("LIFEBENEFIT")
        vxmlJointElem.appendChild xmlBenefitElem
        xmlBenefitElem.setAttribute "TYPE", vstrCoverType
                
        Set xmlElement = vxmlJointElem.ownerDocument.createElement("MONTHLYCOST")
        xmlElement.Text = Format(dblPremium, "0.00")
        xmlBenefitElem.appendChild xmlElement
        
        Set xmlElement = vxmlJointElem.ownerDocument.createElement("IPT")
        xmlElement.Text = Format(dblIPT, "0.00")
        xmlBenefitElem.appendChild xmlElement
        
    Else
        ' Calculate premium and IPT for Applicant1
        CalculatePremiumAndIPT strCoverTypeValue, strApp1Gender, CStr(intApp1Age), _
                               strTerm, blnIsSmoker, blnIsInGoodHealth, _
                               vdblLifeCoverIPTRate, vdblAmount, dblApp1Premium, dblApp1IPT

        
        ' More than one applicant but not joint life relationship
        If intApp2Age > 0 Then
            
            ' Calculate premium and IPT for Applicant2 and add it to Applicant 1
            CalculatePremiumAndIPT strCoverTypeValue, strApp2Gender, CStr(intApp2Age), _
                                   strTerm, blnIsSmoker, blnIsInGoodHealth, _
                                   vdblLifeCoverIPTRate, vdblAmount, dblApp2Premium, dblApp2IPT
        
            dblPremium = dblApp1Premium + dblApp2Premium
            dblIPT = dblApp1IPT + dblApp2IPT
            
            Set xmlBenefitElem = vxmlJointElem.ownerDocument.createElement("LIFEBENEFIT")
            vxmlJointElem.appendChild xmlBenefitElem
            xmlBenefitElem.setAttribute "TYPE", vstrCoverType
            Set xmlElement = vxmlJointElem.ownerDocument.createElement("MONTHLYCOST")
            xmlElement.Text = Format(dblPremium, "0.00")
            xmlBenefitElem.appendChild xmlElement
            Set xmlElement = vxmlJointElem.ownerDocument.createElement("IPT")
            xmlElement.Text = Format(dblIPT, "0.00")
            xmlBenefitElem.appendChild xmlElement
        Else
            dblPremium = dblApp1Premium
            Set xmlBenefitElem = vxmlApplicant1Elem.ownerDocument.createElement("LIFEBENEFIT")
            vxmlApplicant1Elem.appendChild xmlBenefitElem
            xmlBenefitElem.setAttribute "TYPE", vstrCoverType
            
            Set xmlElement = vxmlApplicant1Elem.ownerDocument.createElement("MONTHLYCOST")
            xmlElement.Text = Format(dblPremium, "0.00")
            xmlBenefitElem.appendChild xmlElement
            
            Set xmlElement = vxmlApplicant1Elem.ownerDocument.createElement("IPT")
            xmlElement.Text = Format(dblApp1IPT, "0.00")
            xmlBenefitElem.appendChild xmlElement
        End If
    End If
        
    ' SR - create elements 'BENEFITTYPE', 'APPLICANTS and append
    Dim intBenefitType As Integer
'MS 19/07/00 Modified for performance enhancements
'    intBenefitType = GetComboValue(vstrCoverType, "LifeCoverType")
    intBenefitType = CInt(objComboDO.GetFirstComboValueId("LifeCoverType", vstrCoverType))
    
    Set xmlElement = vxmlJointElem.ownerDocument.createElement("BENEFITTYPE")
    xmlElement.Text = intBenefitType
    xmlBenefitElem.appendChild xmlElement
        
    Set xmlElement = vxmlJointElem.ownerDocument.createElement("APPLICANTS")
    xmlElement.Text = UCase(strRelationship)
    xmlBenefitElem.appendChild xmlElement
    
    
    CalculatePremiumsForCoverType = dblPremium

'MS 19/07/00 Modified for performance enhancements
'    Set xmlIn = Nothing
    Set xmlElement = Nothing
    Set xmlBenefitElem = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objErrAssist = Nothing
     
    Exit Function

CalculatePremiumsForCoverTypeVbErr:

'MS 19/07/00 Modified for performance enhancements
'    Set xmlIn = Nothing
    Set xmlElement = Nothing
    Set xmlBenefitElem = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
'    Set objErrAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function CalculatePremiums(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the life cover premiums for each applicant and joint premiums
'               if required
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'
'
'   Request Format:
'
'                   <LIFECOVER>
'                       <ENDOWMENTNEED>Value</ENDOWMENTNEED>
'                       <TERMASSURANCENEED>Value</TERMASSURANCENEED>
'                       <INTERESTONLYTERM>Value</INTERESTONLYTERM>
'                       <CAPITALANDINTERESTTERM>Value</CAPITALANDINTERESTTERM>
'                       <RELATIONSHIP>Value</RELATIONSHIP>
'                       <APPLICANT1SMOKERSTATUS>Value</APPLICANT1SMOKERSTATUS>
'                       <APPLICANT1GOODHEALTH>Value</APPLICANT1GOODHEALTH>
'                       <APPLICANT1GENDER>Value</APPLICANT1GENDER>
'                       <APPLICANT1AGE>Value</APPLICANT1AGE>
'                       <APPLICANT2SMOKERSTATUS>Value</APPLICANT2SMOKERSTATUS>
'                       <APPLICANT2GOODHEALTH>Value</APPLICANT2GOODHEALTH>
'                       <APPLICANT2GENDER>Value</APPLICANT2GENDER>
'                       <APPLICANT2AGE>Value</APPLICANT2AGE>
'                   </LIFECOVER>
'
'------------------------------------------------------------------------------------------
On Error GoTo CalculatePremiumsVbErr

    Const strFunctionName As String = "CalculatePremiums"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method
    
    Set xmlTempResponseNode = ILifeCoverBO_CalculatePremiums(xmlIn.documentElement)

    ' Check response
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    CalculatePremiums = xmlResponseElem.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalculatePremiumsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Exit Function

CalculatePremiumsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculatePremiums = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    ' go to clean-up section
    Resume CalculatePremiumsExit
End Function


