Attribute VB_Name = "Lookups"
'Workfile:      Lookups.bas
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
'RF     24/08/01    Created.
'DS     21/03/02    Some changes to the lookup for contact method. These are temporary and need to
'                   be checked with the spec. SYS4306.
'STB    17/06/02    SYS4837 Preferred communication method updated to match mappings.
'PSC    19/01/2007  EP2_928 Amend lookupInterestRateType to return hard coded values
'------------------------------------------------------------------------------------------
Option Explicit

Enum CONVERSIONDIRECTION
    cdOmigaToOptimus
    cdOptimusToOmiga
End Enum

Enum CONTACTMETHOD
    cmHome = 1
    cmWork = 2
    cmMobile = 3
    cmFax = 4
    cmEmail = 5
    cmUndefined = 99
End Enum

Public Function lookupTitle( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for ContactDetail.NameDetail.title
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupTitle = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupTitle"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupTitle = strConvertedValue
End Function

Public Function lookupIndemnityCompanyName( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for MortgageKey.Charge.MortgageInsuranceType
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupIndemnityCompanyName = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupIndemnityCompanyName"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupIndemnityCompanyName = strConvertedValue
End Function

Public Function lookupGender( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for PrimaryCustomer.gender
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupGender = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupGender"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupGender = strConvertedValue
    
End Function


Public Function lookupPreferredContactMethod( _
    ByVal vstrOptimusValue As String) _
    As CONTACTMETHOD
' header ----------------------------------------------------------------------------------
' description:
'   lookup for PrimaryCustomer.preferredCommunicationMethod:
'   return contact method corresponding to vstrOptimusValue
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo lookupPreferredContactMethodVbErr
    
    Const strFunctionName = "lookupPreferredContactMethod"
        
    Dim cmPreferred As CONTACTMETHOD

    Select Case vstrOptimusValue
        Case "1"
            cmPreferred = cmHome
        Case "2"
            cmPreferred = cmWork
        Case "3"
            cmPreferred = cmMobile
        Case "4"
            cmPreferred = cmFax
        Case "5"
            cmPreferred = cmEmail
        Case Else
            cmPreferred = cmUndefined
    End Select
    
    If vstrOptimusValue = "" Then
        cmPreferred = "0"
    Else
        cmPreferred = vstrOptimusValue
    End If
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupPreferredContactMethod = cmPreferred
    Exit Function
    
    If Len(vstrOptimusValue) > 0 Then
        ' fixme - need real values
        If vstrOptimusValue = "1" Then
            cmPreferred = cmHome
        Else
            'etc
        End If
    End If
    
    lookupPreferredContactMethod = cmPreferred
    
    Exit Function
    
lookupPreferredContactMethodVbErr:
    
    ' re-raise error
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function lookupCustomerRoleType( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert Optimus CISearchDescription.expirySuffixDescription to or from
'   Omiga CUSTOMERROLETYPE.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "lookupCustomerRoleType"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
        
            If vstrOriginalValue = "41" Then
                strConvertedValue = GetFirstComboValueId("CustomerRoleType", "G")
            Else
                If vstrOriginalValue = "40" Then
                    strConvertedValue = GetFirstComboValueId("CustomerRoleType", "A")
                Else
                    ' fixme - logic for "Other" not yet defined
                    errThrowError strFunctionName, oeNotImplemented
                End If
            End If
        End If
    End If
    
    lookupCustomerRoleType = strConvertedValue
    
End Function

Public Function lookupPropertyDescription( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupPropertyDescription = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupPropertyDescription"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupPropertyDescription = strConvertedValue
    
End Function

Public Function lookupRedemptionStatus( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert PIComponent.status to or from an Omiga "RedemptionStatus" combo id.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupRedemptionStatus = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupRedemptionStatus"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupRedemptionStatus = strConvertedValue
    
End Function

Public Function lookupRepaymentType( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert PIComponent.paymentType to or from an Omiga combo id.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupRepaymentType = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupRepaymentType"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupRepaymentType = strConvertedValue
    
End Function

Public Function lookupPaymentMethod( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert PaymentSourceKey.paymentMethod to or from an Omiga combo id.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupPaymentMethod = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupPaymentMethod"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupPaymentMethod = strConvertedValue
    
End Function

Public Function lookupMortgageProductCode( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert PIComponent.mortgageProductCode to a combo id
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "lookupMortgageProductCode"
    Dim objMort As MortgageProductBO
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    Dim strConvertedValue As String
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseDoc As New FreeThreadedDOMDocument40
    Dim xmlRe As IXMLDOMNode
    Dim vstrXMLRequest As String
    Dim vstrXmlResponse As String
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    
    On Error GoTo errorHandler
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            Set objMort = New MortgageProductBO
            
            Set xmlElem = xmlRequestDoc.createElement("REQUEST")
            Set xmlTempNode = xmlRequestDoc.appendChild(xmlElem)
            
            Set xmlElem = xmlRequestDoc.createElement("MORTGAGEPRODUCTLANGUAGE")
            Call xmlTempNode.appendChild(xmlElem)
            vstrXmlResponse = objMort.GetMortgageProductLanguage(xmlRequestDoc.xml)
            If errCheckXMLResponse(vstrXmlResponse) = 0 Then
                Call xmlResponseDoc.loadXML(vstrXmlResponse)
                Set xmlResponseNode = xmlResponseDoc.selectSingleNode(".//MORTGAGEPRODUCTLANGUAGE[MORTGAGEPRODUCTCODE=" & vstrOriginalValue & "]/PRODUCTNAME")
                If Not (xmlResponseNode Is Nothing) Then
                    strConvertedValue = xmlResponseNode.Text
                Else
                    strConvertedValue = ""
                End If
            Else
            '
            End If
        End If
    End If
    
    lookupMortgageProductCode = strConvertedValue
    
    Exit Function
    
errorHandler:
    
    errThrowError strFunctionName, Err.Number

End Function

Public Function lookupHomeInsuranceType( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for PeripheralSecurity.entitySuffix
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupHomeInsuranceType = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupHomeInsuranceType"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupHomeInsuranceType = strConvertedValue
    
End Function

Public Function lookupPurposeOfLoan( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupPurposeOfLoan = vstrOriginalValue
    Exit Function
    
    Const strFunctionName As String = "lookupPurposeOfLoan"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupPurposeOfLoan = strConvertedValue
    
End Function

Public Function lookupCounty( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for AddressDetail.location
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupCounty = vstrOriginalValue
    Exit Function
        
    Const strFunctionName As String = "lookupCounty"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupCounty = strConvertedValue
    
End Function

Public Function lookupCountry( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for AddressDetail.countryCode
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupCountry = vstrOriginalValue
    Exit Function
        
    Const strFunctionName As String = "lookupCountry"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            ' fixme
            strConvertedValue = "1"
        End If
    End If
    
    lookupCountry = strConvertedValue
    
End Function

Public Function lookupPropertyTenure( _
    ByVal vstrOriginalValue As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for RealEstateImpl.freeholdTitle
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    
    ' AS 30/10/01 Conversion now done by ODIConverter.
    lookupPropertyTenure = vstrOriginalValue
    Exit Function
            
    Const strFunctionName As String = "lookupPropertyTenure"
    
    Dim strConvertedValue As String
    
    If Len(vstrOriginalValue) = 0 Then
        strConvertedValue = ""
    Else
        If vConversionDirection = cdOmigaToOptimus Then
            errThrowError strFunctionName, oeNotImplemented
        Else
            If OptimusBooleanToOmiga(vstrOriginalValue) = "1" Then
                strConvertedValue = GetFirstComboValueId("PropertyTenure", "F")
            Else
                strConvertedValue = ""
            End If
        End If
    End If
    
    lookupPropertyTenure = strConvertedValue
    
End Function

'PSC 19/01/2007 EP2_928 - Start
Public Function lookupInterestRateType( _
    ByVal vstrRateChangeFrequency As String, _
    ByVal vstrMinimumRate As String, _
    ByVal vstrMaximumRate As String, _
    ByVal vConversionDirection As CONVERSIONDIRECTION) As String
' header ----------------------------------------------------------------------------------
' description:
'   lookup for PIComponent.maximumRate and PIComponent.rateChangeFrequency
'   The logic is:
'   If maximumRate Is Not "" Then
'       INTERESTRATETYPE = Omiga combo id for group [to be defined] where validation value = "C" (Capped/Floored).
'   Else
'       If rateChangeFrequency = "00" Then
'           INTERESTRATETYPE = Omiga combo id for group [to be defined] where validation value = "F" (Fixed)
'       Else
'           INTERESTRATETYPE = Omiga combo id for group [to be defined] where validation value = [to be defined] (Variable)
'       End If
'   End If
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "lookupInterestRateType"
    
    Dim strConvertedValue As String
    
    If vConversionDirection = cdOmigaToOptimus Then
        errThrowError strFunctionName, oeNotImplemented
    Else
        If vstrRateChangeFrequency = "00" Then
            strConvertedValue = "Fixed"
        ElseIf Len(vstrMinimumRate) > 0 Or Len(vstrMaximumRate) > 0 Then
            strConvertedValue = "Capped/Collared"
        Else
            strConvertedValue = "Variable"
        End If
    End If
    
    lookupInterestRateType = strConvertedValue
    
End Function
'PSC 19/01/2007 EP2_928 - End

