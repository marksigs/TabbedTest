Attribute VB_Name = "Transformations"
'Workfile:      Transformations.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'   Convert Optimus XML to Omiga XML and vice versa.
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     24/08/01    Expanded methods created by LD.
'RF     17/09/01    Handle TITLE correctly.
'AS     20/09/01    Added Exit Functions.
'RF     26/09/01    Handle COUNTY and COUNTRY as lookups.
'DS     23/03/02    Fixes to XML handling in addresses. SYS4306.
'DS     25/03/02    Added a function to convert Omiga to Optimus dates.
'SG     07/02/03    SYS6050 Check for blank title in OptimusNameToOmiga
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub OptimusNameToOmiga( _
    ByVal vnodeOmiga As IXMLDOMNode, _
    ByVal vnodeOptimusNameDetail As IXMLDOMNode, _
    ByVal vblnIncludeSalutation As Boolean, _
    ByVal vblnIncludeTitle As Boolean)
' header ----------------------------------------------------------------------------------
' description:
'   set the following name attributes on vnodeOmiga:
'   SURNAME
'   FIRSTFORENAME
'   SECONDFORENAME
'   OTHERFORENAMES
'   TITLE                       {is vblnIncludeTitle = true)
'   CORRESPONDENCESALUTATION    {if vblnIncludeSalutation = true)
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo OptimusNameToOmigaExit

    Const strFunctionName = "OptimusNameToOmiga"
    
    Dim strSurname As String
    Dim strFIRSTFORENAME As String
    Dim strOptimusTitleValue As String
    Dim strTitleAsOmigaComboId As String
    
    If vnodeOmiga Is Nothing Then
        errThrowError strFunctionName, oeMissingParameter, "Omiga node missing"
    End If
    
    If vnodeOptimusNameDetail Is Nothing Then
        errThrowError strFunctionName, oeMissingParameter, "Optimus NameDetail Node missing"
    End If
    
    strSurname = xmlGetNodeText(vnodeOptimusNameDetail, ".//SURNAME/@DATA")
    xmlSetAttributeValue vnodeOmiga, "SURNAME", strSurname

    strFIRSTFORENAME = xmlGetNodeText(vnodeOptimusNameDetail, ".//GIVENNAME1/@DATA")
    xmlSetAttributeValue vnodeOmiga, "FIRSTFORENAME", strFIRSTFORENAME

    xmlSetAttributeValue vnodeOmiga, "SECONDFORENAME", _
        xmlGetNodeText(vnodeOptimusNameDetail, ".//GIVENNAME2/@DATA")
    
    xmlSetAttributeValue vnodeOmiga, "OTHERFORENAMES", _
        xmlGetNodeText(vnodeOptimusNameDetail, ".//GIVENNAME3/@DATA")
    
    'RF 17/09/01: Handle TITLE correctly
    
    ' title and customer salutation
    If vblnIncludeTitle = True Or vblnIncludeSalutation = True Then
        strOptimusTitleValue = xmlGetNodeText(vnodeOptimusNameDetail, ".//TITLE/@DATA")
        strTitleAsOmigaComboId = lookupTitle(strOptimusTitleValue, cdOptimusToOmiga)
    
        If vblnIncludeTitle = True Then
            xmlSetAttributeValue vnodeOmiga, "TITLE", strTitleAsOmigaComboId
        End If
    
        If vblnIncludeSalutation = True Then
        
            'SG 07/02/03 SYS6050 Start
            If strTitleAsOmigaComboId = "" Then
                xmlSetAttributeValue vnodeOmiga, "CORRESPONDENCESALUTATION", _
                    strFIRSTFORENAME & " " & _
                    strSurname
            Else
                'Existing code
                xmlSetAttributeValue vnodeOmiga, "CORRESPONDENCESALUTATION", _
                    GetComboText("Title", strTitleAsOmigaComboId) & " " & _
                    strFIRSTFORENAME & " " & _
                    strSurname
            End If
            'SG 07/02/03 SYS6050 End
        End If
    End If
    
OptimusNameToOmigaExit:

    errCheckError strFunctionName
    
End Sub

Public Sub OptimusAddressToOmiga( _
    ByVal OmigaNode As IXMLDOMNode, _
    ByVal OptimusAddressNode As IXMLDOMNode, _
    blnIncludeDeliveryPointSuffix As Boolean, _
    blnIncludeMailSortCode As Boolean)
    
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo OptimusAddressToOmigaExit

    Const strFunctionName = "OptimusAddressToOmiga"
    
    If OmigaNode Is Nothing Then
        errThrowError strFunctionName, oeMissingParameter, "OmigaNode missing"
    End If
    
    If OptimusAddressNode Is Nothing Then
        errThrowError strFunctionName, oeMissingParameter, "OptimusAddressNode missing"
    End If
    
    Dim strTemp As String
    
    xmlSetAttributeValue OmigaNode, "BUILDINGORHOUSENAME", _
        xmlGetNodeText(OptimusAddressNode, "HOUSENAME/@DATA")
    xmlSetAttributeValue OmigaNode, "BUILDINGORHOUSENUMBER", _
        xmlGetNodeText(OptimusAddressNode, "HOUSENUMBER/@DATA")
    xmlSetAttributeValue OmigaNode, "FLATNUMBER", _
        xmlGetNodeText(OptimusAddressNode, "SUITENUMBER/@DATA")
    xmlSetAttributeValue OmigaNode, "STREET", _
        xmlGetNodeText(OptimusAddressNode, "STREETNAME/@DATA")
    xmlSetAttributeValue OmigaNode, "DISTRICT", _
        xmlGetNodeText(OptimusAddressNode, "DISTRICTNAME/@DATA")
    xmlSetAttributeValue OmigaNode, "TOWN", _
        xmlGetNodeText(OptimusAddressNode, "CITYNAME/@DATA")
        
    strTemp = xmlGetNodeText(OptimusAddressNode, "LOCATION/@DATA")
    xmlSetAttributeValue OmigaNode, "COUNTY", lookupCounty(strTemp, cdOptimusToOmiga)

    strTemp = xmlGetNodeText(OptimusAddressNode, "COUNTRYCODE/@DATA")
    xmlSetAttributeValue OmigaNode, "COUNTRY", lookupCountry(strTemp, cdOptimusToOmiga)
        
    xmlSetAttributeValue OmigaNode, "POSTCODE", _
        xmlGetNodeText(OptimusAddressNode, "POSTALCODE/@DATA")
    If blnIncludeDeliveryPointSuffix Then
        xmlSetAttributeValue OmigaNode, "DELIVERYPOINTSUFFIX", ""
    End If
    If blnIncludeMailSortCode Then
        xmlSetAttributeValue OmigaNode, "MAILSORTCODE", ""
    End If

OptimusAddressToOmigaExit:

    errCheckError strFunctionName

End Sub

Public Function OptimusDateToOmiga( _
    ByVal vstrOptimusDate As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert date from Optimus CCYYMMDD format to Omiga DD/MM/YY format
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo OptimusDateToOmigaErr

    Const strFunctionName = "OptimusDateToOmiga"
    
    Dim strConvertedValue As String
    
    If Len(vstrOptimusDate) = 0 Then
        strConvertedValue = ""
    Else
        If Len(vstrOptimusDate) <> Len("CCYYMMDD") Then
            errThrowError strFunctionName, oeInvalidParameter, _
                "Invalid date format: " & vstrOptimusDate
        Else
        
            strConvertedValue = _
                Mid(vstrOptimusDate, 7, 2) & "/" & _
                Mid(vstrOptimusDate, 5, 2) & "/" & _
                Left(vstrOptimusDate, 4)
        
        End If
    
    End If
    
    OptimusDateToOmiga = strConvertedValue
    
    Exit Function
        
OptimusDateToOmigaErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function OmigaDateToOptimus( _
    ByVal vstrOmigaDate As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert date from Omiga DD/MM/YY format to Optimus CCYYMMDD format
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo OmigaDateToOptimusErr

    Const strFunctionName = "OmigaDateToOptimus"
    
    Dim strConvertedValue As String
    
    If Len(vstrOmigaDate) = 0 Then
        strConvertedValue = ""
    Else
        If Len(vstrOmigaDate) <> Len("DD/MM/YYYY") Then
            errThrowError strFunctionName, oeInvalidParameter, _
                "Invalid date format: " & vstrOmigaDate
        Else
        
            strConvertedValue = _
                Mid(vstrOmigaDate, 7, 4) & _
                Mid(vstrOmigaDate, 4, 2) & _
                Left(vstrOmigaDate, 2)
        End If
    
    End If
    
    OmigaDateToOptimus = strConvertedValue
    
    Exit Function
        
OmigaDateToOptimusErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function OptimusBooleanToOmiga( _
    ByVal vstrOptimusBoolean As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert Boolean from Optimus true/false format to Omiga 0/1 format
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo OptimusBooleanToOmigaErr

    Const strFunctionName = "OptimusBooleanToOmiga"
    
    Dim strConvertedValue As String
    
    If Len(vstrOptimusBoolean) = 0 Then
        strConvertedValue = ""
    Else
        If UCase(vstrOptimusBoolean) = "TRUE" Then
            strConvertedValue = "1"
        Else
            If UCase(vstrOptimusBoolean) = "FALSE" Then
                strConvertedValue = "0"
            Else
                errThrowError strFunctionName, oeInvalidParameter, _
                    "Invalid Boolean format: " & vstrOptimusBoolean
            End If
        End If
    End If
    
    OptimusBooleanToOmiga = strConvertedValue
    
    Exit Function
        
OptimusBooleanToOmigaErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

