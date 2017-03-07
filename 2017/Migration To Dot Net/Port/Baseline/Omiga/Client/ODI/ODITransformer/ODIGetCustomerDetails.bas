Attribute VB_Name = "ODIGetCustomerDetails"
'Workfile:      ODIGetCustomerDetails.bas
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
'RF     24/08/01    Expanded method created by LD.
'AS     20/09/01    Added OPTIMUS_DBX compile time switch.
'AS     19/10/01    Removed OPTIMUS_DBX compile time switch.
'DS     20/03/02    Have added some code to facilitate FindAccountDetails setting the Lender code
'                   This is a bit hacky and has been done due to time constraints.
'DS     23/03/02    Fixes to make it work. SYS4306.
'DS     05/04/02    Removed the lendercode hack due to spec change.
'DS     17/04/02    Send out a TELEPHONENUMBERLIST even if it is empty.
'DS     16/05/02    Correct mappings for phone numbers.
'JLD    22/05/02    SYS4623 set the channelId on the created customer.
'STB    17/06/02    SYS4837 Preferred communication method updated to match mappings.
'DS     27/06/02    SYS4974 Set maritalStatus
'STB    11/07/02    SYS4610 Prefix the area code with a zero. Optimus strips it.
'PSC    21/03/07    EP2_1717 Add leading zero to mobile area code
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub GetCustomerDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description: If the three optional parameters are specified, strLenderCode is set as specced in
'               FindAccountDetails spec.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    On Error GoTo GetCustomerDetailsExit
    
    Const strFunctionName As String = "GetCustomerDetails"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim strCustomerNumber As String
    Dim nodePrimaryCustomer As IXMLDOMNode
    Dim strTemp As String
    Dim strDOB As String
    Dim nodeTemp As IXMLDOMNode
               
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey
    '------------------------------------------------------------------------------------------
    
    Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    strCustomerNumber = xmlGetMandatoryAttributeText(nodeOmigaCustomer, "CUSTOMERNUMBER")
    
    Set nodeTemp = CreatePrimaryCustomerKey(strCustomerNumber)
    
    Set nodeConverterResponse = PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey( _
        vobjODITransformerState, nodeTemp)
    
    CheckConverterResponse nodeConverterResponse, True
    
    Set nodeTemp = Nothing
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    Set nodePrimaryCustomer = xmlGetMandatoryNode( _
        nodeConverterResponse, _
        "ARGUMENTS/OBJECT/PRIMARYCUSTOMERIMPL", _
        oerecordnotfound)
        
    ' customer response node
    Set nodeOmigaCustomer = vxmlResponseNode.ownerDocument.createElement("CUSTOMER")
    
    xmlSetAttributeValue nodeOmigaCustomer, "CUSTOMERNUMBER", _
        xmlGetNodeText(nodePrimaryCustomer, "PRIMARYKEY/PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA")
    
    OptimusNameToOmiga _
        nodeOmigaCustomer, _
        nodePrimaryCustomer.selectSingleNode( _
            "CONTACTDETAIL/CONTACTDETAIL/NAMEDETAIL/NAMEDETAIL"), _
        vblnIncludeSalutation:=False, _
        vblnIncludeTitle:=True
    
    strDOB = xmlGetNodeText(nodePrimaryCustomer, "BIRTHDATE/@DATA")
    
    'convert data format
    strDOB = OptimusDateToOmiga(strDOB)
    xmlSetAttributeValue nodeOmigaCustomer, "DATEOFBIRTH", strDOB
    
    ' derive age from date of birth
    If Len(strDOB) > 0 Then
        strTemp = GetAge(strDOB)
    Else
        strTemp = ""
    End If
    xmlSetAttributeValue nodeOmigaCustomer, "AGE", strTemp
        
    strTemp = xmlGetNodeText(nodePrimaryCustomer, "GENDER/@DATA")
    xmlSetAttributeValue nodeOmigaCustomer, "GENDER", _
        lookupGender(strTemp, cdOptimusToOmiga)
    
    xmlSetAttributeValue nodeOmigaCustomer, "MOTHERSMAIDENNAME", _
        xmlGetNodeText(nodePrimaryCustomer, "SECURITYAREA1/@DATA")
    xmlSetAttributeValue nodeOmigaCustomer, "CONTACTEMAILADDRESS", _
        xmlGetNodeText(nodePrimaryCustomer, "EMAILADDRESS/@DATA")
    
    'SYS4974
    xmlSetAttributeValue nodeOmigaCustomer, "MARITALSTATUS", xmlGetNodeText(nodePrimaryCustomer, "MARITALSTATUS/@DATA")
    
    xmlSetAttributeValue nodeOmigaCustomer, "NATIONALINSURANCENUMBER", _
        xmlGetNodeText(nodePrimaryCustomer, "SOCIALINSURANCENUMBER/@DATA")
        
    strTemp = xmlGetNodeText(nodePrimaryCustomer, "EMPLOYEE/@DATA")
    strTemp = OptimusBooleanToOmiga(strTemp)
    xmlSetAttributeValue nodeOmigaCustomer, "MEMBEROFSTAFF", strTemp
    
    'JLD SYS4623 Use the channelId of the user from the REQUEST tag
    xmlSetAttributeValue nodeOmigaCustomer, "CHANNELID", xmlGetMandatoryAttributeText(vxmlRequestNode, "CHANNELID")
    xmlSetAttributeValue nodeOmigaCustomer, "OTHERSYSTEMTYPE", ""
    vxmlResponseNode.appendChild nodeOmigaCustomer
    
    ' current address response node
    Set nodeTemp = vxmlResponseNode.ownerDocument.createElement("CURRENTADDRESS")
    
    OptimusAddressToOmiga nodeTemp, nodePrimaryCustomer.selectSingleNode( _
        "CONTACTDETAIL/CONTACTDETAIL/ADDRESSDETAIL/ADDRESSDETAIL"), True, True
    nodeOmigaCustomer.appendChild nodeTemp
      
    '------------------------------------------------------------------------------------------
    ' handle telephone numbers
    '------------------------------------------------------------------------------------------
    
    Dim nodeTelephoneNumberList As IXMLDOMNode
    Dim cmPreferredMethod As CONTACTMETHOD
    Dim strPreferredMethod As String
    Dim strPreferredTime As String
        
    
    Set nodeTelephoneNumberList = _
        vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERLIST")
    
    strPreferredMethod = xmlGetNodeText(nodePrimaryCustomer, ".//PREFERREDCOMMUNICATIONMETHOD/@DATA")
    cmPreferredMethod = lookupPreferredContactMethod(strPreferredMethod)
    
    'SYS4837 - Preferred communication method updated to match mappings.
    If cmPreferredMethod = cmEmail Then
        xmlSetAttributeValue nodeOmigaCustomer, "EMAILPREFERRED", "1"
    Else
        xmlSetAttributeValue nodeOmigaCustomer, "EMAILPREFERRED", "0"
    End If
    
    strPreferredTime = xmlGetNodeText( _
        nodePrimaryCustomer, "PREFERREDCOMMUNICATIONTIME/@DATA")
    
    ' home phone
    
    strTemp = xmlGetNodeText(nodePrimaryCustomer, "HOMEPHONENUMBER/@DATA")
    
    If Len(strTemp) > 0 Then
        Set nodeTemp = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
        xmlSetAttributeValue nodeTemp, "USAGE", GetFirstComboValueId("TelephoneUsage", "H")
        xmlSetAttributeValue nodeTemp, "COUNTRYCODE", _
            xmlGetNodeText(nodePrimaryCustomer, "HOMEPHONECOUNTRYCODE/@DATA")
        
        'SYS4610 - Prefix the area code with a zero. Optimus strips it.
        xmlSetAttributeValue nodeTemp, "AREACODE", "0" & xmlGetNodeText(nodePrimaryCustomer, "HOMEPHONEAREACODE/@DATA")
            
        xmlSetAttributeValue nodeTemp, "TELEPHONENUMBER", strTemp
        xmlSetAttributeValue nodeTemp, "EXTENSIONNUMBER", _
            xmlGetNodeText(nodePrimaryCustomer, "HOMEPHONEEXTNUMBER/@DATA")
        
        If cmPreferredMethod = cmHome Then
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", strPreferredTime
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "1"
        Else
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", ""
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "0"
        End If
        
        nodeTelephoneNumberList.appendChild nodeTemp
    End If
    
    ' work
    
    strTemp = xmlGetNodeText(nodePrimaryCustomer, "BUSINESSPHONENUMBER/@DATA")
    
    If Len(strTemp) > 0 Then
        Set nodeTemp = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
        xmlSetAttributeValue nodeTemp, "USAGE", GetFirstComboValueId("TelephoneUsage", "W")
        xmlSetAttributeValue nodeTemp, "COUNTRYCODE", _
            xmlGetNodeText(nodePrimaryCustomer, "BUSINESSCOUNTRYCODE/@DATA")
        
        'SYS4610 - Prefix the area code with a zero. Optimus strips it.
        xmlSetAttributeValue nodeTemp, "AREACODE", "0" & xmlGetNodeText(nodePrimaryCustomer, "BUSINESSAREACODE/@DATA")
        
        xmlSetAttributeValue nodeTemp, "TELEPHONENUMBER", strTemp
        xmlSetAttributeValue nodeTemp, "EXTENSIONNUMBER", _
            xmlGetNodeText(nodePrimaryCustomer, "BUSINESSEXTNUMBER/@DATA")
        
        If cmPreferredMethod = cmWork Then
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", strPreferredTime
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "1"
        Else
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", ""
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "0"
        End If
        
        nodeTelephoneNumberList.appendChild nodeTemp
    End If
    
    ' fax
    
    strTemp = xmlGetNodeText(nodePrimaryCustomer, "FAXNUMBER/@DATA")
    
    If Len(strTemp) > 0 Then
        Set nodeTemp = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
        xmlSetAttributeValue nodeTemp, "USAGE", GetFirstComboValueId("TelephoneUsage", "F")
        xmlSetAttributeValue nodeTemp, "COUNTRYCODE", _
            xmlGetNodeText(nodePrimaryCustomer, "FAXCOUNTRYCODE/@DATA")
        
        'SYS4610 - Prefix the area code with a zero. Optimus strips it.
        xmlSetAttributeValue nodeTemp, "AREACODE", "0" & xmlGetNodeText(nodePrimaryCustomer, "FAXAREACODE/@DATA")
            
        xmlSetAttributeValue nodeTemp, "TELEPHONENUMBER", strTemp
        xmlSetAttributeValue nodeTemp, "EXTENSIONNUMBER", ""
    
        If cmPreferredMethod = cmFax Then
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", strPreferredTime
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "1"
        Else
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", ""
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "0"
        End If
    
        nodeTelephoneNumberList.appendChild nodeTemp
    End If
    
    ' mobile
    
    strTemp = xmlGetNodeText(nodePrimaryCustomer, "CELLPHONENUMBER/@DATA")
    
    If Len(strTemp) > 0 Then
        Set nodeTemp = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
        xmlSetAttributeValue nodeTemp, "USAGE", GetFirstComboValueId("TelephoneUsage", "M")
        xmlSetAttributeValue nodeTemp, "COUNTRYCODE", _
            xmlGetNodeText(nodePrimaryCustomer, "CELLCOUNTRYCODE/@DATA")
        ' PSC 21/03/2007 EP2_1717
        xmlSetAttributeValue nodeTemp, "AREACODE", _
            "0" & xmlGetNodeText(nodePrimaryCustomer, "CELLAREACODE/@DATA")
        xmlSetAttributeValue nodeTemp, "TELEPHONENUMBER", strTemp
        xmlSetAttributeValue nodeTemp, "EXTENSIONNUMBER", ""
        
        If cmPreferredMethod = cmMobile Then
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", strPreferredTime
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "1"
        Else
            xmlSetAttributeValue nodeTemp, "CONTACTTIME", ""
            xmlSetAttributeValue nodeTemp, "PREFERREDMETHODOFCONTACT", "0"
        End If
        
        nodeTelephoneNumberList.appendChild nodeTemp
    End If
          
    nodeOmigaCustomer.appendChild nodeTelephoneNumberList
    
    ' exceptions
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode
              
GetCustomerDetailsExit:

    Set nodeOmigaCustomer = Nothing
    Set nodeConverterResponse = Nothing
    Set nodeTelephoneNumberList = Nothing
    Set nodePrimaryCustomer = Nothing
    Set nodeTemp = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub


Private Function GetAge(ByVal vstrDateOfBirth As String) As Integer
' header ----------------------------------------------------------------------------------
' description:
'   Calculate Age from date of birth.
'   Code was copied from omCust.CustomerDO and then modified to suit current template style.
' pass:
'   strDateOfBirth
'       date of birth in format dd/mm/yyyy
' return:
'   integer representation of age as calculated from date of birth
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetAgeVbErr

    Const strFunctionName = "GetAge"
    
    If (Len(vstrDateOfBirth) < Len("dd/mm/yyyy")) Or _
       (Not IsDate(vstrDateOfBirth)) Then
        errThrowError strFunctionName, oeInvalidDateTimeFormat, _
            " Expected Date Of Birth in format DD/MM/YYYY, but received: " & vstrDateOfBirth
    End If
    
    Dim dtmDoB As Date
    Dim lngDoB As Long
    Dim lngNow As Long
    
    dtmDoB = vstrDateOfBirth
    
    lngDoB = (Year(dtmDoB) * 10000) + (Month(dtmDoB) * 100) + Day(dtmDoB)
    lngNow = (Year(Now) * 10000) + (Month(Now) * 100) + Day(Now)
    
    GetAge = Fix((lngNow - lngDoB) / 10000)
    
    Exit Function
    
GetAgeVbErr:
    
    ' re-raise error
    Err.Raise Err.Number, Err.Source, Err.Description

End Function




