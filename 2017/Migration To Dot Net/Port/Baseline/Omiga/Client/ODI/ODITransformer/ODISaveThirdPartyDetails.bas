Attribute VB_Name = "ODISaveThirdPartyDetails"
'Workfile:      ODISaveThirdPartyDetails.cls
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
'RF     04/09/01    Expanded class created by LD.
'DM     02/04/02    SYS4350 Implement SaveThirdpartyDetails
'DS     24/04/02    SYS4350 Finish off SaveThirdPartyDetails
'STB    10/05/02    SYS4558 Correct the XML in SaveThirdPartyDetails.
'SG/JR  18/11/02    SYS5765/SYSMCP1256 Amend SaveThirdPartyDetails to update ServerUpdateDate/Time from data passed in.
'SG     03/01/03    SYS5579 Amended SaveThirdPartyDetails, pass ContactTitle to optimus
'DRC -  06/05/03    MSMS102 - change to map CompanyName to OPTIMUS Client surname for Legal Reps
'DRC    23/05/03    MSMS136 - can't have blanks in Searchkey
'------------------------------------------------------------------------------------------
Option Explicit

'DM     02/04/02    SYS4350 Implement SaveThirdpartyDetails
Public Sub SaveThirdPartyDetails( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified CUSTOMER entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
'The XML that we are trying to generate to send to Optimus
'<PRIMARYCUSTOMERIMPL>
'   <PRIMARYKEY>
'       <PRIMARYCUSTOMERKEY>
'           <ENTITYNUMBER DATA=""/>
'       </PRIMARYCUSTOMERKEY>
'   </PRIMARYKEY>
'   <CONTACT>
'       <CONTACT>
'           <ADDRESS>
'               <ADDRESS>
'                   <LINE1 DATA=""/>
'                     ...
'                   <LINE6 DATA=""/>
'               </ADDRESS>
'           </ADDRESS>
'           <NAME>
'               <NAME>
'                   <LINE1 DATA=""/>
'                   <LINE2 DATA=""/>
'               </NAME>
'           </NAME>
'       </CONTACT>
'   </CONTACT>
'   <SERVERUPDATEDATE DATA=""/>
'   <SERVERUPDATETIME DATA=""/>
'   <BRANCHCIFSUFFIX DATA=""/>
'   <BUSINESSAREACODE DATA=""/>
'   <BUSINESSCOUNTRYCODE DATA=""/>
'   <BUSINESSEXTNUMBER DATA=""/>
'   <CONTACTDETAIL>
'       <CONTACTDETAIL>
'           <ADDRESSDETAIL>
'               <ADDRESSDETAIL>
'                   <CITYNAME DATA=""/>
'                   <COUNTRYCODE DATA=""/>
'                   <DISTRICTNAME DATA=""/>
'                   <FORMATADDRESSCHANGEIND DATA=""/>
'                   <HOUSENAME DATA=""/>
'                   <HOUSENUMBER DATA=""/>
'                   <LOCATION DATA =""/>
'                   <POSTALCODE DATA=""/>
'                   <STREETNAME DATA=""/>
'                   <SUITENUMBER DATA=""/>
'               </ADDRESSDETAIL>
'           </ADDRESSDETAIL>
'           <NAMEDETAIL>
'               <NAMEDETAIL>
'                   <FORMATNAMECHANGEIND DATA=""/>
'                   <GIVENNAME1 DATA="" />
'                   <GIVENNAME2 DATA="" />
'                   <GIVENNAME3 DATA="" />
'                   <SURNAME DATA="" />
'                   <TITLE DATA="" />
'               </NAMEDETAIL>
'           </NAMEDETAIL>
'       </CONTACTDETAIL>
'   </CONTACTDETAIL>
'   <EMAILADDRESS DATA=""/>
'   <ENTITYTYPE DATA=""/>
'   <LANGUAGEPREFERENCE DATA=""/>
'   <OCCUPATIONTYPE DATA=""/>
'   <SEARCHKEY DATA=""/>
'   <SPOKENLANGUAGEPREFERENCE DATA=""/>
'</PRIMARYCUSTOMERIMPL>
'-------------------------------------------------------------------------------------------

On Error GoTo SaveThirdPartyDetailsExit
    
    Const strFunctionName As String = "SaveThirdPartyDetails"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim nodeThirdParty As IXMLDOMNode
    Dim nodeIntermediaryIndividual As IXMLDOMNode
    Dim nodeIntermediaryOrganisation As IXMLDOMNode
    Dim nodeAddress As IXMLDOMNode
    Dim nodeContactDetails As IXMLDOMNode
    
    ' Nodes to create the xml we are sending to Optimus
    Dim nodePrimaryCustomerImpl As IXMLDOMNode
    ' These two nodes are used to iterate throught the xml assigning values
    ' as we go along. They will refer to many different nodes throughtout
    ' the life of this method. As a convention I have always used nodeTemp2
    ' as the terminal node i.e. then I assign the attribute value to.
    Dim nodeTemp1 As IXMLDOMNode
    Dim nodeTemp2 As IXMLDOMNode
    
    Dim strForeName As String
    Dim strSurname As String
    Dim strCompanyName As String
    
    'SG 03/01/03 SYS5579
    Dim strContactTitle As String
        
    ' We will need these nodes so lets get them upfront.
    Set nodeThirdParty = xmlGetNode(vxmlRequestNode, "/REQUEST/THIRDPARTY")
    Set nodeIntermediaryIndividual = xmlGetNode(nodeThirdParty, "INTERMEDIARYINDIVIDUAL")
    Set nodeAddress = xmlGetNode(nodeThirdParty, "ADDRESS")
    Set nodeContactDetails = xmlGetNode(nodeThirdParty, "CONTACTDETAILS")
    Set nodeIntermediaryOrganisation = xmlGetNode(nodeThirdParty, "INTERMEDIARYORGANISATION")
        
    'Get the Broker's (individual intermediary) name if relevant.
    If Not nodeIntermediaryIndividual Is Nothing Then
        strForeName = xmlGetAttributeText(nodeIntermediaryIndividual, "FORENAME")
        strSurname = xmlGetAttributeText(nodeIntermediaryIndividual, "SURNAME")
    
    'Otherwise just get the contact details names.
    ElseIf Not nodeContactDetails Is Nothing Then
        strForeName = xmlGetAttributeText(nodeContactDetails, "FORENAME")
        strSurname = xmlGetAttributeText(nodeContactDetails, "SURNAME")
    End If
    strCompanyName = xmlGetAttributeText(nodeThirdParty, "COMPANYNAME")
        
    'SG 03/01/03 SYS5579
    strContactTitle = xmlGetAttributeText(nodeContactDetails, "TITLE")
        
    ' Create the XML structure we are going to send to Optimus
    Set nodePrimaryCustomerImpl = vxmlRequestNode.ownerDocument.createElement("PRIMARYCUSTOMERIMPL")
    
    'BEGIN Set up the PRIMARYCUSTOMERKEY node
    Dim nodePrimaryKey As IXMLDOMNode
    
    'PRIMARYKEY
    Set nodePrimaryKey = vxmlRequestNode.ownerDocument.createElement("PRIMARYKEY")
    nodePrimaryCustomerImpl.appendChild nodePrimaryKey

    'PRIMARYCUSTOMERKEY
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("PRIMARYCUSTOMERKEY")
    nodePrimaryKey.appendChild nodeTemp1
    
    'ENTITYNUMBER
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("ENTITYNUMBER")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeThirdParty, "OTHERSYSTEMNUMBER")
    
    Set nodeTemp1 = Nothing
    Set nodeTemp2 = Nothing
    'END Set up the PRIMARYCUSTOMERKEY node
    
    'BEGIN Set up the CONTACT node
    Dim nodeName As IXMLDOMNode
    Dim nodeContact As IXMLDOMNode
    Dim nodeContact2 As IXMLDOMNode
    
    'CONTACT - parent.
    Set nodeContact = vxmlRequestNode.ownerDocument.createElement("CONTACT")
    nodePrimaryCustomerImpl.appendChild nodeContact
    
    'CONTACT
    Set nodeContact2 = vxmlRequestNode.ownerDocument.createElement("CONTACT")
    nodeContact.appendChild nodeContact2
        
    'ADDRESS.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("ADDRESS")
    nodeContact2.appendChild nodeTemp1
    
    'ADDRESS.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("ADDRESS")
    nodeTemp1.appendChild nodeTemp2
    
    'LINE1.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("LINE1")
    nodeTemp2.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", ""
    
    'LINE2.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("LINE2")
    nodeTemp2.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", ""
    
    'LINE3.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("LINE3")
    nodeTemp2.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", ""
    
    'LINE4.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("LINE4")
    nodeTemp2.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", ""
    
    'LINE5.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("LINE5")
    nodeTemp2.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", ""
    
    'LINE6.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("LINE6")
    nodeTemp2.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", ""
            
    'NAME
    Set nodeName = vxmlRequestNode.ownerDocument.createElement("NAME")
    nodeContact2.appendChild nodeName
    
    'NAME
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("NAME")
    nodeName.appendChild nodeTemp1
    
    'LINE 1
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("LINE1")
    nodeTemp1.appendChild nodeTemp2
    
    'SYS4558 - Both Intermediary organisation name AND individual names go into LINE1.
    If nodeIntermediaryOrganisation Is Nothing Then
        'SR 06/05/2003 - MSMS89
        If xmlGetAttributeText(nodeThirdParty, "THIRDPARTYTYPE") = "61" Then
           'DRC - MSMS102 - change to CompanyName
            xmlSetAttributeValue nodeTemp2, "DATA", strCompanyName
        Else 'SR 06/05/2003 - MSMS89 - End
            xmlSetAttributeValue nodeTemp2, "DATA", strForeName & " " & strSurname
        End If 'SR 06/05/2003 - MSMS89
    Else
        xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeIntermediaryOrganisation, "NAME")
    End If
    
    'LINE2
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("LINE2")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", strForeName & " " & strSurname
    
    Set nodeTemp1 = Nothing
    Set nodeTemp2 = Nothing
    'END Set up the CONTACT node
                    
    'SERVERUPDATEDATE
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("SERVERUPDATEDATE")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    
    'SG 18/11/02 SYS5765
    'xmlSetAttributeValue nodeTemp2, "DATA", ""
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeThirdParty, "SERVERUPDATEDATE")
    
    Set nodeTemp2 = Nothing
        
    'SERVERUPDATETIME
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("SERVERUPDATETIME")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    
    'SG 18/11/02 SYS5765
    'xmlSetAttributeValue nodeTemp2, "DATA", ""
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeThirdParty, "SERVERUPDATETIME")
    
    Set nodeTemp2 = Nothing
            
    'BRANCHCIFSUFFIX.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("BRANCHCIFSUFFIX")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", "0"
    Set nodeTemp2 = Nothing
    
    'TEMP: Val() used here to remove leading zero. Probably not required...
    'BUSINESSAREACODE
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("BUSINESSAREACODE")
    nodePrimaryCustomerImpl.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", Val(xmlGetAttributeText(nodeContactDetails, "AREACODE"))
    
    'BUSINESSCOUNTRYCODE
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("BUSINESSCOUNTRYCODE")
    nodePrimaryCustomerImpl.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", xmlGetAttributeText(nodeContactDetails, "COUNTRYCODE")

    'BUSINESSEXTNUMBER
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("BUSINESSEXTNUMBER")
    nodePrimaryCustomerImpl.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", xmlGetAttributeText(nodeContactDetails, "TELEPHONEEXTENSIONNUMBER")

    'BUSINESSPHONENUMBER
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("BUSINESSPHONENUMBER")
    nodePrimaryCustomerImpl.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", xmlGetAttributeText(nodeContactDetails, "TELEPHONENUMBER")
    
    'BEGIN Set up CONTACTDETAILS node
    Dim nodeAddressDetail As IXMLDOMNode
    Dim nodeContactDetail As IXMLDOMNode
    Dim nodeContactDetail2 As IXMLDOMNode
    
    'CONTACTDETAIL - parent.
    Set nodeContactDetail = vxmlRequestNode.ownerDocument.createElement("CONTACTDETAIL")
    nodePrimaryCustomerImpl.appendChild nodeContactDetail

    'CONTACTDETAIL
    Set nodeContactDetail2 = vxmlRequestNode.ownerDocument.createElement("CONTACTDETAIL")
    nodeContactDetail.appendChild nodeContactDetail2
    
    'ADDRESSDETAIL - parent.
    Set nodeAddressDetail = vxmlRequestNode.ownerDocument.createElement("ADDRESSDETAIL")
    nodeContactDetail2.appendChild nodeAddressDetail
        
    'ADDRESSDETAIL
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("ADDRESSDETAIL")
    nodeAddressDetail.appendChild nodeTemp1
        
    'CITYNAME
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("CITYNAME")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "TOWN")
        
    'COUNTRYCODE
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("COUNTRYCODE")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "COUNTRY")
    
    'DISTRICTNAME
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("DISTRICTNAME")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "DISTRICT")
    
    'FORMATTEDADDRESSCHANGEIND = false.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("FORMATTEDADDRESSCHANGEIND")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", "false"
    
    'HOUSENAME
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("HOUSENAME")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "BUILDINGORHOUSENAME")
    
    'HOUSENUMBER
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("HOUSENUMBER")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "BUILDINGORHOUSENUMBER")
    
    'LOCATION
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("LOCATION")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "COUNTY")
    
    'POSTALCODE
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("POSTALCODE")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "POSTCODE")
    
    'STREETNAME
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("STREETNAME")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "STREET")
    
    'SUITENUMBER
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("SUITENUMBER")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeAddress, "FLATNUMBER")
    
    'NAMEDETAIL.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("NAMEDETAIL")
    nodeContactDetail2.appendChild nodeTemp2
        
    'NAMEDETAIL.
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("NAMEDETAIL")
    nodeTemp2.appendChild nodeTemp1
        
    'FORMATTEDNAMECHANGEIND.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("FORMATTEDNAMECHANGEIND")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", "false"
        
    'GIVENNAME1.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("GIVENNAME1")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", strForeName
        
    'GIVENNAME2.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("GIVENNAME2")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", ""
        
    'GIVENNAME3.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("GIVENNAME3")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", ""
        
    'SURNAME.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("SURNAME")
    nodeTemp1.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", strSurname
        
    'TITLE.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("TITLE")
    nodeTemp1.appendChild nodeTemp2
    'SG 03/01/03 SYS5579
    'xmlSetAttributeValue nodeTemp2, "DATA", ""
    xmlSetAttributeValue nodeTemp2, "DATA", strContactTitle


    Set nodeTemp1 = Nothing
    Set nodeTemp2 = Nothing
    
    'EMAILADDRESS
    Set nodeTemp1 = vxmlRequestNode.ownerDocument.createElement("EMAILADDRESS")
    nodePrimaryCustomerImpl.appendChild nodeTemp1
    xmlSetAttributeValue nodeTemp1, "DATA", xmlGetAttributeText(nodeContactDetails, "EMAILADDRESS")
    
    'ENTITYTYPE.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("ENTITYTYPE")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", xmlGetAttributeText(nodeThirdParty, "THIRDPARTYTYPE")
    Set nodeTemp2 = Nothing
            
    'LANGUAGEPREFERENCE = E.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("LANGUAGEPREFERENCE")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", "E"
    Set nodeTemp2 = Nothing
        
    'OCCUPATIONTYPE = 01.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("OCCUPATIONTYPE")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", "01"
    Set nodeTemp2 = Nothing
            
    'SEARCHKEY
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("SEARCHKEY")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    
    If nodeIntermediaryOrganisation Is Nothing Then
      'DRC - MSMS102 - change to CompanyName for Legal Reps
      If xmlGetAttributeText(nodeThirdParty, "THIRDPARTYTYPE") = "61" Then
          'DRC MSMS136 - can't have blanks in Searchkey
            xmlSetAttributeValue nodeTemp2, "DATA", UCase(Replace(strCompanyName, " ", ""))
        Else
            xmlSetAttributeValue nodeTemp2, "DATA", UCase(Left$(Replace((strSurname & Left$(strForeName, 1)), " ", ""), 10))
      End If       'End DRC - MSMS102
    Else
        xmlSetAttributeValue nodeTemp2, "DATA", UCase(Left$(Replace(xmlGetAttributeText(nodeIntermediaryOrganisation, "NAME"), " ", ""), 10))
    End If
    
    Set nodeTemp2 = Nothing
    
    'SPOKENLANGUAGEPREFERENCE = E.
    Set nodeTemp2 = vxmlRequestNode.ownerDocument.createElement("SPOKENLANGUAGEPREFERENCE")
    nodePrimaryCustomerImpl.appendChild nodeTemp2
    xmlSetAttributeValue nodeTemp2, "DATA", "E"
    Set nodeTemp2 = Nothing
    
    'Send the request to the converter.
    Set vxmlResponseNode = PrimaryCustomerImpl_update(objODITransformerState, nodePrimaryCustomerImpl)
    
    'Ensure no errors were returned.
    CheckConverterResponse vxmlResponseNode, True
    
    '*** Need to check what we need in here.
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oerecordnotfound
    End If
    
SaveThirdPartyDetailsExit:

    Set nodePrimaryCustomerImpl = Nothing
    Set nodeThirdParty = Nothing
    Set nodeIntermediaryIndividual = Nothing
    Set nodeAddress = Nothing
    Set nodeContactDetails = Nothing
    Set nodeIntermediaryOrganisation = Nothing
    Set nodeTemp1 = Nothing
    Set nodeTemp2 = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub

