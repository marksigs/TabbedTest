Attribute VB_Name = "ODIFindCustomerList"
'Workfile:      ODIFindCustomerList.bas
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
'RF     23/08/01    Expanded shell method created by LD.
'AS     02/10/01    Added OPTIMUS_DBX compile time switch.
'RF     09/10/01    DATEOFBIRTH and POSTCODE are optional in request.
'AS     19/10/01    Removed OPTIMUS_DBX compile time switch.
'AS     22/10/01    Fixed XPATH.
'DS     21/02/02    Fix the XML output so it matches the DTD
'DS     23/03/02    Loads of fixes to make it work. SYS4306.
'DS     25/03/02    Fixes to date formats.
'DS     30/04/02    Use FreeThreadedDOMDocument.
'STB    30/04/02    SYS4517 Return the source system for customer searches.
'DS     17/05/02    SYS4624 Fix MothersMaidenName.
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub FindCustomerList( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve CUSTOMER entities based on search criteria.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CUSTOMER node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
    On Error GoTo FindCustomerListExit
    
    Const strFunctionName = "FindCustomerList"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim sAdminSystemName As String
    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodelistOptimusShortcuts As IXMLDOMNodeList
    Dim nodeOptShortcut As IXMLDOMNode
    Dim elemOmigaCustomerList As IXMLDOMElement
    Dim elemOmigaCustomer As IXMLDOMElement
    Dim elemOmigaAddress As IXMLDOMElement
    Dim nodeCustomerPattern As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument
    Dim nodeTemp As IXMLDOMNode
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_searchByPattern_PrimaryCustomerPattern
    '------------------------------------------------------------------------------------------
    
    Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, ".//CUSTOMER")
    Set nodeCustomerPattern = xmlDoc.createElement("PRIMARYCUSTOMERPATTERN")
    
    Set nodeTemp = xmlDoc.createElement("FORENAME")
    xmlSetAttributeValue nodeTemp, "DATA", _
        xmlGetMandatoryAttributeText(nodeOmigaCustomer, "FIRSTFORENAME")
    nodeCustomerPattern.appendChild nodeTemp
    
    Set nodeTemp = xmlDoc.createElement("BIRTHDATE")
    ' RF 09/10/01 DATEOFBIRTH is optional
    Dim strTmp
    strTmp = xmlGetAttributeText(nodeOmigaCustomer, "DATEOFBIRTH")
    xmlSetAttributeValue nodeTemp, "DATA", OmigaDateToOptimus(strTmp)
    nodeCustomerPattern.appendChild nodeTemp
    
    Set nodeTemp = xmlDoc.createElement("SURNAME")
    xmlSetAttributeValue nodeTemp, "DATA", _
        xmlGetMandatoryAttributeText(nodeOmigaCustomer, "SURNAME")
    nodeCustomerPattern.appendChild nodeTemp
    
    Set nodeTemp = xmlDoc.createElement("POSTALCODE")
    ' RF 09/10/01 POSTCODE is optional
    xmlSetAttributeValue nodeTemp, "DATA", _
        xmlGetAttributeText(nodeOmigaCustomer, "POSTCODE")
    nodeCustomerPattern.appendChild nodeTemp

    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_PrimaryCustomerPattern( _
            objODITransformerState, nodeCustomerPattern)
    
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    Set elemOmigaCustomerList = _
        vxmlResponseNode.ownerDocument.createElement("CUSTOMERLIST")
    vxmlResponseNode.appendChild elemOmigaCustomerList
        
    ' AS 22/10/01 Changed from "PRIMARYCUSTOMERSHORTCUT".
    Set nodelistOptimusShortcuts = nodeConverterResponse.selectNodes( _
        ".//OBJECTSHORTCUT")
        
    If nodelistOptimusShortcuts Is Nothing Then
        errThrowError strFunctionName, oerecordnotfound
    End If
            
    'SYS4517 - Obtain the admin system's name to be set against each customer.
    sAdminSystemName = GetGlobalParamString("AdminSystemName")
            
    For Each nodeOptShortcut In nodelistOptimusShortcuts
    
        Set elemOmigaCustomer = _
            vxmlResponseNode.ownerDocument.createElement("CUSTOMER")
        elemOmigaCustomerList.appendChild elemOmigaCustomer
        
        elemOmigaCustomer.setAttribute "OTHERSYSTEMCUSTOMERNUMBER", _
            xmlGetNodeText(nodeOptShortcut, _
                ".//PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA")
        
        OptimusNameToOmiga _
            elemOmigaCustomer, _
            nodeOptShortcut.selectSingleNode( _
                ".//PCSEARCHDESCRIPTION/CONTACTDETAIL/CONTACTDETAIL/NAMEDETAIL/NAMEDETAIL"), _
            vblnIncludeSalutation:=True, _
            vblnIncludeTitle:=False
        
        ' AS 22/10/01 Added "/@DATA"
        elemOmigaCustomer.setAttribute "DATEOFBIRTH", OptimusDateToOmiga(xmlGetNodeText(nodeOptShortcut, ".//PCSEARCHDESCRIPTION/BIRTHDATE/@DATA"))
                            
        'elemOmigaCustomer.appendChild elemOmigaAddress
        
        OptimusAddressToOmiga _
            elemOmigaCustomer, _
            nodeOptShortcut.selectSingleNode( _
                ".//PCSEARCHDESCRIPTION/CONTACTDETAIL/CONTACTDETAIL/ADDRESSDETAIL/ADDRESSDETAIL"), _
            False, False
        
        elemOmigaCustomer.setAttribute "ADDRESSTYPE", "1"
            
        elemOmigaCustomer.setAttribute "CHANNELID", ""
        
        elemOmigaCustomer.setAttribute "OTHERSYSTEMTYPE", ""
        
        elemOmigaCustomer.setAttribute "MOTHERSMAIDENNAME", _
            xmlGetNodeText(nodeOptShortcut, _
                ".//PCSEARCHDESCRIPTION/MOTHERMAIDENNAME/@DATA")
        
        'SYS4517 - Return the source system for this user.
        elemOmigaCustomer.setAttribute "SOURCE", sAdminSystemName
        
    Next nodeOptShortcut

    'Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
            
FindCustomerListExit:

    'Set objODIConverter = Nothing
    Set nodeOmigaCustomer = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub



