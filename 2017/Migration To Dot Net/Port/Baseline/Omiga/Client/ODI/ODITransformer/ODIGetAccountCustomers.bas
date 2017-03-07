Attribute VB_Name = "ODIGetAccountCustomers"
'Workfile:      ODIGetAccountCustomers.cls
'Copyright:     Copyright © 2007 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    16/01/2007  EP2_855 Created
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub GetAccountCustomers( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
' exceptions:
'   oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountCustomersError
    
    Const strFunctionName As String = "GetAccountCustomers"
    
    Dim xmlAccount As IXMLDOMNode
    Dim xmlMortgageKey As IXMLDOMNode
    
    Dim strAccountNo As String
    
    Set xmlAccount = xmlGetMandatoryNode(vxmlRequestNode, "ACCOUNT")
    strAccountNo = xmlGetMandatoryAttributeText(xmlAccount, "ACCOUNTNUMBER")
    Set xmlMortgageKey = CreateMortgageKey(strAccountNo)
    
    AddCustomerDetails vobjODITransformerState, _
                       xmlMortgageKey, _
                       vxmlResponseNode, _
                       xmlGetNodeText(vxmlRequestNode, "//REQUEST/@CHANNELID")

GetAccountCustomersExit:

    Set xmlAccount = Nothing
    Set xmlMortgageKey = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    Exit Sub
    
GetAccountCustomersError:
    errCheckError strFunctionName

End Sub
        
Private Sub AddCustomerDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeMortgageKey As IXMLDOMNode, _
    ByVal vnodeTarget As IXMLDOMNode, _
    strChannelId As String)
' header ----------------------------------------------------------------------------------
' description:
'   Add to vnodeTarget the details for the customers attached to the mortgage with a
'   key of vnodeMortgageKey.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerDetailsExit

    Const strFunctionName As String = "AddCustomerDetails"

    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodelistOptimusCustomers As IXMLDOMNodeList
    Dim strSearchKey As String
    Dim nodePrimaryCustomerKey As IXMLDOMNode
    Dim nodeOmigaCustomerList As IXMLDOMNode
    Dim nodeOmigaCustomer, nodeOptimusCustomer As IXMLDOMNode
    Dim nodeRequest As IXMLDOMNode
    Dim nodeResponse As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim nodeTemp As IXMLDOMNode
    Dim strTemp As String
    Dim index As Integer
    
    
    '------------------------------------------------------------------------------------------
    'Call PlexusHomeImpl.searchByPattern, supplying a CustomerListPattern.
    '------------------------------------------------------------------------------------------

    strSearchKey = xmlGetNodeText(vnodeMortgageKey, "REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER/@DATA") _
                        & "." & xmlGetNodeText(vnodeMortgageKey, "CHARGETYPE/@DATA")

    Set xmlDoc = New FreeThreadedDOMDocument40
    Set nodeResponse = xmlDoc.createElement("RESPONSE")
    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_CustomerListPattern( _
            vobjODITransformerState, strSearchKey)
            
    Set nodelistOptimusCustomers = nodeConverterResponse.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT")
        
    If nodelistOptimusCustomers Is Nothing Then
        errThrowError strFunctionName, oerecordnotfound
    End If
    
    ' customer list node
    Set nodeOmigaCustomerList = _
        vnodeTarget.ownerDocument.createElement("CUSTOMERLIST")
    vnodeTarget.appendChild nodeOmigaCustomerList

    index = 1
    For Each nodeOptimusCustomer In nodelistOptimusCustomers
    
        Set nodePrimaryCustomerKey = xmlGetNode(nodeOptimusCustomer, ".//PRIMARYCUSTOMERKEY")
        Set nodeRequest = xmlDoc.createElement("REQUEST")
        Set nodeTemp = xmlDoc.createElement("CUSTOMER")
        xmlSetAttributeValue nodeTemp, "CUSTOMERNUMBER", _
            xmlGetMandatoryNodeText(nodePrimaryCustomerKey, ".//ENTITYNUMBER/@DATA")
        nodeRequest.appendChild nodeTemp
                        
        'DS Added the channelID to this request xml as a quick workaround for problem caused by SYS4623
                        
        xmlSetAttributeValue nodeRequest, "CHANNELID", strChannelId
                        
        GetCustomerDetails vobjODITransformerState, nodeRequest, nodeResponse
        
        ' add the customer details
    
        Set nodeOmigaCustomer = _
            xmlGetMandatoryNode(nodeResponse, "CUSTOMER")
        
        xmlSetAttributeValue nodeOmigaCustomer, "CUSTOMERORDER", index
        xmlSetAttributeValue nodeOmigaCustomer, "CUSTOMERROLETYPE", xmlGetNodeText(nodeOptimusCustomer, ".//EXPIRYSUFFIX/@DATA")
        
        nodeOmigaCustomerList.appendChild nodeOmigaCustomer
        
        Set nodeRequest = Nothing
        Set nodeTemp = Nothing
         index = index + 1
    
    Next nodeOptimusCustomer
    Set xmlDoc = Nothing
        
AddCustomerDetailsExit:

    Set nodeConverterResponse = Nothing
    Set nodelistOptimusCustomers = Nothing
    Set nodePrimaryCustomerKey = Nothing
    Set nodeOmigaCustomerList = Nothing
    Set nodeOmigaCustomer = Nothing
    Set nodeRequest = Nothing
    Set nodeResponse = Nothing
    Set xmlDoc = Nothing
    Set nodeTemp = Nothing

End Sub


