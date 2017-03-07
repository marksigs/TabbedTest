Attribute VB_Name = "ODIFindCustomer"
'Workfile:      ODIFindCustomer.bas
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
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub FindCustomer( _
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
    On Error GoTo FindCustomerExit
    
    Const strFunctionName = "FindCustomer"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim strCustomerNumber As String
    Dim xmlPrimaryCustomerNode As IXMLDOMNode
    Dim nodelistOptimusShortcuts As IXMLDOMNodeList
    Dim nodeOptShortcut As IXMLDOMNode
    Dim elemOmigaCustomerList As IXMLDOMElement
    Dim elemOmigaCustomer As IXMLDOMElement
    Dim elemOmigaAddress As IXMLDOMElement
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey
    '------------------------------------------------------------------------------------------
    
    Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    strCustomerNumber = xmlGetMandatoryAttributeText(nodeOmigaCustomer, "CUSTOMERNUMBER")
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey( _
            objODITransformerState, strCustomerNumber)
    
    '------------------------------------------------------------------------------------------
    ' parse the response
    '------------------------------------------------------------------------------------------
    
    Set elemOmigaCustomerList = _
        vxmlResponseNode.ownerDocument.createElement("CUSTOMERLIST")
    vxmlResponseNode.appendChild elemOmigaCustomerList
        
    Set nodelistOptimusShortcuts = nodeConverterResponse.selectNodes( _
        "PRIMARYCUSTOMERSHORTCUT")
        
    If nodelistOptimusShortcuts Is Nothing Then
        errThrowError strFunctionName, oerecordnotfound
    End If
            
    For Each nodeOptShortcut In nodelistOptimusShortcuts
    
        Set elemOmigaCustomer = _
            vxmlResponseNode.ownerDocument.createElement("CUSTOMER")
        elemOmigaCustomerList.appendChild elemOmigaCustomer
        
        elemOmigaCustomer.setAttribute "OTHERSYSTEMCUSTOMERNUMBER", _
            xmlGetNodeText(nodeOptShortcut, _
                ".//PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA")
        
        OptimusName2Omiga _
            elemOmigaCustomer, _
            nodeOptShortcut.selectSingleNode( _
                ".//PCSEARCHDESCRIPTION/CONTACTDETAIL/CONTACTDETAIL/NAMEDETAIL/NAMEDETAIL"), _
            blnIncludeSalutation:=True
        
        elemOmigaCustomer.setAttribute "DATEOFBIRTH", _
            xmlGetNodeText(nodeOptShortcut, _
                "//.PCSEARCHDESCRIPTION/BIRTHDATE")
                
        Set elemOmigaAddress = _
            vxmlResponseNode.ownerDocument.createElement("CURRENTADDRESS")
            
        elemOmigaCustomer.appendChild elemOmigaAddress
        
        OptimusAddress2Omiga _
            elemOmigaAddress, _
            nodeOptShortcut.selectSingleNode( _
                ".//PCSEARCHDESCRIPTION/CONTACTDETAIL/CONTACTDETAIL/ADDRESSDETAIL/ADDRESSDETAIL")
        
        ' fixme - are the following required?
        
        elemOmigaCustomer.setAttribute "ADDRESSTYPE", _
            "1"
            
        elemOmigaCustomer.setAttribute "CHANNELID", _
            "1"
        
        elemOmigaCustomer.setAttribute "OTHERSYSTEMTYPE", _
            "1"
        
        elemOmigaCustomer.setAttribute "MOTHERSMAIDENNAME", _
            xmlGetNodeText(nodeOptShortcut, _
                "//.PCSEARCHDESCRIPTION/MOTHERMAIDENNAME")
        
        elemOmigaCustomer.setAttribute "SOURCE", "Optimus"
        
    Next nodeOptShortcut
    
    Set xmlPrimaryCustomerNode = nodeConverterResponse.selectSingleNode("RESPONSE/ARGUMENTS/MSG_ARRAY/PRIMARYCUSTOMERIMPL")
    
    Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
            
FindCustomerExit:

    'Set objODIConverter1 = Nothing
    Set nodeOmigaCustomer = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub



