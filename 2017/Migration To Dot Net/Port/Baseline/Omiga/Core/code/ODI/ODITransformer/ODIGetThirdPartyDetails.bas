Attribute VB_Name = "ODIGetThirdPartyDetails"
'Workfile:      ODIGetThirdPartyDetails.bas
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SG/JR  18/11/2002  SYS5765/SYSMCAP1256 Created.
'------------------------------------------------------------------------------------------

Option Explicit

Public Sub GetThirdPartyDetails( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetThirdPartyDetailsExit
    
    Const strFunctionName As String = "GetThirdPartyDetails"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodePrimaryCustomer As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    Dim strOtherSystemNumber As String
               
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_findByPrimaryKey_ThirdPartyKey
    '------------------------------------------------------------------------------------------
    
    Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, "THIRDPARTY")
    strOtherSystemNumber = xmlGetMandatoryAttributeText(nodeOmigaCustomer, "OTHERSYSTEMNUMBER")
    
    Set nodeTemp = CreatePrimaryCustomerKey(strOtherSystemNumber)
    
    Set nodeConverterResponse = PlexusHomeImpl_findByPrimaryKey_ThirdPartyKey( _
        vobjODITransformerState, nodeTemp)
    
    CheckConverterResponse nodeConverterResponse, True
    
    Set nodeTemp = Nothing
    
    Set nodePrimaryCustomer = xmlGetMandatoryNode( _
        nodeConverterResponse, _
        "ARGUMENTS/OBJECT/PRIMARYCUSTOMERIMPL", _
        oerecordnotfound)
        
    ' ThirdParty response node
    Set nodeOmigaCustomer = vxmlResponseNode.ownerDocument.createElement("THIRDPARTY")
    
    xmlSetAttributeValue nodeOmigaCustomer, "OTHERSYSTEMNUMBER", _
        xmlGetNodeText(nodePrimaryCustomer, "PRIMARYKEY/PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA")

    xmlSetAttributeValue nodeOmigaCustomer, "SERVERUPDATEDATE", _
        xmlGetNodeText(nodePrimaryCustomer, "SERVERUPDATEDATE/@DATA")
    
    xmlSetAttributeValue nodeOmigaCustomer, "SERVERUPDATETIME", _
        xmlGetNodeText(nodePrimaryCustomer, "SERVERUPDATETIME/@DATA")

    vxmlResponseNode.appendChild nodeOmigaCustomer

GetThirdPartyDetailsExit:

    Set nodeOmigaCustomer = Nothing
    Set nodePrimaryCustomer = Nothing
    Set nodeConverterResponse = Nothing
    Set nodeTemp = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName
End Sub
