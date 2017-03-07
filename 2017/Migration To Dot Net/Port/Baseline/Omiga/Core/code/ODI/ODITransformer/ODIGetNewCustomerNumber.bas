Attribute VB_Name = "ODIGetNewCustomerNumber"
'Workfile:      ODIGetNewCustomerNumber.cls
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
'RF     27/08/01    Expanded class created by LD.
'DS     13/12/01    Fix to XML reponse processing.
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub GetNewCustomerNumber( _
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
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
    On Error GoTo GetNewCustomerNumberExit
    
    Const strFunctionName = "GetNewCustomerNumber"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodePrimaryCustomer As IXMLDOMNode
    Dim nodeNumberResponse As IXMLDOMNode
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_create_PrimaryCustomerKey
    '------------------------------------------------------------------------------------------
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_create_PrimaryCustomerKey(vobjODITransformerState)
        
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    Set nodePrimaryCustomer = nodeConverterResponse.selectSingleNode( _
        "//RESPONSE/ARGUMENTS/OBJECT/PRIMARYCUSTOMERIMPL")
        
    Set nodeNumberResponse = _
        vxmlResponseNode.ownerDocument.createElement("NUMBERRESPONSE")
    xmlSetAttributeValue nodeNumberResponse, "OTHERSYSTEMNUMBER", _
        xmlGetNodeText(nodePrimaryCustomer, _
            ".//PRIMARYKEY/PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA")
    vxmlResponseNode.appendChild nodeNumberResponse
                
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode
                          
GetNewCustomerNumberExit:

    Set nodePrimaryCustomer = Nothing
    Set nodeConverterResponse = Nothing
    Set nodeNumberResponse = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub











