Attribute VB_Name = "ODIGetNewAccountNumber"
'Workfile:      ODIGetNewAccountNumber.bas
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
'RF     27/08/01    Expanded method created by LD.
'DS     10/12/2001  Fixed the response processsing to look for correct node
'DRC    30/07/2002  Added an extra arguement for an optional charge type
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub GetNewAccountNumber( _
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
'------------------------------------------------------------------------------------------
    On Error GoTo GetNewAccountNumberExit
    
    Const strFunctionName As String = "GetNewAccountNumber"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim nodeMortgageImpl As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodeNumberResponse As IXMLDOMNode
    'AQR SYS4945 - need to look for charge type type as an extra optional attrib.
    'and send it as an arguement to Plexus
    Dim strChargeType As String
    strChargeType = xmlGetAttributeText(vxmlRequestNode, "CHARGETYPE")
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_create_PrimaryMortgageKey
    '------------------------------------------------------------------------------------------
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_create_PrimaryMortgageKey(vobjODITransformerState, , strChargeType)
   'AQR SYS4945 - end
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    Set nodeMortgageImpl = nodeConverterResponse.selectSingleNode( _
        "//RESPONSE/ARGUMENTS/OBJECT/MORTGAGEIMPL/PRIMARYKEY/MORTGAGEKEY")
        
    Set nodeNumberResponse = _
        vxmlResponseNode.ownerDocument.createElement("NUMBERRESPONSE")
    xmlSetAttributeValue nodeNumberResponse, "OTHERSYSTEMNUMBER", _
        xmlGetNodeText(nodeMortgageImpl, ".//REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER/@DATA") & _
        "." & _
        xmlGetNodeText(nodeMortgageImpl, ".//CHARGETYPE/@DATA")
    
    vxmlResponseNode.appendChild nodeNumberResponse
        
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode
        
GetNewAccountNumberExit:

    Set nodeMortgageImpl = Nothing
    Set nodeConverterResponse = Nothing
    Set nodeNumberResponse = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub









