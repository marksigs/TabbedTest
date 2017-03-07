Attribute VB_Name = "ODISaveBaseRateSet"
'Workfile:      ODISaveBaseRateSet.bas
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
'STB    09-Jul-2002 SYS5101 - Create/update base rate functionality added.
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub SaveBaseRateSet(ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Maps the Omiga XML for the base rate set structure onto OSG equivalent.
'
' pass:
'   objODITransformerState
'   vxmlRequestNode
'   vxmlResponseNode
'------------------------------------------------------------------------------------------
    
    Dim xmlElement As IXMLDOMNode
    Dim xmlRateChange As IXMLDOMNode
    
    Const strFunctionName As String = "SaveBaseRateSet"
    
    On Error GoTo SaveBaseRateSetExit
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
  
    'The structure to map onto is a RateChangeCall: -
    '<RATECHANGECALL>
    '   <EFFECTIVEDATE DATA=""/>
    '   <EXPIRYDATE DATA=""/>
    '   <INTERESTRATE DATA=""/>
    '   <RATECODE DATA=""/>
    '</RATECHANGECALL>
    
    'RATECHANGECALL.
    Set xmlRateChange = vxmlResponseNode.ownerDocument.createElement("RATECHANGECALL")
    
    'EFFECTIVEDATE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("EFFECTIVEDATE")
    xmlRateChange.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(vxmlRequestNode, "BASERATESET/@BASERATESTARTDATE")
    
    'EXPIRYDATE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("EXPIRYDATE")
    xmlRateChange.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(vxmlRequestNode, "BASERATESET/@EXPIRYDATE")
    
    'INTERESTRATE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("INTERESTRATE")
    xmlRateChange.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(vxmlRequestNode, "BASERATESET/@BASEINTERESTRATE")
    
    'RATECODE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("RATECODE")
    xmlRateChange.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(vxmlRequestNode, "BASERATESET/@BASERATESET")
    
    'Create the new base rate.
    Set vxmlResponseNode = PlexusHomeImpl_ProcessRateChange(objODITransformerState, xmlRateChange)
    
    'Ensure no errors were returned.
    CheckConverterResponse vxmlResponseNode, True
    
SaveBaseRateSetExit:
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub


