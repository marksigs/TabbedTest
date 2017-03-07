Attribute VB_Name = "ODIGetAccountBalanceAndStatus"
'Workfile:      ODIGetAccountBalanceAndStatus.cls
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
'RF     05/09/01    Created.
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub GetAccountBalanceAndStatus( _
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
On Error GoTo GetAccountBalanceAndStatusExit
    
    Const strFunctionName As String = "GetAccountBalanceAndStatus"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeConverterResponse As IXMLDOMNode
    Dim strCollateralNumber As String
    Dim nodeMortgageKey As IXMLDOMNode
    
    '------------------------------------------------------------------------------------------
    ' Call PlexusHome.findByPrimaryKey supplying a MortgageKey
    '------------------------------------------------------------------------------------------
    
    strCollateralNumber = xmlGetMandatoryNodeText(vxmlRequestNode, _
        ".//REQUEST/MORTGAGE/@COLLATERALNUMBER")
        
    Set nodeMortgageKey = CreateMortgageKey(strCollateralNumber)
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_findByPrimaryKey_MortgageKey( _
            vobjODITransformerState, nodeMortgageKey)
            
    ' don't check the response - sending it out whole
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
    
    vxmlResponseNode.appendChild xmlGetMandatoryNode( _
        nodeConverterResponse, _
        ".//RESPONSE/ARGUMENTS/OBJECT/MORTGAGEIMPL", _
        oerecordnotfound)
          
    ' exceptions
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode

GetAccountBalanceAndStatusExit:

    Set nodeConverterResponse = Nothing
    Set nodeMortgageKey = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub


