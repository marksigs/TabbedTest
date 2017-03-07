Attribute VB_Name = "ODISaveMortgageProductDetails"
'Workfile:      ODISaveMortgageProductDetails.bas
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:
'
'Dependencies:  List any other dependent components
'               e.g. BuildingsAndContentsSubQuoteTxBO, BuildingsAndContentsSubQuoteDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'STB    21/05/02    SYS4609 Implemented SaveMortgageProductDetails.
'------------------------------------------------------------------------------------------
Option Explicit


Public Sub SaveMortgageProductDetails( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Send a product update/creation through to Optimus.
' pass:
'   objODITransformerState
'       Stores ODI state information.
'   vxmlRequestNode
'       <REQUEST USERID="" UNITID="" MACHINEID="" CHANNELID="" MODE="CREATE|UPDATE" OPERATION="SaveMortgageProductDetails" ADMINSYSTEMSTATE="">
'           <MORTGAGEPRODUCT MORTGAGEPRODUCTCODE="" MORTGAGEPRODUCTNAME=""/>
'       </REQUEST>
'   vxmlResponseNode
'       XML RESPONSE node (standard + ADMINSYSTEMSTATE).
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim xmlElement As IXMLDOMNode
    Dim sMortgageProductCode As String
    Dim xmlMortgageProductImpl As IXMLDOMNode
    
    Const strFunctionName As String = "SaveMortgageProductDetails"

    On Error GoTo SaveMortgageProductDetailsExit
        
    'Get the Product Code.
    sMortgageProductCode = xmlGetNodeText(vxmlRequestNode, "//MORTGAGEPRODUCT/@MORTGAGEPRODUCTCODE")
        
    'TODO: There is no definition yet for MORTGAGEPRODUCTIMPL. This is all guesswork.
    
    'The structure to create/update is as follows: -
    '    <MORTGAGEPRODUCTIMPL>
    '       <PRIMARYKEY DATA=""/>
    '       <NAME DATA=""/>
    '    </MORTGAGEPRODUCTIMPL>

    'MORTGAGEPRODUCTIMPL.
    Set xmlMortgageProductImpl = vxmlResponseNode.ownerDocument.createElement("MORTGAGEPRODUCTIMPL")
    
    'PRIMARYKEY.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("PRIMARYKEY")
    xmlMortgageProductImpl.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(vxmlRequestNode, "//MORTGAGEPRODUCT/@MORTGAGEPRODUCTCODE")
    
    'NAME.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("NAME")
    xmlMortgageProductImpl.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(vxmlRequestNode, "//MORTGAGEPRODUCT/@MORTGAGEPRODUCTNAME")
    
    'If mode is create then create the product in Optimus.
    If xmlGetNodeText(vxmlRequestNode, "//REQUEST/@MODE") = "CREATE" Then
        'Supply the product key value.
        Set vxmlResponseNode = PlexusHomeImpl_create_MortgageProductKey(objODITransformerState, sMortgageProductCode)
        
        'Ensure no errors were returned.
        CheckConverterResponse vxmlResponseNode, True
    End If
    
    'Update the mortgage product with its name.
    Set vxmlResponseNode = MortgageProductImpl_update(objODITransformerState, xmlMortgageProductImpl)
    
    'Ensure no errors were returned.
    CheckConverterResponse vxmlResponseNode, True
    
    'Ensure we have a response.
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oerecordnotfound
    End If
    
SaveMortgageProductDetailsExit:
    Set xmlElement = Nothing
    Set xmlMortgageProductImpl = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub
