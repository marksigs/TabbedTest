Attribute VB_Name = "ODIFindBusinessForCustomer"
'Workfile:      ODIFindBusinessForCustomer.bas
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
'RF     28/08/01    Expanded code created by LD.
'DS     23/03/02    Loads of fixes to make it work. SYS 4306.
'DRC    20/06/03    MSMS171 Temp Fix - plexus doesn't return anything if the balance is zero
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub FindBusinessForCustomer( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieves a list of all active Mortgage accounts for a specified customer.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    On Error GoTo FindBusinessForCustomerExit
    
    Const strFunctionName As String = "FindBusinessForCustomer"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim nodeOmigaCustomer As IXMLDOMNode
    Dim nodeConverterResponse As IXMLDOMNode
    Dim strCustomerNumber As String
    Dim nodelistOptimusShortcuts, nodelistOptimusShortcuts2 As IXMLDOMNodeList
    Dim nodeOptShortcut, nodeOptShortcut2 As IXMLDOMNode
    Dim elemOmigaAccountList As IXMLDOMElement
    Dim elemOmigaAccount As IXMLDOMElement
    Dim strTemp As String
    Dim strAccountNumber As String
    Dim strOutstandingBalance As String
    
    '------------------------------------------------------------------------------------------
    ' call PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern
    '------------------------------------------------------------------------------------------
    
    Set nodeOmigaCustomer = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    strCustomerNumber = xmlGetMandatoryAttributeText(nodeOmigaCustomer, "CUSTOMERNUMBER")
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern( _
            vobjODITransformerState, strCustomerNumber)
    
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' handle the response
    '------------------------------------------------------------------------------------------
        
    Set elemOmigaAccountList = _
        vxmlResponseNode.ownerDocument.createElement("MORTGAGEACCOUNTLIST")
                
    Set nodelistOptimusShortcuts = nodeConverterResponse.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT") '/ARGUMENTS/OBJECT/OBJECTSHORTCUT")
        
    If nodelistOptimusShortcuts Is Nothing Then
        errThrowError strFunctionName, oerecordnotfound
    End If
            
    For Each nodeOptShortcut In nodelistOptimusShortcuts
    
        Set elemOmigaAccount = _
            vxmlResponseNode.ownerDocument.createElement("MORTGAGEACCOUNT")
        elemOmigaAccountList.appendChild elemOmigaAccount
        
        Dim temp, temp2 As String
     
        temp = xmlGetMandatoryNodeText(nodeOptShortcut, "TARGETKEY/MORTGAGEKEY/REALESTATEKEY/REALESTATEKEY/COLLATERALNUMBER/@DATA")
        temp2 = xmlGetMandatoryNodeText(nodeOptShortcut, "TARGETKEY/MORTGAGEKEY/CHARGETYPE/@DATA")
     
        strAccountNumber = temp & "." & temp2
        
        elemOmigaAccount.setAttribute "ACCOUNTNUMBER", strAccountNumber
                
        elemOmigaAccount.setAttribute "BUSINESSTYPEINDICATOR", "M"
        elemOmigaAccount.setAttribute "BUSINESSTYPE", "Mortgage Account"
        
        elemOmigaAccount.setAttribute "STATUS", xmlGetMandatoryNodeText(nodeOptShortcut, "PAYLOAD/CISEARCHDESCRIPTION/STATUS/@DATA")
'MSMS171 Temp Fix  - DRC 20/06/03 - plexus doesn't return anything if the balance is zero
        ' elemOmigaAccount.setAttribute "OUTSTANDINGBALANCE", xmlGetMandatoryNodeText(nodeOptShortcut, "PAYLOAD/CISEARCHDESCRIPTION/CURRENTOUTSTANDINGBALANCE/@DATA")
        strOutstandingBalance = xmlGetNodeText(nodeOptShortcut, "PAYLOAD/CISEARCHDESCRIPTION/CURRENTOUTSTANDINGBALANCE/@DATA")
        If Len(strOutstandingBalance) = 0 Then
          strOutstandingBalance = "0.0"
        End If
        elemOmigaAccount.setAttribute "OUTSTANDINGBALANCE", strOutstandingBalance
'MSMS171 Temp Fix  - End
        Dim nodeConverterResponse2 As IXMLDOMNode
    
        Set nodeConverterResponse2 = PlexusHomeImpl_searchByPattern_CustomerListPattern(vobjODITransformerState, strAccountNumber)
            
        CheckConverterResponse nodeConverterResponse2, True
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = nodeConverterResponse2.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT")
        ' just want to know the number of arguments returned
        If (nodeList Is Nothing) Then
        Else
            elemOmigaAccount.setAttribute "NUMBEROFAPPLICANTS", nodeList.length
        End If
                
        strTemp = xmlGetMandatoryNodeText(nodeOptShortcut, "PAYLOAD/CISEARCHDESCRIPTION/FIRSTADVANCEDATE/@DATA")
        strTemp = OptimusDateToOmiga(strTemp)
        elemOmigaAccount.setAttribute "DATECREATED", strTemp
        
        
        Set nodelistOptimusShortcuts2 = nodeConverterResponse2.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT")
        If nodelistOptimusShortcuts2 Is Nothing Then
            errThrowError strFunctionName, oerecordnotfound
        End If
        
        For Each nodeOptShortcut2 In nodelistOptimusShortcuts2
            If xmlGetMandatoryNodeText(nodeOptShortcut2, "TARGETKEY/PRIMARYCUSTOMERKEY/ENTITYNUMBER/@DATA") = strCustomerNumber Then
            strTemp = xmlGetMandatoryNodeText(nodeOptShortcut2, "PAYLOAD/CLSEARCHDESCRIPTION/EXPIRYSUFFIX/@DATA")
            elemOmigaAccount.setAttribute "CUSTOMERROLETYPE", strTemp
    
            End If
        Next nodeOptShortcut2
    
    Next nodeOptShortcut
    
    vxmlResponseNode.appendChild elemOmigaAccountList
    
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode

FindBusinessForCustomerExit:

    Set nodeOmigaCustomer = Nothing
    Set nodelistOptimusShortcuts = Nothing
    Set elemOmigaAccountList = Nothing
    Set elemOmigaAccount = Nothing
    Set nodeConverterResponse = Nothing
    Set nodeOptShortcut = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub

Private Function GetNumberOfApplicants( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vstrAccountNumber As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Gets a count of all the applicants for a specified account number.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfApplicantsExit

    Const cstrFunctionName As String = "GetNumberOfApplicants"
    
    Dim nodeConverterResponse As IXMLDOMNode
    
    Set nodeConverterResponse = _
        PlexusHomeImpl_searchByPattern_CustomerListPattern( _
            vobjODITransformerState, vstrAccountNumber)
            
    CheckConverterResponse nodeConverterResponse, True
    
    ' just want to know the number of arguments returned
    GetNumberOfApplicants = nodeConverterResponse.selectNodes("ARGUMENTS/OBJECT/OBJECTSHORTCUT").length
        
GetNumberOfApplicantsExit:

    Set nodeConverterResponse = Nothing
    
End Function

