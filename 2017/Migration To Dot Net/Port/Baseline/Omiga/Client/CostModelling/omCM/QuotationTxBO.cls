VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "QuotationTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuotationTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Quotation Business Object which requires transactions
'
'Dependencies:  QuotationDO
'
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    09/11/99    Created
'JLD    28/04/00    typo in SaveOneOffCostDetails.
'                   Always check the response from Save methods called because if the method failed
'                   then the context will have been aborted. If we then set context to complete we get a lovelly automation error
'MS     09/08/00    SYS0834 - Save correct PPSUBQUOTENUMBER
'MS     14/06/00    If we haven't got a PPSUBQUOTENUMBER we need to create one
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     16/09/2002  BMIDS00432  Removed LifeSubQuote functionality from IQuotationTxBO_SaveOneOffCostDetails
'------------------------------------------------------------------------------------------
'Mars Specific History:
'
'Prog   Date        AQR         Description
'GHun   10/07/2006  MARS1893    Removed LifeSubQuote functionality from CreateMortgageAndLifeSubQuotes
'------------------------------------------------------------------------------------------

Option Explicit

Implements IQuotationTxBO
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrROOT_NODE_NAME = "QUOTATION"

Private Enum SUBQUOTEINSTANCE
    sqiFirst
    sqiNew
End Enum

Private Function CreateBCSubQuote(ByVal vxmlQuotationElement As IXMLDOMElement, _
                                  ByVal sqiInstance As SUBQUOTEINSTANCE) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the first sub quote or new sub quote for Buildings and Contents
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBCSubQuoteVbErr

    Const strFunctionName As String = "CreateBCSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlActionNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim objIQuotationDO As IQuotationDO
    Dim objIBCSubQuoteBO As IBuildingsAndContentsSubQuoteBO
    Dim xmlTempResponsElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    ' Create a document with the request node to pass to other BOs
    Set xmlRequestNode = objXmlAssist.GetRequestNode(vxmlQuotationElement.ownerDocument)
    xmlRequestDoc.appendChild xmlRequestNode
    
    Set xmlActionNode = xmlRequestDoc.createElement("CREATE")
    xmlRequestNode.appendChild xmlActionNode
        
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
        Set objIBCSubQuoteBO = New BuildingsAndContentsSubQuoteBO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
        Set objIBCSubQuoteBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteBO")
    End If
    
   ' Copy the current instance of data to put into the request
    Set xmlTableNode = vxmlQuotationElement.cloneNode(True)
    xmlActionNode.appendChild xmlTableNode
                         
    objXmlAssist.ChangeNodeName xmlTableNode, "QUOTATION", "BCSUBQUOTEDETAILS"
                                     
    ' Create the Buildings and contents sub quote
    If sqiInstance = sqiFirst Then
        Set xmlTempResponsElem = objIBCSubQuoteBO.CreateFirstSubQuote(xmlRequestNode)
    Else
        Set xmlTempResponsElem = objIBCSubQuoteBO.CreateNewSubQuote(xmlRequestNode)
    End If
        
    objErrAssist.CheckResponse xmlTempResponsElem, xmlResponseElem, True
    
    Dim strBCSubQuoteNo As String
    
    Set xmlElement = xmlTempResponsElem.getElementsByTagName("BCSUBQUOTENUMBER").Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlResponseElem.appendChild xmlElement.cloneNode(True)
        strBCSubQuoteNo = xmlElement.Text
    End If

    If Len(strBCSubQuoteNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "BCSUBQUOTENUMBER not generated"
    End If
                          
    ' If we are doing a new sub quote then a BCSubQuoteNumber tag will be present
    ' so update it with the number from the newly created sub quote to update the
    ' quotation. Otherwise if creating the first sub quote one will not be there
    ' so we need to create one
    Dim strTagName As String
    
    strTagName = "BCSUBQUOTENUMBER"
    
    Set xmlElement = xmlRequestDoc.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = xmlRequestDoc.createElement(strTagName)
        xmlTableNode.appendChild xmlElement
    End If
    
    xmlElement.Text = strBCSubQuoteNo
    
    objXmlAssist.ChangeNodeName xmlTableNode, "BCSUBQUOTEDETAILS", "QUOTATION"

    objIQuotationDO.Update xmlTableNode
    
    Set CreateBCSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBCSubQuoteExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlActionNode = Nothing
    Set xmlTableNode = Nothing
    Set objIQuotationDO = Nothing
    Set objIBCSubQuoteBO = Nothing
    Set xmlTempResponsElem = Nothing
    
    Exit Function

CreateBCSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set CreateBCSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBCSubQuoteExit

End Function

Private Function CreatePPSubQuote(ByVal vxmlQuotationElement As IXMLDOMElement, _
                                  ByVal sqiInstance As SUBQUOTEINSTANCE) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the first sub quote or new sub quote
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreatePPSubQuoteVbErr

    Const strFunctionName As String = "CreatePPSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlActionNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim objIQuotationDO As IQuotationDO
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    
    Dim strTagName As String
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strPPSubQuoteNo As String
    Dim strResponse As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Create a document with the request node to pass to other BOs
    Set xmlRequestNode = objXmlAssist.GetRequestNode(vxmlQuotationElement.ownerDocument)
    xmlRequestDoc.appendChild xmlRequestNode
    
    Set xmlActionNode = xmlRequestDoc.createElement("CREATE")
    xmlRequestNode.appendChild xmlActionNode
        
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
        Set objPPSubQuoteBO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteBO")
    End If
    
   ' Copy the current instance of data to put into the request
    Set xmlTableNode = vxmlQuotationElement.cloneNode(True)
    xmlActionNode.appendChild xmlTableNode
                 
    objXmlAssist.ChangeNodeName xmlTableNode, "QUOTATION", "PAYMENTPROTECTIONSUBQUOTE"
        
    If sqiInstance = sqiFirst Then
        strResponse = objPPSubQuoteBO.CreateFirstSubQuote(xmlRequestNode.xml)
    Else
        strResponse = objPPSubQuoteBO.CreateNewSubQuote(xmlRequestNode.xml)
    End If
        
    objErrAssist.CheckXMLResponse strResponse, True
    
    Set xmlResponseDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
    ' APS 24/11/99 - Return the response back to the QuotationBO
    Set xmlElement = xmlResponseDoc.getElementsByTagName("PPSUBQUOTENUMBER").Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlResponseElem.appendChild xmlElement.cloneNode(True)
        strPPSubQuoteNo = xmlElement.Text
    End If
                      
        
    If Len(strPPSubQuoteNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "PPSUBQUOTENUMBER not generated"
    End If
 
    
    ' If we are doing a new sub quote then a PPSubQuoteNumber tag will be present
    ' so update it with the number from the newly created sub quote to update the
    ' quotation. Otherwise if creating the first sub quote one will not be there
    ' so we need to create one
    
    ' APS 03/03/00 - Commented out line of code and added missing tagname
    'Set xmlElement = xmlRequestDoc.getElementsByTagName(strTagName).Item(0)
    strTagName = "PPSUBQUOTENUMBER"
    'MS     09/08/00    SYS0834 - Save correct PPSUBQUOTENUMBER
    Set xmlElement = xmlTableNode.selectSingleNode(".//PPSUBQUOTENUMBER")
    'MS     09/08/00    SYS0834 - Save correct PPSUBQUOTENUMBER
    'Dont need to create tag as already have one - just update the
    ' quotation with the latest number
    
    'Set xmlElement = xmlRequestDoc.createElement(strTagName)
    If Not xmlElement Is Nothing Then
'        xmlTableNode.appendChild xmlElement
        xmlElement.Text = strPPSubQuoteNo
    Else
        'MS 14/06/00 If we haven't got a PPSUBQUOTENUMBER we need to create one
        
        Set xmlElement = xmlRequestDoc.createElement("PPSUBQUOTENUMBER")
        xmlElement.Text = strPPSubQuoteNo
        xmlTableNode.appendChild xmlElement
    End If
    
    objXmlAssist.ChangeNodeName xmlTableNode, "PAYMENTPROTECTIONSUBQUOTE", "QUOTATION"

    objIQuotationDO.Update xmlTableNode
    
    Set CreatePPSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreatePPSubQuoteExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlResponseDoc = Nothing
    Set objIQuotationDO = Nothing
    Set objPPSubQuoteBO = Nothing
    Set xmlActionNode = Nothing
    
    Exit Function

CreatePPSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set CreatePPSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreatePPSubQuoteExit

End Function

Private Function CreateMortgageAndLifeSubQuotes(ByVal vxmlQuotationElement As IXMLDOMElement, _
                                                ByVal sqiInstance As SUBQUOTEINSTANCE) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the first sub quote or new sub quote for mortgage and life
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageAndLifeSubQuotesVbErr

    Const strFunctionName As String = "CreateMortgageAndLifeSubQuotes"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlActionNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim objIQuotationDO As IQuotationDO
    Dim objIMortSubQuoteBO As IMortgageSubQuoteBO
    'Dim objLifeSubQuoteBO As LifeSubQuoteBO
    Dim xmlTempResponsElem As IXMLDOMElement
    Dim xmlMortSubQuoteElem As IXMLDOMElement
    
    Dim strMortSubQuoteNo As String
    'Dim strLifeSubQuoteNo As String
    'Dim strResponse As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
   
    ' Create a document with the request node to pass to other BOs
    Set xmlRequestNode = objXmlAssist.GetRequestNode(vxmlQuotationElement.ownerDocument)
    xmlRequestDoc.appendChild xmlRequestNode
    
    Set xmlActionNode = xmlRequestDoc.createElement("CREATE")
    xmlRequestNode.appendChild xmlActionNode
        
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
        Set objIMortSubQuoteBO = New MortgageSubQuoteBO
        'Set objLifeSubQuoteBO = New LifeSubQuoteBO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
        Set objIMortSubQuoteBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
        'Set objLifeSubQuoteBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteBO")
    End If
    
   ' Copy the current instance of data to put into the request
    Set xmlTableNode = vxmlQuotationElement.cloneNode(True)
    xmlActionNode.appendChild xmlTableNode
                         
    objXmlAssist.ChangeNodeName xmlTableNode, "QUOTATION", "MORTGAGESUBQUOTE"
                                     
    ' Create the mortgage sub quote
    If sqiInstance = sqiFirst Then
        Set xmlTempResponsElem = objIMortSubQuoteBO.CreateFirstSubQuote(xmlRequestNode)
    Else
        Set xmlTempResponsElem = objIMortSubQuoteBO.CreateNewSubQuote(xmlRequestNode)
    End If
        
    objErrAssist.CheckResponse xmlTempResponsElem, xmlResponseElem, True
    
    Set xmlMortSubQuoteElem = xmlTempResponsElem.getElementsByTagName("MORTGAGESUBQUOTE").Item(0)
            
    If Not xmlMortSubQuoteElem Is Nothing Then
        Set xmlMortSubQuoteElem = xmlResponseElem.appendChild(xmlMortSubQuoteElem.cloneNode(True))
        strMortSubQuoteNo = objXmlAssist.GetTagValue(xmlMortSubQuoteElem, "MORTGAGESUBQUOTENUMBER")
    End If
    
    If Len(strMortSubQuoteNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "MORTGAGESUBQUOTENUMBER not generated"
    End If

    'MAR1893 GHun No longer required
    'objXmlAssist.ChangeNodeName xmlTableNode, "MORTGAGESUBQUOTE", "LIFESUBQUOTE"
    '
    '' Create the life sub quote
    'If sqiInstance = sqiFirst Then
    '    strResponse = objLifeSubQuoteBO.CreateFirstSubQuote(xmlRequestNode.xml)
    'Else
    '    strResponse = objLifeSubQuoteBO.CreateNewSubQuote(xmlRequestNode.xml)
    'End If
    '
    'objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
    '
    'Set xmlResponseDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    '
    'Set xmlElement = xmlResponseDoc.getElementsByTagName("LIFESUBQUOTENUMBER").Item(0)
    '
    'If Not xmlElement Is Nothing Then
    '    xmlMortSubQuoteElem.appendChild xmlElement.cloneNode(True)
    '    strLifeSubQuoteNo = xmlElement.Text
    'End If
    '
    'If Len(strLifeSubQuoteNo) = 0 Then
    '    objErrAssist.ThrowError TypeName(Me), _
    '                            strFunctionName, _
    '                            oeMissingPrimaryTag, _
    '                            "LIFESUBQUOTENUMBER not generated"
    'End If
    'MAR1893 End

    ' If we are doing a new sub quote then a MortgageSubQuoteNumber tag will be present
    ' so update it with the number from the newly created sub quote to update the
    ' quotation. Otherwise if creating the first sub quote one will not be there
    ' so we need to create one
    Dim strTagName As String
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    
    Set xmlElement = xmlRequestDoc.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = xmlRequestDoc.createElement(strTagName)
        xmlTableNode.appendChild xmlElement
    End If
    
    xmlElement.Text = strMortSubQuoteNo
    
    'MAR1893 GHun No longer required
    '' If we are doing a new sub quote then a LifeSubQuoteNumber tag will be present
    '' so update it with the number from the newly created sub quote to update the
    '' quotation. Otherwise if creating the first sub quote one will not be there
    '' so we need to create one
    'strTagName = "LIFESUBQUOTENUMBER"
    '
    'Set xmlElement = xmlRequestDoc.getElementsByTagName(strTagName).Item(0)
    '
    'If xmlElement Is Nothing Then
    '    Set xmlElement = xmlRequestDoc.createElement(strTagName)
    '    xmlTableNode.appendChild xmlElement
    'End If
    '
    'xmlElement.Text = strLifeSubQuoteNo
    '
    'objXmlAssist.ChangeNodeName xmlTableNode, "LIFESUBQUOTE", "QUOTATION"
    
    objXmlAssist.ChangeNodeName xmlTableNode, "MORTGAGESUBQUOTE", "QUOTATION"
    'MAR1893 End
    
    objIQuotationDO.Update xmlTableNode
            
    Set CreateMortgageAndLifeSubQuotes = xmlResponseElem.cloneNode(True)
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateMortgageAndLifeSubQuotesExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlActionNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlResponseDoc = Nothing
    Set objIQuotationDO = Nothing
    Set objIMortSubQuoteBO = Nothing
    'Set objLifeSubQuoteBO = Nothing
    Set xmlTempResponsElem = Nothing
    Set xmlMortSubQuoteElem = Nothing
    Set objContext = Nothing
    
    Exit Function

CreateMortgageAndLifeSubQuotesVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set CreateMortgageAndLifeSubQuotes = objErrAssist.CreateErrorResponseEx()
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageAndLifeSubQuotesExit

End Function

Private Function SetupCreateQuotationRequest(ByRef rxmlRequestElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  creates the request to create a quotation which is based on whether the
'               quotation is a copy of an existing one
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SetupCreateQuotationRequestVbErr

    Const strFunctionName As String = "SetupCreateQuotationRequest"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    Dim xmlQuotationElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim strQuotationNumber As String
    
    strQuotationNumber = objXmlAssist.GetTagValue(rxmlRequestElement, "QUOTATIONNUMBER")
    
    ' If a quote number is present then get that quote and create a new one based on it
    If Len(strQuotationNumber) > 0 Then
    
        If objContext Is Nothing Then
            Set objIQuotationDO = New QuotationDO
        Else
            Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
        End If
         
        Set xmlQuotationElem = objIQuotationDO.GetData(rxmlRequestElement)
        
        Dim strTagName As String
        
        strTagName = "QUOTATIONCOMPLETE"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = "0"
                
        strTagName = "LOANSANDLIABILITIES"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = ""
        
        strTagName = "MONTHLYMORTGAGEPAYMENTS"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = ""
        
        strTagName = "MORTGAGERELATEDINSURANCE"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = ""
        
        strTagName = "OTHEROUTGOINGS"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = ""
        
        strTagName = "TOTALMONTHLYINCOME"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = ""
        
        strTagName = "TOTALQUOTATIONCOST"
        Set xmlElement = xmlQuotationElem.getElementsByTagName(strTagName).Item(0)
        xmlElement.Text = ""
      
        Set rxmlRequestElement = xmlQuotationElem
    End If
    
    strTagName = "DATEANDTIMEGENERATED"
    Set xmlElement = rxmlRequestElement.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = rxmlRequestElement.ownerDocument.createElement(strTagName)
        rxmlRequestElement.appendChild xmlElement
    End If
    
    xmlElement.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")

    Set SetupCreateQuotationRequest = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SetupCreateQuotationRequestExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIQuotationDO = Nothing
    Set xmlQuotationElem = Nothing
    
    Exit Function

SetupCreateQuotationRequestVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set SetupCreateQuotationRequest = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SetupCreateQuotationRequestExit

End Function

Private Function IQuotationTxBO_CreateFirstBCSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the buildings and contents sub quote
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstBCSubQuoteVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateFirstBCSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlBCResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
              
    Set xmlBCResponseNode = CreateBCSubQuote(xmlElement, sqiFirst)
    objErrAssist.CheckResponse xmlBCResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlBCResponseNode
    
    Set IQuotationTxBO_CreateFirstBCSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstBCSubQuoteExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlBCResponseNode = Nothing
    
    Exit Function

CreateFirstBCSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateFirstBCSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstBCSubQuoteExit

End Function

Private Function IQuotationTxBO_CreateFirstMortgageLifeSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the first mortgage subquote and its linked life subquote
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstMortgageLifeSubQuotesVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateFirstMortgageLifeSubQuotes"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlMortLifeResponseNode As IXMLDOMNode
    
    Set objErrAssist = New ErrAssist
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
          
    Set xmlMortLifeResponseNode = CreateMortgageAndLifeSubQuotes(xmlElement, sqiFirst)
    objErrAssist.CheckResponse xmlMortLifeResponseNode, xmlResponseElem, True
    Set objXmlAssist = New XMLAssist
    objXmlAssist.AttachResponseData xmlResponseElem, xmlMortLifeResponseNode
    
    Set IQuotationTxBO_CreateFirstMortgageLifeSubQuotes = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstMortgageLifeSubQuotesExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlMortLifeResponseNode = Nothing
    Set objContext = Nothing
    
    Exit Function

CreateFirstMortgageLifeSubQuotesVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateFirstMortgageLifeSubQuotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstMortgageLifeSubQuotesExit

End Function

Private Function IQuotationTxBO_CreateFirstPPSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the first payment protection subquote
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstPPSubQuoteVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateFirstPPSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlPPResponseNode As IXMLDOMNode
        
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
           
    Set xmlPPResponseNode = CreatePPSubQuote(xmlElement, sqiFirst)
    objErrAssist.CheckResponse xmlPPResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlPPResponseNode
    
    Set IQuotationTxBO_CreateFirstPPSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstPPSubQuoteExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPPResponseNode = Nothing
    
    Exit Function

CreateFirstPPSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateFirstPPSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstPPSubQuoteExit

End Function

Private Function IQuotationTxBO_CreateNewBCSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create the buildings and contents sub quote
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewBCSubQuoteVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateNewBCSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlBCResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
          
    Set xmlBCResponseNode = CreateBCSubQuote(xmlElement, sqiNew)
    objErrAssist.CheckResponse xmlBCResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlBCResponseNode
    
    Set IQuotationTxBO_CreateNewBCSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewBCSubQuoteExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlBCResponseNode = Nothing
    
    Exit Function

CreateNewBCSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateNewBCSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewBCSubQuoteExit

End Function

Private Function IQuotationTxBO_CreateNewMortgageLifeSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create new mortgage and life subquotes based on existing quotes
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewMortgageLifeSubQuotesVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateNewMortgageLifeSubQuotes"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlMortLifeResponseNode As IXMLDOMNode
    
    Set objErrAssist = New ErrAssist
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
       
    Set xmlMortLifeResponseNode = CreateMortgageAndLifeSubQuotes(xmlElement, sqiNew)
    objErrAssist.CheckResponse xmlMortLifeResponseNode, xmlResponseElem, True
    Set objXmlAssist = New XMLAssist
    objXmlAssist.AttachResponseData xmlResponseElem, xmlMortLifeResponseNode
    
    Set IQuotationTxBO_CreateNewMortgageLifeSubQuotes = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewMortgageLifeSubQuotesExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlMortLifeResponseNode = Nothing
    Set objContext = Nothing
    
    Exit Function

CreateNewMortgageLifeSubQuotesVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateNewMortgageLifeSubQuotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewMortgageLifeSubQuotesExit

End Function

Private Function IQuotationTxBO_CreateNewPPSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create new payment protection subquote based on an existing one
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPPSubQuoteVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateNewPPSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlPPResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
         
    Set xmlPPResponseNode = CreatePPSubQuote(xmlElement, sqiNew)
    objErrAssist.CheckResponse xmlPPResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlPPResponseNode
    
    Set IQuotationTxBO_CreateNewPPSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewPPSubQuoteExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlPPResponseNode = Nothing
    
    Exit Function

CreateNewPPSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateNewPPSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPPSubQuoteExit

End Function

Private Function IQuotationTxBO_CreateNewQuotation(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  creates a new quotation
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewQuotationVbErr

    Const strFunctionName As String = "IQuotationTxBO_CreateNewQuotation"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    Dim xmlRequestElement As IXMLDOMElement
    Dim xmlSetupReponseNode As IXMLDOMNode
    Dim xmlQuoteNoElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
        
    Set xmlRequestElement = xmlElement.cloneNode(True)
    Set xmlSetupReponseNode = SetupCreateQuotationRequest(xmlRequestElement)
    objErrAssist.CheckResponse xmlSetupReponseNode, xmlResponseElem, True
    Set xmlQuoteNoElem = objIQuotationDO.Create(xmlRequestElement)
    xmlResponseElem.appendChild xmlQuoteNoElem
    
    Set IQuotationTxBO_CreateNewQuotation = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewQuotationExit:

    Set objIQuotationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElement = Nothing
    Set xmlSetupReponseNode = Nothing
    
    Exit Function

CreateNewQuotationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_CreateNewQuotation = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit

End Function

Private Function IQuotationTxBO_ReinstateSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Reinstates a particular subquote
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ReinstateSubQuoteVbErr

    Const strFunctionName As String = "IQuotationTxBO_ReinstateSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    Dim xmlUpdateDoc As FreeThreadedDOMDocument40
    Set xmlUpdateDoc = xmlCreateDOMObject()
    Dim xmlUpdateElem As IXMLDOMElement
    'Dim xmlLifeSubQuoteElem As IXMLDOMElement
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Set xmlRequestDoc = xmlCreateDOMObject()
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDoc As FreeThreadedDOMDocument40 'MAR1061
    Dim xmlTempElem As IXMLDOMElement
    Dim objIMortgageSubquoteDO As IMortgageSubQuoteDO
    Dim xmlSubquoteDataNode As IXMLDOMNode
    Dim xmlAppElem As IXMLDOMElement
    Dim objAppBO As Object
    Dim strPurchasePrice As String
    Dim strResponse As String
    
    Set objErrAssist = New ErrAssist
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
        Set objIMortgageSubquoteDO = New MortgageSubQuoteDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
        Set objIMortgageSubquoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
    End If
    
    Set xmlUpdateElem = xmlElement.cloneNode(True)
    xmlUpdateDoc.appendChild xmlUpdateElem
    
    Dim strActiveSubQuoteNo As String
    Dim strSelectedSubQuoteNo As String
    Dim strBusinessType As String
    Dim strTagName As String
    
    Set objXmlAssist = New XMLAssist
    strActiveSubQuoteNo = objXmlAssist.GetTagValue(xmlUpdateElem, "ACTIVESUBQUOTENUMBER")
    strSelectedSubQuoteNo = objXmlAssist.GetTagValue(xmlUpdateElem, "SELECTEDSUBQUOTENUMBER")
    strBusinessType = objXmlAssist.GetTagValue(xmlUpdateElem, "BUSINESSTYPE")
    
    If strSelectedSubQuoteNo = strActiveSubQuoteNo Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 184
    End If
    
    ' Determine the field to update on quotation to show the reinstated sub quote
    If StrComp(strBusinessType, "Mortgage", vbTextCompare) = 0 Then
        strTagName = "MORTGAGESUBQUOTENUMBER"
        'Set xmlLifeSubQuoteElem = vXMLRequest.ownerDocument.createElement("LIFESUBQUOTENUMBER")
        'xmlUpdateElem.appendChild xmlLifeSubQuoteElem
        'xmlLifeSubQuoteElem.Text = strSelectedSubQuoteNo
    ElseIf StrComp(strBusinessType, "BC", vbTextCompare) = 0 Then
        strTagName = "BCSUBQUOTENUMBER"
    ElseIf StrComp(strBusinessType, "PP", vbTextCompare) = 0 Then
        strTagName = "PPSUBQUOTENUMBER"
    Else
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "BUSINESSTYPE = " & strBusinessType
                                
    End If
    
    objXmlAssist.ChangeNodeName xmlUpdateElem, "SELECTEDSUBQUOTENUMBER", strTagName
    
    objIQuotationDO.Update xmlUpdateElem
    
    Set xmlAppDoc = xmlCreateDOMObject()
    
    'MAR1061 update AFF with new purchaseprice
    If StrComp(strBusinessType, "Mortgage", vbTextCompare) = 0 Then
        'Get the MortgageSubquote details
        Set xmlRequestElem = vXMLRequest.cloneNode(True)
        xmlAppDoc.appendChild xmlRequestElem
        Set xmlAppElem = xmlRequestElem.appendChild(xmlAppDoc.createElement("MORTGAGESUBQUOTE"))
        Set xmlTempElem = xmlAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONNUMBER"))
        xmlTempElem.Text = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
        Set xmlTempElem = xmlAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONFACTFINDNUMBER"))
        xmlTempElem.Text = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")
        Set xmlTempElem = xmlAppElem.appendChild(xmlAppDoc.createElement("MORTGAGESUBQUOTENUMBER"))
        xmlTempElem.Text = objXmlAssist.GetTagValue(xmlElement, "SELECTEDSUBQUOTENUMBER")
        
        Set xmlSubquoteDataNode = objIMortgageSubquoteDO.GetData(xmlRequestElem)
        If Not xmlSubquoteDataNode Is Nothing Then
            strPurchasePrice = objXmlAssist.GetTagValue(xmlSubquoteDataNode, "PURCHASEPRICEORESTIMATEDVALUE")
            
            Set xmlRequestElem = vXMLRequest.cloneNode(True)
            Set xmlAppDoc = New FreeThreadedDOMDocument40
            xmlAppDoc.appendChild xmlRequestElem
            Set xmlAppElem = xmlRequestElem.appendChild(xmlAppDoc.createElement("APPLICATIONFACTFIND"))
            Set xmlTempElem = xmlAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONNUMBER"))
            xmlTempElem.Text = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
            Set xmlTempElem = xmlAppElem.appendChild(xmlAppDoc.createElement("APPLICATIONFACTFINDNUMBER"))
            xmlTempElem.Text = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")
            Set xmlTempElem = xmlAppElem.appendChild(xmlAppDoc.createElement("PURCHASEPRICEORESTIMATEDVALUE"))
            xmlTempElem.Text = strPurchasePrice
            
            Set objAppBO = objContext.CreateInstance("omApp.ApplicationBO")
            strResponse = objAppBO.UpdateApplicationFactFind(xmlAppDoc.xml)
            objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem    'MAR1893 GHun
        End If
    End If
    
    Set IQuotationTxBO_ReinstateSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ReinstateSubQuoteExit:
    
    Set objIQuotationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlUpdateDoc = Nothing
    Set xmlUpdateElem = Nothing
    'Set xmlLifeSubQuoteElem = Nothing
    Set objContext = Nothing
    
    Exit Function

ReinstateSubQuoteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_ReinstateSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateSubQuoteExit

End Function

Private Function IQuotationTxBO_SaveOneOffCostDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
    On Error GoTo SaveOneOffCostDetailsVbErr

    Const strFunctionName As String = "IQuotationTxBO_SaveOneOffCostDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    'Dim objILifeSubQuoteBO As ILifeSubQuoteBO
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlClonedElement As IXMLDOMElement
    Dim xmlClonedRequestElem As IXMLDOMElement
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        'Set objILifeSubQuoteBO = New LifeSubQuoteBO
    Else
        Set objIMortgageSubQuoteBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
        'Set objILifeSubQuoteBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteBO")
    End If
    
    Set xmlClonedRequestElem = objXmlAssist.GetRequestNode(vXMLRequest.ownerDocument)
    Set xmlClonedElement = xmlElement.cloneNode(True)
    
    xmlRequestDoc.appendChild xmlClonedRequestElem
    xmlClonedRequestElem.appendChild xmlClonedElement
    
    objXmlAssist.ChangeNodeName xmlClonedElement, "QUOTATION", "MORTGAGESUBQUOTE"
    Set xmlTempResponseNode = objIMortgageSubQuoteBO.SaveOneOffCostDetails(xmlClonedRequestElem)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True  'We need to know if a save method fails
    
    'MV - 16/09/2002 -  BMIDS00432  Removed LifeSubQuote functionality
    'objXmlAssist.ChangeNodeName xmlClonedElement, "MORTGAGESUBQUOTE", "LIFESUBQUOTE"
    'Set xmlTempResponseNode = objILifeSubQuoteBO.SaveRecostedBenefits(xmlClonedRequestElem)
    'objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True 'We need to know if a save method fails
    
    Set IQuotationTxBO_SaveOneOffCostDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveOneOffCostDetailsExit:
    
    Set objIMortgageSubQuoteBO = Nothing
    'Set objILifeSubQuoteBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

SaveOneOffCostDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_SaveOneOffCostDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOneOffCostDetailsExit

End Function

Private Function IQuotationTxBO_StoreQuotation(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Stores a quotation and sets it to complete
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo StoreQuotationVbErr

    Const strFunctionName As String = "IQuotationTxBO_StoreQuotation"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    Dim xmlQuoteCompleteElem As IXMLDOMElement
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlRequestElem As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlRequestElem = vXMLRequest.cloneNode(True)
    xmlRequestDoc.appendChild xmlRequestElem
          
    If xmlRequestElem.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = xmlRequestElem.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    
    Dim strTagName As String
    strTagName = "QUOTATIONCOMPLETE"
                
    ' Get the Quotation Complete element
    Set xmlQuoteCompleteElem = xmlElement.getElementsByTagName(strTagName).Item(0)
    
    ' Create one if not found
    If xmlQuoteCompleteElem Is Nothing Then
        Set xmlQuoteCompleteElem = xmlRequestElem.ownerDocument.createElement(strTagName)
        xmlElement.appendChild xmlQuoteCompleteElem
    End If
    
    ' Set to complete
    xmlQuoteCompleteElem.Text = "1"
    
    objIQuotationDO.Update xmlElement
    
    Set IQuotationTxBO_StoreQuotation = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
StoreQuotationExit:
    
    Set objIQuotationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRequestElem = Nothing
    
    Exit Function

StoreQuotationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_StoreQuotation = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume StoreQuotationExit

End Function

Private Function IQuotationTxBO_Update(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IQuotationTxBO_Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    
    objIQuotationDO.Update xmlElement
    
    Set IQuotationTxBO_Update = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIQuotationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationTxBO_Update = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

