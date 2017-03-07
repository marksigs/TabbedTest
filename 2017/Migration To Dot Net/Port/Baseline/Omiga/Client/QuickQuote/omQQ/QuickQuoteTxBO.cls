VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "QuickQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuoteTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Transaction object for Quick Quote
'
'Dependencies:  List any other dependent components
'               QuickQuoteBO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    22/03/00    Created
'PSC    07/04/00    Create objects correctly under MTS in CalculateMortgageCosts
'JLD    15/05/00    SYS0715 Update deposit on Reset.
'PSC    30/06/00    SYS1012 Remove ProcessAddedOneOffCosts
'SG     28/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IQuickQuoteTxBO
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXMLAssist As XMLAssist             ' Remove if not required

'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MV     18/07/2002  BMIDS00179 - Core Upgrade Rollback - Modified IQuickQuoteTxBO_ResetMortgageSubQuote
'DPF    31/07/2002  BMIDS00266 - Have removed call to save routine for life cover
'MV     12/08/2002  BMIDS00311 - Core Ref : SYS1705 Removed SafeRef; Modified IQuickQuoteTxBO_CalculateMortgageCosts
'------------------------------------------------------------------------------------------
Private Function IQuickQuoteTxBO_CalculateAndSaveBCSubQuote( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAndSaveBCSubQuoteVbErr

    Const strFunctionName As String = "IQuickQuoteTxBO_CalculateAndSaveBCSubQuote"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Dim strDOB As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = _
            m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If

    ' Fetch input parameters into local variables
    Dim strCustNumber1 As String, strCustNumber2 As String
    Dim strCustVersionNumber1 As String, strCustVersionNumber2 As String
    Dim strApplNumber As String, strAFFNumber As String
    
    strCustNumber1 = vxmlRequest.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustVersionNumber1 = vxmlRequest.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    
    Set xmlNode = vxmlRequest.selectNodes(".//CUSTOMER").Item(1)
    If Not xmlNode Is Nothing Then
        strCustNumber2 = xmlNode.selectSingleNode(".//CUSTOMERNUMBER").Text
        strCustVersionNumber2 = xmlNode.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    End If
    
    strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text

    Set xmlNode = objICostModelAssistBO.CalculateAndSaveBCSubQuote( _
                       vxmlRequest, strCustNumber1, strCustNumber2, strCustVersionNumber1, _
                        strApplNumber, strAFFNumber, "QUICK QUOTE")
                        
    m_objErrAssist.CheckResponse xmlNode, xmlResponseElem, True
    
    Set IQuickQuoteTxBO_CalculateAndSaveBCSubQuote = xmlNode
    
CalculateAndSaveBCSubQuoteExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set objICostModelAssistBO = Nothing

   Exit Function
CalculateAndSaveBCSubQuoteVbErr:
        If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteTxBO_CalculateAndSaveBCSubQuote = _
                            m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit

End Function


Private Function IQuickQuoteTxBO_CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement, _
                                                              xmlLoanCompositionNode As IXMLDOMNode, _
                                                              xmlAppNo As IXMLDOMNode, _
                                                              xmlAppFactFindNo As IXMLDOMNode, _
                                                              xmlLifeSubQuoteNumber As IXMLDOMNode, _
                                                              xmlMortgageSubQuoteNumber As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo CreateVbErr

    Const strFunctionName As String = "IQuickQuoteTxBO_CalculateMortgageCosts"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlMortgageCostsResp As IXMLDOMNode
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objIQuickQuoteBO As IQuickQuoteBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objIQuickQuoteBO = New QuickQuoteBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".CostModelAssistBO")
        Set objIQuickQuoteBO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".QuickQuoteBO")
    End If

    ' Start transaction
    Set xmlMortgageCostsResp = objICostModelAssistBO.CalculateMortgageCosts(vxmlRequest, _
                                                                            xmlLoanCompositionNode, _
                                                                            xmlAppNo, _
                                                                            xmlAppFactFindNo, _
                                                                            xmlLifeSubQuoteNumber, _
                                                                            xmlMortgageSubQuoteNumber)
                                                                            
    m_objErrAssist.CheckResponse xmlMortgageCostsResp, xmlResponseElem, True

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'BMIDS00266 - DPF 31/07/02 - THIS BIT TAKEN OUT AS WE DON'T OFFER LIFE COVER
    ' Calculate and save Life Cover Costs
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Dim xmlLifeCoverReq As IXMLDOMNode
    'Dim xmlLifeCoverResp As IXMLDOMNode
    
    'Set xmlLifeCoverReq = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    'Set xmlDOMDoc = xmlLifeCoverReq.ownerDocument
    
    'm_objXMLAssist.AttachResponseData xmlLifeCoverReq, xmlMortgageCostsResp
    
    ' Call back...
    ' Set xmlLifeCoverResp = objIQuickQuoteBO.CalculateAndSaveLifeCoverCosts(xmlLifeCoverReq)
    'm_objErrAssist.CheckResponse xmlLifeCoverResp, xmlResponseElem, True
    ' End transaction

    

    Set IQuickQuoteTxBO_CalculateMortgageCosts = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:


    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteTxBO_CalculateMortgageCosts = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit

End Function


Private Function IQuickQuoteTxBO_CreateNewQuotation( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateVbErr

    Const strFunctionName As String = "IQuickQuoteTxBO_CreateNewQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim strNewQuotationNumber As String
    Dim xmlTableElem As IXMLDOMNode, xmlElement As IXMLDOMNode
    Dim xmlRootElement As IXMLDOMNode
    
    Dim objICostModelAssistDO As ICostModelAssistDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem

    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = "QUICKQUOTE" Then
        Set xmlRootElement = vxmlRequest
    Else
        Set xmlRootElement = vxmlRequest.getElementsByTagName("QUICKQUOTE").Item(0)
    End If
    
    If xmlRootElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                  "QUICKQUOTE" & " tag not found"
    End If


    If m_objContext Is Nothing Then
        Set objICostModelAssistDO = New CostModelAssistDO
    Else
        Set objICostModelAssistDO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".CostModelAssistDO")
    End If

    ' Call the function in CostModelAssist and add the Quote number to Response
    strNewQuotationNumber = objICostModelAssistDO.CreateNewQuotation(vxmlRequest)
    
    ' Add the new quotation number to Response
    Set xmlTableElem = xmlOut.createElement("QUOTATION")
    Set xmlElement = xmlOut.createElement("NEWQUOTATIONNUMBER")
    xmlElement.Text = strNewQuotationNumber
    xmlResponseElem.appendChild xmlElement
    
    Set IQuickQuoteTxBO_CreateNewQuotation = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:
    Set objICostModelAssistDO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlRootElement = Nothing
    
    Exit Function
CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteTxBO_CreateNewQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IQuickQuoteTxBO_RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  RecostMortgageComponentAndLifeCover multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest
'
'    <QUICKQUOTE>
'        <APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER>
'        <CONTEXT>
'        <INTERESTONLYAMOUNT>
'        <CAPITALANDINTERESTAMOUNT>
'        <APPLICATIONDATE>
'        <ONEOFFCOSTLIST>
'            ...
'        </ONEOFFCOSTLIST>
'        <CUSTOMERLIST>
'            <CUSTOMERNUMBER>
'            <CUSTOMERVERSIONNUMBER>
'            <CUSTOMERNUMBER>
'            <CUSTOMERVERSIONNUMBER>
'        </CUSTOMERLIST>
'        <LOANCOMPONENT>
'            <LOANCOMPONENTSEQUENCENUMBER>
'        </LOANCOMPONENT>
'        <LOANCOMPONENTLIST>
'            ...
'        </LOANCOMPONENTLIST>
'    </QUICKQUOTE>
'
'
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo RecostMortgageComponentAndLifeCoverErr

    Const strFunctionName As String = "IQuickQuoteTxBO_RecostMortgageComponentAndLifeCover"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlRootElement As IXMLDOMElement
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = "QUICKQUOTE" Then
        Set xmlRootElement = vxmlRequest
    Else
        Set xmlRootElement = vxmlRequest.getElementsByTagName("QUICKQUOTE").Item(0)
    End If

    If xmlRootElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                  "QUICKQUOTE" & " tag not found"
    End If

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    Set xmlTempResponseNode = _
            objICostModelAssistBO.RecostMortgageComponentAndLifeCover(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    '--------------------------------------------------------------------------------
    ' Re-cost the associated Life Cover
    '--------------------------------------------------------------------------------
    Set xmlTempResponseNode = RecostLifeCover(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
        
    Set IQuickQuoteTxBO_RecostMortgageComponentAndLifeCover = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RecostMortgageComponentAndLifeCoverExit:
    
    Set xmlRootElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objICostModelAssistBO = Nothing
    
    Exit Function

RecostMortgageComponentAndLifeCoverErr:
     If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteTxBO_RecostMortgageComponentAndLifeCover = _
                                        m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit
End Function

Private Function RecostLifeCover(ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMElement
'--------------------------------------------------------------------------------------------
'Description : Recosts the loan component for ApplicationQuote
'              Called from - RecostMortgageComponentAndLifeCover
'--------------------------------------------------------------------------------------------
On Error GoTo RecostVbErr
    
    Const strFunctionName As String = "RecostLifeCover"
    
    Dim objApplicationBO As ApplicationBO
    Dim objQuickQuoteBO As QuickQuoteBO
    Dim objLifeCoverBO As LifeCoverBO
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCustRoleElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMNode, xmlCustNode As IXMLDOMNode
    
    Dim strCust1Number As String, strCust1VersionNumber As String
    Dim strCust2Number As String, strCust2VersionNumber As String
    Dim strApplNumber As String, strAFFNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    '-------------------------------------------------------------------------
    ' Get Main Applicants and build input for 'BuildLifeCoverCalculationInput'
    '-------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlTableElem
    
    Dim strTag As String, strResponse As String
    strTag = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    strApplNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlElement.Text = strApplNumber
    xmlTableElem.appendChild xmlElement
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    strAFFNumber = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlElement.Text = strAFFNumber
    xmlTableElem.appendChild xmlElement
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    strResponse = objApplicationBO.FindMainApplicants(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set xmlCustRoleElem = xmlDoc.selectNodes("//CUSTOMERROLE").Item(0)
    strCust1Number = xmlCustRoleElem.selectSingleNode("CUSTOMERNUMBER").Text
    strCust1VersionNumber = xmlCustRoleElem.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
    
    Set xmlCustRoleElem = xmlDoc.selectNodes("CUSTOMERROLE").Item(1)
    If Not xmlCustRoleElem Is Nothing Then
        strCust2Number = xmlCustRoleElem.selectSingleNode("CUSTOMERNUMBER").Text
        strCust2VersionNumber = xmlCustRoleElem.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
    End If
    
    '-----------------------------------------------------------------------------------
    'Build Life Cover Calculations Input Interface
    '-----------------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("LIFECOVERCOSTS")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlListElem = vxmlRequest.selectSingleNode(".//LOANCOMPONENTLIST").cloneNode(True)
    xmlTableElem.appendChild xmlListElem
    
    Set xmlListElem = xmlOut.createElement("CUSTOMERLIST")
    xmlTableElem.appendChild xmlListElem
    
    Set xmlCustNode = xmlOut.createElement("CUSTOMER")
    xmlListElem.appendChild xmlCustNode
    
    Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
    xmlElement.Text = strCust1Number
    xmlCustNode.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
    xmlElement.Text = strCust1VersionNumber
    xmlCustNode.appendChild xmlElement
    
    If strCust2Number <> "" Then
        Set xmlCustNode = xmlOut.createElement("CUSTOMER")
        xmlListElem.appendChild xmlCustNode
    
        Set xmlElement = xmlOut.createElement("CUSTOMERNUMBER")
        xmlElement.Text = strCust2Number
        xmlCustNode.appendChild xmlElement
    
        Set xmlElement = xmlOut.createElement("CUSTOMERVERSIONNUMBER")
        xmlElement.Text = strCust2VersionNumber
        xmlCustNode.appendChild xmlElement
    End If
    
    If m_objContext Is Nothing Then
        Set objQuickQuoteBO = New QuickQuoteBO
    Else
        Set objQuickQuoteBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteBO")
    End If
    
    strResponse = objQuickQuoteBO.BuildLifeCoverCalcInputInterface(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    '---------------------------------------------------------------------------
    'Calculate Life Cover Premiums
    '----------------------------------------------------------------------------
    If m_objContext Is Nothing Then
        Set objLifeCoverBO = New LifeCoverBO
    Else
        Set objLifeCoverBO = m_objContext.CreateInstance(gstrLIFECOVER_COMPONENT & ".LifeCoverBO")
    End If
    
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    m_objXMLAssist.AttachResponseData xmlNewRequest, xmlDoc.documentElement
    
    strResponse = objLifeCoverBO.CalculatePremiums(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set RecostLifeCover = xmlDoc.documentElement
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostVbExit:
    Set objApplicationBO = Nothing
    Set objQuickQuoteBO = Nothing
    Set objLifeCoverBO = Nothing
    
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlCustRoleElem = Nothing
    Set xmlListElem = Nothing
    Set xmlCustNode = Nothing
    
    Exit Function
RecostVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set RecostLifeCover = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostVbExit

End Function

Private Function IQuickQuoteTxBO_ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) _
    As IXMLDOMNode

On Error GoTo ReinstateQuotationVbErr

    Const strFunctionName As String = "IQuickQuoteTxBO_ReinstateQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim strQuotationNumber As String
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".CostModelAssistBO")
    End If
    
    Set xmlTempResponseNode = objICostModelAssistBO.ReinstateQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
'    ' On successful execution add message 217 as warning
'    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 217
    
    Set IQuickQuoteTxBO_ReinstateQuotation = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ReinstateQuotationExit:
    Set objICostModelAssistBO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

ReinstateQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteTxBO_ReinstateQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume ReinstateQuotationExit

End Function

Private Function IQuickQuoteTxBO_ResetMortgageSubQuote( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

    On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "IQuickQuoteTxBO_ResetMortgageSubQuote"

    Dim xmlOut As FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlTableElem As IXMLDOMNode
    Dim objIQuickQuoteBO As IQuickQuoteBO
    Dim objIApplicationDO As IApplicationDO
    Dim objMSQBO As MortgageSubQuoteBO
    Dim objConvertAssist As ConvertAssist
    Set objConvertAssist = New ConvertAssist
    Dim lngDeposit As Long
    Dim strAmtRequested As String
    
    Dim lngLTV As Long
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '----------------------------------------------------
    ' Recalculate LTV using the new amount requested
    '----------------------------------------------------
    If m_objContext Is Nothing Then
        Set objIQuickQuoteBO = New QuickQuoteBO
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIQuickQuoteBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteBO")
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
    End If
    
    Set xmlNewRequest = vxmlRequest.cloneNode(True)
        
    m_objXMLAssist.ChangeNodeName xmlNewRequest, "QUICKQUOTE", "LTV"
        
    'Note : objIQuickQuoteBO.CalcCostModelLTV is assuming the Request to be attached to FreeThreadedDOMDocument40.
    '       (in Search Pattern in the method CostModelAssistBO.CalcCostModelLTV)
    Set xmlDoc = m_objXMLAssist.load(xmlNewRequest.xml, TypeName(Me), strFunctionName)
    
    Set xmlTempResponseNode = objIQuickQuoteBO.CalcCostModelLTV(xmlDoc.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    lngLTV = m_objXMLAssist.GetMandatoryElementText(xmlTempResponseNode, ".//LTV")
    
    ' attach LTV Element to Response
    xmlResponseElem.appendChild xmlTempResponseNode.selectSingleNode(".//LTV")
    
    '-------------------------------------------------------------------------------
    ' Recalculate Deposit           15/05/00 JLD SYS0715
    '-------------------------------------------------------------------------------
    Set xmlTempResponseNode = objIApplicationDO.GetApplicationData(vxmlRequest)
    Set xmlTempNode = xmlTempResponseNode.selectSingleNode(".//PURCHASEPRICEORESTIMATEDVALUE")
    lngDeposit = 0
    If Not xmlTempNode Is Nothing Then
        strAmtRequested = vxmlRequest.selectSingleNode(".//AMOUNTREQUESTED").Text
        If objConvertAssist.CSafeLng(xmlTempNode.Text) > objConvertAssist.CSafeLng(strAmtRequested) Then
            lngDeposit = objConvertAssist.CSafeLng(xmlTempNode.Text) - objConvertAssist.CSafeLng(strAmtRequested)
        End If
    End If
    
    '------------------------------------------------------------------------------
    ' Reset the mortgage sub-quote
    '-------------------------------------------------------------------------------
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//AMOUNTREQUESTED").cloneNode(True)
    xmlTableElem.appendChild xmlElement
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
    'SG 28/05/02 SYS4767 START
    'JLD MSMS0034 pass through the typeofapplciation and typeofBuyer to be reset too.
    'Set xmlElement = vxmlRequest.selectSingleNode(".//TYPEOFAPPLICATION").cloneNode(True)
    'xmlTableElem.appendChild xmlElement
    'Set xmlElement = vxmlRequest.selectSingleNode(".//TYPEOFBUYER").cloneNode(True)
    'xmlTableElem.appendChild xmlElement
    'SG 28/05/02 SYS4767 END
    
    Set xmlElement = xmlOut.createElement("LTV")
    xmlElement.Text = lngLTV
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("DEPOSIT")
    xmlElement.Text = lngDeposit
    xmlTableElem.appendChild xmlElement
    
    If m_objContext Is Nothing Then
        Set objMSQBO = New MortgageSubQuoteBO
    Else
        Set objMSQBO = _
            m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
    Dim strResponse As String
    strResponse = objMSQBO.ResetMortgageSubQuote(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set IQuickQuoteTxBO_ResetMortgageSubQuote = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ResetMortgageSubQuoteExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set objIQuickQuoteBO = Nothing
    Set objMSQBO = Nothing
    Set objIApplicationDO = Nothing
    
    Exit Function

ResetMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteTxBO_ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXMLAssist = New XMLAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXMLAssist = Nothing

End Sub

