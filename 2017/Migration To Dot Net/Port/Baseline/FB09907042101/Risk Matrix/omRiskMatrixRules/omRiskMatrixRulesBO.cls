VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omRiskMatrixRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omRiskMatrixRulesBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Risk Matrix Rules Business Object translates determines risk rules
'               for Risk Matrix Component.
'               contains (implied) standard interface:
'               Function omRiskMatrixRulesRequest(ByVal vstrXMLRequest As String) As String
'Dependencies:  RiskMatrixRulesBO
'               OmApp.ApplicationBO
'               omQQ.QuickQuoteApplicantDetailsBO
'               omBase.GlobalBandedParameterBO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   CreateDate                          Package Description
'CWP    Tuesday, May 22, 2001 1:56:02 PM    04D01.T65

Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function omRiskMatrixRulesRequest(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:
'   Risk Matrix Rules Request, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo omRunRiskMatrixErr
    
    Const strFunctionName As String = "omRiskMatrixRulesRequest"
    
    Dim xmlInDoc As New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Dim xmlOutDoc As New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' will be superceded by any subsequent SetAbort
    gobjContext.SetComplete

    ' ensure global objects released before running request
            
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    ' load XML REQUEST
    xmlInDoc.loadXML vstrXMLRequest
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        ' multiple operations
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                ' copy attributes from REQUEST node to OPERATION node
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If

    xmlResponseElem.setAttribute "TYPE", "SUCCESS" ' if we get here, everything is OK
    
    omRiskMatrixRulesRequest = xmlOutDoc.xml
        
omRunRiskMatrixErr: 'release objects
        
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

OmRequestVbErr:
        
    omRiskMatrixRulesRequest = errCreateErrorResponse() ' create ERROR response block from Err object
        
    GoTo omRunRiskMatrixErr 'perform clean up

End Function

Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Request broker, private interface, performs single REQUEST OPERATION
'   Any results of operation (data retrieved, error info etc.) appended to vxmlResponseNode
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       raises oeNotImplemented if OPERATION value not in list of known functions.
'------------------------------------------------------------------------------------------

    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
        
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text ' get operation (function) name from REQUEST/@OPERATION
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text ' get operation (function) name from OPERATION/@NAME
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
    
        Case "RUNRISKMATRIX"
            RiskMatrixAdjustmentPercent vxmlRequestNode, vxmlResponseNode
                        
        Case Else
            errThrowError strFunctionName, oeNotImplemented, vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    
    End Select
    
DoRequestExit:
        
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub RiskMatrixAdjustmentPercent(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve Calculate Adjustment Percent
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    
    Dim strRequest As String
    Dim strResponse As String
    Dim strIncomeScore As String
    Dim strCustomerOrder As String
    Dim strCustomerNumber As String
    Dim strGetHighestEarners As String
    Dim strApplicationMode As String
    Dim strCustomerIncome As String
    Dim strCustomerAIncome As String
    Dim strCustomerBIncome As String
    Dim strCustomerRoleType As String
    Dim strApplicationNumber As String
    Dim strCustomerAnnualIncome As String
    Dim strCustomerVersionNumber As String
    Dim strApplicationFactFindNumber As String
    Dim dblTotalGrossIncomeMonthly As Double
    Dim objApplicationBO As Object
    Dim objGlobalBandParameterBO As Object
    Dim objQuickQuoteApplicationBO As Object
    Dim xmlElementReq As IXMLDOMNode
    Dim xmlRiskMatrixNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseDocument As New FreeThreadedDOMDocument40
    xmlResponseDocument.validateOnParse = False
    xmlResponseDocument.setProperty "NewParser", True
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
            
    Const strFunctionName As String = "RiskMatrixAdjustmentPercent"
            
    On Error GoTo RiskMatrixAdjustmentPercentExit
        
    Set xmlRiskMatrixNode = xmlGetMandatoryNode(vxmlRequestNode, "CALCADJUSTMENTPERCENT") 'check for mandatory node, if exists then
    xmlCheckMandatoryAttribute xmlRiskMatrixNode, "APPLICATIONMODE" 'check for mandatory node, if exists then
    
    strApplicationMode = UCase(xmlGetAttributeText(xmlRiskMatrixNode, "APPLICATIONMODE")) 'find out which mode operating in
    Set xmlElementReq = xmlMakeNodeElementBased(vxmlRequestNode, True, "REQUEST") 'convert xml from attribute to element format
    
    Select Case strApplicationMode
            
            Case "COST MODELLING"
                                
                If objApplicationBO Is Nothing Then
                    Set objApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO") 'create object
                End If
                strGetHighestEarners = objApplicationBO.GetHighestEarners(xmlElementReq.xml)
                
                Set xmlTempResponseNode = xmlLoad(strGetHighestEarners, strFunctionName)  'get highest earners
                errCheckXMLResponse xmlTempResponseNode.xml, True 'if no error in response
                                                                
                Set xmlOperationNodeList = xmlTempResponseNode.selectNodes(".//EMPLOYMENTLIST")
                For Each xmlOperationNode In xmlOperationNodeList 'for each customer list in XML string
                    
                    For Each xmlElement In xmlOperationNode.childNodes 'for each child
                        strCustomerIncome = xmlGetNodeText(xmlElement, "YEARLYTOTALAMOUNT")  'get income value
                        dblTotalGrossIncomeMonthly = dblTotalGrossIncomeMonthly + CDbl(strCustomerIncome)
                    Next xmlElement
                    
                Next xmlOperationNode
                
                dblTotalGrossIncomeMonthly = dblTotalGrossIncomeMonthly / 12 'derive monthly income amount
                                                
            Case Else
                                                                        
                Set xmlTempResponseNode = xmlGetMandatoryNode(xmlElementReq, "CALCADJUSTMENTPERCENT") 'check for mandatory node, if exists then
                strApplicationNumber = xmlGetNodeText(xmlTempResponseNode, "APPLICATIONNUMBER") 'get application number
                strApplicationFactFindNumber = xmlGetNodeText(xmlTempResponseNode, "APPLICATIONFACTFINDNUMBER")
                
                Set xmlTempResponseNode = xmlGetMandatoryNode(xmlElementReq, ".//CUSTOMER") 'check for mandatory node, if exists then
                strCustomerNumber = xmlGetNodeText(xmlTempResponseNode, "CUSTOMERNUMBER") 'get customer number
                strCustomerVersionNumber = xmlGetNodeText(xmlTempResponseNode, "CUSTOMERVERSIONNUMBER") 'get customer version number
                strCustomerRoleType = xmlGetNodeText(xmlTempResponseNode, "CUSTOMERROLETYPE") 'get customer role type
                strCustomerOrder = xmlGetNodeText(xmlTempResponseNode, "CUSTOMERORDER") 'get customer order
                                
                'Organise response XML string for querying Quick Quote Application Details
                strRequest = "<REQUEST>"
                strRequest = strRequest & "<CUSTOMERROLELIST>"
                
                Set xmlOperationNodeList = xmlElementReq.selectNodes(".//CUSTOMERLIST")
                For Each xmlOperationNode In xmlOperationNodeList 'for each customer list in XML string
                    
                    For Each xmlElement In xmlOperationNode.childNodes 'for each child
                        
                        strRequest = strRequest & "<CUSTOMERROLE>"
                        strRequest = strRequest & "<APPLICATIONNUMBER>" & strApplicationNumber & "</APPLICATIONNUMBER>"
                        strRequest = strRequest & "<APPLICATIONFACTFINDNUMBER>" & strApplicationFactFindNumber & "</APPLICATIONFACTFINDNUMBER>"
                        
                        strCustomerNumber = xmlGetNodeText(xmlElement, "CUSTOMERNUMBER") 'get customer number
                        strRequest = strRequest & "<CUSTOMERNUMBER>" & strCustomerNumber & "</CUSTOMERNUMBER>"
                        
                        strRequest = strRequest & "<CUSTOMERVERSIONNUMBER>" & strCustomerVersionNumber & "</CUSTOMERVERSIONNUMBER>"
                        strRequest = strRequest & "<CUSTOMERROLETYPE>" & strCustomerRoleType & "</CUSTOMERROLETYPE>"
                        strRequest = strRequest & "<CUSTOMERORDER>" & strCustomerOrder & "</CUSTOMERORDER>"
                        strRequest = strRequest & "</CUSTOMERROLE>"
                        
                    Next xmlElement
                    
                Next xmlOperationNode
                
                'Close response XML string
                strRequest = strRequest & "</CUSTOMERROLELIST>"
                strRequest = strRequest & "</REQUEST>"
                                                                                        
                Set objQuickQuoteApplicationBO = gobjContext.CreateInstance("omQQ.QuickQuoteApplicantDetailsBO") 'create object
                Set xmlTempResponseNode = xmlLoad(objQuickQuoteApplicationBO.GetQuickQuoteData(strRequest), strFunctionName) 'call quick quote application BO
                errCheckXMLResponse xmlTempResponseNode.xml, True 'if no error in response then
                
                xmlCheckMandatoryNode xmlTempResponseNode, ".//QUICKQUOTEAPPLICANTDETAILSLIST" 'check mandatory node exists
                    
                Set xmlOperationNodeList = xmlTempResponseNode.selectNodes(".//QUICKQUOTEAPPLICANTDETAILSLIST")
                For Each xmlOperationNode In xmlOperationNodeList 'for each quick quote application
                     For Each xmlElement In xmlOperationNode.childNodes 'for each child
                        strCustomerAnnualIncome = xmlGetNodeText(xmlElement, "ANNUALINCOMEORNETPROFIT") 'get annual income
                        dblTotalGrossIncomeMonthly = dblTotalGrossIncomeMonthly + CDbl(strCustomerAnnualIncome) 'derive total gross income
                     Next xmlElement
                Next xmlOperationNode
                dblTotalGrossIncomeMonthly = dblTotalGrossIncomeMonthly / 12 'derive total gross monthly income

    End Select
            
    'Call Global Banded Parameters with Income and derive Score
    
    If objGlobalBandParameterBO Is Nothing Then
        Set objGlobalBandParameterBO = gobjContext.CreateInstance("omBase.GlobalBandedParameterBO") 'create object
    End If
    Set xmlTempResponseNode = xmlLoad(objGlobalBandParameterBO.GetCurrentParameter("IncomeScore", CStr(dblTotalGrossIncomeMonthly)), strFunctionName)  'calculate score
    errCheckXMLResponse xmlTempResponseNode.xml, True 'if no error in response
    
    strIncomeScore = xmlGetNodeText(xmlTempResponseNode, ".//PERCENTAGE") 'find percentage value in response
        
    Set xmlResponseElement = xmlResponseDocument.createElement("RISKMATRIXADJUSTMENTPERCENT") 'create XML result string
    xmlResponseElement.setAttribute "CALCULATEDADJUSTMENT", strIncomeScore 'set value to XML result string
    vxmlResponseNode.appendChild xmlResponseElement 'append result to response document
        
RiskMatrixAdjustmentPercentExit:

    'Release Objects
    Set xmlElement = Nothing
    Set xmlAttrib = Nothing
    Set xmlElementReq = Nothing
    Set xmlOperationNode = Nothing
    Set xmlRiskMatrixNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlResponseDocument = Nothing
    Set xmlResponseElement = Nothing
    Set xmlResponseNode = Nothing
        
    If Not objApplicationBO Is Nothing Then
        Set objApplicationBO = Nothing
    End If
    
    If Not objQuickQuoteApplicationBO Is Nothing Then
        Set objQuickQuoteApplicationBO = Nothing
    End If
    
    If Not objGlobalBandParameterBO Is Nothing Then
        Set objGlobalBandParameterBO = Nothing
    End If
  
    errCheckError strFunctionName, TypeName(Me)
        
End Sub

