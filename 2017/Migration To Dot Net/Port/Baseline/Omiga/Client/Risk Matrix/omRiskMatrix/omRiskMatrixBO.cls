VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omRiskMatrixBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omRiskMatrixBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Risk Matrix Business Object translates user requests for reduction
'               in loan component interest. It calls Risk Matrix Rules Object for
'               client specific implementation rules.
'               contains (implied) standard interface:
'               Public Function omRiskMatrixRequest(ByVal vstrXMLRequest As String) As String
'Dependencies:  RiskMatrixDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   CreateDate                          Package Description
'CWP    Tuesday, May 22, 2001 1:56:02 PM    04D01.T65
'CL     13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'

Option Explicit

Implements ObjectControl

Private gobjDO As IomRiskMatrixDO
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

Public Function omRiskMatrixRequest(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:
'   Risk Matrix Request, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo omRunRiskMatrixExit
    
    Const strFunctionName As String = "omRiskMatrixRequest"
    
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
    Set gobjDO = Nothing
        
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
        DoRequest xmlRequestNode, xmlResponseNode ' OPERATION attribute on REQUEST node do single operation
    Else ' multiple operations
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
    
    omRiskMatrixRequest = xmlOutDoc.xml
        
omRunRiskMatrixExit: 'release objects
        
    Set gobjDO = Nothing
    
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
        
    omRiskMatrixRequest = errCreateErrorResponse() ' create ERROR response block from Err object
        
    GoTo omRunRiskMatrixExit 'perform clean up

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
    
        Case "CALCULATEADJUSTMENTPERCENT"
            CalcAdjustmentPercent vxmlRequestNode, vxmlResponseNode
            
        Case "MANUALINTERESTRATEOVERRIDE"
            ManualInterestRateOverride vxmlRequestNode, vxmlResponseNode
                        
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CalcAdjustmentPercent(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempNode2 As IXMLDOMNode
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMElement
    Dim xmlTempDocument As New FreeThreadedDOMDocument40
    xmlTempDocument.validateOnParse = False
    xmlTempDocument.setProperty "NewParser", True
    Dim objRiskMatrixDO As IomRiskMatrixDO
    Dim objRunRiskMatrixRules As Object
    Dim strMortgageProductCode As String
    Dim strRiskMatrixRulesPercent As String
    Dim strRiskMatrixRulesResponse As String
    Dim strRiskMatrixAdjustmentPercent As String
    Dim strManualDecreaseLimitPercent As String
    Dim strCalculatedAdjustmentPercent As String
    
    On Error GoTo CalcAdjustmentPercentExit
    
    Const strFunctionName As String = "CalcAdjustmentPercent"
        
    Set objRunRiskMatrixRules = gobjContext.CreateInstance("RiskMatrixRules.omRiskMatrixRulesBO")
    
    Set xmlTempResponse = xmlTempDocument.createElement("RESPONSE")
    xmlTempResponse.setAttribute "TYPE", "SUCCESS"
                    
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "RUNRISKMATRIX"
    strRiskMatrixRulesResponse = objRunRiskMatrixRules.omRiskMatrixRulesRequest(vxmlRequestNode.xml)
    
    Set xmlTempNode = xmlLoad(strRiskMatrixRulesResponse, strFunctionName)
    errCheckXMLResponse xmlTempNode.xml, True
                
    Set xmlTempNode2 = xmlGetMandatoryNode(xmlTempNode, ".//RISKMATRIXADJUSTMENTPERCENT")
    xmlCheckMandatoryAttribute xmlTempNode2, "CALCULATEDADJUSTMENT"
    strCalculatedAdjustmentPercent = xmlGetAttributeText(xmlTempNode2, "CALCULATEDADJUSTMENT")
    
    strManualDecreaseLimitPercent = ManualDecreaseLimitPercent(vxmlRequestNode)
    
    If Val(strCalculatedAdjustmentPercent) > 0 Then
        strRiskMatrixAdjustmentPercent = strCalculatedAdjustmentPercent
    Else
        If Abs(Val(strCalculatedAdjustmentPercent)) >= Abs(Val(strManualDecreaseLimitPercent)) Then
            strRiskMatrixAdjustmentPercent = strManualDecreaseLimitPercent
        Else
            strRiskMatrixAdjustmentPercent = strCalculatedAdjustmentPercent
        End If
    End If
        
    Set xmlTempElement = xmlTempDocument.createElement("RISKMATRIXADJUSTMENTPERCENT") 'create XML result string
    xmlTempElement.setAttribute "CALCADJUSTMENTPERCENT", strRiskMatrixAdjustmentPercent 'set value to XML result string
    vxmlResponseNode.appendChild xmlTempElement 'append result to response document
                        
CalcAdjustmentPercentExit:

    Set xmlTempNode = Nothing
    Set xmlTempNode2 = Nothing
    Set xmlTempElement = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempDocument = Nothing
    
    If Not objRiskMatrixDO Is Nothing Then
        Set objRiskMatrixDO = Nothing
    End If
    
    If Not objRunRiskMatrixRules Is Nothing Then
        Set objRunRiskMatrixRules = Nothing
    End If
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ManualInterestRateOverride(vxmlRequestNode As IXMLDOMNode, vxmlResponseNode As IXMLDOMNode)
        
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMElement
    Dim xmlTempDocument As New FreeThreadedDOMDocument40
    xmlTempDocument.validateOnParse = False
    xmlTempDocument.setProperty "NewParser", True
    Dim objRiskMatrixDO As IomRiskMatrixDO
    Dim strCompetencyType As String
    Dim strUserToDecreaseValue As String
    Dim strManualAdjustmentValue As String
    Dim strMaxUserProductDecrease As String
    Dim strManualAdjustmentPercent As String
    Dim strUserDecreaseLimitPercent As String
    Dim strManualDecreaseLimitPercent As String
    Dim strRiskMatrixAdjustmentPercent As String
    Dim strManualInterestRateOverride As String
            
    On Error GoTo ManualInterestRateOverrideExit
    
    Const strFunctionName As String = "ManualInterestRateOverride"
        
    If objRiskMatrixDO Is Nothing Then
        Set objRiskMatrixDO = gobjContext.CreateInstance("omRiskMatrix.omRiskMatrixDO")
    End If

    Set xmlTempResponse = xmlTempDocument.createElement("RESPONSE")
    xmlTempResponse.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempNode = xmlGetMandatoryNode(vxmlRequestNode, "MANUALINTERESTRATEOVERRIDE")
    xmlCheckMandatoryAttribute xmlTempNode, "COMPETENCYTYPE"
    strCompetencyType = xmlGetAttributeText(xmlTempNode, "COMPETENCYTYPE") 'get and assign competency value
    
    xmlCheckMandatoryAttribute xmlTempNode, "MANUALADJUSTMENTVALUE"
    strManualAdjustmentValue = xmlGetAttributeText(xmlTempNode, "MANUALADJUSTMENTVALUE") 'get and assign manual adjustment value
    
    xmlCheckMandatoryAttribute xmlTempNode, "RISKMATRIXADJUSTMENTPERCENT"
    strRiskMatrixAdjustmentPercent = xmlGetAttributeText(xmlTempNode, "RISKMATRIXADJUSTMENTPERCENT") 'get and assign risk matrix adjustment percent
    
    Set xmlTempRequest = xmlTempDocument.createElement("COMPETENCY") 'setup request for competency
    xmlTempRequest.setAttribute "COMPETENCYTYPE", strCompetencyType 'set attribute value
    
    objRiskMatrixDO.GetUserDecreaseLimitPercent xmlTempRequest, xmlTempResponse 'get user decrease limit percent value
    errCheckXMLResponse xmlTempResponse.xml, True 'if no error in response
    
    Set xmlTempNode = xmlGetMandatoryNode(xmlTempResponse, "GETUSERDECREASELIMITPERCENT")
    xmlCheckMandatoryAttribute xmlTempNode, "USERDECREASELIMITPERCENT"
    strUserDecreaseLimitPercent = xmlGetAttributeText(xmlTempNode, "USERDECREASELIMITPERCENT") 'calculate userdecreaselimitpercent
       
    strManualDecreaseLimitPercent = ManualDecreaseLimitPercent(vxmlRequestNode) 'calculate product decrease
        
    If Abs(Val(strUserDecreaseLimitPercent)) <= Abs(Val(strManualDecreaseLimitPercent)) Then
        strMaxUserProductDecrease = strUserDecreaseLimitPercent
    Else
        strMaxUserProductDecrease = strManualDecreaseLimitPercent
    End If
    
    strUserToDecreaseValue = Val(strMaxUserProductDecrease) + Val(strRiskMatrixAdjustmentPercent)
    strManualAdjustmentPercent = -(Val(strUserToDecreaseValue))
            
    If Abs(strManualAdjustmentValue) < Abs(strManualAdjustmentPercent) Then 'return strManualAdjustmentValue
        strManualInterestRateOverride = strManualAdjustmentValue 'raise error errThrowError strFunctionName, 567
    Else 'return strManualAdjustmentPercent
        strManualInterestRateOverride = strManualAdjustmentPercent
    End If
    
    Set xmlTempElement = xmlTempDocument.createElement("MANUALINTERESTRATEOVERRIDE") 'create XML result string
    xmlTempElement.setAttribute "MANUALADJUSTMENTPERCENT", strManualInterestRateOverride 'set value to XML result string
    vxmlResponseNode.appendChild xmlTempElement 'append result to response document
             
ManualInterestRateOverrideExit:

    Set xmlTempNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempDocument = Nothing
    
    If Not objRiskMatrixDO Is Nothing Then
        Set objRiskMatrixDO = Nothing
    End If
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Function ManualDecreaseLimitPercent(ByVal vxmlRequestNode As IXMLDOMNode) As String

    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMElement
    Dim xmlTempDocument As New FreeThreadedDOMDocument40
    xmlTempDocument.validateOnParse = False
    xmlTempDocument.setProperty "NewParser", True
    Dim objRiskMatrixDO As IomRiskMatrixDO
    Dim strMortgageProductCode As String
    
    On Error GoTo ManualDecreaseLimitPercentExit
    
    Const strFunctionName As String = "ManualDecreaseLimitPercent"
        
    If objRiskMatrixDO Is Nothing Then
        Set objRiskMatrixDO = gobjContext.CreateInstance("omRiskMatrix.omRiskMatrixDO")
    End If
    
    Set xmlTempResponse = xmlTempDocument.createElement("RESPONSE")
    xmlTempResponse.setAttribute "TYPE", "SUCCESS"
        
    Set xmlTempNode = xmlGetMandatoryNode(vxmlRequestNode, "MORTGAGEPRODUCT") 'check if mandatory node exits
    xmlCheckMandatoryAttribute xmlTempNode, "MORTGAGEPRODUCTCODE" 'check if product code exits
    strMortgageProductCode = xmlGetAttributeText(xmlTempNode, "MORTGAGEPRODUCTCODE") 'get and assign product code
    
    'setup the request for mortgageproduct
    Set xmlTempRequest = xmlTempDocument.createElement("MORTGAGEPRODUCT") 'create element
    xmlTempRequest.setAttribute "MORTGAGEPRODUCTCODE", strMortgageProductCode 'with attribute value
    
    objRiskMatrixDO.GetManualDecreaseLimitPercent xmlTempRequest, xmlTempResponse 'execute request
    errCheckXMLResponse xmlTempResponse.xml, True 'if no error in response
    
    Set xmlTempNode = xmlGetMandatoryNode(xmlTempResponse, "GETMORTGAGEPRODUCTDETAILS") 'check response with mandatory node
    xmlCheckMandatoryAttribute xmlTempNode, "MANUALDECREASELIMITPERCENT" 'check mandatory attibute
    
    ManualDecreaseLimitPercent = xmlGetAttributeText(xmlTempNode, "MANUALDECREASELIMITPERCENT") 'get and assign function with percent

ManualDecreaseLimitPercentExit:
    
    Set xmlTempNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempDocument = Nothing
    
    If Not objRiskMatrixDO Is Nothing Then
        Set objRiskMatrixDO = Nothing
    End If
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        strFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    errCheckError strFunctionName, TypeName(Me)

End Function


