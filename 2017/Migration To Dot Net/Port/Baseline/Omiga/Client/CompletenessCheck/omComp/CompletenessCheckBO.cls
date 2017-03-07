VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CompletenessCheckBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CompletenessCheckBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling

'Description:   Runs the completeness check rules
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AY     17/05/00    Created
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'IK     08/03/01    re-structured as phase 2 component
'------------------------------------------------------------------------------------------

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

Public Function RunCompletenessCheck(ByVal strXmlIn As String) As String

    On Error GoTo RunCompletenessCheckVbErr

    Const cstrFunctionName As String = "RunCompletenessCheck"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRunRulesDoc As FreeThreadedDOMDocument40
    Dim xmlRunRulesRequestNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlCompCheckRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode

    ' *************************************************************************
    ' validate request block
    ' *************************************************************************
    Set xmlIn = xmlLoad(strXmlIn, cstrFunctionName)
    Set xmlRequestNode = xmlGetMandatoryNode(xmlIn, "REQUEST")
    Set xmlCompCheckRequestNode = xmlGetMandatoryNode(xmlRequestNode, "COMPLETENESSCHECK")
    xmlCheckMandatoryAttribute xmlCompCheckRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlCompCheckRequestNode, "APPLICATIONFACTFINDNUMBER"
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    
    Set xmlRunRulesDoc = New FreeThreadedDOMDocument40
    xmlRunRulesDoc.validateOnParse = False
    xmlRunRulesDoc.setProperty "NewParser", True
    Set xmlRunRulesRequestNode = xmlRunRulesDoc.appendChild(xmlRequestNode.cloneNode(False))

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    
    GetApplicationData xmlCompCheckRequestNode, xmlRunRulesRequestNode
    
    If GetStageOptions(xmlRunRulesRequestNode) Then
        RunRules xmlRunRulesRequestNode, xmlResponseNode
    End If
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    RunCompletenessCheck = xmlOut.xml

RunCompletenessCheckExit:

    Debug.Print RunCompletenessCheck

    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    
    Exit Function

RunCompletenessCheckVbErr:

    RunCompletenessCheck = errCreateErrorResponse()

    '   go to clean-up section
    Resume RunCompletenessCheckExit

End Function

Private Sub GetApplicationData( _
    ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
    ByVal vxmlParentNode As IXMLDOMNode)
    
    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib
        
    ' IK_MCAP_20020524
    ' add RB_TEMPLATE attrib
    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
    xmlAttrib.Value = "CCRBTemplate"
    xmlNode.Attributes.setNamedItem xmlAttrib

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONNUMBER").Text
    
    If vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlNode.appendChild xmlElem
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    xmlRequestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetApplicationData response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError _
            "GetApplicationData", _
            oeXMLMissingElement, _
            "No APPLICATION detail"
    End If
    
    ' IK_MCAP_20020524
    ' append ALL RESPONSE child Nodes
    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
        vxmlParentNode.appendChild xmlNode.cloneNode(True)
    Next
    ' IK_MCAP_20020524 ends
    
GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Function GetStageOptions(ByVal vxmlRulesRequestNode As IXMLDOMNode) As Boolean

    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim intThisStage As Integer
    
    intThisStage = _
        xmlGetNodeAsInteger(vxmlRulesRequestNode, "APPLICATION/STAGE/@STAGESEQUENCENO")
        
    If intThisStage = CInt(GetGlobalParamAmount("TMCustRegStage")) Then
        Set xmlAttrib = _
            vxmlRulesRequestNode.ownerDocument.createAttribute("CUSTREGRULESAPPLY")
        xmlAttrib.Text = "Y"
        vxmlRulesRequestNode.Attributes.setNamedItem xmlAttrib
        GetStageOptions = True
    ElseIf intThisStage = CInt(GetGlobalParamAmount("TMQuickQuoteStage")) Then
        Set xmlAttrib = _
            vxmlRulesRequestNode.ownerDocument.createAttribute("CUSTREGRULESAPPLY")
        xmlAttrib.Text = "Y"
        vxmlRulesRequestNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = _
            vxmlRulesRequestNode.ownerDocument.createAttribute("QUICKQUOTERULESAPPLY")
        xmlAttrib.Text = "Y"
        vxmlRulesRequestNode.Attributes.setNamedItem xmlAttrib
        GetStageOptions = True
    ElseIf intThisStage >= CInt(GetGlobalParamAmount("TMAIPStage")) Then
        Set xmlAttrib = _
            vxmlRulesRequestNode.ownerDocument.createAttribute("AIPRULESAPPLY")
        xmlAttrib.Text = "Y"
        vxmlRulesRequestNode.Attributes.setNamedItem xmlAttrib
        GetStageOptions = True
    End If
    
    Set xmlAttrib = Nothing

End Function

Private Sub RunRules( _
    ByVal vxmlRulesRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo RunRulesExit
    Const cstrFunctionName As String = "RunRules"
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseNode As IXMLDOMNode
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    
    Dim objCompRules As Object
    
    Set objCompRules = gobjContext.CreateInstance("OmCompRules.CompRulesBO")
    xmlThisResponseDoc.loadXML objCompRules.RunRules(vxmlRulesRequestNode.xml)
    Set objCompRules = Nothing
    
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    For Each xmlThisResponseNode In xmlThisResponseDoc.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild _
            xmlThisResponseNode.cloneNode(False)
    Next
    
RunRulesExit:
    
    Set xmlThisResponseDoc = Nothing
    Set xmlThisResponseNode = Nothing
    
    errCheckError cstrFunctionName

End Sub
