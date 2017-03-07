VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ConditionsTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------------
'BMIDS Specific History:

'Prog   Date        AQR         Description
'MV     14/08/2002  BMIDS00333  Modified GetApplicationData() - Core Ref AQR: SYS4919
'GD     03/10/2002  BMIDS00313  APWP2 - Added RunResolveDataRules, Changed RunConditionRules, Changed I..CreateDefaultApplicationConditions
'GD     16/10/2002  BMIDS00650  Enhancements to BMIDS00313 in CreateDefaultApplicationConditions()
'AW     29/11/02    BMIDS01113  Use CONDITIONNAME not truncated CONDITIONDESCRIPTION
'HA     08/05/03    BM0518      Default Satisfy Status to Yes on Conditions.
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
'MARS Specific History:

'Prog   Date        AQR     Description
'MV     14/09/2005  MAR35   Amended IConditionsTxBO_CreateDefaultApplicationConditions
'BC     06/04/2006  MAR1425 Don't add Conditions automatically, if that Condition already exists
'----------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IConditionsTxBO

Private gobjContext As ObjectContext
Private Const cstrApplicationConditions = "APPLICATIONCONDITIONS"



Private Sub IConditionsTxBO_CreateApplicationConditions( _
                ByVal vxmlRequestParentNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateVbErr

    Const cstrFunctionName As String = "IConditionsTxBO_CreateApplicationConditions"
    
    Dim objDO As IConditionsDO
    Dim xmlNodeList As IXMLDOMNodeList
   
    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")

    If xmlNodeList.length > 0 Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
        objDO.CreateApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    End If
    
    gobjContext.SetComplete

CreateVbExit:
    Set objDO = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateVbErr:
    gobjContext.SetAbort
    GoTo CreateVbExit

End Sub


Private Sub IConditionsTxBO_CreateDefaultApplicationConditions( _
                                                        ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                                        ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateVbErr

    Const cstrFunctionName As String = "IConditionsTxBO_CreateDefaultApplicationConditions"
    
    Dim objDO As IConditionsDO
    Dim objBO As IConditionsBO
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlTableNode As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlCondition As IXMLDOMElement

    Dim xmlConditionsDataNode As IXMLDOMNode
    Dim xmlRuleResponseNode As IXMLDOMNode
    Dim xmlResponseList As IXMLDOMNodeList
    Dim xmlApplicationConditionsElem As IXMLDOMElement
    Dim xmlSchemaDataElem As IXMLDOMElement
    Dim xmlTableNodeRequest As IXMLDOMElement
    Dim xmlRulesRequestElem As IXMLDOMElement
    Dim xmlResolveDataRuleRequest As IXMLDOMElement
    Dim xmlResolveDataRuleRequestCondList As IXMLDOMElement
    Dim xmlResolveDataRuleResponse As IXMLDOMElement
    
    Dim strChannelId As String
    Dim strApplicationNumber As String
    Dim strUserId As String
    Dim strUnitId As String
    Dim strConditionRef As String
    Dim strApplicationFFNumber As String
    Dim strConditionDescription As String 'BC MAR1425
    
    Set objBO = gobjContext.CreateInstance(App.Title & ".ConditionsBO")
    Set objDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlTableNodeRequest = xmlDoc.createElement("REQUEST")
    Set xmlTableNode = xmlDoc.createElement("CONDITIONS")
    xmlSetAttributeValue xmlTableNode, "DELETEFLAG", 0
    
    xmlDoc.appendChild xmlTableNode
        
    strChannelId = xmlGetAttributeText(vxmlRequestParentNode, "CHANNELID")
    strUserId = xmlGetAttributeText(vxmlRequestParentNode, "USERID")
    strUnitId = xmlGetAttributeText(vxmlRequestParentNode, "UNITID")
    strApplicationNumber = xmlGetAttributeText(vxmlRequestParentNode.selectSingleNode(".//APPLICATIONCONDITIONS"), "APPLICATIONNUMBER")
    strApplicationFFNumber = xmlGetAttributeText(vxmlRequestParentNode.selectSingleNode(".//APPLICATIONCONDITIONS"), "APPLICATIONFACTFINDNUMBER")
    
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlTableNode, "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
    
    Set xmlTempResponseNode = xmlDoc.createElement("RESPONSE")

'Step:

    'Find Application Product Conditions List
    xmlTableNodeRequest.appendChild xmlTableNode
    
    objBO.FindAppProductConditionsList xmlTableNodeRequest, xmlTempResponseNode
    
'Step:

    'Call GetApplicationData to return xmlSchema data - ONLY MAKE THE CALL ONCE - Performance implications
    Set xmlApplicationConditionsElem = vxmlRequestParentNode.selectSingleNode("APPLICATIONCONDITIONS")
    Set xmlSchemaDataElem = xmlDoc.createElement("RESPONSE")
    
    objBO.GetApplicationData xmlApplicationConditionsElem, xmlSchemaDataElem

    Set xmlSchemaDataElem = xmlSchemaDataElem.selectSingleNode(".//APPLICATION")
    xmlChangeNodeName xmlSchemaDataElem, "APPLICATION", "SCHEMADATA"
    
'Step:
    
    'Run Condition Rules
    Set xmlRuleResponseNode = xmlDoc.createElement("RESPONSE")
    Set xmlRulesRequestElem = vxmlRequestParentNode.selectSingleNode("//REQUEST")
    xmlRulesRequestElem.appendChild xmlSchemaDataElem
    
    RunConditionRules vxmlRequestParentNode, xmlRuleResponseNode 'Method to change no call to GetApplicationData

    'If response ok then
    Set xmlResponseList = xmlRuleResponseNode.selectNodes("CONDITIONS")
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    
    Set xmlResponseList = xmlRuleResponseNode.selectNodes("CONDITIONSRULES")
    If xmlResponseList.length > 0 Then
    
        'Build the request and create application conditions
        Set xmlTableNode = vxmlRequestParentNode.selectSingleNode("APPLICATIONCONDITIONS")
        strApplicationNumber = xmlGetAttributeText(xmlTableNode, "APPLICATIONNUMBER")
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestParentNode)
        
        For Each xmlCondition In xmlResponseList
        
            strConditionRef = xmlGetAttributeText(xmlCondition, "CONTEXT")
            
            'Get Conditions details from the FindConditionsList response node
            Set xmlConditionsDataNode = xmlTempResponseNode.selectSingleNode(".//CONDITIONS[@CONDITIONREFERENCE='" & strConditionRef & "']")
            If Not xmlConditionsDataNode Is Nothing Then
                
                'Create the application conditions with the found conditions data
                 Set xmlTableNode = xmlDoc.createElement("APPLICATIONCONDITIONS")
                 xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
                 xmlSetAttributeValue xmlTableNode, "CONDITIONREFERENCE", _
                                     xmlGetAttributeText(xmlConditionsDataNode, "CONDITIONREFERENCE")
                 xmlSetAttributeValue xmlTableNode, "CONDITIONDESCRIPTION", _
                                 xmlGetAttributeText(xmlConditionsDataNode, "CONDITIONDESCRIPTION")
                 xmlSetAttributeValue xmlTableNode, "CONDITIONNAME", _
                                 xmlGetAttributeText(xmlConditionsDataNode, "CONDITIONNAME")
                 xmlSetAttributeValue xmlTableNode, "CONDITIONTYPE", _
                                 xmlGetAttributeText(xmlConditionsDataNode, "CONDITIONTYPE")
                 xmlSetAttributeValue xmlTableNode, "EDIT", _
                                 xmlGetAttributeText(xmlConditionsDataNode, "EDITABLEIND")
                 xmlSetAttributeValue xmlTableNode, "FREEFORMAT", _
                                 xmlGetAttributeText(xmlConditionsDataNode, "FREEFORMATIND")

                 xmlSetAttributeValue xmlTableNode, "SATISFYSTATUS", 1
                 xmlSetAttributeValue xmlTableNode, "USERID", strUserId
                 xmlSetAttributeValue xmlTableNode, "UNITID", strUnitId
                 
                 xmlNewRequest.appendChild xmlTableNode
            
            End If
        Next

'Step:

        'Run RunResolveDataRules
        Dim xmlAppCondList As IXMLDOMNodeList
        Dim xmlNewRequestCloneAppCondList As IXMLDOMNodeList
        Dim xmlNewRequestCloneElem As IXMLDOMElement
        Dim xmlAppCondElem As IXMLDOMElement
        Dim xmlCondElem As IXMLDOMElement
        Dim xmlApplicationCond As IXMLDOMNode 'BC MAR1425 05/04/2006
        Dim xmlCondList As IXMLDOMNodeList
        Dim iIndex As Integer
        Dim xmlNewRequestClone As IXMLDOMElement
        
        Set xmlNewRequestClone = xmlNewRequest.cloneNode(True)
        Set xmlAppCondList = xmlNewRequest.selectNodes(".//APPLICATIONCONDITIONS")
        If xmlAppCondList.length > 0 Then
        
            Set xmlResolveDataRuleRequestCondList = xmlDoc.createElement("CONDITIONSLIST")
            
            Set xmlResolveDataRuleRequest = xmlDoc.createElement("REQUEST")
            xmlResolveDataRuleRequest.appendChild xmlResolveDataRuleRequestCondList
            xmlResolveDataRuleRequest.appendChild xmlSchemaDataElem
            
            Set xmlResolveDataRuleResponse = xmlDoc.createElement("RESPONSE")
            
            For iIndex = 0 To (xmlAppCondList.length - 1)
                Set xmlAppCondElem = xmlAppCondList.Item(iIndex)
                xmlChangeNodeName xmlAppCondElem, "APPLICATIONCONDITIONS", "CONDITIONS"
                xmlResolveDataRuleRequestCondList.appendChild xmlAppCondElem
            Next iIndex
            
            RunResolveDataRules xmlResolveDataRuleRequest, xmlResolveDataRuleResponse
            
            Set xmlCondList = xmlResolveDataRuleResponse.selectNodes(".//CONDITIONS")
'Step:
            
            Dim xmlModifiedConditionsList As IXMLDOMNodeList
            Dim strPattern  As String
            
            'Set xmlModifiedConditionsList = xmlResolveDataRuleResponse.selectNodes("//CONDITIONS[@USERMODIFIED='1']")
            Set xmlModifiedConditionsList = xmlResolveDataRuleResponse.selectNodes("//CONDITIONS")
         
'Step:
            If xmlModifiedConditionsList.length <= 0 Then
                
                Set xmlNewRequestCloneAppCondList = xmlNewRequestClone.selectNodes(".//APPLICATIONCONDITIONS")
            
                For iIndex = 0 To (xmlCondList.length - 1)
                    Set xmlCondElem = xmlCondList.Item(iIndex) 'Resolved
                    Set xmlNewRequestCloneElem = xmlNewRequestCloneAppCondList.Item(iIndex) 'Unresolved
                    xmlSetAttributeValue xmlNewRequestCloneElem, "CONDITIONNAME", xmlGetAttributeText(xmlCondElem, "CONDITIONNAME")
                    xmlSetAttributeValue xmlNewRequestCloneElem, "CONDITIONDESCRIPTION", xmlGetAttributeText(xmlCondElem, "CONDITIONDESCRIPTION")
                Next iIndex
                
            Else
                
                For iIndex = 0 To (xmlModifiedConditionsList.length - 1)
                    Set xmlCondElem = xmlModifiedConditionsList.Item(iIndex)
                    
                    strConditionRef = xmlGetAttributeText(xmlCondElem, "CONDITIONREFERENCE")
                    strPattern = "//APPLICATIONCONDITIONS[@CONDITIONREFERENCE='" & strConditionRef & "']"
                    strConditionDescription = xmlGetAttributeText(xmlCondElem, "CONDITIONDESCRIPTION") 'BC MAR1425
                    
                    Set xmlCondElem = xmlNewRequestClone.selectSingleNode(strPattern)
                    
                    If Not xmlCondElem Is Nothing Then
                        xmlSetAttributeValue xmlCondElem, "CONDITIONDESCRIPTION", strConditionDescription 'BC MAR1425
                        Set xmlCondElem = Nothing
                    End If
                
                Next iIndex
                
            End If
            
            'BC MAR1425 05/04/2006 Begin
            'Remove Condition from xmlNewRequestClone if already in xmlSchemaDataElem i.e if already exists on Application
            
            Set xmlNewRequestCloneAppCondList = xmlNewRequestClone.selectNodes(".//APPLICATIONCONDITIONS")
            
            For iIndex = 0 To (xmlNewRequestCloneAppCondList.length - 1)
                    
                    Set xmlCondElem = xmlNewRequestCloneAppCondList.Item(iIndex)
                    
                    strConditionRef = xmlGetAttributeText(xmlCondElem, "CONDITIONREFERENCE")
                    strPattern = "//APPLICATIONCONDITIONS[@CONDITIONREFERENCE='" & strConditionRef & "']"
                    
                    Set xmlApplicationCond = xmlSchemaDataElem.selectSingleNode(strPattern)
                    
                    'BC MAR1425 if ConditionReference in both xmlModifiedConditionsList xmlSchemaDataElem
                    'BC MAR1425 Remove from xmlModifiedConditionsList
                    
                    If Not xmlApplicationCond Is Nothing Then
                        xmlNewRequestClone.removeChild xmlCondElem
                        Set xmlCondElem = Nothing
                    End If

            Next iIndex
            'BC MAR1425 05/04/2006 End
            
            objDO.CreateApplicationConditions xmlNewRequestClone, vxmlResponseNode
            
        End If

    End If
    
    gobjContext.SetComplete

CreateVbExit:
    Set objDO = Nothing
    Set xmlDoc = Nothing
    Set xmlTableNode = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlRuleResponseNode = Nothing
    Set xmlConditionsDataNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCondition = Nothing
    Set xmlResponseList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateVbErr:
    gobjContext.SetAbort
    GoTo CreateVbExit

End Sub


Private Sub IConditionsTxBO_DeleteAllApplicationConditions( _
        ByVal vxmlRequestParentNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteVbErr

    Const cstrFunctionName As String = "IConditionsTxBO_DeleteAllApplicationConditions"
    
    Dim objDO As IConditionsDO
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    
    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")

    If xmlNodeList.length > 0 Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
        For Each xmlNode In xmlNodeList
            objDO.DeleteAllApplicationConditions xmlNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete

DeleteVbExit:
    Set objDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

DeleteVbErr:
    gobjContext.SetAbort
    GoTo DeleteVbExit

End Sub

Private Sub IConditionsTxBO_DeleteApplicationConditions( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteVbErr

    Const cstrFunctionName As String = "IConditionsTxBO_DeleteApplicationConditions"
    
    Dim objDO As IConditionsDO
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    
    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")

    If xmlNodeList.length > 0 Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
        For Each xmlNode In xmlNodeList
            objDO.DeleteApplicationConditions xmlNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete

DeleteVbExit:
    Set objDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

DeleteVbErr:
    gobjContext.SetAbort
    GoTo DeleteVbExit
End Sub

Private Sub IConditionsTxBO_UpdateApplicationConditions( _
            ByVal vxmlRequestParentNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateVbErr

    Const cstrFunctionName As String = "IConditionsTxBO_UpdateApplicationConditions"
    
    Dim objDO As IConditionsDO
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    
    Set xmlNodeList = vxmlRequestParentNode.selectNodes("APPLICATIONCONDITIONS")

    If xmlNodeList.length > 0 Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".ConditionsDO")
        objDO.UpdateApplicationConditions vxmlRequestParentNode, vxmlResponseNode
    End If
    
    gobjContext.SetComplete

UpdateVbExit:
    Set objDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

UpdateVbErr:
    gobjContext.SetAbort
    GoTo UpdateVbExit
End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'Private Sub GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
'                               ByVal vxmlParentNode As IXMLDOMNode)
'
'    On Error GoTo GetApplicationDataExit
'    Const cstrFunctionName As String = "GetApplicationData"
'
'    ' <VSA> Visual Studio Analyser Support
'    #If USING_VSA Then
'        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & cstrFunctionName)
'    #End If
'
'    Dim xmlRequestDoc As FreeThreadedDOMDocument40
'
'    Dim xmlElem As IXMLDOMElement
'    Dim xmlNode As IXMLDOMNode
'    Dim xmlAttrib As IXMLDOMAttribute
'
'    'GD BMIDS00650
'    Dim xmlComboElement As IXMLDOMElement
'    Dim xmlComboAttribute As IXMLDOMAttribute
'
'    Dim objRB As Object
'
'    Set xmlRequestDoc = New FreeThreadedDOMDocument40
'    xmlRequestDoc.async = False
'
'    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
'    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
'    'GD BMIDS00650 - change combolookup to YES
'    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
'    xmlAttrib.Text = "YES"
'    xmlNode.Attributes.setNamedItem xmlAttrib
'
'    'GD BMIDS00650
'    'Add SpecialSchemes combo to request
'    Set xmlComboElement = xmlRequestDoc.createElement("COMBO")
'
'    Set xmlComboAttribute = xmlRequestDoc.createAttribute("_SCHEMA_")
'    xmlComboAttribute.Text = "COMBO"
'    xmlComboElement.Attributes.setNamedItem xmlComboAttribute
'
'    Set xmlComboAttribute = xmlRequestDoc.createAttribute("GROUPNAME")
'    xmlComboAttribute.Text = "SpecialSchemes"
'    xmlComboElement.Attributes.setNamedItem xmlComboAttribute
'
'    'Now xmlComboElement = <COMBO _SCHEMA_="COMBO" GROUPNAME="SpecialSchemes"/>
'
'    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
'    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
'    xmlAttrib.Text = "APCondRBTemplate"
'    xmlNode.Attributes.setNamedItem xmlAttrib
'
'    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
'
'    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
'
'    xmlElem.setAttribute _
'        "APPLICATIONNUMBER", _
'        vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONNUMBER").Text
'
'    If vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER") Is Nothing Then
'        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
'    Else
'        xmlElem.setAttribute _
'            "APPLICATIONFACTFINDNUMBER", _
'            vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER").Text
'    End If
'
'    xmlNode.appendChild xmlElem
'    xmlNode.appendChild xmlComboElement
'
'    Debug.Print "GetApplicationData request"
'    Debug.Print xmlRequestDoc.xml
'    Dim strTemp As String
'
'    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
'    strTemp = objRB.OmDataRequest(xmlRequestDoc.xml)
'    xmlRequestDoc.loadXML strTemp
'    Set objRB = Nothing
'
'    Debug.Print "GetApplicationData response"
'    Debug.Print xmlRequestDoc.xml
'
'    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
'
'    If xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
'        errThrowError _
'            "GetApplicationData", _
'            oeXMLMissingElement, _
'            "No APPLICATION detail"
'    End If
'    ' START GD BMIDS00650 Add all COMBO nodes underneath the APPLICATION NODE
'    Dim xmlComboNodeList As IXMLDOMNodeList
'    Dim iComboCount As Integer
'    Dim iIndex As Integer
'
'    Set xmlComboNodeList = xmlRequestDoc.selectNodes(".//RESPONSE/COMBO")
'    iComboCount = xmlComboNodeList.length
'
'    Set xmlElem = xmlRequestDoc.selectSingleNode(".//APPLICATION")
'    For iIndex = 0 To (iComboCount - 1)
'        Set xmlComboElement = xmlComboNodeList.Item(iIndex)
'        xmlElem.appendChild xmlComboElement
'    Next
'    ' END
'    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
'    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
'        vxmlParentNode.appendChild xmlNode.cloneNode(True)
'    Next
'
'GetApplicationDataExit:
'
'    Set xmlElem = Nothing
'    Set xmlNode = Nothing
'    Set xmlAttrib = Nothing
'    Set xmlRequestDoc = Nothing
'
'    ' <VSA> Visual Studio Analyser Support
'    #If USING_VSA Then
'        Set VSA = Nothing
'    #End If
'
'    errCheckError cstrFunctionName
'
'End Sub

Private Sub RunConditionRules( _
        ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlRuleResponseNode As IXMLDOMNode)
'GD BMIDS00313 - removed call to GetApplicationData

    On Error GoTo RunConditionRulesExit
    Const strFunctionName As String = "RunConditionRules"

    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40, _
        xmlThisResponseDoc As FreeThreadedDOMDocument40, _
        xmlThisRequestNode As IXMLDOMNode, _
        xmlApplicationConditionsNode As IXMLDOMNode, _
        xmlResponseList As IXMLDOMNodeList, _
        xmlThisNode As IXMLDOMNode, _
        xmlThisElem As IXMLDOMElement

    Dim objAPRules As Object, _
        gobjAppBOProc As Object
            
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    
    xmlChangeNodeName vxmlRequestNode, "SCHEMADATA", "APPLICATION"
    Set xmlThisRequestNode = vxmlRequestNode.selectSingleNode(".//APPLICATION")
    'Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(True))
    
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "RunConditionRules"
    
    Set xmlApplicationConditionsNode = vxmlRequestNode.selectSingleNode("APPLICATIONCONDITIONS")
    

    'BMIDS00313
    'GetApplicationData xmlApplicationConditionsNode, xmlThisRequestNode
    'xmlThisRequestDoc.appendChild xmlThisRequestNode
    Set objAPRules = gobjContext.CreateInstance("omAPRules.APRulesBO")
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    Dim strTemp As String
    strTemp = objAPRules.RunRules(xmlThisRequestDoc.xml)
    xmlThisResponseDoc.loadXML strTemp
    Set objAPRules = Nothing
      
    Set xmlResponseList = xmlThisResponseDoc.selectNodes("RESPONSE/CONDITIONSRULES")
    If xmlResponseList.length > 0 Then
        For Each xmlThisNode In xmlResponseList
            vxmlRuleResponseNode.appendChild xmlThisNode
        Next
    End If

RunConditionRulesExit:

    Set objAPRules = Nothing
    Set gobjAppBOProc = Nothing
    
    Set xmlResponseList = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationConditionsNode = Nothing
    Set xmlThisNode = Nothing
    Set xmlThisElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlResponseList = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Public Sub RunResolveDataRules( _
        ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlRuleResponseNode As IXMLDOMNode)
'GD BMIDS00313

    'On Error GoTo RunResolveDataRulesExit
    Const strFunctionName As String = "RunResolveDataRules"

    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40, _
        xmlThisResponseDoc As FreeThreadedDOMDocument40, _
        xmlThisRequestNode As IXMLDOMNode, _
        xmlApplicationConditionsNode As IXMLDOMNode, _
        xmlResponseList As IXMLDOMNodeList, _
        xmlThisNode As IXMLDOMNode, _
        xmlThisElem As IXMLDOMElement

    Dim objAPRules As Object, _
        gobjAppBOProc As Object
            
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlRequestOnly As IXMLDOMElement
    Dim xmlConditionsListElem As IXMLDOMElement
    Dim xmlSchemaDataElem As IXMLDOMElement
    Dim strTemp As String
    Dim strConditionsErrorFilePath As String
    Dim xmlConditionsLogFileElem As IXMLDOMElement
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
   
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    
    'xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "RunResolveDataRules"
    If Not (vxmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing) Then
        vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "RunResolveDataRules"
    Else
        'xmlThisRequestNode.attributes.setNamedItem(
        
        xmlSetAttributeValue vxmlRequestNode, "OPERATION", "RunResolveDataRules"
    End If
    'Set xmlRequestOnly = vxmlRequestNode.selectSingleNode("//REQUEST")
    
    Set xmlConditionsListElem = vxmlRequestNode.selectSingleNode(".//CONDITIONSLIST")
    Set xmlSchemaDataElem = vxmlRequestNode.selectSingleNode(".//SCHEMADATA")
    Set xmlApplicationElem = xmlThisRequestDoc.createElement("APPLICATION")
    Set xmlConditionsLogFileElem = xmlThisRequestDoc.createElement("CONDITIONRULESLOGFILE")
    strConditionsErrorFilePath = GetGlobalParamString("OfferConditionResolveFailure")
   
    xmlConditionsLogFileElem.setAttribute "FILENAME", strConditionsErrorFilePath
    
    vxmlRequestNode.appendChild xmlApplicationElem
    xmlApplicationElem.appendChild xmlConditionsListElem
    xmlApplicationElem.appendChild xmlSchemaDataElem
    xmlApplicationElem.appendChild xmlConditionsLogFileElem
    
    'Set objAPRules = gobjContext.CreateInstance("omAPRules.APRulesBO")
    Set objAPRules = CreateObject("omAPRules.APRulesBO")
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False


    strTemp = objAPRules.RunRules(vxmlRequestNode.xml)
    'Check the response that is returned from the Rules component
    'errCheckXMLResponse strTemp, True, vxmlRuleResponseNode
    
    xmlThisResponseDoc.loadXML strTemp
    errCheckXMLResponse strTemp, True, vxmlRuleResponseNode
    'Dim lngTemp As Long
    'lngTemp = errCheckXMLResponse(strTemp, False)
    'errCheckXMLResponse
    'errCheckError "RunResosolveDataRules", "omAppProc.ConditionsTxBO", xmlThisResponseDoc
    'errThrowError "Run Resolvedata rules", 502, "NO RULE"
    Set objAPRules = Nothing
      
    Set xmlResponseList = xmlThisResponseDoc.selectNodes(".//CONDITIONS")
    If xmlResponseList.length > 0 Then
        For Each xmlThisNode In xmlResponseList
            vxmlRuleResponseNode.appendChild xmlThisNode
        Next
    End If

RunResolveDataRulesExit:

    Set objAPRules = Nothing
    Set gobjAppBOProc = Nothing
    
    Set xmlResponseList = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationConditionsNode = Nothing
    Set xmlThisNode = Nothing
    Set xmlThisElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlResponseList = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
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
    
    vxmlParentNode.appendChild _
        xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION").cloneNode(True)
    
GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub
