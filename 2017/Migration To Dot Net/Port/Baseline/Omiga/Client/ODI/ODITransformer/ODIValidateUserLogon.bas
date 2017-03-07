Attribute VB_Name = "ODIValidateUserLogon"
'Workfile:      ODIValidateUserLogon.bas
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
'RF     24/08/01    Created.
'RF     12/09/01    Allow for ODIConverter interface change.
'RF     13/11/01    Host and Environment optional in ValidateUserLogon request.
'DS     30/04/02    Use FreeThreadedDOMDocument40.
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub ValidateUserLogon( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Logs a user on to the AS/400 and creates a Session.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    On Error GoTo ValidateUserLogonExit
    
    Const strFunctionName As String = "ValidateUserLogon"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim nodeConverterResponse As IXMLDOMNode
    Dim nodeUser As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim nodeSignOnProfile As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    'Dim strPort As String
    Dim strEnv As String
    
    '------------------------------------------------------------------------------------------
    ' Call PlexusServerGateway_startNewSession to return a new Session
    '------------------------------------------------------------------------------------------
    
    Set nodeUser = xmlGetMandatoryNode(vxmlRequestNode, ".//USER")
    Set nodeSignOnProfile = xmlDoc.createElement("SIGNONPROFILE")
    
    Set nodeTemp = xmlDoc.createElement("HOST")
    xmlSetAttributeValue nodeTemp, "DATA", xmlGetAttributeText(nodeUser, "HOST")
    nodeSignOnProfile.appendChild nodeTemp
    
    Set nodeTemp = xmlDoc.createElement("ENVIRONMENT")
    xmlSetAttributeValue nodeTemp, "DATA", xmlGetAttributeText(nodeUser, "ENVIRONMENT")
    nodeSignOnProfile.appendChild nodeTemp
    
    Set nodeTemp = xmlDoc.createElement("USERNAME")
    xmlSetAttributeValue nodeTemp, "DATA", _
        xmlGetMandatoryAttributeText(nodeUser, "USERNAME")
    nodeSignOnProfile.appendChild nodeTemp
    
    Set nodeTemp = xmlDoc.createElement("PASSWORD")
    xmlSetAttributeValue nodeTemp, "DATA", _
        xmlGetMandatoryAttributeText(nodeUser, "PASSWORDVALUE")
    nodeSignOnProfile.appendChild nodeTemp
        
'    ' set port value
'    strPort = xmlGetMandatoryNodeText(vxmlRequestNode, _
'        ".//ODIINITIALISATION/@PORT")
'    objODITransformerState.SetPort strPort
    
    ' set ODIENVIRONMENT
    strEnv = xmlGetMandatoryNodeText(vxmlRequestNode, _
        ".//ODIINITIALISATION/@ODIENVIRONMENT")
    objODITransformerState.SetODIEnvironment strEnv
    
    Set nodeConverterResponse = PlexusServerGateway_startNewSession( _
        objODITransformerState, nodeSignOnProfile)
    
    CheckConverterResponse nodeConverterResponse, True
    
    '------------------------------------------------------------------------------------------
    ' save the state info
    '------------------------------------------------------------------------------------------
    
    objODITransformerState.Save nodeConverterResponse
    
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode
            
ValidateUserLogonExit:

    Set nodeConverterResponse = Nothing
    Set nodeUser = Nothing
    Set xmlDoc = Nothing
    Set nodeSignOnProfile = Nothing
    Set nodeTemp = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub




