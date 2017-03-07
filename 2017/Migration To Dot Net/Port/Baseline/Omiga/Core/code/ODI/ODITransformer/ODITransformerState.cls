VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ODITransformerState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ODITransformerState.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Holds state information.
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     23/08/01    Expanded class created by LD.
'RF     12/09/01    Allow for ODIConverter interface change.
'RF     17/09/01    Fix to GetODIEnvironment.
'AS     18/09/01    Renamed SESSIONID to ID to match ObjectMapOSG.xml used by ODIConverter.
'DS     30/04/02    Use FreeThreadedDOMDocument.
'DS     11/07/02    Raise an error if there is no AdminSystemState - unless this is a login.
'------------------------------------------------------------------------------------------
Option Explicit

'   State Format:
'   <ODISTATE>
'       <STATEINFO>
'           <SESSION>
'               <SESSIONIMPL>
'                   <ACCOUNTINGDATE DATA=""/>
'                   <EFFECTIVEDATE DATA=""/>
'                   <ID DATA=""/>
'                   <USERKEY DATA=""/>
'                   <OBJECT>
'                       <PLEXUSLOCALE>
'                           <COUNTRY DATA=""/>
'                           <LANGUAGE DATA=""/>
'                       </PLEXUSLOCALE>
'                   </OBJECT>
'                   <PORT DATA=""/>
'               </SESSIONIMPL>
'           </SESSION>
'           <ODIENVIRONMENT DATA=""/>
'       </STATEINFO>
'   </ODISTATE>
'Dim m_nodeState As IXMLDOMNode
Private m_nodeState As IXMLDOMNode

Public Sub Initialise(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Get the state info from a client's xml request.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo InitialiseExit

    Const strFunctionName = "Initialise"
    
    Dim nodeRequest As IXMLDOMNode
    Dim nodeStateInfo As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument
    Dim elemODIState As IXMLDOMElement
    Dim strAdminSystemStateXml As String
    Dim strOp As String
        
    ' get the value of the ADMINSYSTEMSTATE attribute, if present
    Set nodeRequest = xmlGetMandatoryNode(vxmlRequestNode, ".//REQUEST")
    strAdminSystemStateXml = xmlGetAttributeText(nodeRequest, "ADMINSYSTEMSTATE")
    strOp = UCase(xmlGetAttributeText(nodeRequest, "OPERATION"))
    
    If strAdminSystemStateXml <> "" Then
        
        ' load the ADMINSYSTEMSTATE attribute value into a DOM
        Set xmlDoc = xmlLoad(strAdminSystemStateXml, strFunctionName)
        
'        Set nodeStateInfo = xmlGetMandatoryNode(xmlDoc.documentElement, ".//SESSION")
        Set nodeStateInfo = xmlGetMandatoryNode( _
            xmlDoc.documentElement, "//ADMINSYSTEMSTATE/STATEINFO")
    
        Set elemODIState = xmlDoc.createElement("ODISTATE")
        elemODIState.appendChild nodeStateInfo.cloneNode(True)
        Set m_nodeState = elemODIState
    
    Else
        
        'SYS4795 If we are logging in, allow a null session, otherwise throw an error
        If strOp = "VALIDATEUSERLOGON" Then

            Dim strSessionXml As String
    
            strSessionXml = _
                "<ODISTATE>" & _
                "<STATEINFO>" & _
                "<SESSION>" & _
                "<SESSIONIMPL>" & _
                "<ACCOUNTINGDATE DATA=""" & _
                """/><EFFECTIVEDATE DATA=""" & _
                """/><ID DATA=""" & _
                """/><USERKEY DATA=""" & _
                """/></SESSIONIMPL>" & _
                "</SESSION>" & _
                "</STATEINFO>" & _
                "</ODISTATE>"
    
            Set m_nodeState = xmlLoad(strSessionXml, strFunctionName).documentElement
        
            ' initialise port
            SetPort "0"
        Else
            errThrowError "ODITransformer.Initialise", 4501, "Please login to the admin system"
        End If
    End If

InitialiseExit:

    Set nodeStateInfo = Nothing
    Set xmlDoc = Nothing
    Set elemODIState = Nothing
    
    errCheckError strFunctionName

End Sub

Public Sub Save(ByVal vxmlConverterResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Get the session info from a client's xml request.
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo SaveExit

    Const strFunctionName = "Save"
    
    Dim nodeNewSessionImpl As IXMLDOMNode
    Dim nodeMsgSession As IXMLDOMNode
    
    Set nodeMsgSession = xmlGetMandatoryNode(m_nodeState, ".//SESSION")
    
    Set nodeNewSessionImpl = xmlGetMandatoryNode( _
        vxmlConverterResponseNode, _
        ".//ARGUMENTS/OBJECT/SESSIONIMPL")
        
    nodeMsgSession.replaceChild nodeNewSessionImpl, nodeMsgSession.firstChild()
        
SaveExit:

    Set nodeNewSessionImpl = Nothing
    Set nodeMsgSession = Nothing
    
    errCheckError strFunctionName

End Sub
            
Public Sub AddToResponseNode(ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Add state XML to response node
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo AddToResponseNodeExit

    Const strFunctionName = "AddToResponseNode"
    
    Dim xmlAdminSystemStateNode As IXMLDOMNode
    Set xmlAdminSystemStateNode = _
        vxmlResponseNode.ownerDocument.createElement("ADMINSYSTEMSTATE")
    
    xmlAdminSystemStateNode.appendChild GetStateInfo()
    
    vxmlResponseNode.appendChild xmlAdminSystemStateNode

AddToResponseNodeExit:

    Set xmlAdminSystemStateNode = Nothing
    
    errCheckError strFunctionName

End Sub

Public Function GetSession() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo GetSessionExit

    Const strFunctionName = "GetSession"

    Dim nodeSession As IXMLDOMNode

'    Set GetSession = _
'        m_nodeState.selectSingleNode(".//SESSION/SESSIONIMPL").cloneNode(True)
    
    Set nodeSession = m_nodeState.selectSingleNode(".//SESSION")
    
    If nodeSession Is Nothing Then
        ' "Invalid Administration System state information"
        errThrowError strFunctionName, 4501, "SESSION not found"
    End If
    
    Set GetSession = nodeSession.cloneNode(True)

GetSessionExit:

    Set nodeSession = Nothing
    
    errCheckError strFunctionName

End Function


Public Function GetODIEnvironment() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo GetODIEnvironmentExit

    Const strFunctionName = "GetODIEnvironment"
    
'    GetODIEnvironment = xmlGetMandatoryNodeText( _
'        m_nodeState, "//ODISTATE/STATEINFO/ODIENVIRONMENT/@DATA")
    GetODIEnvironment = xmlGetMandatoryNodeText( _
        m_nodeState, ".//STATEINFO/ODIENVIRONMENT/@DATA")

GetODIEnvironmentExit:

    errCheckError strFunctionName

End Function

'Public Function GetPort() As Long
Public Function GetPort() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo GetPortExit

    Const strFunctionName = "GetPort"
    
'    GetPort = CSafeLng( _
'        xmlGetMandatoryNodeText(m_nodeState, ".//SESSION/SESSIONIMPL/PORT/@DATA"))
    GetPort = xmlGetMandatoryNodeText( _
        m_nodeState, ".//SESSION/SESSIONIMPL/PORT/@DATA")

GetPortExit:

    errCheckError strFunctionName

End Function

Private Function SetPort(ByVal vstrNewValue)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo SetPortExit

    Const strFunctionName = "SetPort"
    Dim nodeSession As IXMLDOMNode
    Dim nodePort As IXMLDOMNode

    Set nodePort = xmlGetNode(m_nodeState, ".//SESSION/SESSIONIMPL/PORT")

    If nodePort Is Nothing Then
        Set nodeSession = xmlGetMandatoryNode( _
            m_nodeState, "..//SESSION/SESSIONIMPL")
        Set nodePort = nodeSession.appendChild( _
            m_nodeState.ownerDocument.createElement("PORT"))
    End If

    xmlSetAttributeValue nodePort, "DATA", vstrNewValue

SetPortExit:

    errCheckError strFunctionName

End Function

Public Function SetODIEnvironment(ByVal vstrNewValue)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo SetODIEnvironmentExit

    Const strFunctionName = "SetODIEnvironment"
    Dim nodeState As IXMLDOMNode
    Dim nodeEnv As IXMLDOMNode

    Set nodeEnv = xmlGetNode(m_nodeState, ".//ODISTATE/ODIENVIRONMENT")

    If nodeEnv Is Nothing Then
        ' node does not yet exist so create it
        Set nodeState = xmlGetMandatoryNode( _
            m_nodeState, "//ODISTATE/STATEINFO")
        Set nodeEnv = nodeState.appendChild( _
            m_nodeState.ownerDocument.createElement("ODIENVIRONMENT"))
    End If

    xmlSetAttributeValue nodeEnv, "DATA", vstrNewValue

SetODIEnvironmentExit:

    Set nodeState = Nothing
    Set nodeEnv = Nothing

    errCheckError strFunctionName

End Function

Public Function GetStateInfo() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo GetStateInfoExit

    Const strFunctionName = "GetStateInfo"

    Dim nodeStateInfo As IXMLDOMNode

    Set nodeStateInfo = m_nodeState.selectSingleNode(".//STATEINFO")
    
    If nodeStateInfo Is Nothing Then
        ' "Invalid Administration System state information"
        errThrowError strFunctionName, 4501, "STATEINFO not found"
    End If
    
    Set GetStateInfo = nodeStateInfo.cloneNode(True)

GetStateInfoExit:

    Set nodeStateInfo = Nothing
    
    errCheckError strFunctionName

End Function


