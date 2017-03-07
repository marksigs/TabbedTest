VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TemplateHandlerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      TemplateHandlerBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'Description:
'       Handles print template database.
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     02/11/00    Created
'IK     17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements ITemplateHandlerBO
'=============================================
'Constant Declaration Section
'=============================================
Private Const cstrROOT_NODE_NAME = "TEMPLATE"
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Public Function FindAvailableTemplates(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   return all the templates with:
'       a security level <= to the supplied security level
'       a stage >= to the supplied stage
'       a language matching the supplied language
' fixme - also match group name and activity number?
' pass:
'   vstrXmlIn
'       XML request data stream containing security level, app stage
'       <REQUEST>
'           <TEMPLATE>
'               <SECURITYLEVEL>
'               <STAGE>
'               <LANGUAGE>
'           </TEMPLATE>
'       </REQUEST>
' return:
'   XML containing list of template ids, names and descriptions
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "FindAvailableTemplates"

Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
    
    On Error GoTo FindAvailableTemplatesVbErr
    ' ik_20030210
    m_objContext.SetComplete
    
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                        
    ' Delegate to FreeThreadedDomDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ITemplateHandlerBO_FindAvailableTemplates(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindAvailableTemplates = xmlResponseElem.xml
    ' ik_20030210
    ' m_objContext.SetComplete
FindAvailableTemplatesExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    gobjTrace.TraceResponse FindAvailableTemplates
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
FindAvailableTemplatesVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    FindAvailableTemplates = m_objErrAssist.CreateErrorResponse
    
    gobjTrace.TraceIdentErrorResponse FindAvailableTemplates
    ' ik_20030210
    ' m_objContext.SetComplete
    ' go to clean-up section
    Resume FindAvailableTemplatesExit
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    ' ik_20030211 traceAssist support
    gobjTrace.StartTrace
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
Public Function GetTemplate(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a single instance of the persistant data associated with this
'   business object
' pass:
'   vstrXmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "GetTemplate"

Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
    
    On Error GoTo GetTemplateVbErr
    ' ik_20030210
    m_objContext.SetComplete
    
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                    
    ' Delegate to FreeThreadedDomDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ITemplateHandlerBO_GetTemplate(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetTemplate = xmlResponseElem.xml
    ' ik_20030210
    ' m_objContext.SetComplete
GetTemplateExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    gobjTrace.TraceResponse GetTemplate
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function

GetTemplateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    GetTemplate = m_objErrAssist.CreateErrorResponse
    ' ik_20030210
    ' m_objContext.SetComplete
    ' go to clean-up section
    Resume GetTemplateExit
End Function
Private Function ITemplateHandlerBO_GetTemplate(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
    
Const cstrFunctionName As String = "ITemplateHandlerBO_GetTemplate"

Dim xmlRequestNode As IXMLDOMNode
Dim xmlDataNode As IXMLDOMNode
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim objITemplateHandlerDO As ITemplateHandlerDO

    On Error GoTo GetTemplateVbErr:
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objITemplateHandlerDO = m_objContext.CreateInstance(App.Title & ".TemplateHandlerDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objITemplateHandlerDO.GetTemplate(xmlRequestNode)

'TW 26/7/2004 BBG 1070
'Temporary fix - this code ran OK in debug mode but failed with "The Parameter is incorrect" when compiled
'    xmlResponseElem.appendChild xmlDataNode
Dim strTempFix As String
    strTempFix = "<RESPONSE TYPE=""SUCCESS"">" & xmlDataNode.xml & "</RESPONSE>"
    xmlOut.setProperty "NewParser", True
    xmlOut.loadXML strTempFix
    Set xmlResponseElem = xmlOut.selectSingleNode("RESPONSE")
'TW End 26/7/2004 BBG 1070
    
    Set ITemplateHandlerBO_GetTemplate = xmlResponseElem
    ' ik_20030210
    ' m_objContext.SetComplete
GetTemplateExit:
    
    Set xmlRequestNode = Nothing
    Set objITemplateHandlerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function

GetTemplateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    Set ITemplateHandlerBO_GetTemplate = m_objErrAssist.CreateErrorResponseEx()
    ' ik_20030210
    ' m_objContext.SetComplete
    ' go to clean-up section
    Resume GetTemplateExit
End Function
Private Function ITemplateHandlerBO_FindAvailableTemplates( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get all instances of the persistant data associated with this
'   business object.
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "ITemplateHandlerBO_FindAvailableTemplates"

Dim objITemplateHandlerDO As ITemplateHandlerDO

Dim xmlDataNode As IXMLDOMNode
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseElem As IXMLDOMElement

    On Error GoTo FindAvailableTemplatesVbErr
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objITemplateHandlerDO = m_objContext.CreateInstance(App.Title & ".TemplateHandlerDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objITemplateHandlerDO.FindAvailableTemplates(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    ' ik_20030210
    ' m_objContext.SetComplete
    Set ITemplateHandlerBO_FindAvailableTemplates = xmlResponseElem

FindAvailableTemplatesExit:
    Set xmlRequestNode = Nothing
    Set objITemplateHandlerDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
FindAvailableTemplatesVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    Set ITemplateHandlerBO_FindAvailableTemplates = m_objErrAssist.CreateErrorResponseEx()
    ' ik_20030210
    ' m_objContext.SetComplete
    ' go to clean-up section
    Resume FindAvailableTemplatesExit
End Function
' ik_20030206
' tracing interface
Public Sub InitialiseTraceInterface( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStart As Single)
    
    If Not vblnTraceOn Then
        Exit Sub
    End If
    
    gobjTrace.TraceInitialiseFromParent vblnTraceOn, vstrFileName, vintStart
End Sub
