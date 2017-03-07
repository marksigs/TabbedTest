VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MQL.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     10/01/01    Created
'AD     06/02/01    Put in link to MQL Configure
'------------------------------------------------------------------------------------------

Option Explicit

Public m_SnapInData As SnapInData

Public Function MQLConfigureGetQueueInfo(ByVal strComputerName As String, ByVal strQueueName) As String

    On Error GoTo ErrTrap_MQLConfigureGetQueueInfo

    Dim strXMLIn As String
    strXMLIn = "<?xml version=""1.0""?>"
    strXMLIn = strXMLIn & "<REQUEST ACTION=""GET"">"
    strXMLIn = strXMLIn & "<QUEUELIST>"
    strXMLIn = strXMLIn & "<QUEUE>"
    strXMLIn = strXMLIn & "<NAME>" + strQueueName + "</NAME>"
    strXMLIn = strXMLIn & "<TYPE>MSMQ1</TYPE>"
    strXMLIn = strXMLIn & "</QUEUE>"
    strXMLIn = strXMLIn & "</QUEUELIST>"
    strXMLIn = strXMLIn & "</REQUEST>"
    
    Dim strXMLOut As String
    
    strXMLOut = MQLConfigure(strXMLIn, strComputerName)
    MQLConfigureGetQueueInfo = strXMLOut

    Exit Function
    
' Error Handler for this method
ErrTrap_MQLConfigureGetQueueInfo:
    m_SnapInData.Error.DisplayError "MQLConfigureGetQueueInfo"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function MQLConfigureGetAllQueuesInfo(ByVal strComputerName As String) As String

    On Error GoTo ErrTrap_MQLConfigureGetAllQueuesInfo
    
    Dim strXMLIn As String
    strXMLIn = "<?xml version=""1.0""?>"
    strXMLIn = strXMLIn & "<REQUEST ACTION=""GET"">"
    strXMLIn = strXMLIn & "<QUEUELIST>"
    strXMLIn = strXMLIn & "</QUEUELIST>"
    strXMLIn = strXMLIn & "</REQUEST>"
    
    Dim strXMLOut As String
    strXMLOut = MQLConfigure(strXMLIn, strComputerName)
    
    MQLConfigureGetAllQueuesInfo = strXMLOut

    Exit Function
' Error Handler for this method
ErrTrap_MQLConfigureGetAllQueuesInfo:
    m_SnapInData.Error.DisplayError "MQLConfigureGetAllQueuesInfo"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function MQLCreateQueue(ByVal dataqpgNewQueue As datawpgNewQueue) As Boolean
    
    ' send create queue request to MQL - if successful then return true
    
    On Error GoTo ErrTrap_MQLCreateQueue
    
    Dim strXMLIn As String
    strXMLIn = "<?xml version=""1.0""?>"
    strXMLIn = strXMLIn & "<REQUEST ACTION=""CREATE"">"
    strXMLIn = strXMLIn & "<QUEUELIST>"
    strXMLIn = strXMLIn & "<QUEUE>"
    
    ' unpack dataqpgNewQueue structure
    
    strXMLIn = strXMLIn & "<NAME>" & dataqpgNewQueue.m_strQueueName & "</NAME>"
    strXMLIn = strXMLIn & "<TYPE>" & dataqpgNewQueue.m_strQueueType & "</TYPE>"
    
    If dataqpgNewQueue.m_strQueueType = "OMMQ1" Then
        strXMLIn = strXMLIn & "<CONNECTIONSTRING>" & dataqpgNewQueue.m_strConnectionString & "</CONNECTIONSTRING>"
        strXMLIn = strXMLIn & "<POLLINTERVAL>" & dataqpgNewQueue.m_nPollingInterval & "</POLLINTERVAL>"
    End If
    
    
    strXMLIn = strXMLIn & "<THREADSLIST>"
    strXMLIn = strXMLIn & "<THREADS>"
    strXMLIn = strXMLIn & "<NUMBER>" & dataqpgNewQueue.m_nQueueThreads & "</NUMBER>"
    strXMLIn = strXMLIn & "</THREADS>"
    strXMLIn = strXMLIn & "</THREADSLIST>"
    strXMLIn = strXMLIn & "</QUEUE>"
    strXMLIn = strXMLIn & "</QUEUELIST>"
    strXMLIn = strXMLIn & "</REQUEST>"
    
    Dim strXMLOut As String
    
    strXMLOut = MQLConfigure(strXMLIn, dataqpgNewQueue.m_strComputerName)
    
    ' check result
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNode As IXMLDOMNode
    
    If XMLDocument.documentElement.getAttribute("TYPE") <> "SUCCESS" Then
        MQLCreateQueue = False
        ' report error
    Else
        MQLCreateQueue = True
    End If

    Exit Function

ErrTrap_MQLCreateQueue:
    m_SnapInData.Error.DisplayError "MQLCreateQueue"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function MQLCreateQueueEvent(ByVal datawpgNewQEvent As datawpgNewQEvent) As Boolean
    
' AD 04/04/01 - this needs to be finished and tested - prop sheet needed.
    
'    On Error GoTo ErrTrap_MQLCreateQueueEvent
    
'    Dim strXMLIn As String
'    strXMLIn = "<?xml version=""1.0""?>"
'    strXMLIn = strXMLIn & "<REQUEST ACTION=""CREATE"">"
'    strXMLIn = strXMLIn & "<QUEUELIST>"
'    strXMLIn = strXMLIn & "<QUEUE>"
    
    ' unpack dataqpgNewQueue structure
    
'    strXMLIn = strXMLIn & "<NAME>" & dataqpgNewQueue.m_strQueueName & "</NAME>"
'    strXMLIn = strXMLIn & "<TYPE>" & dataqpgNewQueue.m_strQueueType & "</TYPE>"
    
    
    
'    datawpgNewQEvent.m_strKey = "??" ' to be returned from service - guid
    
'    Dim strXMLOut As String
    
'    strXMLOut = MQLConfigure(strXMLIn, dataqpgNewQueue.m_strComputerName)
    
    ' check result
'    Dim XMLDocument As DOMDocument
'    Set XMLDocument = New DOMDocument
'    XMLDocument.loadXML strXMLOut
'    Dim XMLNode As IXMLDOMNode
    
'    If XMLDocument.documentElement.getAttribute("TYPE") <> "SUCCESS" Then
        MQLCreateQueueEvent = False
        ' report error
'    Else
'        MQLCreateQueueEvent = True
'    End If

    Exit Function
    
'ErrTrap_MQLCreateQueueEvent:
'    m_SnapInData.Error.DisplayError "MQLCreateQueueEvent"
'    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function MQLStallComponent(ByVal datawpgStallComp As datawpgStallComp, strQueueType As String) As Boolean
    
    ' at the moment only sends 1 component name across but can handle more
    
    On Error GoTo ErrTrap_MQLStallComponent
    
    Dim strXMLIn As String
    strXMLIn = "<?xml version=""1.0""?>"
    strXMLIn = strXMLIn & "<REQUEST ACTION=""UPDATE"">"
    strXMLIn = strXMLIn & "<QUEUELIST>"
    strXMLIn = strXMLIn & "<QUEUE>"
    strXMLIn = strXMLIn & "<NAME>" & datawpgStallComp.m_strQueueName & "</NAME>"
    strXMLIn = strXMLIn & "<TYPE>" & strQueueType & "</TYPE>"
    strXMLIn = strXMLIn & "<TASK>ADDSTALLEDCOMPONENTS</TASK>"
    strXMLIn = strXMLIn & "<COMPONENTLIST>"
    strXMLIn = strXMLIn & "<COMPONENT>" & datawpgStallComp.m_strProgID & "</COMPONENT>"
    strXMLIn = strXMLIn & "</COMPONENTLIST>"
    strXMLIn = strXMLIn & "</QUEUE>"
    strXMLIn = strXMLIn & "</QUEUELIST>"
    strXMLIn = strXMLIn & "</REQUEST>"

    
    Dim strXMLOut As String
    
    strXMLOut = MQLConfigure(strXMLIn, datawpgStallComp.m_strComputerName)
    
    ' check result
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNode As IXMLDOMNode
    
    If XMLDocument.documentElement.getAttribute("TYPE") <> "SUCCESS" Then
        MQLStallComponent = False
        ' report error
    Else
        MQLStallComponent = True
    End If
    
    Exit Function

ErrTrap_MQLStallComponent:
    m_SnapInData.Error.DisplayError "MQLStallComponent"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function MQLStartComponent(ByVal strComputerName As String, ByVal strQueueName As String, ByVal strProgID As String, ByVal strType As String) As Boolean
        
    ' at the moment only sends 1 component name across but can handle more
    On Error GoTo ErrTrap_MQLStartComponent
    
    Dim strXMLIn As String
    strXMLIn = "<?xml version=""1.0""?>"
    strXMLIn = strXMLIn & "<REQUEST ACTION=""UPDATE"">"
    strXMLIn = strXMLIn & "<QUEUELIST>"
    strXMLIn = strXMLIn & "<QUEUE>"
    strXMLIn = strXMLIn & "<NAME>" & strQueueName & "</NAME>"
    strXMLIn = strXMLIn & "<TYPE>" & strType & "</TYPE>"
    strXMLIn = strXMLIn & "<TASK>RESTARTCOMPONENT</TASK>"
    strXMLIn = strXMLIn & "<COMPONENTLIST>"
    strXMLIn = strXMLIn & "<COMPONENT>" & strProgID & "</COMPONENT>"
    strXMLIn = strXMLIn & "</COMPONENTLIST>"
    strXMLIn = strXMLIn & "</QUEUE>"
    strXMLIn = strXMLIn & "</QUEUELIST>"
    strXMLIn = strXMLIn & "</REQUEST>"

    
    Dim strXMLOut As String
    
    strXMLOut = MQLConfigure(strXMLIn, strComputerName)
    
    ' check result
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML strXMLOut
    Dim XMLNode As IXMLDOMNode
    
    If XMLDocument.documentElement.getAttribute("TYPE") <> "SUCCESS" Then
        MQLStartComponent = False
        ' report error
    Else
        MQLStartComponent = True
    End If
    Exit Function

ErrTrap_MQLStartComponent:
    m_SnapInData.Error.DisplayError "MQLStartComponent"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function MQLConfigure(strXMLIn As String, ByVal strComputerName As String) As String

    On Error GoTo ErrTrap_MQLConfigure
        
    ' call the MessageQueueListenerService
    Dim WbemServices As SWbemServices
    Set WbemServices = m_SnapInData.WMI.WMIGetWbemServices(strComputerName, "root/cimv2/Marlborough_Stirling")
    Dim WbemObject As SWbemObject
    Set WbemObject = WbemServices.Get("MessageQueueListenerWMI")
    Dim WbemMethod As SWbemMethod
    Set WbemMethod = WbemObject.Methods_("Configure")
    Dim inParam As SWbemObject
    Set inParam = WbemMethod.InParameters.SpawnInstance_()
    inParam.strXMLIn = strXMLIn
    Dim outParam As SWbemObject
    Set outParam = WbemObject.ExecMethod_("Configure", inParam)
        
    Select Case outParam.returnValue
        Case 0
            MQLConfigure = outParam.strXMLOut
        Case 1
            m_SnapInData.SnapIn.ConsoleMsgBox "MessageQueueListenerWMI - Failed to getting XML in"
        Case 2
            m_SnapInData.SnapIn.ConsoleMsgBox "MessageQueueListenerWMI - Failed to create MessageQueueListener object"
        Case 3
            m_SnapInData.SnapIn.ConsoleMsgBox "MessageQueueListenerWMI - Failed to invoke configure routine"
        Case 4
            m_SnapInData.SnapIn.ConsoleMsgBox "MessageQueueListenerWMI - Failed setting XML out"
        Case 5
            m_SnapInData.SnapIn.ConsoleMsgBox "MessageQueueListenerWMI - Unknown failure"
        Case Else
            Debug.Assert 0
    End Select
    
    ' display the description of the error if any
    Dim XMLDocument As DOMDocument
    Set XMLDocument = New DOMDocument
    XMLDocument.loadXML MQLConfigure
    If XMLDocument.documentElement.getAttribute("TYPE") <> "SUCCESS" Then
        Dim Node As IXMLDOMNode
        Set Node = XMLDocument.documentElement.selectSingleNode(".//ERROR")
        Dim strDescription As String
        strDescription = Node.selectSingleNode("DESCRIPTION").Text
        m_SnapInData.SnapIn.ConsoleMsgBox strDescription
    End If
    
    Exit Function
    
' Error Handler for this method
ErrTrap_MQLConfigure:
    m_SnapInData.Error.DisplayError "MQLConfigure"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

