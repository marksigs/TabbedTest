VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omMessageQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      C:\Projects\dev\omBatch\omMessageQueue.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'
'Description:   Omiga support for MSMQ
'
'Dependencies:  omMSMQ, omOMMQ
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     21/06/00    Created
'PSC    26/04/01    SYS2292 Amend to call correct methodand return a node rather than a
'                   document
'PSC    02/05/01    Amend to send the correct messgage data to AsyncSend
'DM     20/09/01    SYS2716 Removed if and added case to IomMessageQueue_SendToQueue
'------------------------------------------------------------------------------------------
Option Explicit

Implements IomMessageQueue
Implements ObjectControl

Private gobjContext As ObjectContext


'-------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'GHun   27/08/2003  BMIDS622 Changed SendToQueue to allow MessageQueueType to be passed in, so it doesn't
'                   always have to get the value from the database
'-------------------------------------------------------------------------------------------------------

Private Function IomMessageQueue_SendToQueue( _
    ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
    
    'Declare Variables
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlMessageQueue As IXMLDOMNode
    Dim xmlMessage As IXMLDOMNode
    
    Dim intMessageQueueType As Integer
    Dim strXMLResponse As String
    
    Dim objMessageQueue As IOmigaToMessageQueue
        
    'Error Handler
    On Error GoTo SendToQueuevbErr

    'Constant
    Const cstrFunctionName As String = "IomMessageQueue_SendToQueue"
    
    'BMIDS622 GHun 27/08/2003
    intMessageQueueType = xmlGetAttributeAsInteger(vxmlRequestNode.selectSingleNode("MESSAGEQUEUE"), "QUEUETYPE", "-1")
    'If the MessageQueueType has not been passed in, get it from the global parameter
    If intMessageQueueType = -1 Then
    'BMIDS622 End
        'Get MessageQueueType GlobalParameter
        intMessageQueueType = GetGlobalParamAmount("MessageQueueType")
    End If 'BMIDS622
    
    ' There are three types of Queue available at the moment.
    ' Must ensure that the correct Global parameter is set up for the Queue type.
    ' SYS2716
    Select Case intMessageQueueType
        Case 1 ' SQL Server MSMQ
                Set objMessageQueue = gobjContext.CreateInstance("omToMSMQ.OmigaToMessageQueue")
        Case 2, 3 'SQL Server OMMQ Oracle OMMQ
                Set objMessageQueue = gobjContext.CreateInstance("omToOMMQ.OmigaToMessageQueue")
        Case Else
                ' Error Message Queue type not supported.
                errThrowError "IomMessageQueue_SendToQueue", oeInvalidMessageQueueType, "Message Queue Type not supported, check global parameter MessageQueueType"
    End Select
           
    ' Make Request element based
    Set xmlNode = xmlCreateElementRequestFromNode(vxmlRequestNode, "MESSAGEQUEUE", True)
    
    ' Get the message queue element and remove the actual message text to pass down
    ' separately to AsyncSend
    Set xmlMessageQueue = xmlGetMandatoryNode(xmlNode, ".//MESSAGEQUEUE")
    Set xmlMessage = xmlGetMandatoryNode(xmlMessageQueue, "XML")
    xmlMessageQueue.removeChild xmlMessage
    
    ' PSC 26/04/01
    strXMLResponse = objMessageQueue.AsyncSend(xmlNode.xml, xmlMessage.Text)
    
    Set xmlDoc = xmlLoad(strXMLResponse, cstrFunctionName)
    
    ' RF 11/01/02
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    Set IomMessageQueue_SendToQueue = xmlDoc.documentElement
    
    gobjContext.SetComplete
    
SendToQueueExit:
    
    Set objMessageQueue = Nothing
    Set xmlDoc = Nothing
    Set xmlMessageQueue = Nothing
    Set xmlMessage = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
SendToQueuevbErr:

    Set objMessageQueue = Nothing
    Set xmlDoc = Nothing
    Set xmlMessageQueue = Nothing
    Set xmlMessage = Nothing
    Set xmlNode = Nothing
    
    gobjContext.SetAbort
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

