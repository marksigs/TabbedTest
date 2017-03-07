VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "omMQTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omMessageQueueTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Public Sub SendToQueue(vstrXMLRequest As String) As String

'Dependencies:  omMessageQueueDO, omMessageQueueBO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AD     30/10/01    Created - restructing of the component, still using the MessageQueue
'                   class at the bottom to maintain compatibility.
'RF     01/03/02    SYS4215 omMQ.SendToQueue is producing "method ~ of object ~ failed"
'                   in release. Various improvements identified.
'GHun   27/07/2006  CORE287 Commented out unnecessary App.LogEvent lines in SendToQueue
'------------------------------------------------------------------------------------------
Option Explicit
 
Implements IomMQTxBO

Private Function IomMQTxBO_SendToQueue( _
    ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequestNode  xml Request element
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SendToQueuevbErr

    Const strFunctionName As String = "SendToQueue"

    'CORE287 GHun
    'App.LogEvent "Start omMQTxBO " & strFunctionName, vbLogEventTypeInformation
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New DOMDocument
    Dim elemResponse As IXMLDOMElement
    Dim nodeMessageQueueResponse As IXMLDOMNode

    Dim objMessageQueue As IomMessageQueue

    Set elemResponse = xmlOut.createElement("RESPONSE")
    elemResponse.setAttribute "TYPE", "SUCCESS"
    xmlOut.appendChild elemResponse
   
    If objContext Is Nothing Then
        errThrowError strFunctionName, oeMTSNotFound
    End If

    Set objMessageQueue = objContext.CreateInstance(App.Title & ".omMessageQueue")
    
    Set nodeMessageQueueResponse = objMessageQueue.SendToQueue(vxmlRequestNode)

    ' RF 11/01/02
    errCheckXMLResponseNode nodeMessageQueueResponse, , True
    
    Set IomMQTxBO_SendToQueue = nodeMessageQueueResponse.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SendToQueueExit:

    'CORE287 GHun
    'App.LogEvent "Exit omMQTxBO " & strFunctionName, vbLogEventTypeInformation
    
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set elemResponse = Nothing
    Set objMessageQueue = Nothing
    
    Exit Function

SendToQueuevbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning elemResponse
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IomMQTxBO_SendToQueue = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SendToQueueExit

End Function


