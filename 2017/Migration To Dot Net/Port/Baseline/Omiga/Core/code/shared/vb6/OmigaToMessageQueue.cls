VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OmigaToMessageQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      C:\Projects\dev\MessageQueue\omToMSMQ\OmigaToMessageQueue.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'
'Description:   Omiga support for OMMQ
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     16/10/00    Created
'AW     28/03/01    Changed order of last 2 parameters of ADO command to match sp signature
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     19/06/01    SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'LD     12/10/01    SYS2705 Support for SQL Server added
'LD     12/10/01    SYS2708 Add optional element CONNECTIONSTRING
'DS     04/12/01    SYS3298 Add LOGICALQUEUNAME processing
'LD     02/05/02    SYS4414 Amendments for BLOB support, and execute after date, and the threshold
'                   of large message size for sql server
'------------------------------------------------------------------------------------------
Option Explicit

Implements IOmigaToMessageQueue
Implements ObjectControl

Private m_objContext As ObjectContext


Public Function CSafeDate(ByVal vvntExpression As Variant) As Date
' header ----------------------------------------------------------------------------------
' description:
'   creates a date representation of an expression
' pass:
'   vvntExpression  Expression to be converted to a date
' Returns:
'   CSafeDate   Converted Expression
'------------------------------------------------------------------------------------------
On Error GoTo CSafeDateVbErr

    Dim strFunctionName As String
    strFunctionName = "CSafeDate"
    
    Dim dteRetValue As Date

    If Len(vvntExpression) > 0 Then
        dteRetValue = CDate(vvntExpression)
    End If
    
    CSafeDate = dteRetValue

    Exit Function

CSafeDateVbErr:

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function IOmigaToMessageQueue_AsyncSend( _
    ByVal strXMLRequest As String, ByVal strMessage As String) As String
' header ----------------------------------------------------------------------------------
' description:
'    Sends an an asynchronous message to a component via OMMQ
'
' pass:         strMessage containing the message to be sent
'               strXMLRequest containing the queuename and ProgID of the component to be called
'   <REQUEST>
'        <MESSAGEQUEUE>
'               <QUEUENAME>.\Localx</QUEUNAME>
'               <PROGID>omApp.omClass</PROGID>
'               <PRIORITY>3</PRIORITY>
'               <EXECUTEAFTERDATE>2001-03-02 12:55</EXECUTEAFTERDATE>
'               <CONNECTIONSTRING>Provider=MSDAORA;Data Source=om4ltest;User ID=production;Password=production;</CONNECTIONSTRING>
'       </MESSAGEQUEUE>
'   </REQUEST>
'
'   Note that PRIORITY, EXECUTEAFTERDATE and CONNECTIONSTRING are optional elements.  If CONNECTIONSTRING
'   is not supplied, then the Omiga 4 connection string is used (i.e. from the registry)
'
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo IOmigaToMessageQueue_AsyncSendVbErr

    Const strFunctionName As String = "IOmigaToMessageQueue_AsyncSend"
    
    'Extract the request details
    Dim objXMLRequest As DOMDocument
    Set objXMLRequest = xmlLoad(strXMLRequest, strFunctionName)
    
    Dim objXmlRequestElem As IXMLDOMElement
    Set objXmlRequestElem = xmlGetMandatoryNode(objXMLRequest, "REQUEST/MESSAGEQUEUE")
    
    Dim strQueueName As String
    Dim strLogicalQueueName As String
    strLogicalQueueName = xmlGetNodeText(objXmlRequestElem, "LOGICALQUEUENAME")
     
    If strLogicalQueueName <> "" Then
        strQueueName = GetGlobalParamString(strLogicalQueueName)
    Else
        strQueueName = xmlGetMandatoryNodeText(objXmlRequestElem, "QUEUENAME")
    End If
    
    Dim strProgId As String
    strProgId = xmlGetMandatoryNodeText(objXmlRequestElem, "PROGID")
         
    Dim strGuid As String
    strGuid = CreateGUID()
    
    Dim szPriority As String
    Dim nPriority As Integer
    
    szPriority = xmlGetNodeText(objXmlRequestElem, "PRIORITY")
    
    ' 3 is the default
    
    If (szPriority = "") Then
        nPriority = 3
    Else
        nPriority = CInt(szPriority)
    End If
    
    Dim szEXECUTEAFTERDATE As String
    szEXECUTEAFTERDATE = xmlGetNodeText(objXmlRequestElem, "EXECUTEAFTERDATE")
    
    Dim strConnectionString As String
    Dim enmDBProvider As DBPROVIDER
    strConnectionString = xmlGetNodeText(objXmlRequestElem, "CONNECTIONSTRING")
    If strConnectionString = "" Then
        strConnectionString = adoGetDbConnectString
        enmDBProvider = adoGetDbProvider()
    Else
        If InStr(strConnectionString, "SQLOLEDB") > 0 Then
            enmDBProvider = omiga4DBPROVIDERSQLServer
        Else
            enmDBProvider = omiga4DBPROVIDEROracle
        End If
    End If
    
    Dim lngMessageSize As Long
    lngMessageSize = Len(strMessage)
    Dim bLargeMessage As Boolean
    If (lngMessageSize > 2000) Then ' threshold for large message
        bLargeMessage = True
    Else
        bLargeMessage = False
    End If
    Dim bMSDAORAChunking As Boolean
    If (enmDBProvider = omiga4DBPROVIDEROracle And lngMessageSize > 32000) Then ' threshold for large message
        bMSDAORAChunking = True
    Else
        bMSDAORAChunking = False
    End If
    
    Dim adoCommand As New ADODB.Command
    If enmDBProvider = omiga4DBPROVIDEROracle Then
        adoCommand.CommandType = adCmdText
        If (bLargeMessage And bMSDAORAChunking = False) Then
            adoCommand.CommandText = "{call sp_MQLOMMQ.SendMessageLONG(?,?,?,?,?,?)}"
        Else
            adoCommand.CommandText = "{call sp_MQLOMMQ.SendMessage(?,?,?,?,?,?)}"
        End If
    ElseIf enmDBProvider = omiga4DBPROVIDERSQLServer Then
        adoCommand.CommandType = adCmdText
        If (bLargeMessage) Then
            adoCommand.CommandText = "{call USP_MQLOMMQSENDMESSAGENTEXT(?,?,?,?,?,?)}"
        Else
            adoCommand.CommandText = "{call USP_MQLOMMQSENDMESSAGE(?,?,?,?,?,?)}"
        End If
    End If
      
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
    adoCommand.Parameters.Append adoParameter
    adoCommand(0) = GuidStringToByteArray(strGuid)

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strQueueName), strQueueName)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoParameter.Size = Len(strProgId)
    adoCommand(2) = strProgId
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bMSDAORAChunking Then
        adoCommand(3) = Null
    Else
        adoParameter.Size = Len(strMessage)
        adoCommand(3) = strMessage
    End If
    
    ' new bind parameters
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(4) = nPriority
    
    Set adoParameter = adoCommand.CreateParameter(, adDate, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If (szEXECUTEAFTERDATE = "") Then
        ' default to null (i.e. will be converted to the database date)
        adoCommand(5) = Null
    Else
        adoCommand(5) = CSafeDate(szEXECUTEAFTERDATE)
    End If
          
    Dim adoConnection As ADODB.Connection
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = strConnectionString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    
    If (bMSDAORAChunking) Then
        Dim adoRecordset As ADODB.Recordset
        Set adoRecordset = New ADODB.Recordset
        adoRecordset.CursorType = adOpenKeyset
        adoRecordset.LockType = adLockOptimistic
        adoRecordset.Open "SELECT * FROM MQLOMMQLONG WHERE MESSAGEID IS NULL", adoConnection, , , adCmdText
        
        adoRecordset.AddNew
        adoRecordset!MessageId = strGuid
        
        Const lngChunkSize = 32000
        Dim strChunk As String
        Dim lngOffset As Long
        lngOffset = 0 ' Reset offset.
        Do While lngOffset < lngMessageSize
           strChunk = Left(Right(strMessage, lngMessageSize - lngOffset), lngChunkSize)
           adoRecordset!xml.AppendChunk strChunk
           lngOffset = lngOffset + lngChunkSize
        Loop
        adoRecordset.Update
        adoRecordset.Close
    End If
    
    ' send the message
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.Execute , , adExecuteNoRecords
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close
    
    IOmigaToMessageQueue_AsyncSend = "<RESPONSE TYPE=""SUCCESS""/>"
    m_objContext.SetComplete

CleanupOnExit:

    Set adoCommand = Nothing
    Set adoConnection = Nothing
    Set objXMLRequest = Nothing
    Set objXmlRequestElem = Nothing
    
    Exit Function
    
IOmigaToMessageQueue_AsyncSendVbErr:

    ' RF 11/01/02 create ERROR response block from Err object
    IOmigaToMessageQueue_AsyncSend = errCreateErrorResponse()
    
    m_objContext.SetAbort
    
    'Resume CleanupOnExit
    GoTo CleanupOnExit
    
End Function
   
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing

End Sub



