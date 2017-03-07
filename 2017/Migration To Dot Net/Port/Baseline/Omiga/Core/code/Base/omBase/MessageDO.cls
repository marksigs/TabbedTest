VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "MessageDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'Workfile:      MessageDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Message Data Object
'
'Dependencies:
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    21/10/99    Created
'LD     07/11/00    Explicity close recordsets
' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog  Date     Description
'TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
 '=============================================
 'Variable Declaration Section
 '=============================================
 '=============================================
 'Constant Declaration Section
 '=============================================
Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    Dim strFunctionName As String
    strFunctionName = "GetData"
'    Dim objErrAssist As New ErrAssist
                        
    Dim objDOAssist As DOAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
            
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
            
    GetData = objDOAssist.GetData(vstrXMLRequest, LoadMessageData)
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetDataVbErr:
        
        
    Set objDOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
'    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function LoadMessageData()
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
Dim XMLstr As String
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MESSAGE"
    strXML = strXML & "<PRIMARYKEY>MESSAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>MESSAGETEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MESSAGETYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
LoadMessageData = strXML
End Function
Public Function AddDerivedData(ByVal vstrXMLData As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vstrXMLData         base XML data stream
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:
'   AddDerivedData      base XML data stream plus any derived values
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr
    Dim strFunctionName As String
    strFunctionName = "AddDerivedData"
'    Dim objErrAssist As New ErrAssist
    AddDerivedData = vstrXMLData
'    Set objErrAssist = Nothing
    Exit Function
            
'    Set objErrAssist = Nothing
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    Exit Function
AddDerivedDataVBErr:
'    Set objErrAssist = Nothing
    AddDerivedData = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetMessageDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'                   string containing XML data stream representation of
'                   data retrieved
'------------------------------------------------------------------------------------------
Dim lngErrorNumber As Long
Dim strErrorSource As String
Dim strErrorDescription As String
Dim strErrorHelpContext As String
Dim strErrorHelpFile As String
lngErrorNumber = Err.Number
strErrorSource = Err.Source
strErrorDescription = Err.Description
strErrorHelpContext = Err.HelpContext
strErrorHelpFile = Err.HelpFile
    
    Dim strFunctionName As String
    strFunctionName = "GetMessageDetails"
        
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim strSQL As String
    Dim strXML As String
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
'    Dim objXmlElem As IXMLDOMElement
'    Dim objXmlListNode As IXMLDOMNode
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
    Dim rstThisRecSet As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim strResponse As String
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Dim strMessageNumber As String
        
    strMessageNumber = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, "MESSAGENUMBER")
    If Len(strMessageNumber) = 0 Then
        strResponse = CreateErrorDocument(strMessageNumber, _
                                          strErrorSource, _
                                          strErrorHelpContext, _
                                          strErrorHelpFile)
    Else
        If m_objContext Is Nothing Then
            Set objADOAssist = New ADOAssist
            Set objDOAssist = New DOAssist
        Else
            Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
            Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
        End If
        strSQL = "SELECT * FROM MESSAGE WHERE"
        strSQL = strSQL & " MESSAGENUMBER = " & strMessageNumber
                    
        cmd.ActiveConnection = objADOAssist.GetConnStr
        cmd.CommandText = strSQL
                    
        Set rstThisRecSet = cmd.Execute(, , adCmdText)
                    
        If rstThisRecSet.EOF Then
            strResponse = CreateErrorDocument(strMessageNumber, _
                                              strErrorSource, _
                                              strErrorHelpContext, _
                                              strErrorHelpFile)
        Else
            Dim strClassDef As String
            strClassDef = LoadMessageData()
            ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
            ' rstThisRecSet.MoveFirst
            ' loop through the record set
            While Not rstThisRecSet.EOF
                strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strClassDef)
                Set objXmlDoc = m_objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
                Set objXmlNode = objXmlOut.appendChild(objXmlDoc.documentElement)
                rstThisRecSet.MoveNext
            Wend
            strResponse = objXmlOut.xml
            Err.Number = lngErrorNumber
            Err.Source = strErrorSource
            Err.Description = strErrorDescription
            Err.HelpContext = strErrorHelpContext
            Err.HelpFile = strErrorHelpFile
        End If
    End If
    rstThisRecSet.Close
    GetMessageDetails = strResponse
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlOut = Nothing
    Set objXmlDoc = Nothing
    Set objXmlIn = Nothing
    Set objXmlNode = Nothing
'    Set objXmlElem = Nothing
'    Set objXmlListNode = Nothing
    Set objADOAssist = Nothing
    Set objDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
            
    Exit Function
End Function
Private Function CreateErrorDocument(ByVal vstrMessageNumber As String, _
                                     ByVal vstrErrorSource As String, _
                                     ByVal vstrErrorHelpContext As String, _
                                     ByVal vstrErrorHelpFile As String) As String
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlTableElem As IXMLDOMElement
    Dim objXmlElem As IXMLDOMElement
    Dim strMessageText As String
    strMessageText = "Message number " & vstrMessageNumber & " not found"
    Set objXmlTableElem = objXmlOut.createElement("MESSAGE")
    objXmlOut.appendChild objXmlTableElem
    Set objXmlElem = objXmlOut.createElement("MESSAGENUMBER")
    objXmlElem.Text = vstrMessageNumber
    objXmlTableElem.appendChild objXmlElem
    Set objXmlElem = objXmlOut.createElement("MESSAGETEXT")
    objXmlElem.Text = strMessageText
    objXmlTableElem.appendChild objXmlElem
    Set objXmlElem = objXmlOut.createElement("MESSAGETYPE")
    objXmlElem.Text = "Error"
    objXmlTableElem.appendChild objXmlElem
    Err.Number = vbObjectError + 512 + omiga4InvalidMessageNo
    Err.Source = vstrErrorSource
    Err.Description = strMessageText
    Err.HelpContext = vstrErrorHelpContext
    Err.HelpFile = vstrErrorHelpFile
    CreateErrorDocument = objXmlOut.xml
    
    Set objXmlOut = Nothing
    Set objXmlTableElem = Nothing
    Set objXmlElem = Nothing
    
End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
