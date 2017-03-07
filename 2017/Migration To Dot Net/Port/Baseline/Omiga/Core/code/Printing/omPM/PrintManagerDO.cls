VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PrintManagerDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PrintManagerDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Created:       26/06/2001
'Author:        Mike Kennedy
'Description:   DO for the Print Manager component
'Dependencies:  IPrintManagerDO.cls
'               omFVS - File Versioning System
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MK     26/06/01    Created
'MK     04/10/01    Updated for SQL Server
'DR     28/01/02    DMSSYS0005 I've had to add:
'                       1. sqlFormatDate() from sqlAssist.bas from omRB to get the code to compile,
'                       2. a wrapper function GetDbProvider() to get the code to compile,
'                       3. Enumerated type DATETIMEFORMAT from sqlAssist.bas to get the code to compile,
'                       4. Enumerated type ERRNUMBER from ErrAssist.bas to get the code to compile,
'                       5. Added errRaiseError() from ErrAssist.bas to get the code to compile.
'DR     31/01/02    DMSSYS0005 I've amended the sqlFormatDate() function so that it doesnt add an equals sign
'                   and it now uses 102 rather than 103.
'DR     10/02/02    DMSSYS0005 IPrintManagerDO_GetDocumentHistoryList() If searchKey1 is not present then
'                   dont attempt to use it. Also fixed AMD typos in SQL.
'DR     11/02/02    DMSSYS0005 IPrintManagerDO_GetDocumentHistoryList() CREATETIMESTAMP is now part of
'                   the response XML.
'DR     27/02/02    DMSSYS0055 Removed Enum Errors.
'DR     25/03/02    Standardised vxmlRequestNode naming, added debug info.
'SG     23/04/02    SYS4407 Added EventDate to IPrintManagerDO_GetDocumentHistoryList output XML
'TW     17/09/04    General Tidy up
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'Prog   Date        Description
'DPF    19/11/2002  Have added HOSTTEMPLATEID to response from IPrintManagerDO_GetDocumentHistoryList
'IK     06/12/2002  re-apply BMIDS00885
'IK     17/02/2003  BM0200 - add TraceAssist support
'GHun   09/04/2003  BM0508 - Amend CreateDocumentDetailRecord to handle apostrophes in SQL
'------------------------------------------------------------------------------------------
'BBG Specific History:
'Prog   Date        Description
'TK    08/09/2004   BBG879 - Amend IPrintManagerDO_GetDocumentHistoryList and IPrintManagerDO_GetDocumentEventList
'TK    24/09/2004   BBG1317 - Amend IPrintManagerDO_GetDocumentHistoryList - populate data via store proc
'------------------------------------------------------------------------------------------
'Core Specific History:
'AS     04/05/2005  CORE111 Ensure DELIVERYENGINETYPE and DELIVERYTYPE are set in the DocumentAuditDetails record.
'TW     24/10/2005  CORE208/MAR285 Changes for MARS project
'TW     25/10/2005  CORE210/MAR294 Changes for MARS project
'PSC    10/11/2005  CORE196 Amend UpdateDocumentDetailRecord to update the Document Detail Record
'PSC    03/02/2006  CORE239 Amend CreateEventDetailRecord to use milliseconds for event date
'GHun   02/03/2006  CORE252 Fixes for recategorisation
'GHun   02/03/2006  CORE252 Changed UpdateDocumentDetailRecord and GetDocumentEventList
'GHun   06/03/2006  CORE252 Changed GetDocumentHistoryList
'AS     15/12/2006  CORE324 Add support for DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS.
'AS     12/01/2007  CORE331 Removed GetDocumentHistoryListXml - replaced with omCRUD call.
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IPrintManagerDO
Private gobjContext As ObjectContext
Private m_objXmlAssist As XMLAssist
'DR DMSSYS0005
Public Enum DATETIMEFORMAT
    dtfDateTime = 0      ' date and time
    dtfDate = 1          ' date only
    dtfTime = 2          ' time only
End Enum
'DR
Private Const cstrClassName = "omPM::PrintManagerDO"



Public Sub errRaiseError( _
    ByVal vlngErrNo As Long, _
    ByVal vstrSource As String, _
    Optional ByVal vstrDesc As String = "")
    Dim strDesc As String
    If vstrDesc = Empty Then
        strDesc = errGetMessageText(vlngErrNo)
    Else
        strDesc = vstrDesc
    End If
    Err.Raise vlngErrNo, vstrSource, strDesc
End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    Set m_objXmlAssist = New XMLAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
 
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
    Set m_objXmlAssist = Nothing
End Sub
Public Sub IPrintManagerDO_PutAuditTrail( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

Dim strDocumentGuid As String

Dim xmlNode As IXMLDOMElement

    On Error GoTo CreateAuditTrail_Exit
    
    ' ik_20030211
    gobjContext.SetComplete
    ' create element
    Const cstrFunctionName As String = "PutAuditTrail"
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Create the audit trail here!
    Set xmlNode = xmlGetNode(vxmlRequestNode, "EVENTDETAIL")
    
    Select Case CSafeInt(xmlGetAttributeText(xmlNode, "EVENTKEY"))
        Case EVENTKEY_CREATED ' processing a new document
            ' Create a document detail....
            strDocumentGuid = CreateGUID()
            Set xmlNode = xmlGetNode(vxmlRequestNode, "DOCUMENTDETAILS")
            xmlSetAttributeValue xmlNode, "DOCUMENTGUID", strDocumentGuid
            CreateDocumentDetailRecord vxmlRequestNode
            ' and an event detail record ...
            Set xmlNode = vxmlRequestNode
            xmlSetAttributeValue xmlNode, "DOCUMENTGUID", strDocumentGuid
            'TW 31/10/2005 CORE210 return DOCUMENTGUID for use in PACK processing
            xmlSetAttributeValue vxmlResponseNode, "DOCUMENTGUID", strDocumentGuid
            'TW 31/10/2005 CORE210 End
            CreateEventDetailRecord xmlNode
        
        'AS 15/12/2006 CORE324 Added support new events.
        Case _
            EVENTKEY_EDITED, _
            EVENTKEY_RECATEGORISATION, _
            EVENTKEY_NOTAPPROVED, _
            EVENTKEY_APPROVED, _
            EVENTKEY_GEMINIPRINTED
            
            ' Update the DOCUMENTAUDITDETAILS record.
            UpdateDocumentDetailRecord vxmlRequestNode
            ' Create an EVENTAUDITDETAIL record.
            CreateEventDetailRecord vxmlRequestNode
            
        Case _
            EVENTKEY_VIEWED, _
            EVENTKEY_REPRINTED, _
            EVENTKEY_FULFILLMENTSEND, _
            EVENTKEY_FULFILLMENTRESEND, _
            EVENTKEY_FULFILLMENTCANCEL, _
            EVENTKEY_SMS, _
            EVENTKEY_EMAIL
            
            ' Create an EVENTAUDITDETAIL record.
            CreateEventDetailRecord vxmlRequestNode
        
        Case Else
            ' In case of any errors
    End Select
    ' If any errors then return error (funnily enough!)...
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
CreateAuditTrail_Exit:
    ' And tell MTS we're done!!!
    ' ik_20030211
    ' gobjContext.SetComplete
    Set xmlNode = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
' Creates a summary list pf documents printed for the requested ApplicationNumber and any
' additional search criteria
Public Sub IPrintManagerDO_GetDocumentHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetDocumentHistoryList_Exit

    Dim xmlDocDetailNode As IXMLDOMNode
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim strApplicationNumber As String
    Dim strSearchCriteria1 As String
    Dim strSearchCriteria2 As String
    Dim strSearchCriteria3 As String
    Dim strData As String

    Const cstrFunctionName As String = "GetDocumentHistoryList"

    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName

    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream

    ' Need to get the app number and the search keys from the XML In stream
    strApplicationNumber = xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
    strSearchCriteria1 = xmlGetAttributeText(vxmlRequestNode, "SEARCHKEY1")
    strSearchCriteria2 = xmlGetAttributeText(vxmlRequestNode, "SEARCHKEY2")
    strSearchCriteria3 = xmlGetAttributeText(vxmlRequestNode, "SEARCHKEY3")

    'Application Number
    'CORE252 GHun Pass through the correct length rather than padding with spaces
    Set adoParameter = adoCommand.CreateParameter("@pApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    'Search Key1
    'CORE252 GHun Pass through the correct length rather than padding with spaces
    Set adoParameter = adoCommand.CreateParameter("@p_SearchCriteria1", adBSTR, adParamInput, Len(strSearchCriteria1), strSearchCriteria1)
    If Len(strSearchCriteria1) = 0 Then
        adoParameter.Attributes = adParamNullable
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    'Not in used - Search Key2
    'Set adoParameter = adoCommand.CreateParameter("@p_SearchCriteria2", adBSTR, adParamInput, Len(strSearchCriteria2), strSearchCriteria2)
    'adoCommand.Parameters.Append adoParameter
    'Not in used - Search Key3
    'Set adoParameter = adoCommand.CreateParameter("@p_SearchCriteria3", adBSTR, adParamInput, Len(strSearchCriteria3), strSearchCriteria3)
    'adoCommand.Parameters.Append adoParameter

    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .open
    End With

    adoStream.open
    With adoCommand
        .CommandText = "USP_CURRENTVERSION"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "DOCUMENTHISTORYLIST"
        .Execute , , adExecuteStream
    End With

    strData = adoStream.ReadText()
    adoConnection.Close
    adoStream.Close

    Set xmlDocDetailNode = m_objXmlAssist.load(strData, TypeName(Me), cstrFunctionName)

    vxmlResponseNode.appendChild xmlDocDetailNode.selectSingleNode("DOCUMENTHISTORYLIST")
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
        
GetDocumentHistoryList_Exit:
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlDocDetailNode = Nothing
    Set xmlDocDetailNode = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

' Creates a summary list of document events for the requested search criteria
Public Sub IPrintManagerDO_GetDocumentEventList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

Const cstrFunctionName As String = "GetDocumentEventList"

Dim adoCommand As ADODB.Command
Dim adoConnection As ADODB.Connection
Dim adoRecordSet As Recordset

Dim strApplicationNumber As String
Dim strDateEvent As String
Dim strDocumentGuid As String
Dim strDocumentVersion As String
Dim strEvent As String
Dim strFileGUID As String
Dim strNumCopies As String
Dim strPrintLocation As String
Dim strSQL As String
Dim strUnitname As String
Dim strUsername As String

Dim xmlDocDetailNode As IXMLDOMNode
Dim xmlNewChildNode As IXMLDOMNode    ' A node that we can reuse for each history list entry

    On Error GoTo GetDocumentEventList_Exit
    
    ' ik_20030211
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Need to get the app number and the search keys from the XML In stream
    strApplicationNumber = xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
    strDocumentGuid = xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID")
    ' Create s SQL statement to get back a recordset of audi trail entries
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "SELECT * " & _
                 "FROM EventAuditDetail " & _
                 "WHERE applicationnumber = '" & strApplicationNumber & "'" & _
                 " AND DocumentGUID = '" & strDocumentGuid & "'" & _
                 " ORDER BY EventDate"
    Else
        strSQL = "SELECT * " & _
                 "FROM EventAuditDetail " & _
                 "WHERE applicationnumber = '" & strApplicationNumber & "'" & _
                 " AND DocumentGUID = 0x" & strDocumentGuid & _
                 " ORDER BY EventDate"
    End If
    Debug.Print strSQL
    ' Open a database connection
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.open
    
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = strSQL
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "adoCommand.Execute call"
    Set adoRecordSet = adoCommand.Execute
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "adoCommand.Execute return"
    ' Now check to see if we have any records returned......
    If adoRecordSet.EOF Then
        errThrowError cstrFunctionName, oeRecordNotFound
    Else
        adoRecordSet.MoveFirst
    End If
    ' And if we have let's start at the beginning and add them all to the XML response
    While Not adoRecordSet.EOF
        ' Get the values out of the recordset
        strApplicationNumber = adoRecordSet.fields("ApplicationNumber").Value
        strDocumentVersion = adoRecordSet.fields("FileVersion").Value
        strDateEvent = adoRecordSet.fields("EVENTDATE").Value
        ' GUID is a byte array and needs to be converted so we can put it into a string
        strFileGUID = GuidToString(adoRecordSet.fields("FileGUID").Value)
        Select Case adoRecordSet.fields("EventKey").Value
            Case EVENTKEY_CREATED
                strEvent = "CREATED"
            Case EVENTKEY_EDITED
                strEvent = "EDITED"
            Case EVENTKEY_VIEWED
                strEvent = "VIEWED"
            Case EVENTKEY_REPRINTED
                strEvent = "REPRINTED"
'TW 25/10/2005 CORE210/MAR294
            Case EVENTKEY_FULFILLMENTSEND
                strEvent = "FulfilmentSend"
            Case EVENTKEY_FULFILLMENTRESEND
                strEvent = "FulfilmentReSend"
            Case EVENTKEY_FULFILLMENTCANCEL
                strEvent = "FulfilmentCancel"
            Case EVENTKEY_SMS
                strEvent = "SMS"
            Case EVENTKEY_EMAIL
                strEvent = "Email"
'TW 25/10/2005 CORE210/MAR294 End
            'CORE252 GHun
            Case EVENTKEY_RECATEGORISATION
                strEvent = "Recategorisation"
            'CORE252 End
            'AS 15/12/2006 CORE324 Start
            Case EVENTKEY_NOTAPPROVED
                strEvent = "Not approved"
            Case EVENTKEY_APPROVED
                strEvent = "Approved"
            Case EVENTKEY_GEMINIPRINTED
                strEvent = "Gemini printed"
            'AS 15/12/2006 CORE324 End
        End Select
                    
        strUsername = adoRecordSet.fields("UserID").Value
        strUnitname = adoRecordSet.fields("UnitID").Value
        strPrintLocation = adoRecordSet.fields("PrintLocation").Value
        strNumCopies = adoRecordSet.fields("NumberOfCopies").Value
        ' Create a top level element
        Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("DOCEVENTLIST")
        ' And all all the values to the new element
        xmlSetAttributeValue xmlNewChildNode, "APPLICATIONNUMBER", strApplicationNumber
        Set xmlDocDetailNode = vxmlResponseNode.ownerDocument.createElement("DOCUMENTDETAILS")
        xmlSetAttributeValue xmlDocDetailNode, "DOCUMENTVERSION", strDocumentVersion
        xmlSetAttributeValue xmlDocDetailNode, "EVENTTIMESTAMP", strDateEvent
        xmlSetAttributeValue xmlDocDetailNode, "FILEGUID", strFileGUID
        xmlSetAttributeValue xmlDocDetailNode, "EVENT", strEvent
        xmlSetAttributeValue xmlDocDetailNode, "USERNAME", strUsername
        xmlSetAttributeValue xmlDocDetailNode, "UNITNAME", strUnitname
        xmlSetAttributeValue xmlDocDetailNode, "PRINTLOCATION", strPrintLocation
        xmlSetAttributeValue xmlDocDetailNode, "NUMCOPIES", strNumCopies
            
        xmlNewChildNode.appendChild xmlDocDetailNode
        ' And add the node to the response
        vxmlResponseNode.appendChild xmlNewChildNode
        ' next record...
        adoRecordSet.MoveNext
    Wend
    ' Close down
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    ' if we get here, everything is OK
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
GetDocumentEventList_Exit:
    ' And tell MTS we're done!!!
    ' ik_20030211
    ' gobjContext.SetComplete
    Set xmlNewChildNode = Nothing
    Set xmlDocDetailNode = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CreateEventDetailRecord(ByVal vxmlRequestNode As IXMLDOMNode)
Const cstrFunctionName As String = "CreateEventDetailRecord"

Dim adoCommand As New ADODB.Command
Dim adoConnection As New ADODB.Connection
Dim adoRecordSet As Recordset

Dim nNumRecords As Long

Dim strEventDate As String
Dim strEventGUID As String
Dim strFileVersion As String
Dim strSQL As String

Dim xmlNode As IXMLDOMNode
    
    On Error GoTo CreateEventDetailRecordExit
    ' ik_20030211
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & "() Called. With vxmlRequestNode.xml: " & vxmlRequestNode.xml
    Set xmlNode = xmlGetNode(vxmlRequestNode, "EVENTDETAIL")
    
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.open
        
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    ' If the event was a reprint then we have to find the previous print location...
    ' ik_bmids00885
    ' not if we have CONTROLDATA on REQUEST
    If vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "REPRINTDOCUMENT" Then
        
        xmlSetAttributeValue _
            xmlNode, _
            "PRINTLOCATION", _
            vxmlRequestNode.selectSingleNode("CONTROLDATA/OUTPUTTYPE/PRINTER/PRINTERNAME").Text
        ' ok, a bit naughty
        xmlSetAttributeValue xmlNode, "NUMCOPIES", "1"
        xmlSetAttributeValue xmlNode, "PRINTERTYPE", "PRINTER"
    Else
'TW 25/10/2005 CORE210/MAR294
'        If (xmlGetAttributeText(xmlNode, "EVENTKEY") = "3") Or (xmlGetAttributeText(xmlNode, "EVENTKEY") = "2") Then
        Select Case CSafeInt(xmlGetAttributeText(xmlNode, "EVENTKEY"))
            Case _
                EVENTKEY_VIEWED, _
                EVENTKEY_REPRINTED, _
                EVENTKEY_FULFILLMENTSEND, _
                EVENTKEY_FULFILLMENTRESEND, _
                EVENTKEY_FULFILLMENTCANCEL, _
                EVENTKEY_SMS, _
                EVENTKEY_EMAIL, _
                EVENTKEY_RECATEGORISATION, _
                EVENTKEY_NOTAPPROVED, _
                EVENTKEY_APPROVED, _
                EVENTKEY_GEMINIPRINTED

'TW 25/10/2005 CORE210/MAR294 End
                If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                    strSQL = "SELECT * FROM EventAuditDetail WHERE " & _
                             "APPLICATIONNUMBER = '" & xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER") & "'" & _
                             " AND DOCUMENTGUID = '" & xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & "'" & _
                             " AND FILEVERSION = '" & xmlGetAttributeText(xmlNode, "DOCUMENTVERSION") & "'" & _
                             " AND FILEGUID = '" & xmlGetAttributeText(xmlNode, "FILEGUID") & "'" & _
                             " ORDER BY EVENTDATE"
                Else
                    strSQL = "SELECT * FROM EventAuditDetail WHERE " & _
                             "APPLICATIONNUMBER = '" & xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER") & "'" & _
                             " AND DOCUMENTGUID = 0x" & xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & _
                             " AND FILEVERSION = '" & xmlGetAttributeText(xmlNode, "DOCUMENTVERSION") & "'" & _
                             " AND (FILEGUID = 0x" & xmlGetAttributeText(xmlNode, "FILEGUID") & " OR FILEGUID = 0x00000000000000000000000000000000)" & _
                             " ORDER BY EVENTDATE"
                End If
                Debug.Print "Selecting SQL from 'EventAuditDetail' : " & strSQL
                adoCommand.CommandText = strSQL
                ' And a recordset to return all our entries...
                Set adoRecordSet = adoCommand.Execute
                ' Go to the 1st record and grab all the printer values and set them in the node for the insert
                adoRecordSet.MoveFirst
                
                'CORE252 GHun Set default values if the existing values are Null
                If IsNull(adoRecordSet.fields("PRINTLOCATION").Value) Then
                    xmlSetAttributeValue xmlNode, "PRINTLOCATION", ""
                Else
                    xmlSetAttributeValue xmlNode, "PRINTLOCATION", adoRecordSet.fields("PRINTLOCATION").Value
                End If
                
                If IsNull(adoRecordSet.fields("NUMBEROFCOPIES").Value) Then
                    xmlSetAttributeValue xmlNode, "NUMCOPIES", "0"
                Else
                    xmlSetAttributeValue xmlNode, "NUMCOPIES", adoRecordSet.fields("NUMBEROFCOPIES").Value
                End If
                
                If IsNull(adoRecordSet.fields("PRINTERTYPE").Value) Then
                    xmlSetAttributeValue xmlNode, "PRINTERTYPE", ""
                Else
                    xmlSetAttributeValue xmlNode, "PRINTERTYPE", adoRecordSet.fields("PRINTERTYPE").Value
                End If
                'CORE252 End
                
                Set adoRecordSet = Nothing
'TW 25/10/2005 CORE210/MAR294
'        End If
        End Select
'TW 25/10/2005 CORE210/MAR294 End
    End If
    ' ik_bmids00885_ends
    ' Mike - BUG found 10/9/2001 - we need to also have a numeric version filed as we found that oracle thinks V9 is greater
    ' than V10 - so we need to cut off the V and then store that number into the "FILEVERSIONNUM" filed of the database.
    ' NOTE - this will cause problems later on when we look at branching...
    
    strFileVersion = Mid$(xmlGetAttributeText(xmlNode, "DOCUMENTVERSION"), 2)
    
    ' And we have a new GUID for each new event so create it now...
    strEventGUID = CreateGUID()
    
    strEventDate = Format$(Now, "dd mmm yyyy hh:mm:ss")
    
'TW 24/10/2005 CORE208/MAR285
    strSQL = "INSERT INTO EventAuditDetail (APPLICATIONNUMBER, DOCUMENTGUID, FILEVERSION, EVENTGUID, EVENTDATE, FILEGUID, " & _
             "EVENTKEY, UNITID, USERID, PRINTLOCATION, NUMBEROFCOPIES, PRINTERTYPE, FILEVERSIONNUM"
    If xmlGetAttributeText(xmlNode, "PACKFULFILLMENTGUID") <> "" Then
        strSQL = strSQL & ", PACKFULFILLMENTGUID, HOSTTEMPLATEID, FILENETIMAGEREF"
    End If
    strSQL = strSQL & ") Values ("
'TW 24/10/2005 CORE208/MAR285 End
    
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = strSQL & "'" & xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER") & "', " & _
                 "'" & xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "DOCUMENTVERSION") & "', " & _
                 "'" & strEventGUID & "', " & _
                 "TO_DATE('" & strEventDate & "', 'DD-MON-YYYY HH24:MI:SS'), " & _
                 "'" & xmlGetAttributeText(xmlNode, "FILEGUID") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "EVENTKEY") & "', " & _
                 "'" & xmlGetAttributeText(vxmlRequestNode, "UNITID") & "', " & _
                 "'" & xmlGetAttributeText(vxmlRequestNode, "USERID") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "PRINTLOCATION") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "NUMCOPIES") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "PRINTERTYPE") & "', " & _
                 "'" & strFileVersion & "'"
'TW 24/10/2005 CORE208/MAR285
        If xmlGetAttributeText(xmlNode, "PACKFULFILLMENTGUID") <> "" Then
            strSQL = strSQL & ", '" & xmlGetAttributeText(xmlNode, "PACKFULFILLMENTGUID") & "'"
            strSQL = strSQL & ", '" & xmlGetAttributeText(xmlNode, "HOSTTEMPLATEID") & "'"
            If Len(xmlGetAttributeText(xmlNode, "FILENETIMAGEREF")) = 32 Then
                strSQL = strSQL & ", " & xmlGetAttributeText(xmlNode, "FILENETIMAGEREF")
            Else
                strSQL = strSQL & ", null"
            End If
        End If
'TW 24/10/2005 CORE208/MAR285 End
    Else
        ' PSC 03/02/2006 CORE239
        strSQL = strSQL & "'" & xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER") & "', " & _
                 "0x" & xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & ", " & _
                 "'" & xmlGetAttributeText(xmlNode, "DOCUMENTVERSION") & "', " & _
                 "0x" & strEventGUID & ", " & _
                 "GetDate()" & ", " & _
                 "0x" & xmlGetAttributeText(xmlNode, "FILEGUID") & ", " & _
                 "'" & xmlGetAttributeText(xmlNode, "EVENTKEY") & "', " & _
                 "'" & xmlGetAttributeText(vxmlRequestNode, "UNITID") & "', " & _
                 "'" & xmlGetAttributeText(vxmlRequestNode, "USERID") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "PRINTLOCATION") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "NUMCOPIES") & "', " & _
                 "'" & xmlGetAttributeText(xmlNode, "PRINTERTYPE") & "', " & _
                 "'" & strFileVersion & "'"
'TW 24/10/2005 CORE208/MAR285
        If xmlGetAttributeText(xmlNode, "PACKFULFILLMENTGUID") <> "" Then
            strSQL = strSQL & ", 0x" & xmlGetAttributeText(xmlNode, "PACKFULFILLMENTGUID")
            strSQL = strSQL & ", '" & xmlGetAttributeText(xmlNode, "HOSTTEMPLATEID") & "'"
            If Len(xmlGetAttributeText(xmlNode, "FILENETIMAGEREF")) = 32 Then
                strSQL = strSQL & ", 0x" & xmlGetAttributeText(xmlNode, "FILENETIMAGEREF")
            Else
                strSQL = strSQL & ", null"
            End If
        End If
'TW 24/10/2005 CORE208/MAR285 End
    End If
    strSQL = strSQL & ")"
'"'" & xmlGetAttributeText(xmlNode, "TEMPLATEID") & "', " & _
    Debug.Print "Instering SQL into 'EventAuditDetail' : " & strSQL
    adoCommand.CommandText = strSQL
    ' Now we have all the parameters set up, execute the command
    ' (DON'T need a recordset here as we're not returning anything!)
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "adoCommand.Execute call"
    adoCommand.Execute nNumRecords, , adExecuteNoRecords
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "adoCommand.Execute return"
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
CreateEventDetailRecordExit:
    Set xmlNode = Nothing
    Set adoRecordSet = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    
    ' ik_20030211
    ' gobjContext.SetComplete
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub UpdateDocumentDetailRecord(ByVal xmlRequestNode As IXMLDOMNode)
    Const cstrFunctionName As String = "UpdateDocumentDetailRecord"
    
    Dim adoCommand As ADODB.Command
    Dim adoConnection As ADODB.Connection
    Dim adoRecordSet As ADODB.Recordset
    
    Dim strSQL As String
    
    Dim xmlNode As IXMLDOMNode
    
    On Error GoTo UpdateDocumentDetailRecordExit
    ' ik_20030211
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & "() Called. With xmlRequestNode.xml: " & xmlRequestNode.xml
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.open
    ' Need to find the correct record to update
    Set xmlNode = xmlGetNode(xmlRequestNode, "APPLICATIONDETAIL")
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "SELECT * FROM DocumentAuditDetails WHERE APPLICATIONNUMBER = '" & xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER") & "'" & _
                 " AND DOCUMENTGUID = '" & xmlGetAttributeText(xmlRequestNode, "DOCUMENTGUID") & "'"
    Else
        strSQL = "SELECT * FROM DocumentAuditDetails WHERE APPLICATIONNUMBER = '" & xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER") & "'" & _
                 " AND DOCUMENTGUID = 0x" & xmlGetAttributeText(xmlRequestNode, "DOCUMENTGUID")
    End If
    
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = strSQL
        
    ' Now we have all the parameters set up, execute the command
    Set adoRecordSet = adoCommand.Execute
    ' We have the recordset
    ' Now we need to update the values....
    strSQL = "UPDATE DocumentAuditDetails SET "
    ' Now we need to cpmpare the record entries with the XML values and update any changes with the XML value
    'Set xmlNode = xmlGetNode(xmlRequestNode, "APPLICATIONDETAIL")
    If Not xmlNode Is Nothing Then
        If xmlAttributeValueExists(xmlNode, "TEMPLATEID") Then
            If xmlGetAttributeText(xmlNode, "TEMPLATEID") <> adoRecordSet.fields("TEMPLATEID") Then
                ' PSC 10/11/2005 CORE196 - Start
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "TEMPLATEID = '" & xmlGetAttributeText(xmlNode, "TEMPLATEID") & "'"
                ' PSC 10/11/2005 CORE196 - End
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "HOSTTEMPLATEID") Then
            If xmlGetAttributeText(xmlNode, "HOSTTEMPLATEID") <> adoRecordSet.fields("HOSTTEMPLATEID") Then
                ' PSC 10/11/2005 CORE196 - Start
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "HOSTTEMPLATEID = '" & xmlGetAttributeText(xmlNode, "HOSTTEMPLATEID") & "'"
                ' PSC 10/11/2005 CORE196 - End
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "SEARCHKEY1") Then
            If xmlGetAttributeText(xmlNode, "SEARCHKEY1") <> adoRecordSet.fields("SEARCHKEY1") Then
                ' PSC 10/11/2005 CORE196 - Start
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "SEARCHKEY1 = '" & xmlGetAttributeText(xmlNode, "SEARCHKEY1") & "'"
                ' PSC 10/11/2005 CORE196 - End
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "SEARCHKEY2") Then
            If xmlGetAttributeText(xmlNode, "SEARCHKEY2") <> adoRecordSet.fields("SEARCHKEY2") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "SEARCHKEY2 = '" & xmlGetAttributeText(xmlNode, "SEARCHKEY2") & "'"
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "SEARCHKEY3") Then
            If xmlGetAttributeText(xmlNode, "SEARCHKEY3") <> adoRecordSet.fields("SEARCHKEY3") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "SEARCHKEY3 = '" & xmlGetAttributeText(xmlNode, "SEARCHKEY3") & "'"
            End If
        End If
    End If    ' xmlNode is NOT Nothing
    Set xmlNode = xmlGetNode(xmlRequestNode, "DOCUMENTDETAILS")
    If Not xmlNode Is Nothing Then
        If xmlAttributeValueExists(xmlNode, "STAGEID") Then
            If xmlGetAttributeText(xmlNode, "STAGEID") <> adoRecordSet.fields("STAGEID") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "STAGEID = '" & xmlGetAttributeText(xmlNode, "STAGEID") & "'"
            End If
        End If
        'CORE252 GHun
        If xmlAttributeValueExists(xmlNode, "DOCUMENTPURPOSE") Then
            If (Len(xmlGetAttributeText(xmlNode, "DOCUMENTPURPOSE")) > 0 And IsNull(adoRecordSet.fields("DOCUMENTPURPOSE"))) _
                Or (xmlGetAttributeText(xmlNode, "DOCUMENTPURPOSE") <> adoRecordSet.fields("DOCUMENTPURPOSE")) Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "DOCUMENTPURPOSE = '" & xmlGetAttributeText(xmlNode, "DOCUMENTPURPOSE") & "'"
            End If
        End If
        'CORE252 End
        If xmlAttributeValueExists(xmlNode, "DOCUMENTGROUP") Then
            'CORE252 GHun
            If (Len(xmlGetAttributeText(xmlNode, "DOCUMENTGROUP")) > 0 And IsNull(adoRecordSet.fields("DOCUMENTGROUP"))) _
                Or (xmlGetAttributeText(xmlNode, "DOCUMENTGROUP") <> adoRecordSet.fields("DOCUMENTGROUP")) Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "DOCUMENTGROUP = '" & xmlGetAttributeText(xmlNode, "DOCUMENTGROUP") & "'"
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "DOCUMENTNAME") Then
            If xmlGetAttributeText(xmlNode, "DOCUMENTNAME") <> adoRecordSet.fields("DOCUMENTNAME") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "DOCUMENTNAME = '" & xmlGetAttributeText(xmlNode, "DOCUMENTNAME") & "'"
            End If
        End If
        'DR Required for email reprints
        If xmlAttributeValueExists(xmlNode, "DOCUMENTDESCRIPTION") Then
            If xmlGetAttributeText(xmlNode, "DOCUMENTDESCRIPTION") <> adoRecordSet.fields("DOCUMENTDESCRIPTION") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "DOCUMENTDESCRIPTION = '" & xmlGetAttributeText(xmlNode, "DOCUMENTDESCRIPTION") & "'"
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "CUSTOMERNAME") Then
            If xmlGetAttributeText(xmlNode, "CUSTOMERNAME") <> adoRecordSet.fields("CUSTOMERNAME") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "CUSTOMERNAME = '" & xmlGetAttributeText(xmlNode, "CUSTOMERNAME") & "'"
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "RECIPIENTNAME") Then
            If xmlGetAttributeText(xmlNode, "RECIPIENTNAME") <> adoRecordSet.fields("RECIPIENTNAME") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "RECIPIENTNAME = '" & xmlGetAttributeText(xmlNode, "RECIPIENTNAME") & "'"
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "LANGUAGE") Then
            If xmlGetAttributeText(xmlNode, "LANGUAGE") <> adoRecordSet.fields("LANGUAGE") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "LANGUAGE = " & xmlGetAttributeText(xmlNode, "LANGUAGE")
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "SOURCESYSTEM") Then
            If xmlGetAttributeText(xmlNode, "SOURCESYSTEM") <> adoRecordSet.fields("SOURCESYSTEM") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                ' PSC 10/11/2005 CORE196
                strSQL = strSQL & "SOURCESYSTEM = '" & xmlGetAttributeText(xmlNode, "SOURCESYSTEM") & "'"
            End If
        End If
        ' Mike 23/7/2001 - Don't think we should EVER update a creation date.
    '    If xmlGetAttributeText(xmlNode, "CREATIONTIMESTAMP") <> adoRecordSet.Fields("CREATIONDATE") Then
    '        If Right$(strSQL, 4) <> "SET " Then
    '            strSQL = strSQL & ", "
    '        End If
    '
    '        strSQL = strSQL & "CREATIONTIMESTAMP = " & xmlGetAttributeText(xmlNode, "CREATIONDATE")
    '    End If
        
        If xmlAttributeValueExists(xmlNode, "ARCHIVEDATE") Then
            If xmlGetAttributeText(xmlNode, "ARCHIVEDATE") <> adoRecordSet.fields("ARCHIVEDATE") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                Dim strArchiveDate As String
                strArchiveDate = xmlGetAttributeText(xmlNode, "ARCHIVEDATE")
                If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                    strSQL = strSQL & "ARCHIVEDATE = TO_DATE('" & strArchiveDate & "', 'DD-MON-YYYY HH24:MI:SS')"
                Else
                    strSQL = strSQL & "ARCHIVEDATE = " & sqlFormatDate(strArchiveDate)
                End If
            End If
        End If
        If xmlAttributeValueExists(xmlNode, "PRINTDATE") Then
            If xmlGetAttributeText(xmlNode, "PRINTTIMESTAMP") <> adoRecordSet.fields("PRINTDATE") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "PRINTTIMESTAMP = '" & xmlGetAttributeText(xmlNode, "PRINTDATE") & "'"
            End If
        End If
        ' AS 15/12/2006 CORE324 Add support for DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS.
        If xmlAttributeValueExists(xmlNode, "GEMINIPRINTSTATUS") Then
            If IsNull(adoRecordSet.fields("GEMINIPRINTSTATUS")) Or xmlGetAttributeText(xmlNode, "GEMINIPRINTSTATUS") <> adoRecordSet.fields("GEMINIPRINTSTATUS") Then
                If Right$(strSQL, 4) <> "SET " Then
                    strSQL = strSQL & ", "
                End If
                strSQL = strSQL & "GEMINIPRINTSTATUS = " & xmlGetAttributeText(xmlNode, "GEMINIPRINTSTATUS")
            End If
        End If
    End If   ' xmlNode is not Nothing
    ' And add the condition to the SQL statement
    ' Make sure that some of the fields have actually changed.  If the have we set strSQL to NULL so that we don't actually
    ' do anything!
    If Right$(strSQL, 4) = "SET " Then
        strSQL = ""
    Else
        If adoGetDbProvider = omiga4DBPROVIDEROracle Then
            strSQL = strSQL & " WHERE APPLICATIONNUMBER = '" & xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER") & "' AND DOCUMENTGUID = '" & xmlGetAttributeText(xmlRequestNode, "DOCUMENTGUID") & "'"
        Else
            strSQL = strSQL & " WHERE APPLICATIONNUMBER = '" & xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER") & "' AND DOCUMENTGUID = 0x" & xmlGetAttributeText(xmlRequestNode, "DOCUMENTGUID")
        End If
    End If
    Debug.Print "Updating Document Detail with the SQL : " & strSQL
'    Set adoCommand.ActiveConnection = adoConnection
'    adoCommand.CommandType = adCmdText
'    adoCommand.CommandText = strSQL
        
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & " strSQL: " & strSQL
        
        
    ' Now we have all the parameters set up, execute the command
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "adoCommand.Execute call"
    ' PSC 10/11/2005 CORE196 - Start
    If Len(strSQL) > 0 Then     'CORE252 GHun
        adoCommand.CommandText = strSQL
        adoCommand.Execute , , adExecuteNoRecords
    End If
    ' PSC 10/11/2005 CORE196 - End
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "adoCommand.Execute return"
        
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
UpdateDocumentDetailRecordExit:
    Set adoCommand = Nothing
    Set adoRecordSet = Nothing
    Set adoConnection = Nothing
    Set xmlNode = Nothing
    
    ' ik_20030211
    ' gobjContext.SetComplete
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CreateDocumentDetailRecord(ByVal vxmlRequestNode As IXMLDOMNode)

Dim adoCommand As New ADODB.Command
Dim adoConnection As New ADODB.Connection

Dim dteArchive As Date
Dim dteCreate As Date
Dim dtePrint As Date

Dim lNumRecords As Long

Dim xmlNode As IXMLDOMNode
Dim xmlNode2 As IXMLDOMNode

Dim strArchiveDate As String
Dim strCreationDate As String
Dim strDocumentGuid As String
Dim strFreeFormat As String
Dim strPrintDate As String
Dim strSQL As String
Dim columns As String
Dim values As String
    
    On Error GoTo CreateDocumentDetailRecordExit
    ' ik_20030211
    gobjContext.SetComplete
    Const cstrFunctionName As String = "CreateDocumentDetailRecord"
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & "() Called. With vxmlRequestNode.xml: " & vxmlRequestNode.xml
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.open
    ' Must get node here
    Set xmlNode = xmlGetNode(vxmlRequestNode, "APPLICATIONDETAIL")
    Set xmlNode2 = xmlGetNode(vxmlRequestNode, "DOCUMENTDETAILS")
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & "() Event 1 xmlNode2.xml: " & xmlNode2.xml
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & "() Event 1 xmlGetAttributeText(xmlNode2, ""DOCUMENTDESCRIPTION""): " & xmlGetAttributeText(xmlNode2, "DOCUMENTDESCRIPTION")
    ' Create a new GUID for the document
    strDocumentGuid = CreateGUID()
    ' Sort the dates out into the correct format...
'    strCreationDate = Format$(xmlGetAttributeText(xmlNode2, "PRINTTIMESTAMP"), "dd/mm/yyyy hh:mm:ss")
    strCreationDate = Format$(Now, "dd mmm yyyy hh:mm:ss")
    strFreeFormat = xmlGetAttributeText(vxmlRequestNode, "FREEFORMAT")
    strArchiveDate = Format$(xmlGetAttributeText(xmlNode2, "ARCHIVETIMESTAMP"), "dd mmm yyyy hh:mm:ss")
    strPrintDate = Format$(xmlGetAttributeText(xmlNode2, "PRINTTIMESTAMP"), "dd mmm yyyy hh:mm:ss")
    If xmlGetAttributeText(xmlNode2, "CUSTOMERNO") = "" Then
        xmlSetAttributeValue xmlNode2, "CUSTOMERNO", "None"
    End If
    'DR Do the same for reciepient value (wondered where this came from on the GUI)
    If xmlGetAttributeText(xmlNode2, "RECIPIENTNAME") = "" Then
        xmlSetAttributeValue xmlNode2, "RECIPIENTNAME", "None"
    End If
    ' AS 04/05/2005 CORE111 ensure DELIVERYENGINETYPE and DELIVERYTYPE are set in the record.
    If xmlGetAttributeText(xmlNode2, "DELIVERYENGINETYPE") = "" Then
        xmlSetAttributeValue xmlNode2, "DELIVERYENGINETYPE", "null"
    End If
    If xmlGetAttributeText(xmlNode2, "DELIVERYTYPE") = "" Then
        xmlSetAttributeValue xmlNode2, "DELIVERYTYPE", "null"
    End If
    
    columns = _
        "APPLICATIONNUMBER, DOCUMENTGUID, TEMPLATEID, HOSTTEMPLATEID, SEARCHKEY1, STAGEID, " & _
        "DOCUMENTGROUP, DOCUMENTNAME, DOCUMENTDESCRIPTION, CUSTOMERNAME, RECIPIENTNAME, CREATIONDATE, ARCHIVEDATE, PRINTDATE, " & _
        "LANGUAGE, SOURCESYSTEM, DELIVERYENIGNETYPE, DELIVERYTYPE, SEARCHKEY2, SEARCHKEY3"
        
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        values = _
             "'" & xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "DOCUMENTGUID") & "', " & _
             "'" & xmlGetAttributeText(xmlNode, "TEMPLATEID") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "HOSTTEMPLATEID") & "', " & _
             "'" & xmlGetAttributeText(xmlNode, "SEARCHKEY1") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "STAGEID") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "DOCUMENTGROUP") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "DOCUMENTNAME") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "DOCUMENTDESCRIPTION") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "CUSTOMERNO") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "RECIPIENTNAME") & "', " & _
             "TO_DATE ('" & strCreationDate & "', 'DD-MON-YYYY HH24:MI:SS'), " & _
             "TO_DATE ('" & strArchiveDate & "', 'DD-MON-YYYY HH24:MI:SS'), " & _
             "TO_DATE ('" & strPrintDate & "', 'DD-MON-YYYY HH24:MI:SS'), " & _
             "'" & xmlGetAttributeText(xmlNode2, "LANGUAGE") & "', " & _
             "'" & xmlGetAttributeText(xmlNode2, "SOURCESYSTEM") & "', " & _
             xmlGetAttributeText(xmlNode2, "DELIVERYENGINETYPE") & ", " & _
             xmlGetAttributeText(xmlNode2, "DELIVERYTYPE")
    Else
        values = _
            FormatString(xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")) & ", " & _
            "0x" & xmlGetAttributeText(xmlNode2, "DOCUMENTGUID") & ", " & _
            FormatString(xmlGetAttributeText(xmlNode, "TEMPLATEID")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "HOSTTEMPLATEID")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode, "SEARCHKEY1")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "STAGEID")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "DOCUMENTGROUP")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "DOCUMENTNAME")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "DOCUMENTDESCRIPTION")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "CUSTOMERNO")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "RECIPIENTNAME")) & ", " & _
            sqlFormatDate(strCreationDate) & ", " & _
            sqlFormatDate(strArchiveDate) & ", " & _
            sqlFormatDate(strPrintDate) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "LANGUAGE")) & ", " & _
            FormatString(xmlGetAttributeText(xmlNode2, "SOURCESYSTEM")) & ", " & _
            xmlGetAttributeText(xmlNode2, "DELIVERYENGINETYPE") & ", " & _
            xmlGetAttributeText(xmlNode2, "DELIVERYTYPE")
    End If
    
    ' Add Search keys 2 and 3 (unless they don't exist, in which case put nothing there!)
    If xmlAttributeValueExists(xmlNode, "SEARCHKEY2") Then
        values = values & ", " & FormatString(xmlGetAttributeText(xmlNode, "SEARCHKEY2"))
    Else
        values = values & ", ''"
    End If
            
    If xmlAttributeValueExists(xmlNode, "SEARCHKEY3") Then
        values = values & ", " & FormatString(xmlGetAttributeText(xmlNode, "SEARCHKEY3"))
    Else
        values = values & ", ''"
    End If
            
    'AS 15/12/2006 CORE324 Add support for DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS.
    If xmlAttributeValueExists(xmlNode2, "GEMINIPRINTSTATUS") Then
        ' Only insert into GEMINIPRINTSTATUS column if it exists in the request Xml.
        columns = columns & ", GEMINIPRINTSTATUS"
        values = values & ", " & xmlGetAttributeText(xmlNode2, "GEMINIPRINTSTATUS")
    End If

    strSQL = "INSERT INTO DOCUMENTAUDITDETAILS (" & columns & ") VALUES (" & values & ")"
            
    ' TEST display of SQL query statement
    Debug.Print "Inserting SQL into DocumentAuditDetails : " & strSQL & "..."
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = strSQL
        
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & " strSQL: " & strSQL
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "adoCommand.Execute call"
    adoCommand.Execute lNumRecords, , adExecuteNoRecords
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "adoCommand.Execute return"
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
CreateDocumentDetailRecordExit:
    Set xmlNode = Nothing
    Set xmlNode2 = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    
    ' ik_20030211
    ' gobjContext.SetComplete
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
' header ----------------------------------------------------------------------------------
' description:
'   Takes a byte array GUID and converts it to a String
' pass:
'   bytArray
'       array of bytes containing the GUID
' return:
'   String - the GUID in string format
'------------------------------------------------------------------------------------------
    
Private Function GuidToString(bytArray() As Byte) As String
    Dim i As Integer
    Dim strGuid As String
    strGuid = ""
    For i = 0 To 15
        strGuid = strGuid & Right$(Hex$(bytArray(i) + &H100), 2)
    Next i
    GuidToString = strGuid
End Function
'DR DMSSYS0005
Private Function GetDbProvider()
    GetDbProvider = adoGetDbProvider()
End Function
'DR DMSSYS0005
Private Function sqlFormatDate(ByVal vdtmIn As Date, _
    Optional ByVal vdtfDateTimeFormat As DATETIMEFORMAT = dtfDateTime) As String
' header ----------------------------------------------------------------------------------
' description:
'   Formats a date to the appropriate SQL.
' pass:
'   vdtmIn                  date to be formatted
'   vdtfDateTimeFormat      required format
' return:
'   FormatDate              formatted value
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim cstrFunctionName As String
    cstrFunctionName = "sqlFormatDate"
    Dim strDate As String
    Dim eDbEngineType As DBPROVIDER
    eDbEngineType = GetDbProvider
    Select Case vdtfDateTimeFormat
    Case dtfDateTime                ' date and time format
        strDate = Format(vdtmIn, "YYYY-MM-DD HH:MM:SS")
        Select Case eDbEngineType
        Case omiga4DBPROVIDERSQLServer
            'DR DMSSYS0005 Changed from 103 to 102
            strDate = "CONVERT(DATETIME,'" & strDate & "',102)"
        Case omiga4DBPROVIDEROracle
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
            ' not implemented
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist].", cstrFunctionName
        End Select
    Case dtfDate                    ' date only
        strDate = Format(vdtmIn, "YYYY-MM-DD")
        Select Case eDbEngineType
        Case omiga4DBPROVIDERSQLServer
            ' fixme
            ' not implemented
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist].", cstrFunctionName
        Case omiga4DBPROVIDEROracle
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD')"
        End Select
    Case Else
        ' not implemented
        errRaiseError _
            omiga4NotImplemented, _
            "[sqlAssist].", cstrFunctionName
    End Select
    'DR DMSSYS0005 - removed the equals sign to get it to work
    sqlFormatDate = strDate
End Function
'BM0508 GHun 09/04/2003
Private Function FormatString(ByVal vstrIn As String) As String
    FormatString = "'" & Replace(vstrIn, "'", "''") & "'"
End Function
'BM0508 End
