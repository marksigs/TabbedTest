VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "FileVersioningDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'header ----------------------------------------------------------------------------------
'Workfile:      FileVersioningDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Helper functions for error handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog   Date        Description
'LD     26/01/2001  Initial creation
'LD     11/06/2001  SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     27/09/2001  SYS2722 SQL Server Port of FVS
'DM     16/10/2001  SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
'DR     20/03/2002  Adding Compress/uncompress functionality
'IK     04/02/2003  BM0269, improved error handling
'IK     17/02/2003  BM0200 - add TraceAssist support
'IK     05/02/2003  BM0261, more on error handling
'IK     20/02/2003  BM0231, fix error re: filesize > numeric(5)
'AS     06/07/2004  BBG942 Added support for ZLIB compression.
'AS     26/07/2004  BBG1055 Added support for DELIVERYTYPE (Word/PDF).
'AS     20/04/2005  CORE106 Ensure file contents are only compressed once.
'AS     23/05/2005  CORE125 omFVS should be able to retrieve archived documents in compressed format.
'AS     27/05/2005  CORE141 ReadFileContents uses objCompression after it has been set to nothing
'IK     20/04/2006  CORE261 omGemini integration
'AS     28/04/2006  CORE261 Add support for Gemini.
'AS     26/07/2006  CORE286 Get file from archive as a Url and not as bin base 64 string.
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IFileVersioningDO
Private gobjContext As ObjectContext
        
' ik_BM0261
' Private Const CstrClassName = "omDPS::PrintHandlerBO"
Private Const CstrClassName = "omFVS::FileVersioningDO"

'IK_20/04/2006_CORE261
Private Const cstrGeminiQueueName = "GeminiInterface"
Private Const cstrGeminiComponentName = "omGemini.FileVersioningBO"
'IK_20/04/2006_CORE261_ends

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Public Sub IFileVersioningDO_CheckinFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CheckinFileVbErr
    Dim xmlNode As IXMLDOMNode
    Const cstrFunctionName As String = "CheckinFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
        
    ' ik_BM0261
    gobjContext.SetComplete
    Set xmlNode = xmlGetNode(vxmlRequestNode, "CHECKINDATA")
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
'BG 01/10/02 SYS5619 - Compression Code - commented out code
'    ' Need to write the file contents back to the record
'    ' This needs a new GUID.....
'    Dim strFileContentsGuid As String
'    strFileContentsGuid = CreateGUID()
''
'    WriteFileContents adoConnection, strFileContentsGuid, xmlNode
'BG 01/10/02 SYS5619 - Compression Code - END
    ' And check in the file (This takes care of updating the new version)
    ' Need to create some parameters...
    ' Who created the checkin?
    Dim strCreatedBy As String
    strCreatedBy = xmlGetMandatoryAttributeText(xmlNode, "CREATEDBY")
        
    Dim strFileGuid As String
    strFileGuid = xmlGetMandatoryAttributeText(xmlNode, "FILEGUID")
    Dim strFileVersion As String
    strFileVersion = xmlGetMandatoryAttributeText(xmlNode, "FILEVERSION")
    'BG 01/10/02 SYS5619 - Compression Code
    ' Need to write the file contents back to the record
    ' This needs a new GUID.....
    Dim strFileContentsGuid As String
    strFileContentsGuid = CreateGUID()
    WriteFileContents strCreatedBy, adoConnection, strFileContentsGuid, xmlNode
    ' And check in the file (This takes care of updating the new version)
    'BG 01/10/02 SYS5619 - Compression Code - END
    Dim bMetaDataExists As Boolean
    Dim strMetaData As String
    bMetaDataExists = xmlAttributeValueExists(xmlNode, "METADATA")
    If bMetaDataExists = True Then
        strMetaData = xmlGetAttributeText(xmlNode, "METADATA")
    End If
    'BG 30/09/02 SYS5619 Compression code
    Dim lFileSize As Long
    lFileSize = xmlGetAttributeAsLong(xmlNode, "FILESIZE", 0)
    ' ik_bm0231
    lFileSize = 0
    'Dim strFileSize As String
    'strFileSize = xmlGetAttributeAsInteger(xmlNode, "FILESIZE", 0)
    'BG 30/09/02 SYS5619 Compression code - END
    Dim bPermissionsGUIDExists As Boolean
    Dim strPermissionsGUID As String
    bPermissionsGUIDExists = xmlAttributeValueExists(xmlNode, "PERMISSIONSGUID")
    If bPermissionsGUIDExists = True Then
        strPermissionsGUID = xmlGetAttributeText(xmlNode, "PERMISSIONSGUID")
    End If
    Dim strFileContents As String
    strFileContents = strFileContentsGuid 'xmlGetMandatoryAttributeText(xmlNode, "FILECONTENTS")
            
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    ' Create the SQL string for calling the SQL stored procedure
    'PROCEDURE CheckinFile(
    '    szCREATEDBY IN FVFILE.CREATEDBY%TYPE,
    '    rawFILEGUID IN FVFILE.FILEGUID%TYPE,
    '    szFILEVERSION IN FVFILE.FILEVERSION%TYPE,
    '    szMETADATA IN FVFILE.METADATA%TYPE,
    '    nFILESIZE IN FVFILE.FILESIZE%TYPE,
    '    rawPERMISSIONSGUID IN FVFILE.PERMISSIONSGUID%TYPE,
    '    rawFileContentsGuid IN RAW,
    '    tableszCHILDFILEVERSION OUT t_tableszFILEVERSION
    '    )
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVCHECKINFILE(?,?,?,?,?,?,?," _
            & "{resultset 1," _
            & "tableszCHILDFILEVERSION" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVCHECKINFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
        
    ' And set up all the lovely little parameters
    Dim adoParameter As New ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCreatedBy), strCreatedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        'BG 30/09/02 SYS5619 Compression code
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        'Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGUID), strFileGUID)
        'BG 30/09/02 SYS5619 Compression code - END
        adoCommand.Parameters.Append adoParameter
    Else
        'BG 30/09/02 SYS5619 Compression code
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        'Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGUID))
        'BG 30/09/02 SYS5619 Compression code - END
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strMetaData))
    adoParameter.Attributes = adParamNullable     ' Tell the paramter that it is allowed to have NULL values
    adoCommand.Parameters.Append adoParameter
    If bMetaDataExists = True Then
        adoCommand(3) = strMetaData
    Else
        adoCommand(3) = Null
    End If
       
    'BG 30/09/02 SYS5619 Compression code
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, 4, lFileSize)
    'Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileSize), strFileSize)
    'BG 30/09/02 SYS5619 Compression code - END
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPermissionsGUID))
        adoParameter.Attributes = adParamNullable     ' Tell the paramter that it is allowed to have NULL values
        adoCommand.Parameters.Append adoParameter
        If bPermissionsGUIDExists = True Then
            adoCommand(5) = strPermissionsGUID
        Else
            adoCommand(5) = Null
        End If
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileContentsGuid), strFileContentsGuid)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
        adoParameter.Attributes = adParamNullable     ' Tell the paramter that it is allowed to have NULL values
        adoCommand.Parameters.Append adoParameter
        If bPermissionsGUIDExists = True Then
            adoCommand(5) = GuidStringToByteArray(strPermissionsGUID)
        Else
            adoCommand(5) = Null
        End If
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileContentsGuid))
        adoCommand.Parameters.Append adoParameter
    End If
    ' Now we have all the parameters set up, execute the command
    Dim adoRecordSet As ADODB.Recordset
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "USP_FVCHECKINFILE call"
    Set adoRecordSet = adoCommand.Execute
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "USP_FVCHECKINFILE return"
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'adoRecordSet.MoveFirst
    Dim strChildFileVersion As String
    strChildFileVersion = adoRecordSet.fields.Item(0).Value
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("CHECKINFILE")
    xmlSetAttributeValue xmlNewChildNode, "FILEGUID", strFileGuid
    xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", strChildFileVersion
    xmlSetAttributeValue xmlNewChildNode, "ARCHIVETIMESTAMP", Format$(Now, "dd mmm yyyy hh:mm:ss")
    'IK_20/04/2006_CORE261
    xmlSetAttributeValue xmlNewChildNode, "FILECONTENTSGUID", strFileContentsGuid
    'IK_20/04/2006_CORE261_ends
    vxmlResponseNode.appendChild xmlNewChildNode
        
    ' ik_BM0261
    ' gobjContext.SetComplete
CheckinFileExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
CheckinFileVbErr:
    OnErrorDestroyADOObjects adoConnection, adoRecordSet, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo CheckinFileExit
End Sub
Public Sub IFileVersioningDO_CheckoutFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CheckoutFileVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "CheckoutFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Who locked the file?
    Dim strLockedBy As String
    strLockedBy = xmlGetMandatoryAttributeText(vxmlRequestNode, "LOCKEDBY")
    ' GUID of the file?
    Dim strFileGuid As String
    strFileGuid = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEGUID")
    ' File version?
    Dim strFileVersion As String
    strFileVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEVERSION")
    ' File contents?
    Dim bReturnFileContents As Boolean
    bReturnFileContents = True 'xmlGetAttributeAsBoolean(vxmlRequestNode, "FILEVERSION", "Y")
    ' Create an ADO connection
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    ' Create a new ADO command and set the command to be executed to the SQL command set above
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    ' Create the SQL string for calling the SQL stored procedure
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVLOCKFILE(?,?,?," _
            & "{resultset 1," _
            & "tableszFILECONTENTSGUID," _
            & "tableszLOCKEDBY" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVLOCKFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
    ' Create some parameters for the function using the values received from the XML stream above
    ' NOTE : These parameters will match the SQL stored procedure parameters, i.e.
    'PROCEDURE LockFile(
    '    szLOCKEDBY IN FVFOLDER.CREATEDBY%TYPE,
    '    rawFILEGUID IN FVFILE.FILEGUID%TYPE,
    '    szFILEVERSION IN FVFILE.FILEVERSION%TYPE,
    '    tableszFILECONTENTSGUID OUT t_tableszFILECONTENTSGUID,
    '    tableszLOCKEDBY OUT t_tableszLOCKEDBY
    '    )
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strLockedBy), strLockedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
    adoCommand.Parameters.Append adoParameter
    ' Now we have all the parameters set up, execute the command
    Dim adoRecordSet As ADODB.Recordset
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "USP_FVLOCKFILE call"
    Set adoRecordSet = adoCommand.Execute
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "USP_FVLOCKFILE return"
    ' Make sure we have received a pointer to the file record
    If adoRecordSet.EOF Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    ' If we are here, means we have a pointer to the correct record in the file table
    ' Get the GUID for the file contents from the record
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' adoRecordSet.MoveFirst
    Dim strFileContentsGuid As String
    Dim strAlreadyLockedBy As Variant
    strAlreadyLockedBy = adoRecordSet.fields.Item(1).Value
    'BG 30/09/02 SYS5619 Compression code
    If (IsNull(strAlreadyLockedBy)) Then strAlreadyLockedBy = ""
    'BG 30/09/02 SYS5619 Compression code - END
    ' Make sure the file isn't locked out by someone else
    'BG 30/09/02 SYS5619 Compression code
    If (strAlreadyLockedBy = "" Or (LCase(strAlreadyLockedBy) = LCase(strLockedBy))) Then
    'If IsNull(strAlreadyLockedBy) Then
    'BG 30/09/02 SYS5619 Compression code - END
            If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                strFileContentsGuid = adoRecordSet.fields.Item(0).Value
            Else
                strFileContentsGuid = adoGuidToString(adoRecordSet.fields.Item(0).Value)
            End If
        ' lock created, get the contents
        Dim strFileContentsType As String
        Dim strFileContents As String
        
        'AS 26/07/04 BBG1055 Added support for DELIVERYTYPE (Word/PDF).
        Dim nDeliveryType As Integer
        nDeliveryType = DELIVERYTYPE_DOC ' Defaults to Word document.
        
        'AS 26/07/2004 CORE286 Get file from archive as a Url and not as bin base 64 string
        Dim strFileContentsUrl As String
        strFileContentsUrl = ""
        
        If bReturnFileContents = True Then
            'BG 30/09/02 SYS5619 Compression code
            Dim xmlFileContentsNode As IXMLDOMAttribute
            Set xmlFileContentsNode = vxmlResponseNode.ownerDocument.createAttribute("FILECONTENTS")
            ReadFileContents adoConnection, strFileContentsGuid, vxmlRequestNode, strFileContentsType, nDeliveryType, xmlFileContentsNode, strFileContentsUrl
            'BG 30/09/02 SYS5619 Compression code - END
        End If
        ' Now we have locked the file and read the contents of the file so now we can return an xml response
        Dim xmlNewChildNode As IXMLDOMNode
        Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("CHECKOUTFILE")
        If bReturnFileContents = True Then
            xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", strFileVersion
            xmlSetAttributeValue xmlNewChildNode, "FILECONTENTS_TYPE", strFileContentsType
            xmlSetAttributeValue xmlNewChildNode, "DELIVERYTYPE", CStr(nDeliveryType)
            'AS 26/07/2004 CORE286 Get file from archive as a Url and not as bin base 64 string
            If strFileContentsUrl <> "" Then
                xmlSetAttributeValue xmlNewChildNode, "FILECONTENTSURL", strFileContentsUrl
            End If
            'BG 30/09/02 SYS5619 Compression code
            xmlNewChildNode.Attributes.setNamedItem xmlFileContentsNode
            'xmlSetAttributeValue xmlNewChildNode, "FILECONTENTS", strFileContents
            'BG 30/09/02 SYS5619 Compression code - END
        End If
        vxmlResponseNode.appendChild xmlNewChildNode
            
        ' And we're done!
        ' ik_BM0261
        ' gobjContext.SetComplete
    Else
'        strAlreadyLockedBy = adoRecordSet.Fields.Item("tableszLOCKEDBY").Value
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "ERROR"
        ' ik_bm0269
        ' xmlSetAttributeValue vxmlResponseNode, "DESCRIPTION", "File is already locked out by " & strAlreadyLockedBy
        ' gobjContext.SetAbort
        errThrowError cstrFunctionName, oeCommandFailed, "Document locked by user: " & strAlreadyLockedBy
        ' ik_bm0269_ends
    End If
CheckoutFileExit:
    If Not adoRecordSet Is Nothing Then
        adoRecordSet.Close
    End If
    Set adoRecordSet = Nothing
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    If adoConnection <> "" Then
        adoConnection.Close
    End If
    Set adoConnection = Nothing
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
CheckoutFileVbErr:
    OnErrorDestroyADOObjects adoConnection, adoRecordSet, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
'    xmlSetAttributeValue vxmlResponseNode, "TYPE", "ERROR"
'    xmlSetAttributeValue vxmlResponseNode, "DESCRIPTION", "Unknown error in LockFile SQL query"
    
    GoTo CheckoutFileExit
End Sub
Public Sub IFileVersioningDO_CreateFolder( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateFolderVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "CreateFolder"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim strCreatedBy As String
    strCreatedBy = xmlGetMandatoryAttributeText(vxmlRequestNode, "CREATEDBY")
    Dim strParentFolderGUID As String
    strParentFolderGUID = xmlGetAttributeText(vxmlRequestNode, "PARENTFOLDERGUID", FVS_ROOTFOLDERGUID)
    ' mandatory parent folder version if parentgiud supplied
    Dim strParentFolderVersion As String
    If strParentFolderGUID <> FVS_ROOTFOLDERGUID Then
        strParentFolderVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION")
    Else
        strParentFolderVersion = xmlGetAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION", "V1")
    End If
    Dim bMetaDataExists As Boolean
    Dim strMetaData As String
    bMetaDataExists = xmlAttributeValueExists(vxmlRequestNode, "METADATA")
    If bMetaDataExists = True Then
        strMetaData = xmlGetAttributeText(vxmlRequestNode, "METADATA")
    End If
    Dim strFolderName As String
    strFolderName = xmlGetMandatoryAttributeText(vxmlRequestNode, "FOLDERNAME")
    Dim bPermissionsGUIDExists As String
    Dim strPermissionsGUID As String
    bPermissionsGUIDExists = xmlAttributeValueExists(vxmlRequestNode, "PERMISSIONSGUID")
    If bPermissionsGUIDExists = True Then
        strParentFolderVersion = xmlGetAttributeText(vxmlRequestNode, "PERMISSIONSGUID")
    End If
    ' New folder = New GUID
    'BG 30/09/02 SYS5619 Compression code
     Dim strFolderGuid As String
    If xmlAttributeValueExists(vxmlRequestNode, "FOLDERGUID") Then
        strFolderGuid = xmlGetAttributeText(vxmlRequestNode, "FOLDERGUID")
    Else
        strFolderGuid = CreateGUID()
    End If
    'Dim strNewGuid As String
    'strNewGuid = CreateGUID()
    'BG 30/09/02 SYS5619 Compression code - END
    ' New DB connection
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVCREATEFOLDER(?,?,?,?,?,?,?," _
            & "{resultset 1," _
            & "tableszCHILDFOLDERVERSION" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVCREATEFOLDER"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCreatedBy), strCreatedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderGUID), strParentFolderGUID)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strParentFolderGUID))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderVersion), strParentFolderVersion)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strMetaData))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bMetaDataExists = True Then
        adoCommand(3) = strMetaData
    Else
         adoCommand(3) = Null
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFolderName), strFolderName)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPermissionsGUID))
        adoParameter.Attributes = adParamNullable
        adoCommand.Parameters.Append adoParameter
        If bPermissionsGUIDExists = True Then
            adoCommand(5) = strPermissionsGUID
        Else
            adoCommand(5) = Null
        End If
        'BG 30/09/02 SYS5619 Compression code
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFolderGuid), strFolderGuid)
        'Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strNewGuid), strNewGuid)
        'BG 30/09/02 SYS5619 Compression code - END
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
        adoParameter.Attributes = adParamNullable
        adoCommand.Parameters.Append adoParameter
        If bPermissionsGUIDExists = True Then
            adoCommand(5) = GuidStringToByteArray(strPermissionsGUID)
        Else
            adoCommand(5) = Null
        End If
        'BG 30/09/02 SYS5619 Compression code
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFolderGuid))
        'Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strNewGuid))
        'BG 30/09/02 SYS5619 Compression code - END
        adoCommand.Parameters.Append adoParameter
    End If
    Dim adoRecordSet As ADODB.Recordset
    Set adoRecordSet = adoCommand.Execute
    If adoRecordSet.EOF Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' adoRecordSet.MoveFirst
    Dim strChildFolderVersion As String
    strChildFolderVersion = adoRecordSet.fields.Item(0).Value
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("CREATEFOLDER")
    'BG 01/10/02 SYS5619 Compression code
    xmlSetAttributeValue xmlNewChildNode, "FOLDERGUID", strFolderGuid
    'xmlSetAttributeValue xmlNewChildNode, "FOLDERGUID", strNewGuid
    'BG 01/10/02 SYS5619 Compression code
    xmlSetAttributeValue xmlNewChildNode, "FOLDERVERSION", strChildFolderVersion
    vxmlResponseNode.appendChild xmlNewChildNode
        
    ' ik_BM0261
    ' gobjContext.SetComplete
CreateFolderExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
CreateFolderVbErr:
    OnErrorDestroyADOObjects adoConnection, adoRecordSet, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo CreateFolderExit
End Sub
Public Sub IFileVersioningDO_CreateFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateFileVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "CreateFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim xmlNode As IXMLDOMNode
    'DR Error Details
    Dim lngErrorNumber As Long
    Dim strErrorSource As String
    Dim strErrorDescription As String
    Set xmlNode = xmlGetNode(vxmlRequestNode, "CREATEFILEDATA")
    Dim strCreatedBy As String
    strCreatedBy = xmlGetMandatoryAttributeText(xmlNode, "CREATEDBY")
    Dim strParentFolderGUID As String
    strParentFolderGUID = xmlGetAttributeText(xmlNode, "PARENTFOLDERGUID", FVS_ROOTFOLDERGUID)
    ' mandatory parent folder version if parentgiud supplied
    Dim strParentFolderVersion As String
    If strParentFolderGUID <> FVS_ROOTFOLDERGUID Then
        strParentFolderVersion = xmlGetMandatoryAttributeText(xmlNode, "PARENTFOLDERVERSION")
    Else
        strParentFolderVersion = xmlGetAttributeText(xmlNode, "PARENTFOLDERVERSION", "V1")
    End If
    Dim bMetaDataExists As Boolean
    Dim strMetaData As String
    bMetaDataExists = xmlAttributeValueExists(xmlNode, "METADATA")
    If bMetaDataExists = True Then
        strMetaData = xmlGetAttributeText(xmlNode, "METADATA")
    End If
    Dim strFileName As String
    strFileName = xmlGetMandatoryAttributeText(xmlNode, "FILENAME")
    Dim nFileSize As Integer
    nFileSize = xmlGetAttributeAsInteger(xmlNode, "FILESIZE", 0)
    Dim bPermissionsGUIDExists As String
    Dim strPermissionsGUID As String
    bPermissionsGUIDExists = xmlAttributeValueExists(xmlNode, "PERMISSIONSGUID")
    If bPermissionsGUIDExists = True Then
        strParentFolderVersion = xmlGetAttributeText(xmlNode, "PERMISSIONSGUID")
    End If
    'BG 01/10/02 Compression code
    'Dim strNewGuid As String
    'strNewGuid = CreateGUID()
    'BG 01/10/02 Compression code - END
    ' file contents must be written with ADOs AppendChunk
    Dim adoConnection As ADODB.Connection
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
        
    'BG 01/10/02 Compression code
     Dim strFileGuid As String
    If xmlAttributeValueExists(xmlNode, "FILEGUID") Then
        strFileGuid = xmlGetAttributeText(xmlNode, "FILEGUID")
    Else
        strFileGuid = CreateGUID()
    End If
    ' DJP SYSMCP0615, DMS Compression
    WriteFileContents strCreatedBy, adoConnection, strFileGuid, xmlNode
    'WriteFileContents adoConnection, strNewGuid, xmlNode
    'BG 01/10/02 Compression code - END
       
    ' file details
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVCREATEFILE(?,?,?,?,?,?,?,?," _
            & "{resultset 1," _
            & "tableszCHILDFILEVERSION" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVCREATEFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCreatedBy), strCreatedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderGUID), strParentFolderGUID)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strParentFolderGUID))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderVersion), strParentFolderVersion)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strMetaData))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bMetaDataExists = True Then
        adoCommand(3) = strMetaData
    Else
        adoCommand(3) = Null
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileName), strFileName)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(5) = nFileSize
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strPermissionsGUID))
        adoParameter.Attributes = adParamNullable
        adoCommand.Parameters.Append adoParameter
        If bPermissionsGUIDExists = True Then
            adoCommand(6) = strPermissionsGUID
        Else
            adoCommand(6) = Null
        End If
        'BG 01/10/02 Compression code
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        'Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strNewGuid), strNewGuid)
        'BG 01/10/02 Compression code - END
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
        adoParameter.Attributes = adParamNullable
        adoCommand.Parameters.Append adoParameter
        If bPermissionsGUIDExists = True Then
            adoCommand(6) = GuidStringToByteArray(strPermissionsGUID)
        Else
            adoCommand(6) = Null
        End If
        'BG
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        'Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strNewGuid))
        'BG
        adoCommand.Parameters.Append adoParameter
    End If
    Dim adoRecordSet As Recordset
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "USP_FVCREATEFILE call"
    Set adoRecordSet = adoCommand.Execute
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "USP_FVCREATEFILE return"
    If adoRecordSet.EOF Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' adoRecordSet.MoveFirst
    Dim strChildFileVersion As String
    strChildFileVersion = adoRecordSet.fields.Item(0).Value
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("CREATEFILEDATA")
    'BG 01/10/02 Compression code
    xmlSetAttributeValue xmlNewChildNode, "FILEGUID", strFileGuid
    'xmlSetAttributeValue xmlNewChildNode, "FILEGUID", strNewGuid
    'BG 01/10/02 Compression code - END
    xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", strChildFileVersion
    xmlSetAttributeValue xmlNewChildNode, "ARCHIVETIMESTAMP", Now
    
    vxmlResponseNode.appendChild xmlNewChildNode
        
    ' ik_BM0261
    ' gobjContext.SetComplete
CreateFileExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
CreateFileVbErr:
    'Record the error details
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & CstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & lngErrorNumber & ", Err.Source: " & strErrorSource & ", Err.Description: " & strErrorDescription, vbLogEventTypeError
    'Useful debug info
    App.LogEvent CstrClassName & "::" & cstrFunctionName & "(). Debug Information. vxmlRequestNode.xml: " & vxmlRequestNode.xml
    OnErrorDestroyADOObjects adoConnection, adoRecordSet, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo CreateFileExit
End Sub
Public Sub IFileVersioningDO_GetFiles( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetFilesVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "GetFiles"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim strParentFolderGUID As String
    strParentFolderGUID = xmlGetAttributeText(vxmlRequestNode, "PARENTFOLDERGUID", FVS_ROOTFOLDERGUID)
    ' mandatory parent folder version if parentgiud supplied
    Dim strParentFolderVersion As String
    If strParentFolderGUID <> FVS_ROOTFOLDERGUID Then
        strParentFolderVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION")
    Else
        strParentFolderVersion = xmlGetAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION", "V1")
    End If
    Dim strLikeFileName As String
    Dim bLikeFileNameExists As Boolean
    bLikeFileNameExists = xmlAttributeValueExists(vxmlRequestNode, "LIKEFILENAME")
    If bLikeFileNameExists = True Then
        strLikeFileName = xmlGetMandatoryAttributeText(vxmlRequestNode, "LIKEFILENAME")
    End If
    Dim strLikeLatestVersion As String
    Dim bLikeLatestVersionExists As Boolean
    bLikeLatestVersionExists = xmlAttributeValueExists(vxmlRequestNode, "LIKELATESTVERSION")
    If bLikeLatestVersionExists = True Then
        strLikeLatestVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "LIKELATESTVERSION")
    End If
    Dim bIncludeDeletedFiles As Integer
    bIncludeDeletedFiles = xmlGetAttributeAsInteger(vxmlRequestNode, "INCLUDEDELETEDFILES", "0")
    Dim bIncludeNonDeletedFiles As Integer
    bIncludeNonDeletedFiles = xmlGetAttributeAsInteger(vxmlRequestNode, "INCLUDENONDELETEDFILES", "1")
    Dim bIncludeBranchedFiles As Integer
    bIncludeBranchedFiles = xmlGetAttributeAsInteger(vxmlRequestNode, "INCLUDEBRANCHEDFILES", "1")
    Dim bPermissionsGUIDExists As String
    Dim strPermissionsGUID As String
    bPermissionsGUIDExists = xmlAttributeValueExists(vxmlRequestNode, "PERMISSIONSGUID")
    If bPermissionsGUIDExists = True Then
        strParentFolderVersion = xmlGetAttributeText(vxmlRequestNode, "PERMISSIONSGUID")
    End If
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    Dim adoParameter As ADODB.Parameter
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVGETFILES(?,?,?,?,?,?,?," _
            & "{resultset 1000," _
            & "tableszCHILDFILEGUID," _
            & "tableszCHILDFILEVERSION," _
            & "tableszCHILDFILENAME" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderGUID), strParentFolderGUID)
        adoCommand.Parameters.Append adoParameter
    Else
        strSQL = "USP_FVGETFILES"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strParentFolderGUID))
        adoCommand.Parameters.Append adoParameter
    End If
        
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderVersion), strParentFolderVersion)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strLikeFileName))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bLikeFileNameExists = True Then
        adoCommand(2) = strLikeFileName
    Else
        adoCommand(2) = Null
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strLikeLatestVersion))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bLikeLatestVersionExists = True Then
        adoCommand(3) = strLikeLatestVersion
    Else
        adoCommand(3) = Null
    End If
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(4) = bIncludeDeletedFiles
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(5) = bIncludeNonDeletedFiles
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(6) = bIncludeBranchedFiles
        
    Dim adoRecordSet As ADODB.Recordset
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "USP_FVGETFILES call"
    Set adoRecordSet = adoCommand.Execute
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "USP_FVGETFILES return"
    If adoRecordSet.EOF = False Then
        Dim xmlNewChildNode As IXMLDOMNode
        ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
        ' adoRecordSet.MoveFirst
        Do While True
            If adoRecordSet.EOF Then
                Exit Do
            End If
            Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("GETFILE")
            If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                xmlSetAttributeValue xmlNewChildNode, "FILEGUID", adoRecordSet.fields.Item(0).Value
            Else
                xmlSetAttributeValue xmlNewChildNode, "FILEGUID", adoGuidToString(adoRecordSet.fields.Item(0).Value)
            End If
            xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", adoRecordSet.fields.Item(1).Value
            xmlSetAttributeValue xmlNewChildNode, "FILENAME", adoRecordSet.fields.Item(2).Value
            vxmlResponseNode.appendChild xmlNewChildNode
            adoRecordSet.MoveNext
        Loop
    End If
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
       
    ' ik_BM0261
    ' gobjContext.SetComplete
GetFilesExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
GetFilesVbErr:
    OnErrorDestroyADOObjects adoConnection, adoRecordSet, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo GetFilesExit
End Sub
Public Sub IFileVersioningDO_GetFolders( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetFoldersVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "GetFolders"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim strParentFolderGUID As String
    strParentFolderGUID = xmlGetAttributeText(vxmlRequestNode, "PARENTFOLDERGUID", FVS_ROOTFOLDERGUID)
    ' mandatory parent folder version if parentgiud supplied
    Dim strParentFolderVersion As String
    If strParentFolderGUID <> FVS_ROOTFOLDERGUID Then
        strParentFolderVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION")
    Else
        strParentFolderVersion = xmlGetAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION", "V1")
    End If
    Dim strLikeFolderName As String
    Dim bLikeFolderNameExists As Boolean
    bLikeFolderNameExists = xmlAttributeValueExists(vxmlRequestNode, "LIKEFOLDERNAME")
    If bLikeFolderNameExists = True Then
        strLikeFolderName = xmlGetMandatoryAttributeText(vxmlRequestNode, "LIKEFOLDERNAME")
    End If
    Dim strLikeLatestVersion As String
    Dim bLikeLatestVersionExists As Boolean
    bLikeLatestVersionExists = xmlAttributeValueExists(vxmlRequestNode, "LIKELATESTVERSION")
    If bLikeLatestVersionExists = True Then
        strLikeLatestVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "LIKELATESTVERSION")
    End If
    Dim bIncludeDeletedFolders As Integer
    bIncludeDeletedFolders = xmlGetAttributeAsInteger(vxmlRequestNode, "INCLUDEDELETEDFOLDERS", "0")
    Dim bIncludeNonDeletedFolders As Integer
    bIncludeNonDeletedFolders = xmlGetAttributeAsInteger(vxmlRequestNode, "INCLUDENONDELETEDFOLDERS", "1")
    Dim bIncludeBranchedFolders As Integer
    bIncludeBranchedFolders = xmlGetAttributeAsInteger(vxmlRequestNode, "INCLUDEBRANCHEDFOLDERS", "1")
    Dim bPermissionsGUIDExists As String
    Dim strPermissionsGUID As String
    bPermissionsGUIDExists = xmlAttributeValueExists(vxmlRequestNode, "PERMISSIONSGUID")
    If bPermissionsGUIDExists = True Then
        strParentFolderVersion = xmlGetAttributeText(vxmlRequestNode, "PERMISSIONSGUID")
    End If
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    Dim adoParameter As ADODB.Parameter
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVGETFOLDERS(?,?,?,?,?,?,?," _
            & "{resultset 1000," _
            & "tableszCHILDFOLDERGUID," _
            & "tableszCHILDFOLDERVERSION," _
            & "tableszCHILDFOLDERNAME" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderGUID), strParentFolderGUID)
        adoCommand.Parameters.Append adoParameter
    Else
        strSQL = "USP_FVGETFOLDERS"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strParentFolderGUID))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderVersion), strParentFolderVersion)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strLikeFolderName))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bLikeFolderNameExists = True Then
        adoCommand(2) = strLikeFolderName
    Else
        adoCommand(2) = Null
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strLikeLatestVersion))
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    If bLikeLatestVersionExists = True Then
        adoCommand(3) = strLikeLatestVersion
    Else
        adoCommand(3) = Null
    End If
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(4) = bIncludeDeletedFolders
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(5) = bIncludeNonDeletedFolders
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(6) = bIncludeBranchedFolders
        
    Dim adoRecordSet As ADODB.Recordset
    Set adoRecordSet = adoCommand.Execute
    If adoRecordSet.EOF = False Then
        Dim xmlNewChildNode As IXMLDOMNode
        ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
        ' adoRecordSet.MoveFirst
        Do While True
            If adoRecordSet.EOF Then
                Exit Do
            End If
            Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("GETFOLDER")
            If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                xmlSetAttributeValue xmlNewChildNode, "FOLDERGUID", adoRecordSet.fields.Item(0).Value
            Else
                xmlSetAttributeValue xmlNewChildNode, "FOLDERGUID", adoGuidToString(adoRecordSet.fields.Item(0).Value)
            End If
            xmlSetAttributeValue xmlNewChildNode, "FOLDERVERSION", adoRecordSet.fields.Item(1).Value
            xmlSetAttributeValue xmlNewChildNode, "FOLDERNAME", adoRecordSet.fields.Item(2).Value
            vxmlResponseNode.appendChild xmlNewChildNode
            adoRecordSet.MoveNext
        Loop
    End If
    adoRecordSet.Close
    Set adoRecordSet = Nothing
       
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
       
    ' ik_BM0261
    ' gobjContext.SetComplete
GetFoldersExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
GetFoldersVbErr:
    OnErrorDestroyADOObjects adoConnection, adoRecordSet, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo GetFoldersExit
End Sub
Private Sub IFileVersioningDO_UnlockFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UnlockFileVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "UnlockFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim strLockedBy As String
    strLockedBy = xmlGetMandatoryAttributeText(vxmlRequestNode, "LOCKEDBY")
    Dim strFileGuid As String
    strFileGuid = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEGUID")
    Dim strFileVersion As String
    strFileVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEVERSION")
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVUNLOCKFILE(?,?,?)}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVUNLOCKFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strLockedBy), strLockedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
    adoCommand.Parameters.Append adoParameter
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "USP_FVUNLOCKFILE call"
    adoCommand.Execute , , adExecuteNoRecords
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "USP_FVUNLOCKFILE return"
        
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
       
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("UNLOCKFILE")
    vxmlResponseNode.appendChild xmlNewChildNode
    ' ik_BM0261
    ' gobjContext.SetComplete
UnlockFileExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
UnlockFileVbErr:
    OnErrorDestroyADOObjects adoConnection, Nothing, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo UnlockFileExit
End Sub
Private Sub OnErrorDestroyADOObjects(ByRef adoConnection As ADODB.Connection, ByRef adoRecordSet As ADODB.Recordset, ByRef adoCommand As ADODB.Command)
    
    Const cstrFunctionName As String = "OnErrorDestroyADOObjects"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim bErrorsReported As Boolean
    bErrorsReported = False
    ' Close the recordset if still open
    If Not adoRecordSet Is Nothing Then
        If VarType(adoRecordSet.ActiveConnection) = vbObject Then
            If Not adoRecordSet.ActiveConnection Is Nothing Then
                bErrorsReported = ReportADOErrors(adoRecordSet.ActiveConnection.ERRORS)
            End If
        End If
        If adoRecordSet.State = adStateOpen Then
            adoRecordSet.Close
        End If
        Set adoRecordSet = Nothing
    End If
    ' Close the command
    If Not adoCommand Is Nothing Then
        If bErrorsReported = False And VarType(adoCommand.ActiveConnection) = vbObject Then
            If Not adoCommand.ActiveConnection Is Nothing Then
                bErrorsReported = ReportADOErrors(adoCommand.ActiveConnection.ERRORS)
            End If
        End If
        Set adoCommand = Nothing
    End If
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If bErrorsReported = False Then
            bErrorsReported = ReportADOErrors(adoConnection.ERRORS)
        End If
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
        Set adoConnection = Nothing
    End If
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Sub
Private Function ReportADOErrors(ByVal ERRORS As ADODB.ERRORS) As Boolean
    
    Const cstrFunctionName As String = "ReportADOErrors"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ReportADOErrors = False
    Dim Error As ADODB.Error
    For Each Error In ERRORS
        App.LogEvent "NativeError (" + CStr(Error.NativeError) + _
            ")  Number (" + CStr(Error.Number) + _
            ")  Source (" + CStr(Error.Source) + _
            ")  SQLState (" + Error.SQLState + _
            ")  Description (" + Error.Description + _
            ")", _
            vbLogEventTypeWarning
       ReportADOErrors = True
    Next
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Function
Private Sub WriteFileContents( _
    ByVal strUserId As String, _
    ByVal adoConnection As ADODB.Connection, _
    ByVal strFileContentsGuid As String, _
    ByVal vxmlRequestNode As IXMLDOMNode _
    )
    Const cstrFunctionName = "WriteFileContents"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Dim strFileContentsType As String
    
    'IK_20/04/2006_CORE261
    Dim DocumentManagementSystemType As Integer
    DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_DMS
    If Not vxmlRequestNode.Attributes.getNamedItem("DOCUMENTMANAGEMENTSYSTEMTYPE") Is Nothing Then
        DocumentManagementSystemType = vxmlRequestNode.Attributes.getNamedItem("DOCUMENTMANAGEMENTSYSTEMTYPE").Text
    End If
    'IK_20/04/2006_CORE261_ends

    strFileContentsType = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILECONTENTS_TYPE")
    
    'IK_20/04/2006_CORE261
    If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
        strFileContentsType = "GEMINI"
    End If
    'IK_20/04/2006_CORE261_ends
    
    If xmlAttributeValueExists(vxmlRequestNode, "FILECONTENTS") = False Then
        errThrowError "WriteFileContents.xmlAttributeValueExists", Err, "[@FILECONTENTS]"
    End If
    ' compress is compression is available
    Dim varCompressedFileContents As Variant
    Dim nStorageMethod As Integer
    Dim blnCompressed As Boolean
    Dim strCompressionMethod As String
    Dim nDeliveryType As Integer
    Dim objCompression As dmsCompression1
    
    'AS 16/07/04 BBG942 Added support for different document formats
    nDeliveryType = xmlGetAttributeAsInteger(vxmlRequestNode, "DELIVERYTYPE", DELIVERYTYPE_DOC)
    blnCompressed = xmlGetAttributeAsBoolean(vxmlRequestNode, "COMPRESSED", "0")
    strCompressionMethod = xmlGetAttributeText(vxmlRequestNode, "COMPRESSIONMETHOD", "ZLIB")
    
    nStorageMethod = STORAGE_NATIVE
    Dim nodeFileContents As IXMLDOMNode
    Set nodeFileContents = xmlGetAttributeNode(vxmlRequestNode, "FILECONTENTS")
    
    'IK_20/04/2006_CORE261
    If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
        If vxmlRequestNode.Attributes.getNamedItem("COMPRESSIONMETHOD") Is Nothing Then
            vxmlRequestNode.Attributes.setNamedItem (vxmlRequestNode.ownerDocument.createAttribute("COMPRESSIONMETHOD"))
        End If
        If vxmlRequestNode.Attributes.getNamedItem("COMPRESSIONMETHOD").Text = "" Then
            vxmlRequestNode.Attributes.getNamedItem("COMPRESSIONMETHOD").Text = "ZLIB"
        End If
    Else
        If blnCompressed Then
            ' AS 20/04/2005 CORE106 File contents are already compressed, so do not compress again.
            Dim xmlDoc As New FreeThreadedDOMDocument40
            Dim xndConversionNode As IXMLDOMNode
            
            ' Convert the stream from the binary encoding (e.g., bin.base64).
            Set xndConversionNode = xmlDoc.createElement("CONVERSION_NODE")
            xndConversionNode.Text = nodeFileContents.nodeValue
            xndConversionNode.dataType = LCase(strFileContentsType)
            varCompressedFileContents = xndConversionNode.nodeTypedValue
    
            If LCase(strFileContentsType) = "bin.base64" Then
                nStorageMethod = STORAGE_DMSCOMPRESSION1BINBASE64
            ElseIf LCase(strFileContentsType) = "bin.hex" Then
                nStorageMethod = STORAGE_DMSCOMPRESSION1BINHEX
            Else
                nStorageMethod = STORAGE_DMSCOMPRESSION1
            End If
            
            Set xmlDoc = Nothing
            Set xndConversionNode = Nothing
        Else
            ' File contents are not compressed, so compress before saving in database.
            On Error Resume Next
            Set objCompression = CreateObject("DmsCompression.dmsCompression1")
            On Error GoTo 0
            
            If Not objCompression Is Nothing Then
                objCompression.CompressionMethod = strCompressionMethod
                
                'AS 06/07/04 BBG942 Added support for ZLIB compression
                'Compression type defaults to new ZLIB algorithm; can override - set to "COMPAPI" for
                'old algorithm.
                If LCase(strFileContentsType) = "bin.base64" Then
                    nStorageMethod = STORAGE_DMSCOMPRESSION1BINBASE64
                    varCompressedFileContents = objCompression.XMLNODEBINBASE64CompressToSafeArray(nodeFileContents)
                ElseIf LCase(strFileContentsType) = "bin.hex" Then
                    nStorageMethod = STORAGE_DMSCOMPRESSION1BINHEX
                    varCompressedFileContents = objCompression.XMLNODEBINHEXCompressToSafeArray(nodeFileContents)
                Else
                    nStorageMethod = STORAGE_DMSCOMPRESSION1
                    varCompressedFileContents = objCompression.XMLNODECompressToSafeArray(nodeFileContents)
                End If
                Set objCompression = Nothing
            End If
        End If
    'IK_20/04/2006_CORE261
    End If
    
    'If adoGetDbProvider = DBPROVIDEROracle Then
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Dim adoCommand As New ADODB.Command
        Set adoCommand.ActiveConnection = adoConnection
        adoCommand.Properties("SPPrmsLOB") = True
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = "{call PAC_FILEVERSIONING.WriteFileContents(?,?,?,?,?,?)}"
        Dim adoParameter As New ADODB.Parameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strUserId), strUserId)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileContentsGuid), strFileContentsGuid)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, 0, nStorageMethod)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileContentsType), strFileContentsType)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCompressionMethod), strCompressionMethod)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, 0, nDeliveryType)
        adoCommand.Parameters.Append adoParameter
        If nStorageMethod = STORAGE_NATIVE Then
            Set adoParameter = adoCommand.CreateParameter(, adLongVarBinary, adParamInput, LenB(nodeFileContents.nodeValue), nodeFileContents.nodeValue)
        Else
            Set adoParameter = adoCommand.CreateParameter(, adLongVarBinary, adParamInput, LenB(varCompressedFileContents), varCompressedFileContents)
        End If
        adoCommand.Parameters.Append adoParameter
        ' Now we have all the parameters set up, execute the command
        adoCommand.Execute , , adExecuteNoRecords
    Else
        ' ik_BM0200 traceassist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "SELECT * FROM FVFILECONTENTS call"
        Dim adoRecordSet As ADODB.Recordset
        Set adoRecordSet = New ADODB.Recordset
        adoRecordSet.CursorType = adOpenKeyset
        adoRecordSet.LockType = adLockOptimistic
        adoRecordSet.Open "SELECT * FROM FVFILECONTENTS WHERE FILECONTENTSGUID IS NULL", adoConnection, , , adCmdText
        ' ik_BM0200 traceassist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "SELECT * FROM FVFILECONTENTS return"
           
        adoRecordSet.AddNew
        adoRecordSet!FILECONTENTSGUID = strFileContentsGuid
        adoRecordSet!STOREAGEMETHOD = nStorageMethod
        adoRecordSet!FILECONTENTS_TYPE = strFileContentsType
        'AS 06/07/04 BBG942 Added support for ZLIB compression
        adoRecordSet!CompressionMethod = strCompressionMethod
        adoRecordSet!DeliveryType = nDeliveryType
            
        Dim lngOffset As Long
        Dim lngFileContentsSize As Long
        'IK_20/04/2006_CORE261
        If DocumentManagementSystemType <> DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
            If nStorageMethod = STORAGE_NATIVE Then
                If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                    Dim strChunk As String
                    lngFileContentsSize = LenB(nodeFileContents.nodeValue)
                    lngOffset = 0 ' Reset offset.
                    Do While lngOffset < lngFileContentsSize
                        strChunk = LeftB(RightB(nodeFileContents.nodeValue, lngFileContentsSize - lngOffset), FVS_CHUNKSIZE)
                        adoRecordSet!FileContents.AppendChunk strChunk
                        lngOffset = lngOffset + FVS_CHUNKSIZE
                    Loop
                Else
                    'AS BBG1055 No need to chunk with SQL Server.
                    adoRecordSet!FileContents = nodeFileContents.nodeValue
                End If
            Else
                If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                    Dim varChunk As Variant
                    lngFileContentsSize = LenB(varCompressedFileContents)
                    lngOffset = 0 ' Reset offset.
                    Do While lngOffset < lngFileContentsSize
                        varChunk = LeftB(RightB(varCompressedFileContents, lngFileContentsSize - lngOffset), FVS_CHUNKSIZE)
                        adoRecordSet!FileContents.AppendChunk varChunk
                        lngOffset = lngOffset + FVS_CHUNKSIZE
                    Loop
                Else
                    'AS BBG1055 No need to chunk with SQL Server.
                    'Also chunking code was not writing last byte of buffer to database,
                    'causing errors when reading back the data and decompressing it.
                    adoRecordSet!FileContents = varCompressedFileContents
                End If
            End If
        'IK_20/04/2006_CORE261
        End If
        
        ' ik_BM0200 traceassist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "adoRecordSet.Update call"
        adoRecordSet.Update
        ' ik_BM0200 traceassist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "adoRecordSet.Update return"
        adoRecordSet.Close
    End If
'BG 01/10/02 SYS5619 Compression code - END
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
       
End Sub

'AS 28/04/2006 CORE261 Add support for Gemini.
Private Sub ReadFileContents( _
    ByVal adoConnection As ADODB.Connection, _
    ByVal strFileContentsGuid As String, _
    ByVal xmlRequestNode As IXMLDOMNode, _
    ByRef strFileContentsType As String, _
    ByRef nDeliveryType As Integer, _
    ByVal xmlFileContentsNode As IXMLDOMNode, _
    ByRef strFileContentsUrl As String)
    
    Const cstrFunctionName = "ReadFileContents"
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName

    Dim DocumentManagementSystemType As Integer
    DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_DMS
    If Not xmlRequestNode.Attributes.getNamedItem("DOCUMENTMANAGEMENTSYSTEMTYPE") Is Nothing Then
        DocumentManagementSystemType = xmlRequestNode.Attributes.getNamedItem("DOCUMENTMANAGEMENTSYSTEMTYPE").Text
    End If
    If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
        Call ReadFileContentsGemini(strFileContentsGuid, xmlRequestNode, strFileContentsType, nDeliveryType, xmlFileContentsNode, strFileContentsUrl)
    Else
        Call ReadFileContentsDMS(adoConnection, strFileContentsGuid, xmlRequestNode, strFileContentsType, nDeliveryType, xmlFileContentsNode)
    End If

    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Sub

Private Sub ReadFileContentsGemini(ByVal strFileContentsGuid As String, ByVal xmlRequestNode As IXMLDOMNode, ByRef strFileContentsType As String, ByRef nDeliveryType As Integer, ByVal xmlFileContentsNode As IXMLDOMNode, ByRef strFileContentsUrl As String)
    Const cstrFunctionName = "ReadFileContentsGemini"
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName

    xmlRequestNode.Attributes.setNamedItem xmlRequestNode.ownerDocument.createAttribute("FILECONTENTSGUID")
    xmlRequestNode.Attributes.getNamedItem("FILECONTENTSGUID").Text = strFileContentsGuid

    Dim objGeminiBO As Object ' Late bind as omGemini may not be installed.
    Set objGeminiBO = gobjContext.CreateInstance("omGemini.FileVersioningBO")

    Dim xmlResponse As New FreeThreadedDOMDocument40
    xmlResponse.loadXML objGeminiBO.ProcessRequest(xmlRequestNode.xml)
    Set objGeminiBO = Nothing
    
    errCheckXMLResponseNode xmlResponse.documentElement, , True
    
    xmlFileContentsNode.Text = xmlGetAttributeText(xmlResponse.documentElement, "FILECONTENTS")
    strFileContentsType = xmlGetMandatoryAttributeText(xmlResponse.documentElement, "FILECONTENTS_TYPE")
    'AS 26/07/2004 CORE286 Get file from archive as a Url and not as bin base 64 string
    strFileContentsUrl = xmlGetAttributeText(xmlResponse.documentElement, "FILECONTENTSURL")
    nDeliveryType = CSafeInt(xmlGetMandatoryAttributeText(xmlResponse.documentElement, "DELIVERYTYPE"))
    
    Set xmlResponse = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName

End Sub
'AS 28/04/2006 CORE261 End

Private Sub ReadFileContentsDMS( _
    ByVal adoConnection As ADODB.Connection, _
    ByVal strFileContentsGuid As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByRef strFileContentsType As String, _
    ByRef nDeliveryType As Integer, _
    ByVal xmlFileContentsNode As IXMLDOMNode _
    )
    Const cstrFunctionName = "ReadFileContentsDMS"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Dim nStorageMethod As Integer
    Dim strSrcCompressionMethod As String
    Dim strTgtCompressionMethod As String
    Dim adoRecordSet As ADODB.Recordset
    Dim varCompressedFileContents As Variant
        
    strTgtCompressionMethod = xmlGetAttributeText(vxmlRequestNode, "COMPRESSIONMETHOD", "")
        
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
    'If adoGetDbProvider = DBPROVIDEROracle Then
        Dim adoCommand As New ADODB.Command
        Set adoCommand.ActiveConnection = adoConnection
        adoCommand.Properties("SPPrmsLOB") = True
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = "{call PAC_FILEVERSIONING.ReadFileContents(?)}"
        Dim adoParameter As New ADODB.Parameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileContentsGuid), strFileContentsGuid)
        adoCommand.Parameters.Append adoParameter
        Set adoRecordSet = adoCommand.Execute
        strFileContentsType = adoRecordSet.fields.Item(1).Value
        nStorageMethod = adoRecordSet.fields.Item(2).Value
        If nStorageMethod = STORAGE_NATIVE Then
            xmlFileContentsNode.nodeValue = adoRecordSet.fields.Item(0).Value
        Else
            varCompressedFileContents = adoRecordSet.fields.Item(0).Value
        End If
        'AS 06/07/04 BBG942 Added support for ZLIB compression
        strSrcCompressionMethod = adoRecordSet.fields.Item(5).Value
        nDeliveryType = adoRecordSet.fields.Item(6).Value
    Else
        ' ik_BM0200 traceassist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "SELECT * FROM FVFILECONTENTS call"
        Set adoRecordSet = New ADODB.Recordset
        adoRecordSet.CursorType = adOpenForwardOnly
        adoRecordSet.CursorLocation = adUseClient
        ' DJP SYSMCP0615, DMS Compression
        'If adoGetDbProvider = DBPROVIDEROracle Then
        If adoGetDbProvider = omiga4DBPROVIDEROracle Then
            adoRecordSet.Open "SELECT * FROM FVFILECONTENTS WHERE FILECONTENTSGUID = '" & strFileContentsGuid & "'", adoConnection, , , adCmdText
        Else
            adoRecordSet.Open "SELECT * FROM FVFILECONTENTS WHERE FILECONTENTSGUID = 0x" & strFileContentsGuid, adoConnection, , , adCmdText
        End If
        Dim lngFileContentsSize As Long
        Dim lngOffset As Long
        ' file contents must be read with ADOs GetChunk
        lngOffset = 0
        lngFileContentsSize = adoRecordSet!FileContents.ActualSize
        strFileContentsType = adoRecordSet!FILECONTENTS_TYPE
        nStorageMethod = adoRecordSet!STOREAGEMETHOD
        'AS 06/07/04 BBG942 Added support for ZLIB compression
        If IsNull(adoRecordSet!CompressionMethod) Then
            strSrcCompressionMethod = "COMPAPI"
        Else
            strSrcCompressionMethod = adoRecordSet!CompressionMethod
        End If
        If IsNull(adoRecordSet!DeliveryType) Then
            ' Default to Word format.
            nDeliveryType = DELIVERYTYPE_DOC
        Else
            nDeliveryType = adoRecordSet!DeliveryType
        End If

        Dim varChunk As Variant
        If nStorageMethod = STORAGE_NATIVE Then
            Dim varFileContents As Variant
            If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                Do While lngOffset < lngFileContentsSize
                    varChunk = adoRecordSet!FileContents.GetChunk(FVS_CHUNKSIZE)
                    varFileContents = varFileContents & varChunk
                    lngOffset = lngOffset + FVS_CHUNKSIZE
                Loop
                xmlFileContentsNode.nodeValue = varFileContents
            Else
                'AS BBG1055 No need to chunk with SQL Server.
                varFileContents = adoRecordSet!FileContents
            End If
            xmlFileContentsNode.nodeValue = varFileContents
        Else
            If adoGetDbProvider = omiga4DBPROVIDEROracle Then
                Do While lngOffset < lngFileContentsSize
                    varChunk = adoRecordSet!FileContents.GetChunk(FVS_CHUNKSIZE)
                    varCompressedFileContents = varCompressedFileContents & varChunk
                    lngOffset = lngOffset + FVS_CHUNKSIZE
                Loop
            Else
                'AS BBG1055 No need to chunk with SQL Server.
                varCompressedFileContents = adoRecordSet!FileContents
            End If
        End If
        adoRecordSet.Close
        ' ik_BM0200 traceassist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "SELECT * FROM FVFILECONTENTS return"
    End If
    Select Case nStorageMethod
        Case STORAGE_DMSCOMPRESSION1, STORAGE_DMSCOMPRESSION1BINBASE64, STORAGE_DMSCOMPRESSION1BINHEX
            If strTgtCompressionMethod = "" Or strTgtCompressionMethod <> strSrcCompressionMethod Then
                ' Decompress the file.
                ' ik_BM0200 traceassist support
                gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "DmsCompression call"
                Dim objCompression As dmsCompression1
                On Error Resume Next
                Set objCompression = CreateObject("DmsCompression.dmsCompression1")
                On Error GoTo 0
                If Not objCompression Is Nothing Then
                    'AS 06/07/04 BBG942 Added support for ZLIB compression
                    objCompression.CompressionMethod = strSrcCompressionMethod
                    
                    If nStorageMethod = STORAGE_DMSCOMPRESSION1BINBASE64 Then
                        objCompression.SafeArrayDecompressToXMLNODEBINBASE64 varCompressedFileContents, xmlFileContentsNode
                    ElseIf nStorageMethod = STORAGE_DMSCOMPRESSION1BINHEX Then
                        objCompression.SafeArrayDecompressToXMLNODEBINHEX varCompressedFileContents, xmlFileContentsNode
                    Else 'STORAGE_DMSCOMPRESSION1
                        objCompression.SafeArrayDecompressToXMLNODE varCompressedFileContents, xmlFileContentsNode
                    End If
                    
                    If strTgtCompressionMethod <> "" Then
                        ' Target compression method is different to the method used in DMS, so recompress the file using
                        ' the target method.
                        xmlFileContentsNode.dataType = strFileContentsType
                        objCompression.CompressionMethod = strTgtCompressionMethod
                        If nStorageMethod = STORAGE_DMSCOMPRESSION1BINBASE64 Then
                            varCompressedFileContents = objCompression.XMLNODEBINBASE64CompressToSafeArray(xmlFileContentsNode)
                        ElseIf nStorageMethod = STORAGE_DMSCOMPRESSION1BINHEX Then
                            varCompressedFileContents = objCompression.XMLNODEBINHEXCompressToSafeArray(xmlFileContentsNode)
                        Else
                            varCompressedFileContents = objCompression.XMLNODECompressToSafeArray(xmlFileContentsNode)
                        End If
                        xmlFileContentsNode.nodeTypedValue = varCompressedFileContents
                    End If
                    
                    Set objCompression = Nothing
                Else
                    errThrowError cstrFunctionName, Err, "Cannot create object DmsCompression.dmsCompression1"
                End If
                ' ik_BM0200 traceassist support
                gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "DmsCompression return"
                
            Else
                ' Target compression method is the same as the method used in DMS, so return the file without
                ' decompressing it.
                xmlFileContentsNode.dataType = strFileContentsType
                xmlFileContentsNode.nodeTypedValue = varCompressedFileContents
            End If
        Case STORAGE_NATIVE
            ' no further processing is required
        Case Else
            errThrowError cstrFunctionName, Err, "Unrecognised storage method"
    End Select
' BG 01/10/02 SYS5619 Compression code - END
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Sub
Public Sub IFileVersioningDO_ShareFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo ShareFileVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "ShareFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Create a DB connection and open it
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    ' Now we need to set up some parameters to the SQL stored procedure0
    Dim strCreatedBy As String
    strCreatedBy = xmlGetMandatoryAttributeText(vxmlRequestNode, "CREATEDBY")
    Dim strParentFolderGUID As String
    strParentFolderGUID = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERGUID")
    Dim strParentFolderVersion As String
    strParentFolderVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION")
    Dim strFileGuid As String
    strFileGuid = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEGUID")
    Dim strFileVersion As String
    strFileVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEVERSION")
    ' Need an ADO Command
    Dim adoCommand As New ADODB.Command
    adoCommand.ActiveConnection = adoConnection
    ' Create the SQL string for calling the SQL stored procedure
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVSHAREFILE(?,?,?,?,?, " _
            & "{resultset 1," & "tableszFILEVERSION" & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVSHAREFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
    ' Create a parameter object and fill it in
    Dim adoParameter As New ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCreatedBy), strCreatedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderGUID), strParentFolderGUID)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strParentFolderGUID))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderVersion), strParentFolderVersion)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
    adoCommand.Parameters.Append adoParameter
    ' Create a recordset to execute the command...
    Dim adoRecordSet As New ADODB.Recordset
    Set adoRecordSet = adoCommand.Execute
    ' Get the out parameter
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' adoRecordSet.MoveFirst
    Dim strVersion As String
    strVersion = adoRecordSet.fields.Item(0)
    ' And now close everything down nicely!!
    adoRecordSet.Close
    Set adoRecordSet = Nothing
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    ' Now we have shared the file so now we can return an xml response
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("SHAREFILE")
    xmlSetAttributeValue xmlNewChildNode, "FILEGUID", strFileGuid
    xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", strFileVersion
    vxmlResponseNode.appendChild xmlNewChildNode
        
    ' And we're done!
    ' ik_BM0261
    ' gobjContext.SetComplete
ShareFileExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
ShareFileVbErr:
'    OnErrorDestroyADOObjects adoConnection, adoRecordset, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo ShareFileExit
End Sub
Public Sub IFileVersioningDO_RenameFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo ShareFileVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "RenameFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Create a DB connection and open it
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    ' Now we need to set up some parameters to the SQL stored procedure
    Dim strCreatedBy As String
    strCreatedBy = xmlGetMandatoryAttributeText(vxmlRequestNode, "CREATEDBY")
    Dim strParentFolderGUID As String
    strParentFolderGUID = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERGUID")
    Dim strParentFolderVersion As String
    strParentFolderVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "PARENTFOLDERVERSION")
    Dim strFileGuid As String
    strFileGuid = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEGUID")
    Dim strFileVersion As String
    strFileVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEVERSION")
    ' Need an ADO Command
    Dim adoCommand As New ADODB.Command
    adoCommand.ActiveConnection = adoConnection
    ' Create the SQL string for calling the SQL stored procedure
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVRENAMEFILE(?,?,?,?,?, " _
            & "{resultset 1," & "tableszFILEVERSION" & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
    Else
        strSQL = "USP_FVRENAMEFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
    End If
    ' Create a parameter object and fill it in
    Dim adoParameter As New ADODB.Parameter
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strCreatedBy), strCreatedBy)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderGUID), strParentFolderGUID)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strParentFolderGUID))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strParentFolderVersion), strParentFolderVersion)
    adoCommand.Parameters.Append adoParameter
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        adoCommand.Parameters.Append adoParameter
    Else
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        adoCommand.Parameters.Append adoParameter
    End If
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
    adoCommand.Parameters.Append adoParameter
    ' Create a recordset to execute the command...
    Dim adoRecordSet As New ADODB.Recordset
    Set adoRecordSet = adoCommand.Execute
    ' Get the out parameter
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' adoRecordSet.MoveFirst
    Dim strVersion As String
    strVersion = adoRecordSet.fields.Item(0)
    ' And now close everything down nicely!!
    adoRecordSet.Close
    Set adoRecordSet = Nothing
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    ' Now we have shared the file so now we can return an xml response
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("RENAMEFILE")
    xmlSetAttributeValue xmlNewChildNode, "FILEGUID", strFileGuid
    xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", strFileVersion
    vxmlResponseNode.appendChild xmlNewChildNode
        
    ' And we're done!
    ' ik_BM0261
    ' gobjContext.SetComplete
ShareFileExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
ShareFileVbErr:
'    OnErrorDestroyADOObjects adoConnection, adoRecordset, adoCommand
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo ShareFileExit
End Sub
Public Sub IFileVersioningDO_GetFile( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetFileVbErr
        
    ' ik_BM0261
    gobjContext.SetComplete
    Const cstrFunctionName As String = "GetFile"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    'DR Error Details
    Dim lngErrorNumber As Long
    Dim strErrorSource As String
    Dim strErrorDescription As String
    Dim strFileContentsGuid As String
    Dim strFileContentsType As String
    Dim strFileContents As String
    ' GUID of the file?
    Dim strFileGuid As String
    strFileGuid = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEGUID")
    ' File version?
    Dim strFileVersion As String
    strFileVersion = xmlGetMandatoryAttributeText(vxmlRequestNode, "FILEVERSION")
    
    Dim nDeliveryType As Integer
    nDeliveryType = DELIVERYTYPE_DOC ' Defaults to Word document.
    
    'AS 26/07/2004 CORE286 Get file from archive as a Url and not as bin base 64 string
    Dim strFileContentsUrl As String
    strFileContentsUrl = ""
    
    ' Create an ADO connection
    Dim adoConnection As New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open
    ' Create a new ADO command and set the command to be executed to the SQL command set above
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    Dim adoParameter As ADODB.Parameter
    ' need to index fvfile to get the filecontentsguid
    ' Create the SQL string for calling the SQL stored procedure
    Dim strSQL As String
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strSQL = "{call USP_FVGETFILE(?,?," _
            & "{resultset 1," _
            & "tableszFILECONTENTSGUID" _
            & "})}"
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileGuid), strFileGuid)
        adoCommand.Parameters.Append adoParameter
    Else
        strSQL = "USP_FVGETFILE"
        adoCommand.CommandType = adCmdStoredProc
        adoCommand.CommandText = strSQL
        Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        adoCommand.Parameters.Append adoParameter
    End If
    ' Create some parameters for the function using the values received from the XML stream above
    ' NOTE : These parameters will match the SQL stored procedure parameters, i.e.
    'PROCEDURE LockFile(
    '    szLOCKEDBY IN FVFOLDER.CREATEDBY%TYPE,
    '    rawFILEGUID IN FVFILE.FILEGUID%TYPE,
    '    szFILEVERSION IN FVFILE.FILEVERSION%TYPE,
    '    tableszFILECONTENTSGUID OUT t_tableszFILECONTENTSGUID
    '    )
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
    adoCommand.Parameters.Append adoParameter
    ' Now we have all the parameters set up, execute the command
    Dim adoRecordSet As ADODB.Recordset
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "USP_FVGETFILE call"
    Set adoRecordSet = adoCommand.Execute
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "USP_FVGETFILE return"
    ' Make sure we have received a pointer to the file record
    If adoRecordSet.EOF Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    ' If we are here, means we have a pointer to the correct record in the file table
    ' Get the GUID for the file contents from the record
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' adoRecordSet.MoveFirst
    If adoGetDbProvider = omiga4DBPROVIDEROracle Then
        strFileContentsGuid = adoRecordSet.fields.Item(0).Value
    Else
        strFileContentsGuid = adoGuidToString(adoRecordSet.fields.Item(0).Value)
    End If
    adoRecordSet.Close
    Set adoRecordSet = Nothing
    Set adoCommand.ActiveConnection = Nothing
    Set adoCommand = Nothing
    
    ' Get the file contents GUID from the request XML
    'strFileContentsGuid = xmlGetAttributeText(vxmlRequestNode, "FILECONTENTSGUID")
    ' Create an ADO connection
'    adoConnection.ConnectionString = adoGetDbConnectString
'    adoConnection.CursorLocation = adUseClient
'    adoConnection.Open
    ' And recover the contents
   'BG 01/10/02 SYS5619 Compression code
   
    Dim xmlFileContentsNode As IXMLDOMNode
    Set xmlFileContentsNode = vxmlResponseNode.ownerDocument.createAttribute("FILECONTENTS")
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "ReadFileContents call"
    ReadFileContents adoConnection, strFileContentsGuid, vxmlRequestNode, strFileContentsType, nDeliveryType, xmlFileContentsNode, strFileContentsUrl
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "ReadFileContents return"
    'BG 01/10/02 SYS5619 Compression code - END
    adoConnection.Close
    Set adoConnection = Nothing
    ' Now we have locked the file and read the contents of the file so now we can return an xml response
    Dim xmlNewChildNode As IXMLDOMNode
    Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("GETFILE")
    xmlSetAttributeValue xmlNewChildNode, "FILECONTENTS_TYPE", strFileContentsType
    xmlSetAttributeValue xmlNewChildNode, "DELIVERYTYPE", CStr(nDeliveryType)
    'AS 26/07/2004 CORE286 Get file from archive as a Url and not as bin base 64 string
    If strFileContentsUrl <> "" Then
        xmlSetAttributeValue xmlNewChildNode, "FILECONTENTSURL", strFileContentsUrl
    End If

    'BG 01/10/02 SYS5619 Compression code
    xmlNewChildNode.Attributes.setNamedItem xmlFileContentsNode
    'xmlSetAttributeValue xmlNewChildNode, "FILECONTENTS", strFileContents
    'BG 01/10/02 SYS5619 Compression code - END
    vxmlResponseNode.appendChild xmlNewChildNode
        
    ' And we're done!
    ' ik_BM0261
    ' gobjContext.SetComplete
GetFileExit:
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
GetFileVbErr:
    'Record the error details
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & CstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & lngErrorNumber & ", Err.Source: " & strErrorSource & ", Err.Description: " & strErrorDescription, vbLogEventTypeError
    'Useful debug info
    App.LogEvent CstrClassName & "::" & cstrFunctionName & "(). Debug Information. vxmlRequestNode.xml: " & vxmlRequestNode.xml
    'OnErrorDestroyADOObjects adoConnection
    ' ik_BM0261
    ' gobjContext.SetAbort
    GoTo GetFileExit
End Sub

