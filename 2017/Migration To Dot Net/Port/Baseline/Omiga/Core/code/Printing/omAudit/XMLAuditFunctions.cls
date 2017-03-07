VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "XMLAuditFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      XMLAudtFunctions.cls
'Copyright:     Copyright © 2006 Vertex Financial Services
'Description:
'       Handles all aspects of XML and Template Audit Trail for BBG
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date            Description
'TW     02/02/2006      Created
'TW     01/03/2006      BBGRb235 - Problems with xml data containing £ signs. Also added extra tracing
'TW     08/03/2006      CORE253 - Modification to detect template change
'INR    07/09/2006      CORE299 SEARCHKEY1 is not being populated with the TEMPLATEGROUPID.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Private gobjContext As ObjectContext
Private Const cstrClassName = "omAudit::XMLAuditFunctions"

Private Type TemplateDetails
    HostTemplateName As String
    TemplateGUID As String
    DPSTEMPLATEID As String
    TEMPLATEGROUPID As Long
    HOSTTEMPLATEDESCRIPTION As String
    MINROLELEVEL As Long
    MAXCOPIES As Long
    DEFAULTCOPIES As Long
    CUSTOMERSPECIFICIND As Integer
    RECIPIENTTYPE As Long
    PRINTMENUACCESSIND As Integer
    PRINTERDESTINATIONTYPE As Long
    REMOTEPRINTERLOCATION As String
    INACTIVEINDICATOR As Integer
    PDMMETHOD As String
    EDITBEFOREPRINT As Integer
    VIEWBEFOREPRINT As Integer
    DELIVERYTYPE As Long
    URLPOSTIND As Integer
    FIRSTPAGEPRINTERTRAY As Long
    OTHERPAGESPRINTERTRAY As Long
    Response As String
End Type
Private Function IfNull(ByVal f As ADODB.Field, ByVal Default As Variant)
    If IsNull(f) Then
        IfNull = Default
    Else
        IfNull = f
    End If
End Function

Public Function omRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "omRequest"

Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode
Dim xmlAttrib As IXMLDOMAttribute
    
    On Error GoTo omRequestFailed:
    
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    ' will be superceded by any subsequent SetAbort
    gobjContext.SetComplete
    ' create REQUEST & RESPONSE FreeThreadedDOMDocument40s
    
    Set xmlRequestDoc = CreateXMLDoc()
    Set xmlResponseDoc = CreateXMLDoc()
    
    ' create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")
    
    ' load XML REQUEST
'TW 01/03/2006 BBGRb235
    gobjTrace.TraceXML vstrXMLRequest, "omRequest_Request"
'TW 01/03/2006 BBGRb235 End
    
    Debug.Print vstrXMLRequest
    xmlRequestDoc.loadXML vstrXMLRequest
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlRequestDoc, "REQUEST")
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        DoRequest xmlRequestNode, xmlResponseNode
    End If
    ' if we get here, everything is OK
    omRequest = xmlResponseNode.xml
'TW 01/03/2006 BBGRb235
    gobjTrace.TraceXML xmlResponseNode.xml, "omRequest_Response"
'TW 01/03/2006 BBGRb235 End
    
TidyUp:
    
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    gobjTrace.TraceResponse omRequest
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
omRequestFailed:
    
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
'TW 01/03/2006 BBGRb235
    omRequest = errCreateErrorResponse()
'TW 01/03/2006 BBGRb235 End
    
    gobjTrace.TraceIdentErrorResponse omRequest
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print omRequest
    Resume TidyUp:
End Function

Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Private interface, performs single REQUEST OPERATION
'   Any results of operation (data retrieved, error info etc.) appended to vxmlResponseNode
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       raises oeNotImplemented if OPERATION value not in list of known functions.
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "DoRequest"

Dim strOperation As String
    On Error GoTo DoRequestFailed:
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    strOperation = UCase$(vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text)
    Select Case strOperation
        Case "GETXMLFROMFVS"
            Call GetXMLFromFVS(vxmlRequestNode, vxmlResponseNode)
        Case "SAVEXMLTOFVS"
            Call SaveXMLToFVS(vxmlRequestNode, vxmlResponseNode)
        Case "REHYDRATEDOCUMENT"
            Call ReHydrateDocument(vxmlRequestNode, vxmlResponseNode)
        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, _
        strOperation

    End Select
    
DoRequestFailed:
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Public Function CreateTemplateGUID(ByVal strHostTemplateId As String) As String
' This function saves the template pointed to by the TEMPLATE table
' using the DPSTEMPLATEID in the HOSTTEMPLATE table to locate it.
' The TEMPLATE table is updated with the latest GUID of the stored Template
' in FVS.

Const cstrFunctionName = "CreateTemplateGUID"

Dim objFVS As Object

Dim arrTemplateContents() As Byte

Dim strResponse As String
Dim strTemplate As String
Dim strTemplateID As String
Dim strTemplatePath As String

Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlCreateFileDataNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

Dim xmlFileContentsAttribute As IXMLDOMAttribute

Dim f As Integer
Dim X As Integer

Dim conn As ADODB.Connection

Dim recTemp As ADODB.Recordset

    On Error GoTo CreateTemplateGUIDExit:
    
    Set conn = New ADODB.Connection
    Set recTemp = New ADODB.Recordset
    
    conn.ConnectionString = adoGetDbConnectString
    
    conn.Open
    Set recTemp = conn.Execute("Select DPSTEMPLATEID, FILENAME From HOSTTEMPLATE HT Inner Join TEMPLATE T ON HT.DPSTEMPLATEID = T.TEMPLATEID Where HOSTTEMPLATEID = '" & strHostTemplateId & "'")
    If recTemp.EOF Then
        errThrowError "CreateTemplateGUID", 41000
    Else
        strTemplatePath = recTemp!FileName
        strTemplateID = recTemp!DPSTEMPLATEID
    End If
    
    X = InStrRev(strTemplatePath, "\")
    
    f = FreeFile()
    strTemplate = Space$(FileLen(strTemplatePath))
    
    Open strTemplatePath For Binary Access Read As f
    Get f, , strTemplate
    Close f
    
    arrTemplateContents = strTemplate
    
    Set xmlRequestDoc = CreateXMLDoc()
    Set xmlRequestNode = AddNode(xmlRequestDoc, "REQUEST")
    Call AddAttribute(xmlRequestNode, "USERID", "User")
    Call AddAttribute(xmlRequestNode, "UNITID", "Unit")
    Call AddAttribute(xmlRequestNode, "MACHINENAME", "Machine")
    Call AddAttribute(xmlRequestNode, "CHANNELID", "1")
    Call AddAttribute(xmlRequestNode, "USERAUTHORITYLEVEL", "100")
    Call AddAttribute(xmlRequestNode, "APPLICATIONNUMBER", "None")
    Call AddAttribute(xmlRequestNode, "OPERATION", "CREATEFILE")
    
    Set xmlCreateFileDataNode = AddNode(xmlRequestDoc, "CREATEFILEDATA", xmlRequestNode)
    Call AddAttribute(xmlCreateFileDataNode, "CREATEDBY", "User")
    Call AddAttribute(xmlCreateFileDataNode, "FILENAME", Mid$(strTemplatePath, X + 1))
    Call AddAttribute(xmlCreateFileDataNode, "COMPRESSED", "")
    Call AddAttribute(xmlCreateFileDataNode, "COMPRESSIONMETHOD", "ZLIB")
    Call AddAttribute(xmlCreateFileDataNode, "DELIVERYTYPE", "40")
    Call AddAttribute(xmlCreateFileDataNode, "FILEGUID", "")
    Call AddAttribute(xmlCreateFileDataNode, "FILEVERSION", "V1")
    Call AddAttribute(xmlCreateFileDataNode, "FILESIZE", "0")
    Call AddAttribute(xmlCreateFileDataNode, "FILECONTENTS_TYPE", "BIN.BASE64")
    
    Set xmlFileContentsAttribute = xmlRequestDoc.createAttribute("FILECONTENTS")
    xmlFileContentsAttribute.dataType = "bin.base64"
    xmlFileContentsAttribute.nodeTypedValue = arrTemplateContents
    xmlCreateFileDataNode.Attributes.setNamedItem xmlFileContentsAttribute
    
    Set objFVS = CreateObject("omFVS.FileVersioningBO")
    
    strResponse = objFVS.omRequest(xmlRequestDoc.xml)
    Set xmlResponseDoc = CreateXMLDoc()
    xmlResponseDoc.loadXML strResponse
    
    Set xmlResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
    If GetAttribute(xmlResponseNode, "TYPE") = "SUCCESS" Then
        Set xmlCreateFileDataNode = xmlResponseDoc.selectSingleNode("//RESPONSE/CREATEFILEDATA/@FILEGUID")
        conn.Execute "Update TEMPLATE set TEMPLATEGUID = 0x" & xmlCreateFileDataNode.Text & " WHERE TEMPLATEID = '" & strTemplateID & "'"
        CreateTemplateGUID = xmlCreateFileDataNode.Text
    Else
        CreateTemplateGUID = ""
    End If

TidyUp:
    gobjTrace.TraceXML CreateTemplateGUID, "CreateTemplateGUID_response"
    
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    Set xmlCreateFileDataNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    
    Set xmlFileContentsAttribute = Nothing
    
    Set recTemp = Nothing
    Set conn = Nothing
    
    Set objFVS = Nothing
    Exit Function
    
CreateTemplateGUIDExit:
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    CreateTemplateGUID = errCreateErrorResponse()
    
    gobjTrace.TraceIdentErrorResponse CreateTemplateGUID
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print CreateTemplateGUID
    Resume TidyUp:

End Function

Private Function GetTemplateDetails(ByVal strHostTemplateId As String) As TemplateDetails
' This function obtains details of the Host Template and its related Template
' In addition, it checks to see if the template has been updated since
' it was last stored in FVS. If it has, CreateTemplateGUID  is called to
' save the new version of the template to FVS and get a new GUID

Const cstrFunctionName = "GetTemplateDetails"

Dim LocalTemplateDetails As TemplateDetails

Dim conn As ADODB.Connection
Dim recTemp As ADODB.Recordset

'TW 08/03/2006 CORE253
Dim fs As FileSystemObject

Dim strFileName As String

Dim dtTemplateLastModifiedDate As Date
'TW 08/03/2006 CORE253 End

    On Error GoTo CheckTemplateGUIDExit:
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Set conn = New ADODB.Connection
    Set recTemp = New ADODB.Recordset
    
    Call adoBuildDbConnectionString
    
    conn.ConnectionString = adoGetDbConnectString()
    
    conn.Open
'TW 08/03/2006 CORE253
'    Set recTemp = conn.Execute("Select HT.*, Isnull(DBO.FN_VARBINTOGUIDSTR(TEMPLATEGUID), '') AS TEMPLATEGUID From HOSTTEMPLATE HT Inner Join TEMPLATE T ON HT.DPSTEMPLATEID = T.TEMPLATEID Where HOSTTEMPLATEID = '" & strHostTemplateId & "'")
    Set recTemp = conn.Execute("Select HT.*, Isnull(DBO.FN_VARBINTOGUIDSTR(TEMPLATEGUID), '')AS TEMPLATEGUID, T.FILENAME, DateCreated From HOSTTEMPLATE HT Inner Join TEMPLATE T ON HT.DPSTEMPLATEID = T.TEMPLATEID Left Join FVFile FC On T.TEMPLATEGUID = FC.filecontentsguid Where HOSTTEMPLATEID = '" & strHostTemplateId & "'")
'    If Not recTemp.EOF Then
    If recTemp.EOF Then
        Err.Raise 40004, cstrClassName & "::" & cstrFunctionName, "Host Template record not found"
    Else
        strFileName = IfNull(recTemp!FileName, "")
        If Len(strFileName) = 0 Then
            Err.Raise 40005, cstrClassName & "::" & cstrFunctionName, "Template name missing"
        End If
        Set fs = New FileSystemObject
        If fs.FileExists(strFileName) Then
            dtTemplateLastModifiedDate = fs.GetFile(strFileName).DateLastModified
        Else
            Err.Raise 40006, cstrClassName & "::" & cstrFunctionName, "'" & strFileName & "' missing"
        End If
'        LocalTemplateDetails.TemplateGUID = IfNull(recTemp!TemplateGUID, "")
' Check to see if the template has been modified since it was last stored in FVS
        If dtTemplateLastModifiedDate > IfNull(recTemp!DateCreated, 0) Then
' Force a new version to be stored in FVS
            LocalTemplateDetails.TemplateGUID = CreateTemplateGUID(strHostTemplateId)
        Else
            LocalTemplateDetails.TemplateGUID = IfNull(recTemp!TemplateGUID, "")
        End If
'TW 08/03/2006 CORE253 End
        LocalTemplateDetails.HostTemplateName = IfNull(recTemp!HostTemplateName, "")
        
        LocalTemplateDetails.DPSTEMPLATEID = IfNull(recTemp!DPSTEMPLATEID, "")
        LocalTemplateDetails.TEMPLATEGROUPID = IfNull(recTemp!TEMPLATEGROUPID, 0)
        LocalTemplateDetails.HOSTTEMPLATEDESCRIPTION = IfNull(recTemp!HOSTTEMPLATEDESCRIPTION, "")
        LocalTemplateDetails.MINROLELEVEL = IfNull(recTemp!MINROLELEVEL, 0)
        LocalTemplateDetails.MAXCOPIES = IfNull(recTemp!MAXCOPIES, 0)
        LocalTemplateDetails.DEFAULTCOPIES = IfNull(recTemp!DEFAULTCOPIES, 0)
        LocalTemplateDetails.CUSTOMERSPECIFICIND = IfNull(recTemp!CUSTOMERSPECIFICIND, 0)
        LocalTemplateDetails.RECIPIENTTYPE = IfNull(recTemp!RECIPIENTTYPE, 0)
        LocalTemplateDetails.PRINTMENUACCESSIND = IfNull(recTemp!PRINTMENUACCESSIND, 0)
        LocalTemplateDetails.PRINTERDESTINATIONTYPE = IfNull(recTemp!PRINTERDESTINATIONTYPE, 0)
        LocalTemplateDetails.REMOTEPRINTERLOCATION = IfNull(recTemp!REMOTEPRINTERLOCATION, "")
        LocalTemplateDetails.INACTIVEINDICATOR = IfNull(recTemp!INACTIVEINDICATOR, 0)
        LocalTemplateDetails.PDMMETHOD = IfNull(recTemp!PDMMETHOD, "")
        LocalTemplateDetails.EDITBEFOREPRINT = IfNull(recTemp!EDITBEFOREPRINT, 0)
        LocalTemplateDetails.VIEWBEFOREPRINT = IfNull(recTemp!VIEWBEFOREPRINT, 0)
        LocalTemplateDetails.DELIVERYTYPE = IfNull(recTemp!DELIVERYTYPE, 20)
        LocalTemplateDetails.URLPOSTIND = IfNull(recTemp!URLPOSTIND, 0)
        LocalTemplateDetails.FIRSTPAGEPRINTERTRAY = IfNull(recTemp!FIRSTPAGEPRINTERTRAY, 0)
        LocalTemplateDetails.OTHERPAGESPRINTERTRAY = IfNull(recTemp!OTHERPAGESPRINTERTRAY, 0)

    End If
    conn.Close

TidyUp:
'TW 08/03/2006 CORE253
    Set fs = Nothing
'TW 08/03/2006 CORE253 End

    GetTemplateDetails = LocalTemplateDetails
    Set recTemp = Nothing
    Set conn = Nothing
    Exit Function
    
CheckTemplateGUIDExit:
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    LocalTemplateDetails.Response = errCreateErrorResponse()
    
    gobjTrace.TraceIdentErrorResponse LocalTemplateDetails.Response
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print LocalTemplateDetails.Response
    Resume TidyUp:

End Function



Private Sub GetXMLFromFVS(ByVal vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)
Const cstrFunctionName As String = "GetXMLFromFVS"

Dim strDocumentGUID As String
Dim strFileContents As String
Dim strFileGUID As String
Dim strFileVersion As String
Dim strResponse As String

Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlDocumentContentsNode As IXMLDOMNode
Dim xmlGetCriteriaNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

Dim omPM As Object

    Call StartAuditTrace
    On Error GoTo GetXMLFromFVSFailed:
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Set xmlRequestDoc = CreateXMLDoc()
    
    xmlRequestDoc.loadXML vxmlRequestNode.xml
    
    gobjTrace.TraceXML vxmlRequestNode.xml, "GetXMLFromFVS_request"
    
    Set omPM = CreateObject("omPM.PrintManagerBO")
    
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("//REQUEST")
    Call AddAttribute(xmlRequestNode, "OPERATION", "GETDOCUMENTARCHIVE")
    strFileGUID = GetAttribute(xmlRequestNode, "FILEGUID")
    strFileVersion = GetAttribute(xmlRequestNode, "FILEVERSION")
    
    Set xmlGetCriteriaNode = AddNode(xmlRequestDoc, "GETCRITERIA", xmlRequestNode)
    Call AddAttribute(xmlGetCriteriaNode, "EVENTDATE", Format$(Now, "dd/mm/yyyy hh:nn"))
    Call AddAttribute(xmlGetCriteriaNode, "FILEGUID", strFileGUID)
    Call AddAttribute(xmlGetCriteriaNode, "FILEVERSION", strFileVersion)

    strResponse = omPM.omRequest(xmlRequestDoc.xml)
    Set xmlResponseDoc = CreateXMLDoc()
    xmlResponseDoc.loadXML strResponse
    
    Set xmlResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
    If GetAttribute(xmlResponseNode, "TYPE") = "SUCCESS" Then
        Set xmlDocumentContentsNode = xmlResponseDoc.selectSingleNode("//RESPONSE/DOCUMENTCONTENTS/@FILECONTENTS")
        xmlDocumentContentsNode.dataType = "bin.base64"
        
        xmlResponseDoc.loadXML xmlDocumentContentsNode.nodeTypedValue
        
        Set vxmlResponseNode = xmlResponseDoc.documentElement
    
    Else
        Set vxmlResponseNode = xmlResponseDoc
    End If

TidyUp:
    gobjTrace.TraceXML vxmlResponseNode.xml, "GetXMLFromFVS_response"
    
    Set xmlResponseDoc = Nothing
    Set xmlRequestDoc = Nothing
    
    Set xmlDocumentContentsNode = Nothing
    Set xmlGetCriteriaNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    
    Set omPM = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Sub
    
GetXMLFromFVSFailed:
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    
    gobjTrace.TraceIdentErrorResponse vxmlResponseNode.xml
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print vxmlResponseNode.xml
    Resume TidyUp:

End Sub



Private Sub ReHydrateDocument(ByVal vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)
' This function perform the following steps
'   1 Retrieve the xml from FVS
'   2 Retrieve the related template from FVS
'   3 Create a printable document from the xml and rtf retrieved in steps 1 and 2

Const cstrFunctionName = "ReHydrateDocument"
Dim arrData() As Byte

Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlTemplateDoc As FreeThreadedDOMDocument40

Dim xmlGetCriteriaNode As IXMLDOMNode
Dim xmlDocumentContentsNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode
Dim xmlTemplateNode As IXMLDOMNode

Dim conn As ADODB.Connection
Dim recTemp As ADODB.Recordset

Dim strDocumentGUID As String
Dim strFileGUID As String
Dim strFileVersion As String
Dim strHostTemplateId As String
Dim strInput As String
Dim strOutputType As String
Dim strResponse As String
Dim strTemplate As String
Dim strTemplateGUID As String

Dim objCompression As New dmsCompression1

Dim kfiObject As Object
Dim omPM As Object

    On Error GoTo ReHydrateDocumentFailed:
    Call StartAuditTrace
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName

' Retrieve the xml from FVS
    strDocumentGUID = GetAttribute(vxmlRequestNode, "DOCUMENTGUID")
    
    Set xmlResponseDoc = CreateXMLDoc()
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")
    
    Set xmlRequestDoc = CreateXMLDoc()
    xmlRequestDoc.loadXML vxmlRequestNode.xml
    
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("//REQUEST")
    Call AddAttribute(xmlRequestNode, "OPERATION", "GETDOCUMENTARCHIVE")
    strFileGUID = GetAttribute(xmlRequestNode, "FILEGUID")
    strFileVersion = GetAttribute(xmlRequestNode, "FILEVERSION")
    strOutputType = LCase$(GetAttribute(xmlRequestNode, "OUTPUTTYPE"))
    
    Set xmlGetCriteriaNode = AddNode(xmlRequestDoc, "GETCRITERIA", xmlRequestNode)
    Call AddAttribute(xmlGetCriteriaNode, "EVENTDATE", Format$(Now, "dd/mm/yyyy hh:nn"))
    Call AddAttribute(xmlGetCriteriaNode, "FILEGUID", strFileGUID)
    Call AddAttribute(xmlGetCriteriaNode, "FILEVERSION", strFileVersion)

    Set omPM = CreateObject("omPM.PrintManagerBO")
    strResponse = omPM.omRequest(xmlRequestDoc.xml)
    Set xmlResponseDoc = CreateXMLDoc()
    xmlResponseDoc.loadXML strResponse
    
    Set xmlResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
    If GetAttribute(xmlResponseNode, "TYPE") = "SUCCESS" Then
        Set xmlDocumentContentsNode = xmlResponseDoc.selectSingleNode("//RESPONSE/DOCUMENTCONTENTS/@FILECONTENTS")
        xmlDocumentContentsNode.dataType = "bin.base64"
        
        xmlResponseDoc.loadXML xmlDocumentContentsNode.nodeTypedValue
        
    Else
        Set vxmlResponseNode = xmlResponseDoc
    End If
    
' 2 Retrieve the related template from FVS
    Set xmlTemplateDoc = CreateXMLDoc()
    xmlTemplateDoc.loadXML xmlResponseDoc.documentElement.xml
    Set xmlTemplateNode = xmlTemplateDoc.selectSingleNode("//TEMPLATEDATA")
    strTemplateGUID = GetAttribute(xmlTemplateNode, "TEMPLATEGUID")
    strInput = xmlTemplateNode.xml
    
    Set conn = New ADODB.Connection
    Set recTemp = New ADODB.Recordset

    Call adoBuildDbConnectionString

    conn.ConnectionString = adoGetDbConnectString()

    conn.Open

    Set recTemp = conn.Execute("select FileContents from fvfilecontents where filecontentsguid = 0x" & strTemplateGUID)
    arrData = recTemp!FileContents
    
' Set up the Output Type (Default to pdf)
    Select Case strOutputType
        Case "rtf", "pdf", "doc"
        Case Else
            strOutputType = "pdf"
    End Select
'
    Set xmlResponseDoc = CreateXMLDoc()
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")
    
    objCompression.CompressionMethod = "ZLIB"
    strTemplate = objCompression.SafeArrayDecompressToBSTR(arrData)

    conn.Close
' Create a printable document from the xml and template
    Set kfiObject = CreateObject("eKFIEngine.eKFIDelivery")

    Set xmlRequestDoc = CreateXMLDoc()
            
    Set xmlRequestNode = AddNode(xmlRequestDoc, "REQUEST")
    Call AddAttribute(xmlRequestNode, "OutputType", strOutputType)
            
    Set xmlNode = AddNode(xmlRequestDoc, "TEMPLATE", xmlRequestNode)
    xmlNode.Text = Replace(strTemplate, vbCrLf, "")
    
    Set xmlNode = AddNode(xmlRequestDoc, "INPUT", xmlRequestNode)
    xmlNode.Text = strInput
            
    Set xmlResponseDoc = kfiObject.CreateKFI(xmlRequestDoc)
    
    Set vxmlResponseNode = xmlResponseDoc.documentElement

TidyUp:
    
    Set recTemp = Nothing
    Set conn = Nothing
    
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTemplateDoc = Nothing
    
    Set xmlDocumentContentsNode = Nothing
    Set xmlGetCriteriaNode = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTemplateNode = Nothing

    Set kfiObject = Nothing
    Set omPM = Nothing

    Exit Sub
    
ReHydrateDocumentFailed:

    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    
    gobjTrace.TraceIdentErrorResponse vxmlResponseNode.xml
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print vxmlResponseNode.xml
    Resume TidyUp:
End Sub

Private Sub SaveXMLToFVS(ByVal vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

Const cstrFunctionName As String = "SaveXMLToFVS"
Dim arrBytes() As Byte

Dim strApplicationNumber As String
Dim strHostTemplateId As String
Dim strResponse As String
Dim strTemplateGUID As String

Dim omPM As Object

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlControlDataNode As IXMLDOMNode
Dim xmlPrintDataNode As IXMLDOMNode
Dim xmlPrintDocumentDataNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlTemplateDataNode As IXMLDOMNode
Dim xmlTemplateGUIDAttribute As IXMLDOMAttribute

Dim xmlFileContentsAttribute As IXMLDOMAttribute
    
Dim LocalTemplateDetails As TemplateDetails
    
    On Error GoTo SaveXMLToFVSFailed:
'TW 01/03/2006 BBGRb235
    Set xmlResponseDoc = CreateXMLDoc()
'TW 01/03/2006 BBGRb235 End

    Call StartAuditTrace
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName

    Set xmlDoc = CreateXMLDoc()
    Set xmlRequestDoc = CreateXMLDoc()
    
    xmlRequestDoc.loadXML vxmlRequestNode.xml
    
' Set up the REQUEST node to save the data to FVS
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("//REQUEST")
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strHostTemplateId = GetAttribute(xmlRequestNode, "HOSTTEMPLATEID")
    
' Extract the TEMPLATEDATA node from the xml to be saved.
' An error will occur if it is not found

    xmlDoc.loadXML GetAttribute(xmlRequestNode, "XMLPAYLOAD")
'TW 01/03/2006 BBGRb235
    If Len(xmlDoc.xml) = 0 Then
        Err.Raise 40001, cstrClassName & "::" & cstrFunctionName, "XMLPAYLOAD is missing or empty"
    End If
'TW 01/03/2006 BBGRb235 End
    
    Set xmlTemplateDataNode = xmlDoc.selectSingleNode("//TEMPLATEDATA")

'TW 01/03/2006 BBGRb235
    If xmlTemplateDataNode Is Nothing Then
        Err.Raise 40002, cstrClassName & "::" & cstrFunctionName, "No TEMPLATEDATA node in XMLPAYLOAD"
    End If
'TW 01/03/2006 BBGRb235 End
    
' Get the GUID for the relevant TEMPLATE
    LocalTemplateDetails = GetTemplateDetails(strHostTemplateId)
    strTemplateGUID = LocalTemplateDetails.TemplateGUID
    
    If Len(strTemplateGUID) = 0 Then
'TW 08/03/2006 CORE253
'        strTemplateGUID = CreateTemplateGUID(strHostTemplateId)
        Err.Raise 40003, cstrClassName & "::" & cstrFunctionName, "Can't create GUID for template"
'TW 08/03/2006 CORE253 End
    End If
    
' Check to see if there is already a Template GUID in the XML
    Set xmlTemplateGUIDAttribute = xmlTemplateDataNode.Attributes.getNamedItem("TEMPLATEGUID")
    
    If xmlTemplateGUIDAttribute Is Nothing Then
' Template GUID attribute did not exist - create it
        Call AddAttribute(xmlTemplateDataNode, "TEMPLATEGUID", strTemplateGUID)
    Else
' Template GUID attribute exists - update it
        xmlTemplateGUIDAttribute.Text = strTemplateGUID
    End If
    arrBytes = xmlTemplateDataNode.xml
    
' Set up a request block for omPM to save the XML to FVS
    Call AddAttribute(xmlRequestNode, "OPERATION", "SAVEPRINTDOCUMENT")
    
    Set xmlControlDataNode = AddNode(xmlRequestDoc, "CONTROLDATA", xmlRequestNode)
    Call AddAttribute(xmlControlDataNode, "DOCUMENTID", strHostTemplateId)
    Call AddAttribute(xmlControlDataNode, "DPSDOCUMENTID", LocalTemplateDetails.DPSTEMPLATEID)
    Call AddAttribute(xmlControlDataNode, "DELIVERYTYPE", "40")
    Call AddAttribute(xmlControlDataNode, "DELIVERYENGINETYPE", "2")
    Call AddAttribute(xmlControlDataNode, "DESTINATIONTYPE", "W")
    
    Call AddAttribute(xmlControlDataNode, "STAGEID", "0")        '??
    Call AddAttribute(xmlControlDataNode, "DOCUMENTTYPE", "0")   '??
    
    Call AddAttribute(xmlControlDataNode, "TEMPLATEGROUPID", LocalTemplateDetails.TEMPLATEGROUPID)
    'CORE299
    Call AddAttribute(xmlControlDataNode, "SEARCHKEY1", LocalTemplateDetails.TEMPLATEGROUPID)
    Call AddAttribute(xmlControlDataNode, "HOSTTEMPLATENAME", LocalTemplateDetails.HostTemplateName)
    Call AddAttribute(xmlControlDataNode, "HOSTTEMPLATEDESCRIPTION", LocalTemplateDetails.HOSTTEMPLATEDESCRIPTION)

    Set xmlPrintDataNode = AddNode(xmlRequestDoc, "PRINTDATA", xmlRequestNode)
    Call AddAttribute(xmlPrintDataNode, "APPLICATIONNUMBER", strApplicationNumber)

    Set xmlPrintDocumentDataNode = AddNode(xmlRequestDoc, "PRINTDOCUMENTDATA", xmlRequestNode)

    Call AddAttribute(xmlPrintDocumentDataNode, "COMPRESSED", "")
    Call AddAttribute(xmlPrintDocumentDataNode, "COMPRESSIONMETHOD", "ZLIB")
    Call AddAttribute(xmlPrintDocumentDataNode, "DELIVERYTYPE", "40")
    Call AddAttribute(xmlPrintDocumentDataNode, "FILEGUID", "")
    Call AddAttribute(xmlPrintDocumentDataNode, "FILEVERSION", "")
    Call AddAttribute(xmlPrintDocumentDataNode, "FILESIZE", "0")
'TW 08/03/2006 CORE253
'    Call AddAttribute(xmlPrintDocumentDataNode, "FILENAME", "Test of save xml")
    Call AddAttribute(xmlPrintDocumentDataNode, "FILENAME", strApplicationNumber)
'TW 08/03/2006 CORE253 End
    Call AddAttribute(xmlPrintDocumentDataNode, "HOSTTEMPLATEID", strHostTemplateId)

    Set xmlFileContentsAttribute = xmlRequestDoc.createAttribute("FILECONTENTS")
    xmlFileContentsAttribute.dataType = "bin.base64"
    xmlFileContentsAttribute.nodeTypedValue = arrBytes
    xmlPrintDocumentDataNode.Attributes.setNamedItem xmlFileContentsAttribute

    gobjTrace.TraceXML vxmlRequestNode.xml, "SaveXMLToFVS_request"
    
    Set omPM = CreateObject("omPM.PrintManagerBO")
    
    strResponse = omPM.omRequest(xmlRequestDoc.xml)
'TW 01/03/2006 BBGRb235
'    Set xmlResponseDoc = CreateXMLDoc()
'    xmlResponseDoc.loadXML strResponse

'    Set vxmlResponseNode = xmlResponseDoc.documentElement
'TW 01/03/2006 BBGRb235 End

TidyUp:
'TW 01/03/2006 BBGRb235
    xmlResponseDoc.loadXML strResponse
    Set vxmlResponseNode = xmlResponseDoc.documentElement
    
'TW 01/03/2006 BBGRb235 End
    gobjTrace.TraceXML vxmlResponseNode.xml, "SaveXMLToFVS_response"
    
    Set xmlDoc = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    Set xmlControlDataNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlPrintDocumentDataNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlTemplateDataNode = Nothing
    
    Set xmlFileContentsAttribute = Nothing
    Set xmlTemplateGUIDAttribute = Nothing
    
    Set omPM = Nothing
    
    Exit Sub
SaveXMLToFVSFailed:

'TW 01/03/2006 BBGRb235
    strResponse = errCreateErrorResponse()
'TW 01/03/2006 BBGRb235 End

    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    
'TW 01/03/2006 BBGRb235
'    gobjTrace.TraceIdentErrorResponse vxmlResponseNode.xml
'    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
'    Debug.Print vxmlResponseNode.xml
    gobjTrace.TraceIdentErrorResponse strResponse
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print strResponse
'TW 01/03/2006 BBGRb235 End
    Resume TidyUp:
End Sub

Private Sub AddAttribute(ByVal xmlNode As IXMLDOMNode, ByVal strAttributeName As String, ByVal strAttributeValue As String)
Dim xmlAttribute As IXMLDOMAttribute

    Set xmlAttribute = xmlNode.ownerDocument.createAttribute(strAttributeName)
    xmlAttribute.Text = strAttributeValue
    xmlNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = Nothing
End Sub

Private Function AddNode(ByVal xmlDoc As FreeThreadedDOMDocument40, _
                         ByVal strNodeName As String, _
                         Optional xmlParentObject As IXMLDOMNode) As IXMLDOMNode
                         
Dim xmlTempNode As IXMLDOMNode

    Set xmlTempNode = xmlDoc.createNode(NODE_ELEMENT, strNodeName, "")
    If xmlParentObject Is Nothing Then
        xmlDoc.appendChild xmlTempNode
    Else
        xmlParentObject.appendChild xmlTempNode
    End If
    Set AddNode = xmlTempNode
    
    Set xmlTempNode = Nothing
End Function

Private Function GetAttribute(ByVal xmlNode As IXMLDOMNode, ByVal strAttributeName As String) As String
Dim xmlAttribute As IXMLDOMAttribute
    Set xmlAttribute = xmlNode.Attributes.getNamedItem(strAttributeName)
    If Not xmlAttribute Is Nothing Then
        GetAttribute = xmlAttribute.Text
    End If
End Function


Private Function CreateXMLDoc() As FreeThreadedDOMDocument40
Dim xmlTempDoc As FreeThreadedDOMDocument40

    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.async = False
    xmlTempDoc.setProperty "NewParser", True
    Set CreateXMLDoc = xmlTempDoc
    
    Set xmlTempDoc = Nothing
End Function


Private Sub StartAuditTrace()
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
    End If
    If gobjTrace Is Nothing Then
        Set gobjTrace = New traceAssist
    End If
    gobjTrace.StartTrace

End Sub

Private Sub ObjectControl_Activate()
    Call StartAuditTrace
End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub


