VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PackManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Workfile:      PackControl.cls
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:
'       Handles Fulfilment Pack Functions.
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date            Description
'TW     25/08/2005      Created
'TW     19/10/2005      MAR237
'TW     21/10/2005      MAR260 - Corrections for sending to fulfilment
'TW     25/10/2005      MAR294 Changes for fulfilment - CreateAuditTrail
'TW     27/10/2005      MAR211 Changes to handle multiple quotes
'TW     15/11/2005      MAR515 Changes to which documents are sent to FileNet
'                       and representation in Fullfilment
'TW     18/11/2005      MAR579 New method to send a document to FileNet and create audit and event records
'                              New method to get a document from FileNet
'PSC    14/12/2005      MAR802 Make SendPackToEmail and SendPackToSMS customer specific
'PSC    15/12/2005      MAR880 Change to use correct Filenet userid
'BC     16/12/2005      MAR880 Handle errors returned from ProcessPackMember
'HMA    09/01/2005      MAR978 Do not reformat PackID GUID before sending to Fulfilment
'                              Send PackName to Fulfilment
'PSC    11/01/2006      MAR994 Change to use Filenet userid from global parameter
'BC     24/01/2006      MAR1106 Include ApplicationNumber and Quotattion Number in DOCUMENTNAME
'BC     01/02/2006      MAR1106 Also incluse Date/Timestamp
'BC     01/02/2006      MAR1155 Check xmlPackDefinitionDocfor "USEMULTIPLEQUOTES" attribute to process multiple KFI Quotes
'RF     03/02/2006      MAR1191 Add document name to fulfilment request
'PSC    06/02/2006      MAR1197 Add content type for sending to and retrieval from Filenet
'RF     07/02/2006      MAR1216 Ensure document name saved in DMS105
'PE     13/02/02006     MAR1155 USE MULTIPLE QUOTES ON PACK DOES INVOKE CREATING MULTIPLE KFI ON SEND KFI PACK
'BC     14/02/2006      MAR1271 Backout MAR1155
'PSC    15/02/2006      MAR1271 Correct fix for 1155
'DRC    23/02/2006      MAR1193 Put Country into Customer address and check for Scottish Property
'GHun   03/03/2006      MAR1332 Make eventkey consistent
'PE     13/04/2006      MAR1405 CR 141 - ENHANCEMENT KFI document location in FileNet CE
'PE     18/04/2006      MAR1405 CR 141 - ENHANCEMENT KFI document location in FileNet CE
'PSC    20/04/2006      MAR1647 Default document class if not present
'PSC    05/05/2006      MAR1593 Amend SendPackToSMS and SendPackToEmail to cater for multiple customers and
'                               the contact preference
'PE     11/05/2006      MAR1756 Implement Pack Destination (Customer, Solicitor, Solicitor & Customer, New Property Address)
'HMA    12/05/2006      MAR1638 Check for errors returned from omFDM.
'                               Send Customer Number for Cancel and Re-send Pack.
'HMA    23/05/2006      MAR1828 Get Delivery Type from template.
'GHun   07/06/2006      MAR1819 Performance improvements and minor tidy up
'TW     12/06/2006      MAR1867 Performance improvements
'HMA    04/07/2005      MAR1867 Further changes to create audit details correctly.
'TW     12/06/2006      MAR1897 Add quote number to DocumentAuditDetails 'DOCUMENTNAME'
'RF     10/07/2007      MAR1900 ISSUE OFFER TIMEOUT ON NEW OMPACK - Change back to UsesTransaction
'                               (required for updating QUOTATION in SendPack).
'PSC    25/09/2006      MAR1958 Don't get passwords if using windows authentication
'AW     22/12/2006      EP1271  Amendments to SendPack for CC56
'                               Removed eventkey enumeration, now using definitions in constants.bas
'AW     23/12/2006      EP1271  Amended CreateApplicationPack
'AS     27/12/2006      EP1271  More changes.
'AS     08/01/2007      EP1277  Added pack description to SendPack errors.
'AW     18/01/2007      EP1304  Amended GetAddress for Broker/Packager destination packs
'AW     03/04/2007      EP2_2181    Amended SendPack() for alternative document inclusion in packs
'AW     07/04/2007      EP2_2181    Report missing document list
'------------------------------------------------------------------------------------------

Implements ObjectControl
Private gobjContext As ObjectContext

'TW 25/10/2005 MAR294
Enum eventKey
    Created = 0
    Edited = 1
    Viewed = 2
    Reprinted = 3
    FulfilmentSendSuccess = 4
    FulfilmentReSendSuccess = 5
    FulfilmentCancelSuccess = 6
    SMSSuccess = 7
    EmailSuccess = 8
    Recategorisation = 9    'MAR1332 GHun
End Enum
'TW 25/10/2005 MAR294 End

#If FILENET Then
Private Const DEFAULT_FILEVERSION = "1"
Private Const DEFAULT_GEMINIPRINTSTATUS = 0     ' Unknown/Not applicable
#ElseIf Gemini Then
Private Const DEFAULT_FILEVERSION = "V1"
Private Const DEFAULT_GEMINIPRINTSTATUS = GEMINIPRINTSTATUS_APPROVED
#End If

Private Const DEFAULT_FILECONTENTS_TYPE = "BIN.BASE64"

'TW 12/06/2006 MAR1867
Dim objOmFDM As Object
Dim strZeroes As String
'TW 12/06/2006 MAR1867 End

Private Function CreateDocumentAuditDetails(ByVal strApplicationNumber As String, _
                                            ByVal strHostTemplateId As String, _
                                            ByVal strQuoteNumber As String, _
                                            Optional ByVal geminiPrintStatus As Integer = DEFAULT_GEMINIPRINTSTATUS) As String

Const strFunctionName As String = "CreateDocumentAuditDetails"

' TW 07/07/2006 MAR1897 - was
'Private Function CreateDocumentAuditDetails(ByVal strApplicationNumber As String, _
'                                            ByVal strHostTemplateId As String) As String

' Create DocumentAuditDetails record

Dim strDocumentGuid As String
' TW 07/07/2006 MAR1897
Dim strQuotationNumber As String
' TW 07/07/2006 MAR1897 End
    
Dim connOmiga As ADODB.Connection
Dim cmd As ADODB.Command

    On Error GoTo CreateDocumentAuditDetailsError:
    
    Set connOmiga = GetConnection()

    strDocumentGuid = CreateGUID()
    
' TW 07/07/2006 MAR1897
' If no Quote Number supplied, default to 1
    If Len(strQuoteNumber) = 0 Then
        strQuotationNumber = "1"
    Else
        strQuotationNumber = strQuoteNumber
    End If
' TW 07/07/2006 MAR1897 End

    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "Writing DocumentAuditDetails"

    Set cmd = New ADODB.Command
        
    'MAR1867 Use With
    With cmd
        Set .ActiveConnection = connOmiga
        .CommandType = adCmdStoredProc
        .CommandText = "usp_CreateDocumentAuditDetails"
        .Parameters.Append .CreateParameter("@ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        .Parameters.Append .CreateParameter("@DocumentGUID", adVarBinary, adParamInput, 16, GuidStringToByteArray(strDocumentGuid))
        .Parameters.Append .CreateParameter("@HostTemplateID", adBSTR, adParamInput, Len(strHostTemplateId), strHostTemplateId)
' TW 07/07/2006 MAR1897
        .Parameters.Append .CreateParameter("@QuotationNumber", adBSTR, adParamInput, Len(strQuotationNumber), strQuotationNumber)
' TW 07/07/2006 MAR1897 End
#If Gemini Then
        .Parameters.Append .CreateParameter("@GeminiPrintStatus", adInteger, adParamInput, 5, geminiPrintStatus)
#End If

        .Execute , , adExecuteNoRecords
        
        Set .ActiveConnection = Nothing
    End With
    
    CreateDocumentAuditDetails = strDocumentGuid
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "DocumentAuditDetails written"

TidyUp:
    Set connOmiga = Nothing
    Exit Function

CreateDocumentAuditDetailsError:
    App.LogEvent "PackControl.CreateDocumentAuditDetails " & Err.Description, vbLogEventTypeError
    Resume TidyUp

End Function

Private Sub CreateEventAuditDetail(ByVal strApplicationNumber, _
                                    ByVal intEventKey As Integer, _
                                    ByVal strHostTemplateId As String, _
                                    ByVal strDocumentGuid As String, _
                                    ByVal strPackGUID As String, _
                                    ByVal strFileNetImageRef As String, _
                                    ByVal strUnitId As String, _
                                    ByVal strUserId As String, _
                                    Optional ByVal strFileGuid As String = "", _
                                    Optional ByVal strFileVersion As String = DEFAULT_FILEVERSION)

' Create EventAuditDetail record
Const strFunctionName = "CreateEventAuditDetail"
    
Dim connOmiga As ADODB.Connection
Dim cmd As ADODB.Command

    On Error GoTo CreateEventAuditDetailError:
    
    'MAR1867
    If (strFileGuid = "") Then
        strFileGuid = strZeroes
    End If
    
    Set connOmiga = GetConnection()

    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "Writing EventAuditDetail"

    'MAR1867
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = connOmiga
        .CommandType = adCmdStoredProc
        .CommandText = "usp_CreateEventAuditDetail"
        .Parameters.Append .CreateParameter("@ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        .Parameters.Append .CreateParameter("@UnitID", adBSTR, adParamInput, Len(strUnitId), strUnitId)
        .Parameters.Append .CreateParameter("@UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
        .Parameters.Append .CreateParameter("@DocumentGUID", adVarBinary, adParamInput, 16, GuidStringToByteArray(strDocumentGuid))
        .Parameters.Append .CreateParameter("@PackGUID", adVarBinary, adParamInput, 16, GuidStringToByteArray(strPackGUID))
        .Parameters.Append .CreateParameter("@FileNetImageRef", adVarBinary, adParamInput, 16, GuidStringToByteArray(strFileNetImageRef))
        .Parameters.Append .CreateParameter("@HostTemplateID", adBSTR, adParamInput, Len(strHostTemplateId), strHostTemplateId)
        .Parameters.Append .CreateParameter("@EventKey", adInteger, adParamInput, 5, intEventKey)
        .Parameters.Append .CreateParameter("@FileGUID", adVarBinary, adParamInput, 16, GuidStringToByteArray(strFileGuid))
        .Parameters.Append .CreateParameter("@FileVersion", adBSTR, adParamInput, Len(strFileVersion), strFileVersion)
        .Execute , , adExecuteNoRecords
        Set .ActiveConnection = Nothing
    End With
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "EventAuditDetail written"

TidyUp:
    Set connOmiga = Nothing
    Exit Sub
    
CreateEventAuditDetailError:
    App.LogEvent "PackControl.CreateEventAuditDetail " & Err.Description, vbLogEventTypeError
    Resume TidyUp
End Sub

Private Sub DealWithErrorResponse(ByRef xmlResponseDoc As FreeThreadedDOMDocument40, ByVal strSource As String, ByRef ErrorResponse As ErrObject)
Dim xmlResponseNode As IXMLDOMNode
Dim xmlErrorNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode

    Set xmlResponseNode = xmlResponseDoc.selectSingleNode("//RESPONSE")
    If xmlResponseNode Is Nothing Then
        Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")
    End If
    Call AddAttribute(xmlResponseNode, "TYPE", "ERROR")
    Call AddAttribute(xmlResponseNode, "DESCRIPTION", ErrorResponse.Description)
    
    Set xmlErrorNode = AddNode(xmlResponseDoc, "ERROR", xmlResponseNode)
    Set xmlNode = AddNode(xmlResponseDoc, "NUMBER", xmlErrorNode)
    xmlNode.Text = ErrorResponse.Number
    Set xmlNode = AddNode(xmlResponseDoc, "DESCRIPTION", xmlErrorNode)
    xmlNode.Text = ErrorResponse.Description
    Set xmlNode = AddNode(xmlResponseDoc, "SOURCE", xmlErrorNode)
    xmlNode.Text = ErrorResponse.Source
    Set xmlNode = AddNode(xmlResponseDoc, "VERSION", xmlErrorNode)
    xmlNode.Text = App.Comments
    
    App.LogEvent "PackControl." & strSource & " " & ErrorResponse.Description, vbLogEventTypeError
End Sub

Private Function DoRequest(ByVal strFunctionName As String, _
                                     ByVal strApplicationNumber As String, _
                                     ByVal strPackId As String, _
                                     ByVal strUserId As String, _
                                     ByVal xmlNode As IXMLDOMNode, _
                                     Optional ByVal vstrDocumentName As String) As String ' MAR1191

'TW 12/06/2006 MAR1867 Modifications for performance reasons

Dim fso As Scripting.FileSystemObject

Dim boolTraceOn As Boolean
Dim objWshShell As Object

Dim strPassword As String
Dim strRegValue As String
Dim strResponse As String
Dim strTraceFolder As String

    
    On Error Resume Next
    Set objWshShell = CreateObject("WScript.Shell")
    strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\")
    If strRegValue = "1" Or strRegValue = "2" Then
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\FOLDER")
        If Len(strRegValue) > 0 Then
            Set fso = New Scripting.FileSystemObject
            If fso.FolderExists(strRegValue) Then
                boolTraceOn = True
                strTraceFolder = strRegValue & "\"
            End If
        End If
    End If
    On Error GoTo DoRequestError:
    
    ' PSC 25/09/2006 MAR1958 - Start
    ' Get global parameter to see if we need the password
    Dim strSecurityType As String
    strSecurityType = GetGlobalParamString("SecurityCredentialsType")
    
    If StrComp(strSecurityType, "WINDOWSAUTHENTICATION", vbTextCompare) <> 0 Or _
       StrComp(strFunctionName, "getFileNetGUID", vbTextCompare) = 0 Or _
       StrComp(strFunctionName, "getFileNetRecord", vbTextCompare) = 0 Then
        strPassword = GetPassword(strUserId)
    Else
        strPassword = ""
    End If
    ' PSC 25/09/2006 MAR1958 - End
    
    If objOmFDM Is Nothing Then
        Err.Raise 40018, "PackControl." & strFunctionName, "Can't create omFDM.FileNetInterfaceBO"
    End If
    
    objOmFDM.traceOn = boolTraceOn
    objOmFDM.traceLocation = strTraceFolder
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "objOmFDM." & strFunctionName & "  call"
    
    Select Case strFunctionName
        Case "getFileNetGUID"
            ' MAR1191
            ' PSC 06/02/2006 MAR1197
            strResponse = objOmFDM.GetFileNetGUID( _
                strApplicationNumber, _
                xmlNode.Attributes.getNamedItem("FILECONTENTS").nodeTypedValue, _
                strUserId, strPassword, vstrDocumentName, _
                xmlGetAttributeText(xmlNode, "CONTENTTYPE"), xmlGetAttributeText(xmlNode, "DOCUMENTCLASS"))
        Case "initialRequest"
            strResponse = objOmFDM.initialRequest(strUserId, strPassword, xmlNode.xml)
        Case "cancelPack"
            strResponse = objOmFDM.CancelPack(strPackId, strUserId, strPassword, _
                            xmlGetAttributeText(xmlNode, "OTHERSYSTEMCUSTOMERNUMBER"))    ' MAR1638
        Case "reSendPack"
            strResponse = objOmFDM.ReSendPack(strPackId, strUserId, strPassword, _
                            xmlGetAttributeText(xmlNode, "OTHERSYSTEMCUSTOMERNUMBER"))    ' MAR1638
        Case "sendToSMS"
            strResponse = objOmFDM.sendToSMS(strUserId, strPassword, xmlNode.xml)
        Case "sendToEmail"
            strResponse = objOmFDM.sendToEmail(strUserId, strPassword, xmlNode.xml)
        Case "sendDocumentToFileNet"
            strResponse = objOmFDM.SendDocumentToFileNet( _
                strUserId, strPassword, xmlNode.xml, vstrDocumentName) ' MAR1191
        Case "getFileNetRecord"
            strResponse = objOmFDM.GetFileNetRecord(strUserId, strPassword, strApplicationNumber, strPackId)
    End Select
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "objOmFDM." & strFunctionName & " Response : " & strResponse
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.3]", "objOmFDM." & strFunctionName & " return"

TidyUp:
    DoRequest = strResponse
    Set objWshShell = Nothing
    Exit Function
DoRequestError:

    App.LogEvent "PackControl." & strFunctionName & Err.Description, vbLogEventTypeError
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    strResponse = Err.Description                                        ' MAR1638
    
    Resume TidyUp

End Function

Private Function DeformatGUID(ByVal strGuid As String) As String
    DeformatGUID = Replace(Replace(Replace(strGuid, "-", ""), "{", ""), "}", "")
End Function

Private Function FormatGuid(ByVal strGuid As String) As String
Dim strFormattedGUID As String
    strFormattedGUID = "{" & Left$(strGuid, 8) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 9, 4) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 13, 4) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 17, 4) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 21, 12) & "}"
    FormatGuid = strFormattedGUID
End Function


Private Function GetConnection() As Connection
    Set GetConnection = New Connection
    GetConnection.ConnectionString = GetOmiga4ConnectionString()
    GetConnection.Open
End Function

Private Function GetFileNetRecord(ByVal xmlRequestNode As IXMLDOMNode) As String
'TW 12/06/2006 MAR1867 Modifications for performance reasons

Const cstrFunctionName As String = "GetFileNetRecord"   ' PSC 06/02/2006 MAR1197

Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

Dim strApplicationNumber As String
Dim strFileNetGUID As String
Dim strResponse As String
Dim strUserId As String
Dim strFileNetUserId As String      ' PSC 11/01/2005 MAR994

' PSC 06/02/2006 MAR1197 - Start
Dim xmlResponse As FreeThreadedDOMDocument40
Dim xmlFileNetRecord As IXMLDOMNode
' PSC 06/02/2006 MAR1197 - End
    
    On Error GoTo GetFileNetRecordError:
    
' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

    strFileNetGUID = GetAttribute(xmlRequestNode, "FILENETIMAGEREF")
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strUserId = GetAttribute(xmlRequestNode, "USERID")
    
' Call omFDM to contact FileNet to get the document contents
    ' PSC 15/12/2005 MAR880
    ' PSC 11/01/2006 MAR994 - Start
    strFileNetUserId = GetMandatoryGlobalParamString("FileNetUserId")
    strResponse = DoRequest("getFileNetRecord", strApplicationNumber, FormatGuid(strFileNetGUID), strFileNetUserId, xmlRequestNode)
    ' PSC 11/01/2006 MAR994 - End
    
    'MAR1867 omFDM returns either an error message in a string OR the XML document contents in a string.
    If Mid$(strResponse, 1, 1) <> "<" Then
        Err.Raise 40019, "PackControl.GetFileNetRecord", "Failed to get document data from FileNet : " & strResponse
    End If
    
    ' PSC 06/02/2006 MAR1197 - Start
    Set xmlResponse = xmlLoad(strResponse, cstrFunctionName)
    Set xmlFileNetRecord = xmlGetMandatoryNode(xmlResponse, "//FILENETRECORD")
    ' PSC 06/02/2006 MAR1197 - End
    
    Set xmlNode = AddNode(xmlResponseDoc, "STORAGEDETAILS", xmlResponseNode)
    Call AddAttribute(xmlNode, "FOLDER", "//Mortgages/" & strApplicationNumber)
    Call AddAttribute(xmlNode, "USERID", strUserId)
    
    ' PSC 25/09/2006 MAR1958 - Start
    ' Get global parameter to see if we need the password
    Dim strSecurityType As String
    Dim strPassword As String
    
    strSecurityType = GetGlobalParamString("SecurityCredentialsType")
    
    If StrComp(strSecurityType, "WINDOWSAUTHENTICATION", vbTextCompare) <> 0 Then
        strPassword = GetPassword(strUserId)
    Else
        strPassword = ""
    End If
        
    Call AddAttribute(xmlNode, "PASSWORD", strPassword)
    ' PSC 25/09/2006 MAR1958 - End
    
    
    Set xmlNode = AddNode(xmlResponseDoc, "DOCUMENTCONTENTS", xmlResponseNode)
    Call AddAttribute(xmlNode, "FILECONTENTS_TYPE", DEFAULT_FILECONTENTS_TYPE)
    ' PSC 06/02/2006 MAR1197 - Start
    xmlCopyAttribute xmlFileNetRecord, xmlNode, "CONTENTTYPE"
    xmlCopyAttribute xmlFileNetRecord, xmlNode, "FILECONTENTS"
    ' PSC 06/02/2006 MAR1197 - End
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    gobjTrace.TraceXML xmlResponseDoc.xml, "GetFileNetRecord " & strFileNetGUID
    
TidyUp:
    On Error Resume Next
    GetFileNetRecord = xmlResponseDoc.xml
    
    On Error GoTo 0
    Set xmlResponseDoc = Nothing

    Set xmlNode = Nothing
    Set xmlResponseNode = Nothing
    ' PSC 06/02/2006 MAR1197 - Start
    Set xmlResponse = Nothing
    Set xmlFileNetRecord = Nothing
    ' PSC 06/02/2006 MAR1197 - End
    
    Exit Function

GetFileNetRecordError:
    Call DealWithErrorResponse(xmlResponseDoc, cstrFunctionName, Err)
    Resume TidyUp

End Function

Private Function GetPassword(ByVal strUserId As String) As String
Dim connOmiga As Connection

Dim recPassword As ADODB.Recordset

Dim strSQL As String
    
    On Error GoTo GetPasswordError:
    
    strSQL = "select PASSWORDVALUE from PASSWORD where USERID = '" & strUserId & "' " _
    & " AND PASSWORDCREATIONDATE = (SELECT MAX(PASSWORDCREATIONDATE) FROM PASSWORD " _
    & " where USERID = '" & strUserId & "')"
    
    Set connOmiga = GetConnection()
    
    Set recPassword = connOmiga.Execute(strSQL)
    
    If recPassword.EOF Then
        Err.Raise 40022, "PackControl.GetPassword", "Can't find password for user name '" & strUserId & "'"
    End If
    
    GetPassword = Decrypt(recPassword(0))
    recPassword.Close
    connOmiga.Close
    
TidyUp:
    Set connOmiga = Nothing
    Set recPassword = Nothing
    Exit Function
    
GetPasswordError:
    App.LogEvent "PackControl.GetPassword " & Err.Description, vbLogEventTypeError
    Resume TidyUp
End Function

Private Function Decrypt(ByVal strPassword As String) As String
Const cint_MIN = 10
Const cint_MAX = 40

Dim lngChar As Long
Dim intIncrement As Integer
Dim strOut As String

Dim strChar As String
    
    ' initialise the random-number generator
    Rnd -1
    Randomize 1
    
    For lngChar = 1 To Len(strPassword)
        intIncrement = Int(Rnd() * cint_MAX) + cint_MIN
    
        strChar = Mid$(strPassword, lngChar, 1)
        If strChar = "~" Then
            lngChar = lngChar + 1
            strChar = Mid$(strPassword, lngChar, 1)
        Else
            intIncrement = 0 - intIncrement
        End If
    
        strOut = strOut & Chr(Asc(strChar) + intIncrement)
    Next

    Decrypt = strOut
        
End Function


Private Function GetQuoteNumber(ByVal xmlNode As IXMLDOMNode) As String
Dim strAcceptedQuoteNumber As String
Dim strActiveQuoteNumber As String
    
    strAcceptedQuoteNumber = GetAttribute(xmlNode, "ACCEPTEDQUOTENUMBER")
    strActiveQuoteNumber = GetAttribute(xmlNode, "ACTIVEQUOTENUMBER")
    
    If Len(strAcceptedQuoteNumber) > 0 Or Len(strActiveQuoteNumber) > 0 Then
        
        If Len(strAcceptedQuoteNumber) <> 0 Then
            GetQuoteNumber = strAcceptedQuoteNumber
        Else
            GetQuoteNumber = strActiveQuoteNumber
        End If
    Else
        GetQuoteNumber = GetAttribute(xmlNode, "QUOTATIONNUMBER")
    End If
End Function

Public Sub InitialiseTraceInterface( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStartElapsed As Single)
    
    If Not vblnTraceOn Then
        Exit Sub
    End If
    
    gobjTrace.TraceInitialiseFromParent vblnTraceOn, vstrFileName, vintStartElapsed
End Sub


Private Function AddNode(ByVal xmlDoc As FreeThreadedDOMDocument40, _
                         ByVal strNodeName As String, _
                         Optional ByVal xmlParentObject As IXMLDOMNode) As IXMLDOMNode
                         
    Set AddNode = xmlDoc.createNode(NODE_ELEMENT, strNodeName, "")
    If xmlParentObject Is Nothing Then
        xmlDoc.appendChild AddNode
    Else
        xmlParentObject.appendChild AddNode
    End If
End Function

Private Sub CreateAuditTrail(ByVal xmlPackMemberNode As IXMLDOMNode, _
                             ByVal strEventKey As String, _
                             ByVal strApplicationNumber As String, _
                             ByRef strDocumentGuid As String, _
                             Optional ByVal strFileGuid As String = "", _
                             Optional ByVal strFileVersion As String = DEFAULT_FILEVERSION, _
                             Optional ByVal geminiPrintStatus As Integer = DEFAULT_GEMINIPRINTSTATUS)
' Rewritten by TW 12/06/2006 MAR1867  to remove use of omPM
' and to use stored procedures to deal with the audit trail entries
Const strFunctionName = "CreateAuditTrail"

Dim intEventKey As Integer

Dim strFileNetImageRef As String
Dim strHostTemplateId As String
Dim strPackGUID As String
'TW 07/07/2006 MAR1897
Dim strQuoteNumber As String
'TW 07/07/2006 MAR1897 End
Dim strUserId As String
Dim strUnitId As String

Dim xmlPackNode As IXMLDOMNode

    On Error GoTo CreateAuditTrailError:
    
    If xmlPackMemberNode.parentNode.nodeName = "PACKCONTROL" Then
        Set xmlPackNode = xmlPackMemberNode.parentNode
    Else
        Set xmlPackNode = xmlPackMemberNode
    End If

    intEventKey = strEventKey
    
    strFileNetImageRef = GetAttribute(xmlPackMemberNode, "FILENETIMAGEREF")
    
    'MAR1867 Deformat the GUID returned from File Net
    strFileNetImageRef = DeformatGUID(strFileNetImageRef)
    
    strHostTemplateId = GetAttribute(xmlPackMemberNode, "DOCUMENTID")
    strUserId = GetAttribute(xmlPackNode, "USERID")
    strUnitId = GetAttribute(xmlPackNode, "UNITID")
    strPackGUID = GetAttribute(xmlPackNode, "PACKGUID")
    'TW 07/07/2006 MAR1897
    strQuoteNumber = GetAttribute(xmlPackMemberNode, "QUOTATIONNUMBER", "1")
    'TW 07/07/2006 MAR1897
    
    If Len(strPackGUID) = 0 Then
        strPackGUID = strZeroes
    End If
    If Len(strFileNetImageRef) = 0 Then
        strFileNetImageRef = strZeroes
    End If

    Select Case intEventKey
        Case EVENTKEY_CREATED ' processing a new document

            'Create DocumentAuditDetails record
'TW 07/07/2006 MAR1897
'            strDocumentGuid = CreateDocumentAuditDetails(strApplicationNumber, strHostTemplateId)
            strDocumentGuid = CreateDocumentAuditDetails(strApplicationNumber, strHostTemplateId, strQuoteNumber, geminiPrintStatus)
'TW 07/07/2006 MAR1897
        
        Case _
            EVENTKEY_EDITED, _
            EVENTKEY_RECATEGORISATION
            
            Call UpdateDocumentAuditDetails(xmlPackMemberNode, strApplicationNumber, strDocumentGuid)
        
        Case EVENTKEY_VIEWED To EVENTKEY_EMAIL
        Case Else
            ' In case of any errors
    End Select

    'Create EventAuditDetail record
    Call CreateEventAuditDetail(strApplicationNumber, intEventKey, strHostTemplateId, strDocumentGuid, strPackGUID, strFileNetImageRef, strUnitId, strUserId, strFileGuid, strFileVersion)
    
    Exit Sub
    
CreateAuditTrailError:
    App.LogEvent "PackControl.CreateAuditTrail " & Err.Description, vbLogEventTypeError
End Sub

Private Sub UpdateDocumentAuditDetails(ByVal xmlRequestNode As IXMLDOMNode, _
                                       ByVal strApplicationNumber As String, _
                                       ByVal strDocumentGuid As String)

' Update DocumentAuditDetails record

Const strFunctionName = "UpdateDocumentAuditDetails"
    
Dim connOmiga As ADODB.Connection
Dim cmd As ADODB.Command

Dim xmlNode As IXMLDOMNode

Dim strArchiveDate As String
Dim strCustomerName As String
Dim strDocumentDescription As String
Dim strDocumentGroup As String
Dim strDocumentName As String
Dim strDocumentPurpose As String
Dim strHostTemplateId As String
Dim strLanguage As String
Dim strPrintDate As String
Dim strRecipientName As String
Dim strSearchKey1 As String
Dim strSearchKey2 As String
Dim strSearchKey3 As String
Dim strSourceSystem As String
Dim strStageID As String
Dim strTemplateID As String

    On Error GoTo UpdateDocumentAuditDetailsError:
    
    Set connOmiga = GetConnection()

    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "Updating DocumentAuditDetails"
    
    Set xmlNode = xmlGetNode(xmlRequestNode, "APPLICATIONDETAIL")
    strTemplateID = GetAttribute(xmlNode, "TEMPLATEID", "?")
    strHostTemplateId = GetAttribute(xmlNode, "HOSTTEMPLATEID", "?")
    strSearchKey1 = GetAttribute(xmlNode, "SEARCHKEY1", "?")
    strSearchKey2 = GetAttribute(xmlNode, "SEARCHKEY2", "?")
    strSearchKey3 = GetAttribute(xmlNode, "SEARCHKEY3", "?")
    
    Set xmlNode = xmlGetNode(xmlRequestNode, "DOCUMENTDETAILS")
    strStageID = GetAttribute(xmlNode, "STAGEID", "?")
    strDocumentPurpose = GetAttribute(xmlNode, "DOCUMENTPURPOSE", "?")
    strDocumentGroup = GetAttribute(xmlNode, "DOCUMENTGROUP", "?")
    strDocumentName = GetAttribute(xmlNode, "DOCUMENTNAME", "?")
    strDocumentDescription = GetAttribute(xmlNode, "DOCUMENTDESCRIPTION", "?")
    strCustomerName = GetAttribute(xmlNode, "CUSTOMERNAME", "?")
    strRecipientName = GetAttribute(xmlNode, "RECIPIENTNAME", "?")
    strSourceSystem = GetAttribute(xmlNode, "SOURCESYSTEM", "?")
    strArchiveDate = GetAttribute(xmlNode, "ARCHIVEDATE", "?")
    strPrintDate = GetAttribute(xmlNode, "PRINTDATE", "?")
    strLanguage = GetAttribute(xmlNode, "LANGUAGE", "?")

    Set cmd = New ADODB.Command
    
    'MAR1867 Use with. Use correct stored procedure
    With cmd
        Set .ActiveConnection = connOmiga
        .CommandType = adCmdStoredProc
        .CommandText = "usp_UpdateDocumentAuditDetails"
        .Parameters.Append .CreateParameter("@ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        .Parameters.Append .CreateParameter("@DocumentGUID", adVarBinary, adParamInput, 16, GuidStringToByteArray(strDocumentGuid))
        .Parameters.Append .CreateParameter("@TemplateID", adBSTR, adParamInput, Len(strTemplateID), strTemplateID)
        .Parameters.Append .CreateParameter("@HostTemplateID", adBSTR, adParamInput, Len(strHostTemplateId), strHostTemplateId)
        .Parameters.Append .CreateParameter("@SearchKey1", adBSTR, adParamInput, Len(strSearchKey1), strSearchKey1)
        .Parameters.Append .CreateParameter("@SearchKey2", adBSTR, adParamInput, Len(strSearchKey2), strSearchKey2)
        .Parameters.Append .CreateParameter("@SearchKey3", adBSTR, adParamInput, Len(strSearchKey3), strSearchKey3)
        .Parameters.Append .CreateParameter("@StageID", adBSTR, adParamInput, Len(strStageID), strStageID)
        .Parameters.Append .CreateParameter("@DocumentGroup", adBSTR, adParamInput, Len(strDocumentGroup), strDocumentGroup)
        .Parameters.Append .CreateParameter("@DocumentName", adBSTR, adParamInput, Len(strDocumentName), strDocumentName)
        .Parameters.Append .CreateParameter("@DocumentDescription", adBSTR, adParamInput, Len(strDocumentDescription), strDocumentDescription)
        .Parameters.Append .CreateParameter("@CustomerName", adBSTR, adParamInput, Len(strCustomerName), strCustomerName)
        .Parameters.Append .CreateParameter("@RecipientName", adBSTR, adParamInput, Len(strRecipientName), strRecipientName)
        .Parameters.Append .CreateParameter("@ArchiveDate", adBSTR, Len(strArchiveDate), strArchiveDate)
        .Parameters.Append .CreateParameter("@PrintDate", adBSTR, Len(strPrintDate), strPrintDate)
        .Parameters.Append .CreateParameter("@Language", adBSTR, adParamInput, Len(strLanguage), strLanguage)
        .Parameters.Append .CreateParameter("@SourceSystem", adBSTR, adParamInput, Len(strSourceSystem), strSourceSystem)
        .Parameters.Append .CreateParameter("@DocumentPurpose", adBSTR, adParamInput, Len(strDocumentPurpose), strDocumentPurpose)

        cmd.Execute , , adExecuteNoRecords
        Set cmd.ActiveConnection = Nothing
    End With
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "DocumentAuditDetails updated"

TidyUp:
    Set connOmiga = Nothing
    Set xmlNode = Nothing
    
    Exit Sub

UpdateDocumentAuditDetailsError:
    App.LogEvent "PackControl.UpdateDocumentAuditDetails " & Err.Description, vbLogEventTypeError
    Resume TidyUp
    
End Sub


Private Function GetAttribute(ByVal xmlNode As IXMLDOMNode, _
                              ByVal strAttributeName As String, _
                              Optional ByVal strDefault As String = "") As String
Dim xmlAttribute As IXMLDOMAttribute
    
    GetAttribute = strDefault
    If Not xmlNode Is Nothing Then
        Set xmlAttribute = xmlNode.Attributes.getNamedItem(strAttributeName)
        If Not xmlAttribute Is Nothing Then
            GetAttribute = xmlAttribute.Text
        End If
    End If
End Function

Private Sub AddAttribute(ByVal xmlNode As IXMLDOMNode, ByVal strAttributeName As String, ByVal strAttributeValue As String)
Dim xmlAttribute As IXMLDOMAttribute

    Set xmlAttribute = xmlNode.ownerDocument.createAttribute(strAttributeName)
    xmlAttribute.Text = strAttributeValue
    xmlNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = Nothing
End Sub


Public Function AddPackMember(ByVal intPackControlNumber As Integer, _
                              ByVal intPackMemberType As Integer, _
                              ByVal intPackMemberControlNumber As Integer, _
                              ByVal strHostTemplateId As String) As Boolean
' Create a new PACKMEMBER record.
' Return .TRUE. if successful
' intPackControlNumber must be present and exist on database.
' intPackMemberType must be 1 (Pack) or 2 (Template)
' intPackMemberControlNumber must be present and exist in the database if intPackMemberType = 1.
' strHostTemplateID must be present and exist in the database if intPackMemberType = 2.

Dim connOmiga As Connection
Dim cmd As ADODB.Command
Dim recTemp As ADODB.Recordset

Dim i As Integer
Dim intPackMemberNumber As Integer

Dim param(1 To 5) As ADODB.Parameter

Dim strSQL As String
    
    On Error GoTo AddPackMemberError:
    If Len(intPackControlNumber) = 0 Then
        Err.Raise 40004, "PackControl.AddPackMember", "No pack name supplied"
    End If
    
    Set connOmiga = GetConnection()
    
    strSQL = "SELECT Count(*) FROM PACKCONTROL C INNER JOIN PACKMEMBER M ON C.PACKCONTROLNUMBER = M.PACKCONTROLNUMBER WHERE C.PACKCONTROLNUMBER = " & intPackControlNumber
    
    Set recTemp = connOmiga.Execute(strSQL)
    
    intPackMemberNumber = recTemp(0) + 1
    
    Select Case intPackMemberType
        Case 1 'Pack
            If intPackMemberControlNumber = 0 Then
                Err.Raise 40006, "PackControl.AddPackMember", "No pack member number supplied"
            End If
            
            strSQL = "SELECT Count(*) FROM PACKCONTROL WHERE PACKCONTROLNUMBER = " & intPackMemberControlNumber
            
            Set recTemp = connOmiga.Execute(strSQL)
            
            If recTemp(0) = 0 Then
                Err.Raise 40005, "PackControl.AddPackMember", "Could not find PACKCONTROL '" & intPackMemberControlNumber & "'"
            End If
        Case 2 'Template
            If Len(strHostTemplateId) = 0 Then
                Err.Raise 40007, "PackControl.AddPackMember", "No host tmplate id supplied"
            End If
            
            strSQL = "SELECT * FROM HOSTTEMPLATE WHERE HOSTTEMPLATEID = '" & strHostTemplateId & "'"
            Set recTemp = connOmiga.Execute(strSQL)
            If recTemp.EOF Then
                Err.Raise 40008, "PackControl.AddPackMember", "Could not find HOSTTEMPLATE '" & strHostTemplateId & "'"
            End If
            
        Case Else
            Err.Raise 40009, "PackControl.AddPackMember", "Pack member type must be 1 (Pack) or 2 (Template)"
    End Select
    
    Set connOmiga = GetConnection()
    
    strSQL = "INSERT into PACKMEMBER VALUES (?, ?, ?, ?, ?)"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = connOmiga
    cmd.CommandText = strSQL
    
    For i = 1 To 5
        Set param(i) = New ADODB.Parameter
    Next i
    
    param(1).Type = adInteger
    param(1).Value = intPackControlNumber

    param(2).Type = adNumeric
    param(2).Precision = 5
    param(2).Value = intPackMemberNumber

    param(3).Type = adTinyInt
    param(3).Value = intPackMemberType

    Select Case intPackMemberType
        Case 1 'Pack
            param(4).Type = adInteger
            param(4).Value = intPackMemberControlNumber
            
            param(5).Type = adVariant
            param(5).Value = Empty
        Case 2 'Template
            param(4).Type = adVariant
            param(4).Value = Empty
            
            param(5).Type = adVarChar
            param(5).Value = strHostTemplateId
            param(5).Size = Len(strHostTemplateId)
    End Select
    
    For i = 1 To 5
        cmd.Parameters.Append param(i)
    Next i
    
    cmd.Execute
    connOmiga.Close
    AddPackMember = True
    
TidyUp:
    On Error GoTo 0
    Set recTemp = Nothing
    For i = 1 To 5
        Set param(i) = Nothing
    Next i
    Set cmd = Nothing
    Set connOmiga = Nothing
    Exit Function
    
AddPackMemberError:
    AddPackMember = False
    App.LogEvent "PackControl.AddPackMember " & Err.Description, vbLogEventTypeError
    Resume TidyUp

End Function


Private Function BuildPackXML(ByVal intPackControlNumber As Integer) As String
Dim intUseMultipleQuotes As Integer

Dim xmlPackDoc As FreeThreadedDOMDocument40

Dim xmlPackNode As IXMLDOMNode

    On Error GoTo BuildPackXMLError:
    
    Set xmlPackDoc = xmlCreateDOMObject()
    Set xmlPackNode = AddNode(xmlPackDoc, "PACKCONTROL")

    Call AddAttribute(xmlPackNode, "PACKCONTROLNUMBER", intPackControlNumber)
    intUseMultipleQuotes = FollowPackHeirarchy(intPackControlNumber, xmlPackNode)
    Call AddAttribute(xmlPackNode, "USEMULTIPLEQUOTES", intUseMultipleQuotes)
    
' Remove any empty nested packs
    For Each xmlPackNode In xmlPackDoc.selectNodes("//PACKMEMBER[@PACKMEMBERTYPE='1']")
        If xmlPackNode.childNodes.length = 0 Then
            xmlPackNode.parentNode.removeChild xmlPackNode
        End If
    Next
    
    BuildPackXML = xmlPackDoc.xml

    gobjTrace.TraceXML xmlPackDoc.xml, "Build " & intPackControlNumber
TidyUp:
    On Error GoTo 0

    Set xmlPackDoc = Nothing
    Set xmlPackNode = Nothing

    Exit Function

BuildPackXMLError:
    App.LogEvent "PackControl.BuildPackXML " & Err.Description, vbLogEventTypeError
    Resume TidyUp
End Function



Public Function CreatePack(ByVal intPackNumber As Integer, ByVal strPackName As String, ByVal strPackDescription As String, ByVal intUseMultipleQuotes As Integer) As Boolean
' Create a new PACKCONTROL record.
' Return .TRUE. if successful
' strPackName must be present and unique.
' strPackDescription need not have any content.

Dim connOmiga As Connection
Dim cmd As ADODB.Command

Dim i As Integer

Dim param(1 To 4) As ADODB.Parameter

Dim strSQL As String
    
    On Error GoTo CreatePackError:
    If Len(strPackName) = 0 Then
        Err.Raise 40001, "PackControl.CreatePack", "No pack name supplied"
    End If
    
    Set connOmiga = GetConnection()
    
    strSQL = "INSERT into PACKCONTROL (PACKCONTROLNUMBER, PACKCONTROLNAME, PACKCONTROLDESCRIPTION, USEMULTIPLEQUOTES) VALUES (?, ?, ?, ?)"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = connOmiga
    cmd.CommandText = strSQL
    
    For i = 1 To 4
        Set param(i) = New ADODB.Parameter
    Next i
    
    param(1).Type = adInteger
    param(1).Value = intPackNumber
    
    param(2).Type = adVarChar
    param(2).Value = strPackName
    param(2).Size = Len(strPackName)
    
    param(3).Type = adVarChar
    param(3).Value = strPackDescription
    param(3).Size = Len(strPackDescription)
    
    param(4).Type = adInteger
    param(4).Value = intUseMultipleQuotes
    
    For i = 1 To 4
        cmd.Parameters.Append param(i)
    Next i
    
    cmd.Execute
    connOmiga.Close
    CreatePack = True
    
TidyUp:
    On Error GoTo 0
    For i = 1 To 4
        Set param(i) = Nothing
    Next i
    Set cmd = Nothing
    Set connOmiga = Nothing
    Exit Function
    
CreatePackError:
    App.LogEvent "PackControl.CreatePack " & Err.Description, vbLogEventTypeError
    Resume TidyUp

End Function


Private Function FollowPackHeirarchy(ByVal intPackControlNumber As Integer, ByVal xmlPackParentNode As IXMLDOMNode) As Integer
Dim blnFirstRecord As Boolean

Dim connOmiga As Connection

Dim strSQL As String

Dim recField As ADODB.Field
Dim recTemp As ADODB.Recordset

Dim xmlPackMemberNode As IXMLDOMNode

    On Error GoTo FollowPackHeirarchyError:

' Connect to the Omiga database
    Set connOmiga = GetConnection()

    strSQL = "SELECT * from FINDDOCUMENTSFORPACK WHERE PACKCONTROLNUMBER = " & intPackControlNumber

    Set recTemp = connOmiga.Execute(strSQL)
    
    If (recTemp.EOF) Then
        Err.Raise 40021, "PackControl.FollowPackHeirarchy", "Pack " & intPackControlNumber & " not found"
    End If
    blnFirstRecord = True
    FollowPackHeirarchy = recTemp!USEMULTIPLEQUOTES
    
    'MAR978  Save Pack Name
    Call AddAttribute(xmlPackParentNode, "PACKNAME", recTemp!PACKNAME)
    'AS 08/01/2007 EP1277 Ensure pack description is returned.
    Call AddAttribute(xmlPackParentNode, "PACKDESCRIPTION", recTemp!PACKDESCRIPTION)
    
    Do While Not recTemp.EOF

' Only add the Packs or Templates if not already added
        If xmlPackParentNode.ownerDocument.selectSingleNode("//PACKMEMBER[@PACKMEMBERCONTROLNUMBER='" & recTemp!PACKMEMBERCONTROLNUMBER & "']") Is Nothing And _
            xmlPackParentNode.ownerDocument.selectSingleNode("//PACKMEMBER[@HOSTTEMPLATEID='" & recTemp!HostTemplateId & "' and @DPSTEMPLATEID='" & recTemp!dpsTemplateId & "']") Is Nothing Then
            If Not blnFirstRecord Or recTemp!PACKMEMBERTYPE = 1 Or xmlPackParentNode Is xmlPackParentNode.ownerDocument.documentElement Then
                Set xmlPackMemberNode = AddNode(xmlPackParentNode.ownerDocument, "PACKMEMBER", xmlPackParentNode)
                
                Call AddAttribute(xmlPackMemberNode, "PACKMEMBERTYPE", recTemp!PACKMEMBERTYPE)
            End If
            Select Case recTemp!PACKMEMBERTYPE
                Case 1 'Pack
                    Call AddAttribute(xmlPackMemberNode, "PACKMEMBERCONTROLNUMBER", recTemp!PACKMEMBERCONTROLNUMBER)
                    Call FollowPackHeirarchy(recTemp!PACKMEMBERCONTROLNUMBER, xmlPackMemberNode)
                Case 2 'Template
' Don't generate first element (cover letter) for nested packs
                    If Not blnFirstRecord Or xmlPackParentNode Is xmlPackParentNode.ownerDocument.documentElement Then
                        If xmlPackParentNode.ownerDocument.selectSingleNode("//PACKMEMBER[@HOSTTEMPLATEID='" & recTemp!HostTemplateId & "' and @DPSTEMPLATEID='" & recTemp!dpsTemplateId & "']") Is Nothing Then
                            For Each recField In recTemp.Fields
                                Select Case UCase$(recField.Name)
                                    Case "HOSTTEMPLATEID", "DPSTEMPLATEID", "TEMPLATEGROUPID", "PACKMEMBERNUMBER", _
                                         "HOSTTEMPLATENAME", "HOSTTEMPLATEDESCRIPTION", "MINROLELEVEL", _
                                         "MAXCOPIES", "DEFAULTCOPIES", "RECIPIENTTYPE", _
                                         "PRINTERDESTINATIONTYPE", "REMOTEPRINTERLOCATION", "PDMMETHOD", _
                                         "DELIVERYTYPE", "FIRSTPAGEPRINTERTRAY", "OTHERPAGESPRINTERTRAY", _
                                         "DOCUMENTCLASS", _
                                         "PACKDESTINATION", "PRECREATEFORPACK", "GEMINIPRINTMODE", "PACKDOCUMENTTYPE" 'AW EP2_2181

                                            If IsNull(recField.Value) Then
                                                Call AddAttribute(xmlPackMemberNode, recField.Name, "")
                                            Else
                                                Call AddAttribute(xmlPackMemberNode, recField.Name, recField.Value)
                                            End If
                                    Case Else
                                        Debug.Print recField.Name
                                End Select
                            Next
                        End If
                    End If
            End Select
        End If
        blnFirstRecord = False
        recTemp.MoveNext
    Loop
    recTemp.Close
    connOmiga.Close
    
TidyUp:
    On Error GoTo 0

    Set recTemp = Nothing
    Set connOmiga = Nothing
    Set xmlPackMemberNode = Nothing

    Exit Function
    
FollowPackHeirarchyError:
    FollowPackHeirarchy = 0
    App.LogEvent "PackControl.FollowPackHeirarchy " & Err.Description, vbLogEventTypeError
    Resume TidyUp
End Function

Private Function GetOmiga4ConnectionString() As String
Const strFunctionName As String = "GetOmiga4ConnectionString"

Dim objWshShell As Object

Dim strConnection As String
Dim strRegSection As String
    
Dim strUserId As String
Dim strPassword As String

Dim lngErrNo As Long
Dim strSource As String
Dim strDescription As String
    
    On Error GoTo BuildDbConnectionStringVbErr
    Set objWshShell = CreateObject("WScript.Shell")
    
    strRegSection = "HKLM\SOFTWARE\OMIGA4\DATABASE CONNECTION\"
    
    On Error Resume Next
            
    strUserId = objWshShell.RegRead(strRegSection & "User ID")
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
            
    On Error GoTo BuildDbConnectionStringVbErr

    If Err.Number <> -2147024894 And Err.Number <> 0 Then
        Err.Raise lngErrNo, strSource, strDescription
    End If

    On Error Resume Next
            
    strPassword = objWshShell.RegRead(strRegSection & "Password")
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
            
    On Error GoTo BuildDbConnectionStringVbErr

    If Err.Number <> -2147024894 And Err.Number <> 0 Then
        Err.Raise lngErrNo, strSource, strDescription
    End If
    
    strConnection = "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & "Server") & ";" & "database=" & objWshShell.RegRead(strRegSection & "Database Name") & ";"

    ' If User Id is present use SQL Server Authentication else
    ' use integrated security
    If Len(strUserId) > 0 Then
        strConnection = strConnection & "UID=" & strUserId & ";" & "pwd=" & strPassword & ";"
    Else
        strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
    End If
           
    GetOmiga4ConnectionString = strConnection
    
    Set objWshShell = Nothing
    
    Exit Function

BuildDbConnectionStringVbErr:
    
    Set objWshShell = Nothing
    
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function





Public Function FindDocumentsForPack(ByVal intPackControlNumber As Integer) As String
' Get an existing PACKCONTROL record.
' Returns RESPONSE xml string
' intPackControlNumber must be present
Const cstrFunctionName As String = "FindDocumentsForPack"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlResponseNode As IXMLDOMNode
    
    On Error GoTo FindDocumentsForPackError:

    Set xmlResponseDoc = xmlCreateDOMObject()
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

    If intPackControlNumber = 0 Then
        Err.Raise 40002, "PackControl.FindDocumentsForPack", "No pack number supplied"
    End If
    
    Set xmlDoc = xmlCreateDOMObject()
    xmlDoc.loadXML BuildPackXML(intPackControlNumber)

    xmlResponseNode.appendChild xmlDoc.selectSingleNode("//PACKCONTROL").cloneNode(True)
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")

TidyUp:
    On Error Resume Next
    FindDocumentsForPack = xmlResponseDoc.xml

    On Error GoTo 0
    Set xmlDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function
    
FindDocumentsForPackError:
    Call DealWithErrorResponse(xmlResponseDoc, cstrFunctionName, Err)
    Resume TidyUp

End Function


Public Function omRequest(ByVal vstrXMLRequest As String) As String
'TW 12/06/2006 MAR1867 Modifications for performance reasons

Const cstrFunctionName As String = "omRequest"
Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

Dim strResponse As String

    On Error GoTo omRequestError

    gobjContext.SetComplete

' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")
    
' Create an xml document for the request
    Set xmlRequestDoc = xmlCreateDOMObject()
    xmlRequestDoc.loadXML vstrXMLRequest
    
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("//REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise 40024, "PackControl.omRequest", "No REQUEST node supplied"
    End If
    
    Select Case UCase$(GetAttribute(xmlRequestNode, "OPERATION"))
        Case "SENDPACK"
            strResponse = SendPack(xmlRequestNode)
        Case "RESENDPACK"
            strResponse = ReSendPack(xmlRequestNode)
        Case "CANCELPACK"
            strResponse = CancelPack(xmlRequestNode)
        Case "SENDPACKTOSMS"
            strResponse = SendPackToSMS(xmlRequestNode)
        Case "SENDPACKTOEMAIL"
            strResponse = SendPackToEMail(xmlRequestNode)
        Case "SENDDOCUMENTTOFILENET"
            strResponse = SendDocumentToFileNet(xmlRequestNode)
        Case "GETFILENETRECORD"
            strResponse = GetFileNetRecord(xmlRequestNode)
        Case Else
    End Select
    Set xmlResponseDoc = xmlCreateDOMObject()
    xmlResponseDoc.loadXML strResponse
TidyUp:
    On Error Resume Next
    omRequest = xmlResponseDoc.xml
    On Error GoTo 0
    
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function
omRequestError:
    Call DealWithErrorResponse(xmlResponseDoc, cstrFunctionName, Err)
    Resume TidyUp
    
End Function

Private Function ProcessPackMember(ByVal xmlPackAddressNode As IXMLDOMNode, _
                            ByVal strPackControlNumber As String, _
                            ByVal strUserId As String, _
                            ByVal strUnitId As String, _
                            ByVal strMachineID As String, _
                            ByVal strChannelID As String, _
                            ByVal strUserAuthorityLevel As String, _
                            ByVal strQuoteNumber As String, _
                            ByVal xmlPackDefinitionNode As IXMLDOMNode, _
                            ByVal vstrDocumentName As String)

'TW 12/06/2006 MAR1867 Rewritten for performance reasons
Const strFunctionName = "ProcessPackMember"

Dim strApplicationNumber As String
Dim strApplicationFactFindNumber As String
Dim strHostTemplateId As String
Dim strRequest As String
Dim strResponse As String
Dim strCustomerNumber As String

Dim intCustomerVersion As Integer

Dim xmlOmPDMRequestDoc As FreeThreadedDOMDocument40
Dim xmlTempDoc As FreeThreadedDOMDocument40

Dim xmlCustomerNode As IXMLDOMNode
Dim xmlPrintDataNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlTemplateDataNode As IXMLDOMNode

Dim xmlTempAttribute As IXMLDOMAttribute

Dim objomPDM As omPDM.PrintDataManagerBO
Dim objKFI As eKFI.eKFIControl

Dim xmlDocRequest As FreeThreadedDOMDocument40
Dim xmlDocResponse As FreeThreadedDOMDocument40
    
Dim connOmiga As ADODB.Connection
Dim cmd As ADODB.Command
Dim rec As ADODB.Recordset
Dim strSQL As String
    
    On Error GoTo ProcessPackMemberError:
    
    strApplicationNumber = GetAttribute(xmlPackAddressNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetAttribute(xmlPackAddressNode, "APPLICATIONFACTFINDNUMBER")
    strHostTemplateId = GetAttribute(xmlPackDefinitionNode, "HOSTTEMPLATEID")
    
    Set xmlOmPDMRequestDoc = xmlCreateDOMObject()
    
    Set xmlRequestNode = AddNode(xmlOmPDMRequestDoc, "REQUEST")
    
    Set xmlPrintDataNode = AddNode(xmlOmPDMRequestDoc, "PRINTDATA", xmlRequestNode)
    
    Set xmlCustomerNode = xmlPackAddressNode.selectSingleNode("//CUSTOMER")
    strCustomerNumber = GetAttribute(xmlCustomerNode, "CUSTOMERNUMBER")
    intCustomerVersion = Val(Format$(GetAttribute(xmlCustomerNode, "CUSTOMERVERSIONNUMBER")))
    Call AddAttribute(xmlPrintDataNode, "CUSTOMERNUMBER", strCustomerNumber)
    Call AddAttribute(xmlPrintDataNode, "CUSTOMERVERSIONNUMBER", intCustomerVersion)
    
    Set xmlTemplateDataNode = AddNode(xmlOmPDMRequestDoc, "TEMPLATEDATA", xmlRequestNode)
    
    xmlTemplateDataNode.appendChild xmlPackAddressNode.cloneNode(True)
    
' REQUEST attributes
    Call AddAttribute(xmlRequestNode, "USERID", strUserId)
    Call AddAttribute(xmlRequestNode, "UNITID", strUnitId)
    Call AddAttribute(xmlRequestNode, "OPERATION", "PrintDocument")
    
' PRINTDATA attributes
    Call AddAttribute(xmlPrintDataNode, "APPLICATIONNUMBER", strApplicationNumber)
    Call AddAttribute(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber)
    Call AddAttribute(xmlPrintDataNode, "METHODNAME", GetAttribute(xmlPackDefinitionNode, "PDMMETHOD"))

' TEMPLATEDATA attributes
    Call AddAttribute(xmlTemplateDataNode, "QUOTATIONNUMBER", strQuoteNumber)
    
    strRequest = xmlOmPDMRequestDoc.xml

    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "Calling omPDM"

    Set objomPDM = New omPDM.PrintDataManagerBO
    strResponse = objomPDM.omRequest(strRequest)
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "Returned from omPDM"
    gobjTrace.TraceXML strResponse, "Data returned by omPDM " & vstrDocumentName
    
    Set objKFI = New eKFI.eKFIControl

    Set xmlDocResponse = xmlCreateDOMObject()
    
    xmlDocResponse.loadXML strResponse
    Set xmlTemplateDataNode = xmlDocResponse.selectSingleNode("//TEMPLATEDATA")

    Set xmlDocRequest = xmlCreateDOMObject()
    Set xmlRequestNode = AddNode(xmlDocRequest, "REQUEST")
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.3]", "Reading Template Attributes"
    
    Set connOmiga = GetConnection()

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = connOmiga
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_GetPrintAttributes"
    cmd.Parameters.Append cmd.CreateParameter("@HostTemplateID", adBSTR, adParamInput, Len(strHostTemplateId), strHostTemplateId)

    Set rec = cmd.Execute()
    
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.4]", "Read Template Attributes"
    
    Call AddAttribute(xmlRequestNode, "Template", rec!TEMPLATEPATH)
    Call AddAttribute(xmlRequestNode, "Data", xmlTemplateDataNode.xml)
    Call AddAttribute(xmlRequestNode, "OutputType", rec!DOCUMENTDELIVERYFORMAT)
    Call AddAttribute(xmlRequestNode, "PDFFontMode", rec!PDFFONTMODE)

    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.5]", "Calling eKFI"
    strResponse = objKFI.ProcessKFI(xmlRequestNode.xml)
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.6]", "Returned from eKFI"
    gobjTrace.TraceXML strResponse, "Response from eKFI.eKFIControl.ProcessKFI " & vstrDocumentName
    
    Set xmlTempDoc = xmlCreateDOMObject()
    xmlTempDoc.loadXML strResponse
    
    Set xmlTempAttribute = xmlTempDoc.selectSingleNode("//RESPONSE/@Type")
    If xmlTempAttribute.Text <> "SUCCESS" Then
        Err.Raise 40013, "PackControl.ProcessPackMember", "Pack " & strPackControlNumber & " failed - see Event log for details"
    End If
    
    Set xmlTempAttribute = xmlTempDoc.selectSingleNode("//DOCUMENTCONTENTS/@FILECONTENTS")
    If xmlTempAttribute Is Nothing Then
        Err.Raise 40015, "PackControl.ProcessPackMember", "Pack " & strPackControlNumber & " failed - no FILECONTENTS returned from omPDM"
    End If
    
    ProcessPackMember = xmlTempAttribute.Text
TidyUp:
    On Error GoTo 0
    Set xmlOmPDMRequestDoc = Nothing
    Set xmlDocRequest = Nothing
    Set xmlDocResponse = Nothing
    Set xmlTempDoc = Nothing
    
    Set xmlPrintDataNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlCustomerNode = Nothing
    
    Set xmlTempAttribute = Nothing
    
    Set objomPDM = Nothing
    
    Set connOmiga = Nothing
    Set rec = Nothing

    Exit Function

ProcessPackMemberError:
    App.LogEvent "PackControl.ProcessPackMember " & Err.Description, vbLogEventTypeError
    ProcessPackMember = "FAILED"
    Resume TidyUp

End Function

Private Function SendPackToSMS(ByVal xmlRequestNode As IXMLDOMNode) As String
'TW 12/06/2006 MAR1867 Rewritten for performance reasons

' Process an existing PACKCONTROL record.
' Returns RESPONSE xml string
' strPackId must be present
Const strFunctionName As String = "SendPackToSMS"

Dim strApplicationNumber As String
Dim strApplicationFactFindNumber As String
Dim strUserId As String
Dim strUnitId As String
Dim strCustomerNumber As String

Dim strDocumentGuid As String
Dim strResponse As String

Dim strXML As String

Dim xmlPackDataDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlPackAddressList As IXMLDOMNodeList
Dim xmlPackNodeList As IXMLDOMNodeList

Dim xmlErrorNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlPackAddressNode As IXMLDOMNode
Dim xmlPackNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

'PSC 14/12/2005 MAR802 - Start
Dim xmlCustomersToRemove As IXMLDOMNodeList
Dim xmlCustomerToRemove As IXMLDOMNode
'PSC 14/12/2005 MAR802 - End

' PSC 05/05/2006 MAR1593 - Start
Dim strPackId As String
Dim xmlSMSCustomers As IXMLDOMNodeList
Dim xmlSMSCustomer As IXMLDOMNode
Dim xmlCurrentSMSCustomer As IXMLDOMNode
Dim xmlAddress As IXMLDOMNode
' PSC 05/05/2006 MAR1593 - Snd

    On Error GoTo SendPackToSMSError:
    
' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

' Get the parameters from the Request Node
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetAttribute(xmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    strUserId = GetAttribute(xmlRequestNode, "USERID")
    strUnitId = GetAttribute(xmlRequestNode, "UNITID")
    strPackId = xmlGetMandatoryAttributeText(xmlRequestNode, "PACKCONTROLNUMBER")
     
     ' PSC 14/12/2005 MAR802 - Start
    strCustomerNumber = GetAttribute(xmlRequestNode, "CUSTOMERNUMBER")
    ' PSC 14/12/2005 MAR802 - End
   
    'MAR1819 GHun
    strXML = ExecStoredProc("usp_PACKCONTROL", strApplicationNumber, strApplicationFactFindNumber, "PACKDATA")

    If Len(strXML) = 0 Then
        Err.Raise 40027, "PackControl.SendPackToSMS", "Application '" & strApplicationNumber & "' can't find addresses for application '" & strApplicationNumber & "'"
    End If
    
    Set xmlPackDataDoc = xmlCreateDOMObject()
    xmlPackDataDoc.loadXML strXML
    'MAR1819 End

    ' PSC 05/05/2006 MAR1593 - Start
    ' Remove non relevant customers
    If Len(strCustomerNumber) > 0 Then
        ' Any customers that do not have the same customer number or are not SMS
        Set xmlCustomersToRemove = xmlPackDataDoc.selectNodes("//ADDRESS/CUSTOMER[@CUSTOMERNUMBER!='" & strCustomerNumber & "' or @CONTACTPREFERENCE!='S']")
    Else
        ' Any customers that are not SMS or have a phone number the same as a previous customer
        Set xmlCustomersToRemove = xmlPackDataDoc.selectNodes("//ADDRESS/CUSTOMER[@CONTACTPREFERENCE!='S' or concat(@COUNTRYCODE,@AREACODE,@TELEPHONENUMBER,@EXTENSIONNUMBER)=concat(preceding::CUSTOMER/@COUNTRYCODE, preceding::CUSTOMER/@AREACODE, preceding::CUSTOMER/@TELEPHONENUMBER, preceding::CUSTOMER/@EXTENSIONNUMBER)]")
    End If
   
    For Each xmlCustomerToRemove In xmlCustomersToRemove
        xmlCustomerToRemove.parentNode.removeChild xmlCustomerToRemove
    Next
     ' PSC 05/05/2006 MAR1593 - End

    ' PSC 14/12/2005 MAR802
    ' PSC 05/05/2006 MAR1593
    Set xmlPackAddressList = xmlPackDataDoc.selectNodes("//ADDRESS[CUSTOMER]")
    
    For Each xmlPackAddressNode In xmlPackAddressList

        ' Create PACKCONTROL node as a child of the RESPONSE node
        Set xmlPackNode = AddNode(xmlResponseDoc, "PACKCONTROL", xmlResponseNode)

        ' Add all the details for the customers
        xmlPackNode.appendChild xmlPackAddressNode.cloneNode(True)

        ' Get the unique Customer references for the PACK
        Call AddAttribute(xmlPackNode, "PACKGUID", CreateGUID())   ' MAR978
        Call AddAttribute(xmlPackNode, "USERID", strUserId)
        Call AddAttribute(xmlPackNode, "UNITID", strUnitId)
        xmlSetAttributeValue xmlPackNode, "PACKTYPEID", strPackId       ' PSC 05/05/2006 MAR1593
    Next
    
' Now send the packs to Fulfilment
    Set xmlPackNodeList = xmlResponseDoc.selectNodes("//PACKCONTROL")
    For Each xmlPackNode In xmlPackNodeList
        
        ' PSC 05/05/06 MAR1593 - Start
        Set xmlAddress = xmlPackNode.selectSingleNode("ADDRESS")
        Set xmlSMSCustomers = xmlAddress.selectNodes("CUSTOMER")
        
        ' Remove the customers to process one at a time
        For Each xmlSMSCustomer In xmlSMSCustomers
            xmlAddress.removeChild xmlSMSCustomer
        Next
        
        ' Process each customer separately
        For Each xmlSMSCustomer In xmlSMSCustomers
            Set xmlCurrentSMSCustomer = xmlAddress.appendChild(xmlSMSCustomer)
            strResponse = DoRequest("sendToSMS", "", "", strUserId, xmlPackNode)
            'MAR1638
            If strResponse <> "Done" Then
                Err.Raise 40019, "PackControl.SendPackToSMS", "Failed to send SMS Pack to Fullfilment: " & strResponse
            End If
            
            Call CreateAuditTrail(xmlPackNode, EVENTKEY_CREATED, strApplicationNumber, strDocumentGuid)
            Call AddAttribute(xmlPackNode, "DOCUMENTGUID", strDocumentGuid)
            Call CreateAuditTrail(xmlPackNode, EVENTKEY_SMS, strApplicationNumber, strDocumentGuid)
            gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "objOmFDM.sendToSMS Response : " & strResponse
            xmlAddress.removeChild xmlCurrentSMSCustomer
        Next
        ' PSC 05/05/06 MAR1593 - End
    Next
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    gobjTrace.TraceXML xmlResponseDoc.xml, "Processed SMS " & strApplicationNumber

TidyUp:
    On Error Resume Next
    SendPackToSMS = xmlResponseDoc.xml
    
    On Error GoTo 0
    
    Set xmlPackDataDoc = Nothing
    Set xmlResponseDoc = Nothing

    Set xmlErrorNode = Nothing
    Set xmlNode = Nothing
    Set xmlPackAddressNode = Nothing
    Set xmlPackNode = Nothing
    Set xmlResponseNode = Nothing
    
    Set xmlPackAddressList = Nothing
    Set xmlPackNodeList = Nothing

    'PSC 14/12/2005 MAR802 - Start
    Set xmlCustomersToRemove = Nothing
    Set xmlCustomerToRemove = Nothing
    'PSC 14/12/2005 MAR802 - End
    
    ' PSC 05/05/2006 MAR1593 - Start
    Set xmlSMSCustomers = Nothing
    Set xmlSMSCustomer = Nothing
    Set xmlCurrentSMSCustomer = Nothing
    Set xmlAddress = Nothing
    ' PSC 05/05/2006 MAR1593 - Snd

    Exit Function
    
SendPackToSMSError:
    Call DealWithErrorResponse(xmlResponseDoc, strFunctionName, Err)
    Resume TidyUp

End Function

Private Function SendDocumentToFileNet(ByVal xmlRequestNode As IXMLDOMNode) As String
'TW 12/06/2006 MAR1867 Rewritten for performance reasons

Const strFunctionName As String = "SendDocumentToFileNet"

Dim strApplicationNumber As String
Dim strDocumentClass As String
Dim strDocumentGuid As String
Dim strDocumentName As String
Dim strFileContents As String
Dim strFileNetGUID As String
Dim strFileNetUserId As String
Dim strHostTemplateId As String
Dim strPackGUID As String
'TW 07/07/2006 MAR1897
Dim strQuoteNumber As String
'TW 07/07/2006 MAR1897 EndDim strUnitID As String
Dim strUnitId As String
Dim strUserId As String

Dim xmlFileNetDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlAttributesNode As IXMLDOMNode
Dim xmlDocumentDetailsNode As IXMLDOMNode
Dim xmlErrorNode As IXMLDOMNode
Dim xmlFileNetNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

Dim connOmiga As ADODB.Connection
Dim rec As ADODB.Recordset
    
    On Error GoTo SendDocumentToFileNetError:
    
' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

' Get the parameters from the Request Node
    strDocumentName = GetAttribute(xmlRequestNode, "DOCUMENTNAME")
    If Len(strDocumentName) = 0 Then
        strDocumentName = GetAttribute(xmlRequestNode, "HOSTTEMPLATENAME")
    End If
    
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strHostTemplateId = GetAttribute(xmlRequestNode, "HOSTTEMPLATEID")
    strUserId = GetAttribute(xmlRequestNode, "USERID")
    strUnitId = GetAttribute(xmlRequestNode, "UNITID")
'TW 07/07/2006 MAR1897
    strQuoteNumber = GetQuoteNumber(xmlRequestNode)
'TW 07/07/2006 MAR1897 End

    Set xmlDocumentDetailsNode = xmlRequestNode.selectSingleNode("DOCUMENTCONTENTS")

    strFileContents = GetAttribute(xmlDocumentDetailsNode, "FILECONTENTS")

    Set xmlFileNetDoc = xmlCreateDOMObject()
    Set xmlFileNetNode = AddNode(xmlFileNetDoc, "FILECONTENTS")
    Call AddAttribute(xmlFileNetNode, "FILECONTENTS", strFileContents)
    xmlCopyAttribute xmlDocumentDetailsNode, xmlFileNetNode, "CONTENTTYPE"
    
    strDocumentClass = GetAttribute(xmlDocumentDetailsNode, "DOCUMENTCLASS")
    
    If Len(strDocumentClass) > 0 Then
        xmlCopyAttribute xmlDocumentDetailsNode, xmlFileNetNode, "DOCUMENTCLASS"
    Else
        Set connOmiga = GetConnection()
    
        Set rec = connOmiga.Execute("SELECT VALIDATIONTYPE From HOSTTEMPLATE LEFT JOIN COMBOVALIDATION ON HOSTTEMPLATE.PERMANENCY = COMBOVALIDATION.VALUEID AND COMBOVALIDATION.GROUPNAME = 'DocumentPermanency' WHERE HOSTTEMPLATEID = '" & strHostTemplateId & "'")
    
        xmlSetAttributeValue xmlFileNetNode, "DOCUMENTCLASS", Format$(rec(0))
    End If

' Send the document to FileNet
    strFileNetUserId = GetMandatoryGlobalParamString("FileNetUserId")
    strFileNetGUID = DoRequest( _
        "getFileNetGUID", strApplicationNumber, "", _
        strFileNetUserId, xmlFileNetNode, strDocumentName)
    
    
    If strFileNetGUID = "Failed" Then
        Err.Raise 40017, "PackControl.SendDocumentToFileNet", _
            "Failed in omPack.PackManagerBO.SendDocumenttoFilenet, calling omPack.DoRequest.getFileNetGUID"
    End If
    strFileNetGUID = DeformatGUID(strFileNetGUID)
    
' Create the DocumentAuditDetails record
'TW 07/07/2006 MAR1897
'    strDocumentGuid = CreateDocumentAuditDetails(strApplicationNumber, strHostTemplateId)
    strDocumentGuid = CreateDocumentAuditDetails(strApplicationNumber, strHostTemplateId, strQuoteNumber)
'TW 07/07/2006 MAR1897 End
    
' Create the EventAuditDetail record
    strPackGUID = strZeroes
    Call CreateEventAuditDetail(strApplicationNumber, EVENTKEY_CREATED, strHostTemplateId, strDocumentGuid, strPackGUID, strFileNetGUID, strUnitId, strUserId)
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    Call AddAttribute(xmlResponseNode, "DOCUMENTGUID", strDocumentGuid)
    Call AddAttribute(xmlResponseNode, "FILENETIMAGEREF", strFileNetGUID)
    
    gobjTrace.TraceXML xmlResponseDoc.xml, "FileNet Document " & strApplicationNumber

TidyUp:
    On Error Resume Next
    SendDocumentToFileNet = xmlResponseDoc.xml
    On Error GoTo 0
    
    Set connOmiga = Nothing
    Set rec = Nothing
    
    Set xmlFileNetDoc = Nothing
    Set xmlResponseDoc = Nothing

    Set xmlAttributesNode = Nothing
    Set xmlDocumentDetailsNode = Nothing
    Set xmlErrorNode = Nothing
    Set xmlFileNetNode = Nothing
    Set xmlNode = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function
    
SendDocumentToFileNetError:
    Call DealWithErrorResponse(xmlResponseDoc, strFunctionName, Err)
    
    Resume TidyUp

'TW 12/06/2006 MAR1867 End
End Function


Private Function SendPackToEMail(ByVal xmlRequestNode As IXMLDOMNode) As String
'TW 12/06/2006 MAR1867 Rewritten for performance reasons

' Process an existing PACKCONTROL record.
' Returns RESPONSE xml string
' strPackId must be present
Const strFunctionName As String = "SendPackToEMail"

Dim strApplicationNumber As String
Dim strApplicationFactFindNumber As String
Dim strCustomerNumber As String
Dim strDocumentGuid As String
Dim strPackId As String
Dim strResponse As String
Dim strUserId As String
Dim strUnitId As String
Dim strXML As String

Dim xmlPackDataDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlCustomersToRemove As IXMLDOMNodeList
Dim xmlCustomerToRemove As IXMLDOMNode
Dim xmlEmailCustomer As IXMLDOMNode
Dim xmlEmailCustomers As IXMLDOMNodeList
Dim xmlPackAddressList As IXMLDOMNodeList
Dim xmlPackNodeList As IXMLDOMNodeList

Dim xmlAddress As IXMLDOMNode
Dim xmlCurrentEmailCustomer As IXMLDOMNode
Dim xmlPackAddressNode As IXMLDOMNode
Dim xmlPackNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

    On Error GoTo SendPackToEMailError:
    
' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

' Get the parameters from the Request Node
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetAttribute(xmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    strCustomerNumber = GetAttribute(xmlRequestNode, "CUSTOMERNUMBER")
    strPackId = xmlGetMandatoryAttributeText(xmlRequestNode, "PACKCONTROLNUMBER")
    strUnitId = GetAttribute(xmlRequestNode, "UNITID")
    strUserId = GetAttribute(xmlRequestNode, "USERID")
    
' Get the unique addressees for this application
    strXML = ExecStoredProc("usp_PACKCONTROL", strApplicationNumber, strApplicationFactFindNumber, "PACKDATA")

    If Len(strXML) = 0 Then
        Err.Raise 40027, "PackControl.SendPackToEMail", "Application '" & strApplicationNumber & "' can't find addresses for application '" & strApplicationNumber & "'"
    End If
    
    Set xmlPackDataDoc = xmlCreateDOMObject()
    xmlPackDataDoc.loadXML strXML
    
    ' Remove non relevant customers
    If Len(strCustomerNumber) > 0 Then
        ' Any customers that do not have the same customer number or are not Email
        Set xmlCustomersToRemove = xmlPackDataDoc.selectNodes("//ADDRESS/CUSTOMER[@CUSTOMERNUMBER!='" & strCustomerNumber & "' or @CONTACTPREFERENCE!='E']")
    Else
        ' Any customers that are not Email or have an Email address the same as a previous customer
        Set xmlCustomersToRemove = xmlPackDataDoc.selectNodes("//ADDRESS/CUSTOMER[@CONTACTPREFERENCE!='E' or @CONTACTEMAILADDRESS=preceding::CUSTOMER/@CONTACTEMAILADDRESS]")
    End If
   
    For Each xmlCustomerToRemove In xmlCustomersToRemove
        xmlCustomerToRemove.parentNode.removeChild xmlCustomerToRemove
    Next

    Set xmlPackAddressList = xmlPackDataDoc.selectNodes("//ADDRESS[CUSTOMER]")

    For Each xmlPackAddressNode In xmlPackAddressList

        ' Create PACKCONTROL node as a child of the RESPONSE node
        Set xmlPackNode = AddNode(xmlResponseDoc, "PACKCONTROL", xmlResponseNode)

        ' Add all the details for the customers
        xmlPackNode.appendChild xmlPackAddressNode.cloneNode(True)

        ' Get the unique Customer references for the PACK
        Call AddAttribute(xmlPackNode, "PACKGUID", CreateGUID())
        Call AddAttribute(xmlPackNode, "USERID", strUserId)
        Call AddAttribute(xmlPackNode, "UNITID", strUnitId)
        xmlSetAttributeValue xmlPackNode, "PACKTYPEID", strPackId
    Next
    
' Now send the packs to Fulfilment
    Set xmlPackNodeList = xmlResponseDoc.selectNodes("//PACKCONTROL")
    For Each xmlPackNode In xmlPackNodeList
    
        Set xmlAddress = xmlPackNode.selectSingleNode("ADDRESS")
        Set xmlEmailCustomers = xmlAddress.selectNodes("CUSTOMER")
        
        ' Remove the customers to process one at a time
        For Each xmlEmailCustomer In xmlEmailCustomers
            xmlAddress.removeChild xmlEmailCustomer
        Next
        
        ' Process each customer separately
        For Each xmlEmailCustomer In xmlEmailCustomers
            Set xmlCurrentEmailCustomer = xmlAddress.appendChild(xmlEmailCustomer)
            strResponse = DoRequest("sendToEmail", "", "", strUserId, xmlPackNode)
            If strResponse <> "Done" Then
                Err.Raise 40019, "PackControl.SendPackToEMail", "Failed to send Email Pack to Fullfilment: " & strResponse
            End If
            
            Call CreateAuditTrail(xmlPackNode, EVENTKEY_CREATED, strApplicationNumber, strDocumentGuid)
            Call AddAttribute(xmlPackNode, "DOCUMENTGUID", strDocumentGuid)
            Call CreateAuditTrail(xmlPackNode, EVENTKEY_EMAIL, strApplicationNumber, strDocumentGuid)
            
            gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "objOmFDM.InitialRequest Response : " & strResponse
            xmlAddress.removeChild xmlCurrentEmailCustomer
        Next
    Next
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    gobjTrace.TraceXML xmlResponseDoc.xml, "Processed Email " & strApplicationNumber

TidyUp:
    On Error Resume Next
    SendPackToEMail = xmlResponseDoc.xml
    
    On Error GoTo 0
    
    Set xmlPackDataDoc = Nothing
    Set xmlResponseDoc = Nothing

    Set xmlPackAddressNode = Nothing
    Set xmlPackNode = Nothing
    Set xmlResponseNode = Nothing
    
    Set xmlPackAddressList = Nothing
    Set xmlPackNodeList = Nothing

    Set xmlCustomersToRemove = Nothing
    Set xmlCustomerToRemove = Nothing
    
    Set xmlEmailCustomers = Nothing
    Set xmlEmailCustomer = Nothing
    Set xmlCurrentEmailCustomer = Nothing
    Set xmlAddress = Nothing

    Exit Function
    
SendPackToEMailError:
    Call DealWithErrorResponse(xmlResponseDoc, strFunctionName, Err)
    
    Resume TidyUp

End Function


Private Function SendPack(ByVal xmlRequestNode As IXMLDOMNode) As String
' Process an existing PACKCONTROL record.
' Returns RESPONSE xml string
' strPackId must be present
Const strFunctionName As String = "SendPack"

Dim intQuoteNumber As Integer

Dim strApplicationNumber As String
Dim strApplicationFactFindNumber As String
Dim strPackId As String
Dim strPackGUID As String
Dim strUserId As String
Dim strUnitId As String
Dim strMachineID As String
Dim strChannelID As String
Dim strUserAuthorityLevel As String

Dim strDocumentGuid As String
Dim strDocumentName As String ' MAR1191
Dim strHostTemplateName As String
Dim strFileNetGUID As String
Dim strPackName As String
Dim strPackDescription As String
Dim strQuoteArray() As String
'TW 07/07/2006 MAR1897
Dim strQuoteNumber As String
'TW 07/07/2006 MAR1897 End
Dim strSQL As String
Dim strXML As String
Dim strFileNetUserId As String      ' PSC 11/01/2005 MAR994
Dim strTemplateID As String
Dim strResponse As String

Dim xmlPackDataDoc As FreeThreadedDOMDocument40
Dim xmlPropertyDoc As FreeThreadedDOMDocument40
Dim xmlPackDefinitionDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlApplicationPackDetails As FreeThreadedDOMDocument40

#If FILENET Then
    Dim xmlFileNetDoc As FreeThreadedDOMDocument40
    Dim xmlFileNetNode As IXMLDOMNode
#ElseIf Gemini Then
    Dim xmlExistingPackDocs As FreeThreadedDOMDocument40
    Dim xmlExistingPackDocNode As IXMLDOMNode
    Dim xmlPreCreateList As IXMLDOMNodeList
    Dim xmlSavePrintDocumentRequestDoc As FreeThreadedDOMDocument40
    Dim xmlSavePrintDocumentResponseDoc As FreeThreadedDOMDocument40
    Dim xmlSavePrintDocumentRequestNode As IXMLDOMNode
    Dim xmlControlDataNode As IXMLDOMNode
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlPrintDocumentDataNode As IXMLDOMNode
    Dim xmlCreateFileDataNode As IXMLDOMNode
    Dim xmlPreCreatedNode As IXMLDOMNode
    'AW EP2_2181
    Dim xmlOptionalPackDocs As FreeThreadedDOMDocument40
    Dim xmlOptionalPackList As IXMLDOMNodeList
    Dim xmlOptionalItem As IXMLDOMNode
    Dim xmlOptionalPackElem As IXMLDOMElement
    Dim xmlOptionalDocNode As IXMLDOMNode
    Dim xmlTempDocNode As IXMLDOMNode
    Dim colValidationTypes As Collection
    Dim nArrayInex As Integer
    Dim strValidationType As String
    Dim dtDate1 As Date
    Dim dtDate2 As Date
    Dim strTemplateIDTemp As String
    Dim blnIsFirstTemplate As Boolean
    Dim strPackDocumentType As String
    Dim strMissingList As String
    'AW EP2_2181   -   End
#End If
    
Dim xmlPackAddressList As IXMLDOMNodeList
Dim xmlPackCopyNodeList As IXMLDOMNodeList
Dim xmlPackDefinitionNodeList As IXMLDOMNodeList
Dim xmlPackNodeList As IXMLDOMNodeList

Dim xmlErrorNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlPackAddressNode As IXMLDOMNode
Dim xmlPackCopyNode As IXMLDOMNode
Dim xmlPackDefinitionNode As IXMLDOMNode
Dim xmlPackElementNode As IXMLDOMNode
Dim xmlPackMemberNode As IXMLDOMNode
Dim xmlApplicationPackNode As IXMLDOMNode
Dim xmlAppPackMemberNode As IXMLDOMNode

Dim xmlPackNode As IXMLDOMNode
Dim xmlPropertyNode As IXMLDOMNode
Dim xmlQuoteNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

Dim xmlPackDefinitionNodeAttribute As IXMLDOMAttribute
    
Dim blnFirstRecord As Boolean
    
Dim cmd As ADODB.Command
Dim connOmiga As Connection
Dim rec As ADODB.Recordset
Dim strCreationDate As String 'BC MAR1106 01/02/2006
Dim strPackDestination As String 'PE 10/05/2006

    strCreationDate = Format$(Now, "yyyymmdd_hhmmss") 'BC MAR1106 01/02/2006
    On Error GoTo SendPackError:
    
' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

' Get the parameters from the Request Node
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetAttribute(xmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    strPackId = GetAttribute(xmlRequestNode, "PACKCONTROLNUMBER")
    strUserId = GetAttribute(xmlRequestNode, "USERID")
    strUnitId = GetAttribute(xmlRequestNode, "UNITID")
    strMachineID = GetAttribute(xmlRequestNode, "MACHINEID")
    strChannelID = GetAttribute(xmlRequestNode, "CHANNELID")
    strUserAuthorityLevel = GetAttribute(xmlRequestNode, "USERAUTHORITYLEVEL")
    
    If Len(strPackId) = 0 Then
        Err.Raise 40010, "PackControl.SendPack", "No pack control number supplied"
    End If
    
    Set xmlPackDefinitionDoc = xmlCreateDOMObject()
    xmlPackDefinitionDoc.loadXML BuildPackXML(strPackId)
    
    strPackName = GetAttribute(xmlPackDefinitionDoc.selectSingleNode("//PACKCONTROL"), "PACKNAME")
    strPackDescription = GetAttribute(xmlPackDefinitionDoc.selectSingleNode("//PACKCONTROL"), "PACKDESCRIPTION")
    If Len(strPackDescription) <> 0 Then
        strPackDescription = strPackDescription & "(" & strPackId & ")"
    Else
        strPackDescription = strPackId
    End If
    
    ' PE MAR1756 - 10/05/2006
    strPackDestination = Trim$(UCase$(GetAttribute(xmlPackDefinitionDoc.selectSingleNode("//PACKCONTROL/PACKMEMBER"), "PACKDESTINATION")))

#If Gemini Then
    'AW EP1271 - Get existing documents in this pack.
    Set xmlExistingPackDocs = xmlCreateDOMObject()
    strXML = GetExistingPackDocuments(strApplicationNumber, strPackId)
    xmlExistingPackDocs.loadXML "<DOCUMENTLIST>" & strXML & "</DOCUMENTLIST>"
    
    'AW EP2_2181
    Set xmlOptionalPackDocs = xmlCreateDOMObject()
    Set xmlOptionalPackElem = xmlOptionalPackDocs.createElement("LIST")
    xmlOptionalPackDocs.appendChild xmlOptionalPackElem
    nArrayInex = 0
    Set colValidationTypes = New Collection
    'AW EP2_2181   -   End
    
    Set xmlPreCreateList = xmlPackDefinitionDoc.selectNodes("//PACKMEMBER[@PRECREATEFORPACK='1']")
    If xmlPreCreateList.length > 0 Then
        For Each xmlPackMemberNode In xmlPreCreateList
            strTemplateID = GetAttribute(xmlPackMemberNode, "HOSTTEMPLATEID")
            'AW EP2_2181
            strPackDocumentType = GetAttribute(xmlPackMemberNode, "PACKDOCUMENTTYPE")
            
            If strPackDocumentType = "" Then
                If xmlExistingPackDocs.selectSingleNode("//DOCUMENTDETAILS[@HOSTTEMPLATEID='" & strTemplateID & "']") Is Nothing Then
                    Err.Raise 40020, "PackControl.SendPack", "Pack '" & strPackDescription & "' requires the following pre-created document: '" & GetAttribute(xmlPackMemberNode, "HOSTTEMPLATENAME") & "'"
                End If
            Else
            
                'Maintain list of discrete validation types
                If Not IsItemInCollection(colValidationTypes, strPackDocumentType) Then
                    colValidationTypes.Add (strPackDocumentType)
                End If
                
                'Creat temp list of optional documents
                Set xmlTempDocNode = xmlExistingPackDocs.selectSingleNode("//DOCUMENTDETAILS[@HOSTTEMPLATEID='" & strTemplateID & "']")
                If Not xmlTempDocNode Is Nothing Then
                    
                    Set xmlOptionalDocNode = xmlOptionalPackDocs.createElement("ITEM")
                
                    Call AddAttribute(xmlOptionalDocNode, "TEMPLATEID", strTemplateID)
                    Call AddAttribute(xmlOptionalDocNode, "VALIDATIONTYPE", strPackDocumentType)
                    Call AddAttribute(xmlOptionalDocNode, "EVENTDATE", GetAttribute(xmlTempDocNode, "EVENTDATE"))
                    
                    Set xmlTempDocNode = xmlOptionalPackElem.appendChild(xmlOptionalDocNode)
                End If
                
                nArrayInex = nArrayInex + 1
                
            End If
        Next
        
        'At least one document for each validation type must exist
        For nArrayInex = 1 To colValidationTypes.Count
        
            strValidationType = colValidationTypes(nArrayInex)
            
            Set xmlOptionalPackList = xmlOptionalPackDocs.selectNodes("//ITEM[@VALIDATIONTYPE='" & strValidationType & "']")
            If xmlOptionalPackList.length = 0 Then
                'AW EP2_2181
                For Each xmlPackNode In xmlPackDefinitionDoc.selectNodes("//PACKMEMBER[@PACKDOCUMENTTYPE='" & strValidationType & "']")
                    strMissingList = " " & strMissingList & GetAttribute(xmlPackNode, "HOSTTEMPLATENAME") & ","
                Next
                
                Err.Raise 40020, "PackControl.SendPack", "Pack '" & strPackDescription & "' one of the following pre-created documents is missing: '" & strMissingList
            Else
                If xmlOptionalPackList.length = 1 Then
                    
                    Set xmlOptionalItem = xmlOptionalPackList.Item(0)
                    strTemplateID = GetAttribute(xmlOptionalItem, "TEMPLATEID")
                    strPackDocumentType = GetAttribute(xmlOptionalItem, "VALIDATIONTYPE")
                    
                    'remove all other optional documents from pack definition
                    For Each xmlPackNode In xmlPackDefinitionDoc.selectNodes("//PACKMEMBER[@PACKDOCUMENTTYPE='" & strValidationType & "' and @HOSTTEMPLATEID!='" & strTemplateID & "']")
                            xmlPackNode.parentNode.removeChild xmlPackNode
                    Next
                   
                Else
                    'More than 1 of the optional documents are available, have to select
                    'the most recent one then remove the remaining optional docs from pack definition
                    blnIsFirstTemplate = True
                    
                    For Each xmlOptionalItem In xmlOptionalPackList
                        
                        If (blnIsFirstTemplate) Then
                           dtDate1 = CDate(GetAttribute(xmlOptionalItem, "EVENTDATE"))
                           strTemplateID = GetAttribute(xmlOptionalItem, "TEMPLATEID")
                           
                           blnIsFirstTemplate = False
                        Else
                            dtDate2 = CDate(GetAttribute(xmlOptionalItem, "EVENTDATE"))
                            strTemplateIDTemp = GetAttribute(xmlOptionalItem, "TEMPLATEID")
                            
                            If (dtDate2 > dtDate1) Then
                            
                                Set xmlPackNode = xmlPackDefinitionDoc.selectSingleNode("//PACKMEMBER[@HOSTTEMPLATEID='" & strTemplateID & "']")
                                If Not xmlPackNode Is Nothing Then
                                    xmlPackNode.parentNode.removeChild xmlPackNode
                                End If
                                
                                dtDate1 = dtDate2
                                strTemplateID = strTemplateIDTemp
                            
                            End If
                            
                        End If
                        
                    Next xmlOptionalItem
                    'AW EP2_2181
                    'remove all other optional documents from pack definition except our most recent
                    For Each xmlPackNode In xmlPackDefinitionDoc.selectNodes("//PACKMEMBER[@PACKDOCUMENTTYPE='" & strValidationType & "' and @HOSTTEMPLATEID!='" & strTemplateID & "']")
                            xmlPackNode.parentNode.removeChild xmlPackNode
                    Next
                        
                End If
            End If

        Next nArrayInex
        'AW EP2_2181   -   End
    End If
    'AW EP1271 - End
#End If

    ' Get the unique addressees for this application
    ' PE MAR1756 - 10/05/2006
    strXML = GetAddress(strPackDestination, strApplicationNumber, strApplicationFactFindNumber)
    
    If Len(strXML) = 0 Then
        Err.Raise 40020, "PackControl.SendPack", "Pack '" & strPackDescription & "' can't find addresses for application '" & strApplicationNumber & "'"
    End If
    
    Set xmlPackDataDoc = xmlCreateDOMObject()
    xmlPackDataDoc.loadXML "<PACKDATA>" & strXML & "</PACKDATA>"
    
    'MAR1193 DRC - need to see if property location is Scotland
    ' Is property in scotland for this application

    strXML = ExecStoredProc("usp_SCOTTISHPROPERTY", strApplicationNumber, strApplicationFactFindNumber)

    If Len(strXML) = 0 Then
        Err.Raise 40020, "PackControl.SendPack", "Pack '" & strPackDescription & "' can't find property location for application '" & strApplicationNumber & "'"
    End If
    
    Set xmlPropertyDoc = xmlCreateDOMObject()
    xmlPropertyDoc.loadXML strXML
    Set xmlPropertyNode = xmlPropertyDoc.selectSingleNode("//PROPERTY")
    
    'MAR1193 - End

' Get details of the quote(s) required
    ReDim strQuoteArray(0)
    Set xmlNode = xmlPackDataDoc.selectSingleNode("//PACKDATA")
    If (GetAttribute(xmlPackDefinitionDoc.selectSingleNode("//PACKCONTROL"), "USEMULTIPLEQUOTES") = "1") Then ' BC MAR1155 01/02/2006
        
        Set connOmiga = GetConnection()
    
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = connOmiga
        
        strSQL = "SELECT QUOTATIONNUMBER FROM QUOTATION WHERE APPLICATIONNUMBER = '" & strApplicationNumber & "' AND APPLICATIONFACTFINDNUMBER = '" & strApplicationFactFindNumber & "' AND KFIREQUIREDINDICATOR > 0"
        
        Set rec = connOmiga.Execute(strSQL)
        Do While Not rec.EOF
            ReDim Preserve strQuoteArray(0 To intQuoteNumber)
            strQuoteArray(intQuoteNumber) = rec!QUOTATIONNUMBER
            intQuoteNumber = intQuoteNumber + 1
            rec.MoveNext
        Loop
    End If
    If intQuoteNumber = 0 Then
        Set xmlQuoteNode = xmlPackDataDoc.selectSingleNode("//ADDRESS")

        strQuoteNumber = GetQuoteNumber(xmlQuoteNode)
        
        If Len(strQuoteNumber) > 0 Then
            strQuoteArray(0) = strQuoteNumber
        Else
            Err.Raise 40016, "PackControl.SendPack", "Pack '" & strPackDescription & "' no active or accepted quote for application '" & strApplicationNumber & "'"
        End If
    End If

' Release the database objects
    
    If Not connOmiga Is Nothing Then
        If connOmiga.State = adStateOpen Then connOmiga.Close
        Set connOmiga = Nothing
    End If
        
    Set cmd = Nothing
    Set rec = Nothing

    blnFirstRecord = True ' PSC 15/02/2006 MAR1271

    Set xmlPackAddressList = xmlPackDataDoc.selectNodes("//ADDRESS")
    For Each xmlPackAddressNode In xmlPackAddressList
        
    ' Create PACKCONTROL node as a child of the RESPONSE node
        Set xmlPackNode = AddNode(xmlResponseDoc, "PACKCONTROL", xmlResponseNode)

    ' Add all the details for the customers
        xmlPackNode.appendChild xmlPackAddressNode.cloneNode(True)

    ' Get the unique Customer references for the PACK
        strPackGUID = CreateGUID()
        Call AddAttribute(xmlPackNode, "PACKGUID", strPackGUID)       ' MAR978
        Call AddAttribute(xmlPackNode, "USERID", strUserId)
        Call AddAttribute(xmlPackNode, "UNITID", strUnitId)
        Call AddAttribute(xmlPackNode, "PACKCONTROLNUMBER", strPackId)
        Call AddAttribute(xmlPackNode, "PACKNAME", strPackName)        ' MAR978

'MAR1193 Copy the Property in Scotland indicator
        Call AddAttribute(xmlPackNode, "INSCOTLAND", GetAttribute(xmlPropertyNode, "INSCOTLAND"))
        
        Set xmlPackDefinitionNodeList = xmlPackDefinitionDoc.selectNodes("//PACKMEMBER[@PACKMEMBERTYPE='2']")
        For Each xmlPackDefinitionNode In xmlPackDefinitionNodeList

            For intQuoteNumber = 0 To UBound(strQuoteArray, 1)
' Create a PACKMEMBER node as a child of the PACK node
                Set xmlPackMemberNode = AddNode(xmlResponseDoc, "PACKMEMBER", xmlPackNode)
                
                Dim HostTemplateId As String
                Dim dpsTemplateId As String
                Dim templateGroupdId As String
                Dim hostTemplateDescription As String
                Dim precreated As Boolean
                
                HostTemplateId = GetAttribute(xmlPackDefinitionNode, "HOSTTEMPLATEID")
                dpsTemplateId = GetAttribute(xmlPackDefinitionNode, "DPSTEMPLATEID")
                templateGroupdId = GetAttribute(xmlPackDefinitionNode, "TEMPLATEGROUPID")
                hostTemplateDescription = GetAttribute(xmlPackDefinitionNode, "HOSTTEMPLATEDESCRIPTION")
                precreated = False
                
                Call AddAttribute(xmlPackMemberNode, "METHODNAME", GetAttribute(xmlPackDefinitionNode, "PDMMETHOD"))
                Call AddAttribute(xmlPackMemberNode, "DOCUMENTID", HostTemplateId)
                Call AddAttribute(xmlPackMemberNode, "DPSDOCUMENTID", dpsTemplateId)
                Call AddAttribute(xmlPackMemberNode, "TEMPLATEGROUPID", templateGroupdId)
                Call AddAttribute(xmlPackMemberNode, "DOCUMENTDESCRIPTION", hostTemplateDescription)
#If Gemini Then
                Call AddAttribute(xmlPackMemberNode, "PACKMEMBERNUMBER", GetAttribute(xmlPackDefinitionNode, "PACKMEMBERNUMBER"))
                Set xmlPreCreatedNode = xmlPackDefinitionDoc.selectSingleNode("//PACKMEMBER[@HOSTTEMPLATEID='" & HostTemplateId & "' and @PRECREATEFORPACK='1']")
                If Not xmlPreCreatedNode Is Nothing Then
                    ' This document must be pre-created.
                    Set xmlExistingPackDocNode = xmlExistingPackDocs.selectSingleNode("//DOCUMENTDETAILS[@HOSTTEMPLATEID='" & HostTemplateId & "']")
                    If Not xmlExistingPackDocNode Is Nothing Then
                        ' There is an existing pre-created document.
                        precreated = True
                        strDocumentGuid = GetAttribute(xmlExistingPackDocNode, "DOCUMENTGUID")
                    Else
                        Err.Raise 40020, "PackControl.SendPack", "Pack '" & strPackDescription & "' requires the following pre-created document: '" & GetAttribute(xmlPackDefinitionNode, "HOSTTEMPLATENAME") & "'"
                    End If
                End If
#End If
'TW 12/06/2006 MAR1867
                Call AddAttribute(xmlPackMemberNode, "DELIVERYTYPE", GetAttribute(xmlPackDefinitionNode, "DELIVERYTYPE"))
'TW 12/06/2006 MAR1867 End
                
                ' MAR1191 Start
                strHostTemplateName = GetAttribute(xmlPackDefinitionNode, "HOSTTEMPLATENAME")
                strDocumentName = strHostTemplateName & _
                                  "_" & strApplicationNumber & _
                                  "_" & strQuoteArray(intQuoteNumber) & "_" & strCreationDate
                AddAttribute xmlPackMemberNode, "DOCUMENTNAME", strDocumentName ' MAR1216
                ' MAR1191 End

                If Not precreated Then
                    ' Generate the PACKELEMENT document
                    strResponse = ProcessPackMember(xmlPackAddressNode, _
                             strPackId, _
                             strUserId, _
                             strUnitId, _
                             strMachineID, _
                             strChannelID, _
                             strUserAuthorityLevel, _
                             strQuoteArray(intQuoteNumber), _
                             xmlPackDefinitionNode, _
                             strDocumentName) ' MAR1191
                    'BC MAR880
                    If strResponse = "FAILED" Then
                        Err.Raise 40018, "PackControl." & strFunctionName, "Failure in omPACK.ProcessPackMember"
                    End If
                End If
                
                Dim intDeliveryType As Integer
                intDeliveryType = xmlGetAttributeAsInteger(xmlPackDefinitionNode, "DELIVERYTYPE")

                Dim strFileGuid As String
                Dim strFileVersion As String
                Dim geminiPrintStatus As Integer
#If FILENET Then
                Set xmlFileNetDoc = xmlCreateDOMObject()
                Set xmlFileNetNode = AddNode(xmlFileNetDoc, "FILECONTENTS")
                Call AddAttribute(xmlFileNetNode, "FILECONTENTS", strResponse)
                
                ' PSC 06/02/2006 MAR1197 - Start
                xmlSetAttributeValue xmlFileNetNode, "CONTENTTYPE", GetValidationTypeForValueID("DocumentDeliveryType", intDeliveryType)
                ' PSC 06/02/2006 MAR1197 - End

                'PE 13/04/2006 MAR1405
                xmlSetAttributeValue xmlFileNetNode, "DOCUMENTCLASS", GetAttribute(xmlPackDefinitionNode, "DOCUMENTCLASS")

' Now send the document to FileNet
                ' PSC 15/12/2005 MAR880
                ' PSC 11/01/2006 MAR994 - Start
                strFileNetUserId = GetMandatoryGlobalParamString("FileNetUserId")
                strFileNetGUID = DoRequest("getFileNetGUID", strApplicationNumber, "", strFileNetUserId, xmlFileNetNode, strDocumentName) ' MAR1191
                 ' PSC 11/01/2006 MAR994 - End
                gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "objOmFDM.getFileNetGUID Response : " & strFileNetGUID
        
                ' MAR1867 Check that omFDM has returned a GUID (enclosed in {})
                If Mid$(strFileNetGUID, 1, 1) <> "{" Then
                    Err.Raise 40017, "Failed in SendDocumenttoFilenet, calling omPACK.DoRequest.getFileNetGUID"
                End If
                
                Call AddAttribute(xmlPackMemberNode, "FILENETIMAGEREF", strFileNetGUID)

                ''TW 07/07/2006 MAR1897
                Call AddAttribute(xmlPackMemberNode, "QUOTATIONNUMBER", strQuoteArray(intQuoteNumber))
                ''TW 07/07/2006 MAR1897 End
                
                If Not precreated Then
                    Call CreateAuditTrail(xmlPackMemberNode, EVENTKEY_CREATED, strApplicationNumber, strDocumentGuid, strFileGuid, strFileVersion, geminiPrintStatus)
                End If

#ElseIf Gemini Then
                If Not precreated Then
                    'AS 27/12/2006 EP1271 START Archive document in DMS/Gemini.
                    Set xmlSavePrintDocumentRequestDoc = xmlCreateDOMObject()
                    Set xmlSavePrintDocumentRequestNode = AddNode(xmlSavePrintDocumentRequestDoc, "REQUEST")
                    AddAttribute xmlSavePrintDocumentRequestNode, "USERID", strUserId
                    AddAttribute xmlSavePrintDocumentRequestNode, "UNITID", strUnitId
                    AddAttribute xmlSavePrintDocumentRequestNode, "OPERATION", "SAVEPRINTDOCUMENT"
                    Set xmlPrintDataNode = AddNode(xmlSavePrintDocumentRequestDoc, "PRINTDATA", xmlSavePrintDocumentRequestNode)
                    Set xmlControlDataNode = AddNode(xmlSavePrintDocumentRequestDoc, "CONTROLDATA", xmlSavePrintDocumentRequestNode)
                    Set xmlPrintDocumentDataNode = AddNode(xmlSavePrintDocumentRequestDoc, "PRINTDOCUMENTDATA", xmlSavePrintDocumentRequestNode)
                    AddAttribute xmlPrintDataNode, "APPLICATIONNUMBER", strApplicationNumber
                    AddAttribute xmlControlDataNode, "HOSTTEMPLATENAME", strHostTemplateName
                    AddAttribute xmlControlDataNode, "DOCUMENTID", HostTemplateId
                    AddAttribute xmlControlDataNode, "DPSDOCUMENTID", dpsTemplateId
                    AddAttribute xmlControlDataNode, "TEMPLATEGROUPID", templateGroupdId
                    AddAttribute xmlControlDataNode, "HOSTTEMPLATEDESCRIPTION", hostTemplateDescription
                    AddAttribute xmlControlDataNode, "DESTINATIONTYPE", "W" ' Prevents creation of document.
                    AddAttribute xmlControlDataNode, "DELIVERYTYPE", CStr(intDeliveryType)
                    AddAttribute xmlControlDataNode, "PACKFULFILLMENTGUID", strPackGUID
                    AddAttribute xmlControlDataNode, "GEMINIPRINTSTATUS", CStr(DEFAULT_GEMINIPRINTSTATUS)
                    AddAttribute xmlPrintDocumentDataNode, "CREATEDBY", strUserId
                    AddAttribute xmlPrintDocumentDataNode, "FILENAME", strDocumentName
                    AddAttribute xmlPrintDocumentDataNode, "DELIVERYTYPE", CStr(intDeliveryType)
                    AddAttribute xmlPrintDocumentDataNode, "COMPRESSED", "0"
                    AddAttribute xmlPrintDocumentDataNode, "COMPRESSIONMETHOD", ""
                    AddAttribute xmlPrintDocumentDataNode, "FILECONTENTS_TYPE", DEFAULT_FILECONTENTS_TYPE
                    AddAttribute xmlPrintDocumentDataNode, "FILECONTENTS", strResponse
                    
                    Dim objPM As PrintManagerBO
                    If gobjContext Is Nothing Then
                        Set objPM = New PrintManagerBO
                    Else
                        Set objPM = gobjContext.CreateInstance(gstrPRINTMANAGER_COMPONENT & ".PrintManagerBO")
                    End If
                    Dim response As String
                    response = objPM.omRequest(xmlSavePrintDocumentRequestDoc.xml)
                    errCheckXMLResponse response, True
                    Set xmlSavePrintDocumentResponseDoc = xmlCreateDOMObject()
                    xmlSavePrintDocumentResponseDoc.loadXML response
                    strDocumentGuid = GetAttribute(xmlSavePrintDocumentResponseDoc.documentElement, "DOCUMENTGUID")
                    strFileGuid = GetAttribute(xmlSavePrintDocumentResponseDoc.documentElement, "FILEGUID")
                    strFileVersion = DEFAULT_FILEVERSION
                    geminiPrintStatus = DEFAULT_GEMINIPRINTSTATUS
                    Call AddAttribute(xmlPackMemberNode, "QUOTATIONNUMBER", strQuoteArray(intQuoteNumber))
                    'AS 27/12/2006 EP1271 END
                End If
#End If

                Call AddAttribute(xmlPackMemberNode, "DOCUMENTGUID", strDocumentGuid)
#If Gemini Then
                If Not precreated Then
                    'AS 27/12/2006 EP1271 START
                    Call AddAttribute(xmlPackMemberNode, "FILECONTENTS", strResponse)
                    Call AddAttribute(xmlPackMemberNode, "FILECONTENTS_TYPE", DEFAULT_FILECONTENTS_TYPE)
                    Call AddAttribute(xmlPackMemberNode, "COMPRESSIONMETHOD", "")
                    'AS 27/12/2006 EP1271 END
                End If
#End If
                
                If xmlPackDefinitionNode Is xmlPackDefinitionNodeList(0) And intQuoteNumber = 0 Then
                    Call AddAttribute(xmlPackMemberNode, "PRIMARY", "true")
                    'Peter Edney - 13/02/2006
                    'MAR1155 - USE MULTIPLE QUOTES ON PACK DOES INVOKE CREATING MULTIPLE KFI ON SEND KFI PACK
                    'BC MAR1271 Back out MAR1155
                    Exit For
                End If
            Next intQuoteNumber
            If xmlPackDefinitionNode Is xmlPackDefinitionNodeList(0) And Not blnFirstRecord Then
                Exit For
            End If
            
        Next ' For Each xmlPackDefinitionNode In xmlPackDefinitionNodeList
        
        If Not xmlPackAddressNode Is xmlPackAddressList(0) Then

            ' Copy all the documents from the first pack (excluding the cover letter) to all other packs
            Set xmlPackCopyNodeList = xmlResponseNode.selectNodes("//PACKCONTROL[1]/PACKMEMBER")
            For Each xmlPackCopyNode In xmlPackCopyNodeList
                If Not xmlPackCopyNode Is xmlPackCopyNodeList(0) Then
                    xmlPackNode.appendChild xmlPackCopyNode.cloneNode(True)
                End If
            Next
        End If
        blnFirstRecord = False
        
    Next ' For Each xmlPackAddressNode In xmlPackAddressList
    
    gobjTrace.TraceXML xmlResponseDoc.xml, "Before Fullfilment " & strPackId
    
    ' Now send the packs to Fulfilment
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1a]", "Sending packs to Fulfilment"
    Set xmlPackNodeList = xmlResponseDoc.selectNodes("//PACKCONTROL")
    For Each xmlPackNode In xmlPackNodeList
#If FILENET Then
        strResponse = DoRequest("initialRequest", "", "", strUserId, xmlPackNode)
        'MAR1638
        If strResponse <> "Done" Then
            Err.Raise 40019, "PackControl.SendPack", "Failed to send Pack to Fulfilment: " & strResponse
        End If
#ElseIf Gemini Then
        Set xmlApplicationPackDetails = xmlCreateDOMObject()
        Set xmlApplicationPackNode = AddNode(xmlApplicationPackDetails, "APPLICATIONPACK")
        Call AddAttribute(xmlApplicationPackNode, "PACKFULFILLMENTGUID", xmlGetAttributeText(xmlPackNode, "PACKGUID"))
        Call AddAttribute(xmlApplicationPackNode, "APPLICATIONNUMBER", strApplicationNumber)
        Call AddAttribute(xmlApplicationPackNode, "PACKCONTROLNUMBER", strPackId)
        'AS 27/12/2006 EP1271 START
        Call AddAttribute(xmlApplicationPackNode, "USERID", strUserId)
        Call AddAttribute(xmlApplicationPackNode, "UNITID", strUnitId)
        'AS 27/12/2006 EP1271 END
#End If
        
        For Each xmlPackMemberNode In xmlPackNode.selectNodes("./PACKMEMBER")
        
            'MAR1867 Get the Document GUID for each Pack Member
            strDocumentGuid = xmlGetAttributeText(xmlPackMemberNode, "DOCUMENTGUID")
#If FILENET Then
            Call CreateAuditTrail(xmlPackMemberNode, EVENTKEY_FULFILLMENTSEND, strApplicationNumber, strDocumentGuid)
#ElseIf Gemini Then
            Set xmlAppPackMemberNode = AddNode(xmlApplicationPackDetails, "APPLICATIONPACKMEMBER", xmlApplicationPackNode)
            Call AddAttribute(xmlAppPackMemberNode, "DOCUMENTGUID", strDocumentGuid)
            Call AddAttribute(xmlAppPackMemberNode, "PACKMEMBERNUMBER", xmlGetAttributeText(xmlPackMemberNode, "PACKMEMBERNUMBER"))
            Call AddAttribute(xmlAppPackMemberNode, "DELIVERYTYPE", xmlGetAttributeText(xmlPackMemberNode, "DELIVERYTYPE"))
            Dim strFileContents As String
            strFileContents = xmlGetAttributeText(xmlPackMemberNode, "FILECONTENTS")
            If Len(strFileContents) > 0 Then
                Call AddAttribute(xmlAppPackMemberNode, "COMPRESSIONMETHOD", "")
                Call AddAttribute(xmlAppPackMemberNode, "FILECONTENTS_TYPE", xmlGetAttributeText(xmlPackMemberNode, "FILECONTENTS_TYPE"))
                Call AddAttribute(xmlAppPackMemberNode, "FILECONTENTS", xmlGetAttributeText(xmlPackMemberNode, "FILECONTENTS"))
            End If
#End If
        Next
        
#If FILENET Then
        gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "objOmFDM.InitialRequest Response : " & strResponse
#ElseIf Gemini Then
        CreateApplicationPack (xmlApplicationPackDetails.xml)
        Set xmlApplicationPackDetails = Nothing
#End If

    Next
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2a]", "Finished sending packs to Fulfilment"
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    gobjTrace.TraceXML xmlResponseDoc.xml, "Processed " & strPackId

#If FILENET Then
    'AS 27/12/2006 EP1271 Why is this commented out for Gemini?
    Set connOmiga = GetConnection()
    strSQL = "UPDATE QUOTATION SET KFIPRINTEDINDICATOR = 1, KFIREQUIREDINDICATOR = 0 WHERE APPLICATIONNUMBER = '" & strApplicationNumber & "' AND APPLICATIONFACTFINDNUMBER = '" & strApplicationFactFindNumber & "' AND KFIREQUIREDINDICATOR > 0"
    connOmiga.Execute strSQL
#End If
    
TidyUp:
    On Error Resume Next
    SendPack = xmlResponseDoc.xml
    
    On Error GoTo 0
    
    Set xmlPackDataDoc = Nothing
    Set xmlPackDefinitionDoc = Nothing
    Set xmlApplicationPackDetails = Nothing
    Set xmlResponseDoc = Nothing
    
#If FILENET Then
    Set xmlFileNetDoc = Nothing
    Set xmlFileNetNode = Nothing
#ElseIf Gemini Then
    Set xmlExistingPackDocs = Nothing
    Set xmlExistingPackDocNode = Nothing
    Set xmlPreCreateList = Nothing
    Set xmlSavePrintDocumentResponseDoc = Nothing
    Set xmlSavePrintDocumentRequestNode = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlPrintDocumentDataNode = Nothing
    Set xmlCreateFileDataNode = Nothing
    Set xmlPreCreatedNode = Nothing
    Set objPM = Nothing
    'AW EP2_2181
    Set xmlOptionalPackDocs = Nothing
    Set xmlOptionalPackList = Nothing
    Set xmlOptionalItem = Nothing
    Set xmlOptionalPackElem = Nothing
    Set xmlOptionalDocNode = Nothing
    Set xmlTempDocNode = Nothing
    'AW EP2_2181    -End
#End If

    Set xmlErrorNode = Nothing
    Set xmlNode = Nothing
    Set xmlPackAddressNode = Nothing
    Set xmlPackCopyNode = Nothing
    Set xmlPackDefinitionNode = Nothing
    Set xmlPackElementNode = Nothing
    Set xmlPackMemberNode = Nothing
    Set xmlPackNode = Nothing
    Set xmlQuoteNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlPropertyDoc = Nothing
    Set xmlPropertyNode = Nothing
    Set xmlPackAddressList = Nothing
    Set xmlPackCopyNodeList = Nothing
    Set xmlPackDefinitionNodeList = Nothing
    Set xmlPackNodeList = Nothing
    Set xmlAppPackMemberNode = Nothing
    Set xmlApplicationPackNode = Nothing
    Set xmlPackDefinitionNodeAttribute = Nothing
    
    Set connOmiga = Nothing
    Set cmd = Nothing
    
    Set rec = Nothing
    
    Exit Function
    
SendPackError:
    Call DealWithErrorResponse(xmlResponseDoc, strFunctionName, Err)
    
    Resume TidyUp

End Function

Private Function CancelPack(ByVal xmlRequestNode As IXMLDOMNode) As String
'TW 12/06/2006 MAR1867 Rewritten for performance reasons

Const strFunctionName = "CancelPack"

Dim strApplicationNumber As String
Dim strDocumentGuid As String
Dim strFileNetImageRef As String
Dim strHostTemplateId As String
Dim strPackGUID As String
Dim strPackId As String
Dim strResponse As String
Dim strUnitId As String
Dim strUserId As String

Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlEventNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

' Cancels an existing PACK
' Returns RESPONSE xml string
    On Error GoTo CancelPackError:

' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

    Set xmlEventNode = xmlRequestNode.selectSingleNode("//EVENTDETAIL")
    If xmlEventNode Is Nothing Then
        Err.Raise 40025, "PackControl.CancelPack", "No EVENTDETAIL node supplied"
    End If
    
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strHostTemplateId = GetAttribute(xmlEventNode, "HOSTTEMPLATEID")
    strPackId = GetAttribute(xmlEventNode, "PACKFULFILLMENTGUID")
    strUserId = GetAttribute(xmlRequestNode, "USERID")
    strUnitId = GetAttribute(xmlRequestNode, "UNITID")

    strResponse = DoRequest("cancelPack", "", strPackId, strUserId, xmlEventNode)
    If (strResponse <> "Done") Then
        Err.Raise 40023, "PackControl.CancelPack", "Failed to Cancel Pack in Fulfilment: " & strResponse
    End If
    
' Create the EventAuditDetail record
    strPackGUID = strZeroes
    strFileNetImageRef = strZeroes
    strDocumentGuid = GetAttribute(xmlRequestNode, "DOCUMENTGUID")
    
    Call CreateEventAuditDetail(strApplicationNumber, EVENTKEY_FULFILLMENTCANCEL, strHostTemplateId, strDocumentGuid, strPackGUID, strFileNetImageRef, strUnitId, strUserId)
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    gobjTrace.TraceXML xmlResponseDoc.xml, "CancelPack " & strPackId
   
TidyUp:
    On Error Resume Next
    CancelPack = xmlResponseDoc.xml
    On Error GoTo 0
    
    Set xmlResponseDoc = Nothing
    
    Set xmlEventNode = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function
    
CancelPackError:
    
    Call DealWithErrorResponse(xmlResponseDoc, strFunctionName, Err)
    Resume TidyUp

End Function

Private Function ReSendPack(ByVal xmlRequestNode As IXMLDOMNode) As String
'TW 12/06/2006 MAR1867 Rewritten for performance reasons

Const strFunctionName = "ReSendPack"

Dim strApplicationNumber As String
Dim strDocumentGuid As String
Dim strFileNetImageRef As String
Dim strHostTemplateId As String
Dim strPackGUID As String
Dim strPackId As String
Dim strResponse As String
Dim strUnitId As String
Dim strUserId As String

Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlEventNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode

' ReSends an existing PACK
' Returns RESPONSE xml string
' strPackId must be present

    On Error GoTo ReSendPackError:

' Create an xml document for the response
    Set xmlResponseDoc = xmlCreateDOMObject()
    
' Create RESPONSE node
    Set xmlResponseNode = AddNode(xmlResponseDoc, "RESPONSE")

    Set xmlEventNode = xmlRequestNode.selectSingleNode("//EVENTDETAIL")
    If xmlEventNode Is Nothing Then
        Err.Raise 40026, "PackControl.ReSendPack", "No EVENTDETAIL node supplied"
    End If
    
    strApplicationNumber = GetAttribute(xmlRequestNode, "APPLICATIONNUMBER")
    strHostTemplateId = GetAttribute(xmlEventNode, "HOSTTEMPLATEID")
    strPackId = GetAttribute(xmlEventNode, "PACKFULFILLMENTGUID")
    strUnitId = GetAttribute(xmlRequestNode, "UNITID")
    strUserId = GetAttribute(xmlRequestNode, "USERID")

    strResponse = DoRequest("reSendPack", "", strPackId, strUserId, xmlEventNode)
    If strResponse <> "Done" Then
        Err.Raise 40023, "PackControl.ReSendPack", "Failed to ReSend Pack to Fulfilment: " & strResponse
    End If
    
' Create the EventAuditDetail record
    strPackGUID = strZeroes
    strFileNetImageRef = strZeroes
    strDocumentGuid = GetAttribute(xmlRequestNode, "DOCUMENTGUID")
    
    Call CreateEventAuditDetail(strApplicationNumber, EVENTKEY_FULFILLMENTRESEND, strHostTemplateId, strDocumentGuid, strPackGUID, strFileNetImageRef, strUnitId, strUserId)
    
    Call AddAttribute(xmlResponseNode, "TYPE", "SUCCESS")
    
    
    gobjTrace.TraceXML xmlResponseDoc.xml, "ReSendPack " & strPackId
   
TidyUp:
    On Error Resume Next
    ReSendPack = xmlResponseDoc.xml
    On Error GoTo 0
    
    Set xmlResponseDoc = Nothing
    
    Set xmlEventNode = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function
    
ReSendPackError:
    Call DealWithErrorResponse(xmlResponseDoc, strFunctionName, Err)
    Resume TidyUp
    
'TW 12/06/2006 MAR1867 End
End Function


Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    gobjTrace.StartTrace
'TW 12/06/2006 MAR1867
    On Error Resume Next
    
    strZeroes = String$(32, "0")
    
    Set objOmFDM = CreateObject("omFDM.FileNetInterfaceBO")
    If objOmFDM Is Nothing Then
        Err.Raise 40031, "PackControl", "Can't create omFDM object"
    End If
    On Error GoTo 0
'TW 12/06/2006 MAR1867  End
End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
'TW 12/06/2006 MAR1867
    Set objOmFDM = Nothing
'TW 12/06/2006 MAR1867
End Sub

' PE MAR1756 - 10/05/2006
Private Function GetAddress(ByVal strPackDestination As String, ByVal strAppNo As String, ByVal strAppFFNo As String) As String

Const PackDestinationSolicitor As String = "S"
Const PackDestinationSolicitorandCustomer As String = "B"
Const PackDestinationNewProperty As String = "N"
Const PackDestinationCustomer As String = "C"
'AW EP1304
Const PackDestinationBroker As String = "BR"
Const PackDestinationPackager As String = "P"
'AW EP1304 - End
Dim strXML As String
    strXML = vbNullString
    
    'MAR1819 GHun Change calls to ExecStoredProc to pass parameters
    If (strPackDestination = PackDestinationCustomer) Or (strPackDestination = PackDestinationSolicitorandCustomer) Then
        strXML = ExecStoredProc("usp_PACKCONTROL", strAppNo, strAppFFNo)
    End If
        
    If (strPackDestination = PackDestinationNewProperty) Then
        strXML = ExecStoredProc("USP_PACKCONTROLNEWADDRESS", strAppNo, strAppFFNo)

    End If
    
    'AW EP1304
    If (strPackDestination = PackDestinationBroker) Then
        strXML = ExecStoredProc("USP_PACKCONTROLBROKER", strAppNo, strAppFFNo)
    End If
    
    If (strPackDestination = PackDestinationPackager) Then
        strXML = ExecStoredProc("USP_PACKCONTROLPACKAGER", strAppNo, strAppFFNo)
    End If
    'AW EP1304 - End
    
    If (strPackDestination = PackDestinationSolicitor) Or (strPackDestination = PackDestinationSolicitorandCustomer) Then
        strXML = strXML & ExecStoredProc("usp_PACKCONTROLSOLICITOR", strAppNo, strAppFFNo)

    End If
    'MAR1819 End
        
    GetAddress = strXML
    
End Function

'MAR1819 GHun Changed to execute stored procs as stored procs, rather than text
' PE MAR1756 - 10/05/2006
Private Function ExecStoredProc(ByVal vstrProcName As String, ByVal vstrAppNumber As String, ByVal vstrAppFactFindNumber As String, Optional ByVal vstrXMLRoot As String = vbNullString) As String

    Const cstrFunctionName As String = "ExecStoredProc"
    Dim adoStream   As ADODB.Stream
    Dim cmd         As ADODB.Command
    Dim connOmiga   As Connection

On Error GoTo ExecStoredProcVbErr

    Set connOmiga = GetConnection()
    
    Set adoStream = New ADODB.Stream
    adoStream.Open
    
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = connOmiga
        .CommandType = adCmdStoredProc
        .CommandText = vstrProcName
        .Parameters.Append .CreateParameter("@ApplicationNumber", adBSTR, adParamInput, Len(vstrAppNumber), vstrAppNumber)
        .Parameters.Append .CreateParameter("@ApplicationFactFindNumber", adInteger, adParamInput, , CInt(vstrAppFactFindNumber))
    
        .Properties("Output Stream") = adoStream
        If Len(vstrXMLRoot) > 0 Then
            .Properties("XML Root") = vstrXMLRoot
        End If
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
    ExecStoredProc = adoStream.ReadText()

    adoStream.Close
    connOmiga.Close

ExecStoredProcVbErr:
    
    Set connOmiga = Nothing
    Set adoStream = Nothing
    Set cmd = Nothing
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1819 End

#If Gemini Then
'EP1271 - CC56
Private Function GetExistingPackDocuments(ByVal vstrAppNumber As String, ByVal vstrPackNumber As String) As String

    Const cstrFunctionName As String = "GetExistingPackDocuments"
    Dim adoStream   As ADODB.Stream
    Dim cmd         As ADODB.Command
    Dim connOmiga   As Connection

On Error GoTo GetExistingPackDocumentsVbErr

    Set connOmiga = GetConnection()
    
    Set adoStream = New ADODB.Stream
    adoStream.Open
    
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = connOmiga
        .CommandType = adCmdStoredProc
        .CommandText = "usp_GetExistingPackDocuments"
        .Parameters.Append .CreateParameter("@applicationNumber", adBSTR, adParamInput, Len(vstrAppNumber), vstrAppNumber)
        .Parameters.Append .CreateParameter("@PackControlNumber", adInteger, adParamInput, , CInt(vstrPackNumber))
    
        .Properties("Output Stream") = adoStream
        
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
    
    GetExistingPackDocuments = adoStream.ReadText()

    adoStream.Close
    connOmiga.Close

GetExistingPackDocumentsVbErr:
    
    Set connOmiga = Nothing
    Set adoStream = Nothing
    Set cmd = Nothing
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Sub CreateApplicationPack(ByVal vstrPackDetails As String)

    Const cstrFunctionName As String = "CreateApplicationPack"
    
    Dim cmd         As ADODB.Command
    Dim connOmiga   As Connection
    
    Dim xmlApplicationPackDoc As FreeThreadedDOMDocument40
    Dim xmlApplicationPackNode As IXMLDOMNode
    Dim xmlPackMemberNode As IXMLDOMNode
    Dim xmlPackMemberList As IXMLDOMNodeList
    
    Dim intPackControl As Integer, intStatus As Integer
    Dim strPackGUID As String, strAppNum As String
    Dim strUserId As String
    Dim strUnitId As String
    Dim strDocGUID As String
    Dim intPackMember As Integer
    
On Error GoTo CreateApplicationPackVbErr

    Set xmlApplicationPackDoc = xmlCreateDOMObject()
    xmlApplicationPackDoc.loadXML vstrPackDetails
    Set xmlApplicationPackNode = xmlApplicationPackDoc.selectSingleNode("//APPLICATIONPACK")
    
    If Not xmlApplicationPackNode Is Nothing Then
        strPackGUID = GetAttribute(xmlApplicationPackNode, "PACKFULFILLMENTGUID")
        strAppNum = GetAttribute(xmlApplicationPackNode, "APPLICATIONNUMBER")
        strUserId = GetAttribute(xmlApplicationPackNode, "USERID")
        strUnitId = GetAttribute(xmlApplicationPackNode, "UNITID")
        intPackControl = CInt(GetAttribute(xmlApplicationPackNode, "PACKCONTROLNUMBER"))
        intStatus = 10
    End If
    
    Set xmlPackMemberList = xmlApplicationPackDoc.selectNodes("//APPLICATIONPACKMEMBER")
    
    Set connOmiga = GetConnection()
    
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = connOmiga
        .CommandType = adCmdStoredProc
        .CommandText = "usp_CreateApplicationPack"
        .Parameters.Append .CreateParameter("@pPackGuid", adVarBinary, adParamInput, 16, GuidStringToByteArray(strPackGUID))
        .Parameters.Append .CreateParameter("@pPackControlNum", adInteger, adParamInput, , intPackControl)
        .Parameters.Append .CreateParameter("@pApplicationNum", adBSTR, adParamInput, Len(strAppNum), strAppNum)
        .Parameters.Append .CreateParameter("@pStatus", adInteger, adParamInput, , intStatus)
        .Execute , , adExecuteNoRecords
        Set .ActiveConnection = Nothing
    End With

    Dim pack As GeminiPack
    pack.UserId = strUserId
    pack.UnitId = strUnitId
    pack.ApplicationNumber = strAppNum
    pack.PackFulfillmentGuid = strPackGUID

    Dim documentIndex As Integer
    documentIndex = 0
    For Each xmlPackMemberNode In xmlPackMemberList
    
        strDocGUID = GetAttribute(xmlPackMemberNode, "DOCUMENTGUID")
        intPackMember = CInt(GetAttribute(xmlPackMemberNode, "PACKMEMBERNUMBER"))
        
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = connOmiga
            .CommandType = adCmdStoredProc
            .CommandText = "usp_CreateAppPackMember"
            .Parameters.Append .CreateParameter("@pPackGuid", adVarBinary, adParamInput, 16, GuidStringToByteArray(strPackGUID))
            .Parameters.Append .CreateParameter("@pPackMemberNum", adInteger, adParamInput, , intPackMember)
            .Parameters.Append .CreateParameter("@pApplicationNum", adBSTR, adParamInput, Len(strAppNum), strAppNum)
            .Parameters.Append .CreateParameter("@pDocumentGuid", adVarBinary, adParamInput, 16, GuidStringToByteArray(strDocGUID))
            .Execute , , adExecuteNoRecords
            Set .ActiveConnection = Nothing
        End With
        
        ReDim Preserve pack.Documents(0 To documentIndex)
        pack.Documents(documentIndex).DocumentGuid = strDocGUID
        
        Dim strFileContents As String
        strFileContents = GetAttribute(xmlPackMemberNode, "FILECONTENTS")
        If Len(strFileContents) > 0 Then
            pack.Documents(documentIndex).DocumentDetails.CompressionMethod = GetAttribute(xmlPackMemberNode, "COMPRESSIONMETHOD")
            pack.Documents(documentIndex).DocumentDetails.DeliveryType = CSafeInt(GetAttribute(xmlPackMemberNode, "DELIVERYTYPE"))
            pack.Documents(documentIndex).DocumentDetails.FileContentsType = GetAttribute(xmlPackMemberNode, "FILECONTENTS_TYPE")
            pack.Documents(documentIndex).DocumentDetails.FileContents = strFileContents
        End If
        
        documentIndex = documentIndex + 1
    Next
    
    connOmiga.Close

    GeminiSendToFulfillment gobjContext, pack, True, False, True, GetGlobalParamBoolean("GeminiErrorIfGeminiPrinted")

CreateApplicationPackVbErr:
    
    Set xmlApplicationPackDoc = Nothing
    Set xmlPackMemberNode = Nothing
    Set connOmiga = Nothing
    Set cmd = Nothing
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'EP1271 - End
#End If
'AW EP2_2181
#If Gemini Then
Private Function IsItemInCollection(colCollection As Collection, strItem As String) As Boolean

    Dim nIndex As Integer
    
    IsItemInCollection = False
    
    For nIndex = 1 To colCollection.Count
       If colCollection(nIndex) = strItem Then
            IsItemInCollection = True
            Exit For
       End If
    Next nIndex

End Function
#End If
'AW EP2_2181   -   End
