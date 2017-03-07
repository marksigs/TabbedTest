VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PrintManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PrintManagerBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Created:       26/06/2001
'Author:        Mike Kennedy
'Description:   BO for the Print Manager component
'Dependencies:  omDPS - Document Production System
'               omFVS - File Versioning System
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MK     26/06/01    Created
'DR     29/01/02    DMSSYS0005 - Throw an error if omRequest() XML request is not valid
'DR     11/02/02    DMSSYS0005 - Dont throw an error (see above) assume request is from
'                   old style print and perform a simple, straight-through print.
'                   SaveDocument() Default value for freeformat if no value supplied
'                   SaveDocument() 24 typo fixed.
'                   GetPrintLocation() Missing Exit Subs inserted.
'DR     14/02/02    DMSSYS0005 Commented out strange error handling in
'                   IMessageQueueComponentVC1_OnMessage() and reintated old error handling
'                   code.
' DR    12/03/02    SYS4278 Adding errors to the event log/updating error handling.
' CL    13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
' HMA   22/04/04    CORE46  Add Delivery Type and Delivery Engine Type to SavePrintDocument
' HMA   15/06/04    BBG751  Add Delivery Type and Delivery Engine Type to ProcessRequest
' AS    06/07/04    BBG919  Fixed very slow logging to Event log (don't log document - can take secs)
' AS    26/07/04    BBG1055 Pass DELIVERYTYPE to FVS.
' AS    29/07/04    BBG1149 As a temporary measure (until printing of PDFs is added) all
'                   documents should be stored in DMS as Word, not PDFs.
'                   See also BBG1043 in omDPS:PrintHandlerBO:ProcessRequest().
' TK    08/09/04    BBG879 - Amend GetDocumentArchive()
'TW     17/09/04    General Tidy up
' PSC   21/10/2004  BBG1639 Amend SavePrintDocument to have local DPS object
' TK    17/11/2004  BBG1672 Released all used objects and used errCheckXMLResponseNode() instead of errCheckXMLResponse()
' PSC   18/11/2004  BBG1806 Improve error detection
' PSC   20/12/2004  E2EM00003127 Add correct exit processing
' AS    21/04/2005  CORE106 Pass compression attributes into omFVS.
' AS    27/04/2005  CORE108 Fix binary incompatibilities with BBG MX.
' MV    29/04/2005  CORE110 - Amended SavePrintDocument()
' MV    05/05/2005  CORE116 - Undo CORE110.
' AS    17/05/2005  CORE129 Pass delivery type into omFVS.
' AS    19/05/2005  CORE131 Fixed bug introduced by CORE129.
' AS    18/08/2005  CORE187 Pass METADATA attribute into omFVS.
' PSC   10/11/2005  CORE196 Add UpdateDocumentAuditDetails
' TW    29/11/2005  CORE219 Add methods to handle different Document Management System types
' TW    01/12/2005  CORE222 Modified SavePrintDocument to handle sending documents printed on the workstation to FileNet
' BC    01/02/2006  CORE236 Add Quote Number and Date/Time to FILENAME to ensure unique name
' PSC   06/02/2006  CORE240 Amend so that Filenet content type is based on delivey type
'                           Correct CORE236 to look for TEMPLATEDATA and make it optional
' AS    04/04/2006  CORE261 Add support for Gemini.
' IK    26/04/2006  CORE261 Add support for Gemini.
' GHun  28/04/2006  CORE269 Configurably disable mutex in printing
' AS    28/04/2006  CORE261 Add support for Gemini. Made CORE236 FILENET specific.
' AS    05/06/2006  CORE277 omPM SavePrintDocument needs to return FILEGUID
' AS    26/07/2006  CORE286 Get file from archive as a Url and not as bin base 64 string.
' AS    13/11/2006  CORE309 omPM: do not send an archive message to the GeminiInterface queue for scanned documents.
' AS    05/12/2006  CORE312 Gemini Printing
' AS    15/12/2006  CORE324 Add support for DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS
' AS    20/12/2006  CORE325 omPM: Gemini printing pack handling
' AS    10/01/2007  CORE329 omPM: QueueGeminiRequest selecting from wrong omFVS response element
' AS    12/01/2007  CORE331 Removed GetDocumentHistoryListXml - replaced with omCRUD call.
'------------------------------------------------------------------------------------------
Private Const PRINT_MUTEX = "PRINT_MUTEX"
Option Explicit
Implements IMessageQueueComponentVC1    ' for message queue listening
Implements ObjectControl
'=============================================
'Type Definition Declaration Section
'=============================================
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
'=============================================
'API/DLL Procedure Declaration Section
'=============================================
'CORE269 GHun Not used
'Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
'=============================================
'Variable Declaration Section
'=============================================
Private gobjContext As ObjectContext
'TK 15/11/2004 Private m_objContext As ObjectContext
'TK 15/11/2004 Private gobjTxBO As IPrintManagerTxBO
'TK 15/11/2004 Private gobjDO As IPrintManagerDO
'Private gobjDPS As PrintHandlerBO
Private Const cstrClassName = "omPM::PrintManagerBO"

' TW 29/11/2005 CORE219
Dim DocumentManagementSystemType As Integer
Dim strHostTemplateId As String

Private m_blnUseMutex As Boolean    'CORE269 GHun

Private Function SendDocumentToFileNet(ByVal vxmlRequestNode As IXMLDOMNode) As String
' TW 29/11/2005 CORE219
Const cstrFunctionName = "SendDocumentToFileNet"

Dim strResponse As String

Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40

Dim xmlDocumentContents As IXMLDOMNode
Dim xmlPrintDocumentDataNode As IXMLDOMNode
Dim xmlRequest As IXMLDOMNode

Dim objOmPackControl As Object

    On Error GoTo SendDocumentToFileNetError:
    
    ' PSC 06/02/2006 CORE240
    Dim intDeliveryType As Integer

    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlRequest = xmlRequestDoc.createNode(NODE_ELEMENT, "REQUEST", "")
    
    xmlRequestDoc.appendChild xmlRequest

    Set xmlPrintDocumentDataNode = xmlGetNode(vxmlRequestNode, "PRINTDOCUMENTDATA")

    Call AddAttribute(xmlRequest, "OPERATION", "SENDDOCUMENTTOFILENET")
    Call AddAttribute(xmlRequest, "APPLICATIONNUMBER", xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER"))
    Call AddAttribute(xmlRequest, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID"))
    Call AddAttribute(xmlRequest, "USERID", "CEAdmin")
    Call AddAttribute(xmlRequest, "CHANNELID", xmlGetAttributeText(vxmlRequestNode, "CHANNELID"))
    Call AddAttribute(xmlRequest, "HOSTTEMPLATEID", strHostTemplateId)
    
    Call AddAttribute(xmlRequest, "HOSTTEMPLATEDESCRIPTION", xmlGetAttributeText(xmlPrintDocumentDataNode, "FILENAME"))
    Call AddAttribute(xmlRequest, "HOSTTEMPLATENAME", xmlGetAttributeText(xmlPrintDocumentDataNode, "FILENAME"))
    
    Set xmlDocumentContents = xmlRequestDoc.createNode(NODE_ELEMENT, "DOCUMENTCONTENTS", "")
    xmlRequest.appendChild xmlDocumentContents

    Call AddAttribute(xmlDocumentContents, "FILEVERSION", "V1")
    Call AddAttribute(xmlDocumentContents, "FILESIZE", "")
    Call AddAttribute(xmlDocumentContents, "FILECONTENTS_TYPE", "BIN.BASE64")
    Call AddAttribute(xmlDocumentContents, "FILECONTENTS", xmlGetAttributeText(xmlPrintDocumentDataNode, "FILECONTENTS"))
    
    ' PSC 06/02/2006 CORE240 - Start
    intDeliveryType = xmlGetAttributeAsInteger(xmlPrintDocumentDataNode, "DELIVERYTYPE")
    Call AddAttribute(xmlDocumentContents, "CONTENTTYPE", GetValidationTypeForValueID("DocumentDeliveryType", intDeliveryType))
    ' PSC 06/02/2006 CORE240 - End
    
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
    End If
    
    ' Create an instance to the file versioning system
    If objOmPackControl Is Nothing Then
        Set objOmPackControl = gobjContext.CreateInstance("omPack.PackManagerBO")
        
        gobjTrace.TraceInitialiseOffspring objOmPackControl
    End If
    
    gobjTrace.TraceXML xmlRequestDoc.xml, "SendDocumentToFileNet Request"

    strResponse = objOmPackControl.omRequest(xmlRequestDoc.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    If Not xmlResponseDoc Is Nothing Then
        SendDocumentToFileNet = xmlResponseDoc.documentElement.xml
    End If
    
SendDocumentToFileNetError:

    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    Set xmlDocumentContents = Nothing
    Set xmlPrintDocumentDataNode = Nothing
    Set xmlRequest = Nothing
    
    Set objOmPackControl = Nothing
    
' TW 29/11/2005 CORE219 End
End Function

Private Sub AddAttribute(ByVal xmlNode As IXMLDOMNode, ByVal strAttributeName As String, ByVal strAttributeValue As String)
Dim xmlAttribute As IXMLDOMAttribute

    Set xmlAttribute = xmlNode.ownerDocument.createAttribute(strAttributeName)
    xmlAttribute.Text = strAttributeValue
    xmlNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = Nothing
End Sub


' TW 29/11/2005 CORE219 End
'------------------------------------------------------------------------------------------
' BMIDS specific History:
'
' Prog  Date        Description
' IK    28/11/2002  BMIDS00606 - store / audit in DMS before document production
' IK    12/12/2002  BM0037 - added new method UnlockDocument
' IK    04/02/2003  BM0269 - improved error handling
' IK    17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     17/11/2003  CORE8 - Amended
'AS     05/05/2005  CORE114 - Do not audit on getting archived document for workstation printing.
'------------------------------------------------------------------------------------------
'BG BMIDS00592 31/10/02
'Private gobjDPS As Object
'BG BMIDS00592 31/10/02 end
Private Sub ObjectControl_Activate()
    
    Set gobjContext = GetObjectContext()
    
    gobjTrace.StartTrace
    
    'CORE269 GHun Check if a mutex should be used
    On Error Resume Next
    m_blnUseMutex = GetMandatoryGlobalParamBoolean("UseMutexForPrinting")   'CORE269 GHun
    If Err.Number <> 0 Then
         Err.Clear
         m_blnUseMutex = True
    End If
    'CORE269 End
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    
    Set gobjContext = Nothing
End Sub
Public Function IMessageQueueComponentVC1_OnMessage( _
    ByVal vstrXMLRequest As String) As MESSQ_RESP
    On Error Resume Next
    Const cstrFunctionName As String = "IMessageQueueComponentVC1_OnMessage"
'    App.LogEvent "IMessageQueueComponentVC1_OnMessage() Called", vbLogEventTypeInformation
    
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
        gobjTrace.StartTrace
    End If
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Dim MQResp As MESSQ_RESP
    MQResp = MESSQ_RESP_STALL_COMPONENT
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
         
    'DR DMSSYS0005 - Error Details
    Dim lngErrorNumber As Long
    Dim strErrorSource As String
    Dim strErrorDescription As String
     
    ' will be superceded by any subsequent SetAbort
    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If
    
    Set xmlOutDoc = xmlCreateDOMObject
    
    Debug.Print "vstrXMLRequest = " & vstrXMLRequest
    
    Set xmlInDoc = xmlLoad(vstrXMLRequest, cstrFunctionName)
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    Set xmlResponseNode = xmlOutDoc.createElement("RESPONSE")   ' PSC 18/11/2004 BBG1806
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        ' mutex so that only a single instance of word is used
        If m_blnUseMutex Then   'CORE269 GHun
            Dim Mutex As OMMUTEXLib.omMutex1
            Set Mutex = New OMMUTEXLib.omMutex1
            Mutex.AcquireMutex PRINT_MUTEX, False ' low priorty mutex when called via queue
        End If  'CORE269 GHun
        'DR Should go through DoRequest not ProcessRequest
        DoRequest xmlRequestNode, xmlResponseNode
        ' OK, should check the response here and set MQResp to the appropriate value.
        lngErrorNumber = Err.Number
        strErrorSource = Err.Source
        strErrorDescription = Err.Description
        ' release mutex
        If m_blnUseMutex Then 'CORE269 GHun
            Mutex.ReleaseMutex
            Set Mutex = Nothing
        End If 'CORE269 GHun
        If lngErrorNumber = 0 Then
            'OK
'            App.LogEvent "IMessageQueueComponentVC1_OnMessage() MESSQ_RESP_SUCCESS no error", vbLogEventTypeInformation
            MQResp = MESSQ_RESP_SUCCESS
        Else
            MQResp = MESSQ_RESP_RETRY_MOVE_MESSAGE
            App.LogEvent "IMessageQueueComponentVC1_OnMessage() MESSQ_RESP_RETRY_MOVE_MESSAGE. " & _
                            "Err.Number: " & lngErrorNumber & _
                            ", Err.Source: " & strErrorSource & _
                            ", Err.Description: " & strErrorDescription & _
                            ". xmlResponseNode: " & xmlResponseNode.xml, _
                            vbLogEventTypeError
            'DR Turn off the error handling so that error is bubbled up to the calling code
            'DR The problem with raising an error is that the queue components dont handle it :(
            'It just keeps on retrying, over and over again. Will attempt to control queue
            'via return code.
'                Err.Raise lngErrorNumber, strErrorSource, "IMessageQueueComponentVC1_OnMessage() MESSQ_RESP_RETRY_MOVE_MESSAGE. Err.Number: " & lngErrorNumber & _
'                                ", Err.Source: " & Err.Source & _
'                                ", Err.Description: " & Err.Description & _
'                                ". xmlResponseNode: " & xmlResponseNode.xml, _
'                                vbLogEventTypeError
            
            
        End If
    End If
    IMessageQueueComponentVC1_OnMessage = MQResp
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName, CStr(MQResp)
    
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    
End Function
Public Function omRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "OmRequest"

Dim xmlAttrib As IXMLDOMAttribute
Dim xmlInDoc As FreeThreadedDOMDocument40
Dim xmlOperationNode As IXMLDOMNode
Dim xmlOperationNodeList As IXMLDOMNodeList
Dim xmlOutDoc As FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseElem As IXMLDOMElement
Dim xmlResponseNode As IXMLDOMNode
    
    On Error GoTo OmRequestVbErr
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
        gobjTrace.StartTrace
    End If
    
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    'DR DMSSYS0005 - number of operations included in the request XML. This is
    'used to check whether there has been any difficulty processing the XML request node.
    'This approach is used, rather than the more usual mandatory attrib/element checking
    'becuase the operation attribute may be present on the request element - not very sensible
    'way of doing it because it doesnt easily allow for multiple operations.
    '
    'Instead you get the possiblity of using multiple elements or a single attribute.
    
    ' will be superceded by any subsequent SetAbort
    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If
    
    Set xmlOutDoc = xmlCreateDOMObject
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    ' load XML REQUEST
    Debug.Print "vstrXMLRequest = " & vstrXMLRequest
    
    Set xmlInDoc = xmlLoad(vstrXMLRequest, cstrFunctionName)
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        'DR DMSSYS0005
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        ' multiple operations
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        If xmlOperationNodeList.length = 0 Then
            'Problem - no valid DMS2 operations in XML request
            'It could be that this an OLD style Omiga4 print request - assume it is.
                
            'Needs an OPERATION attribute if it's to be processed - simple, straight
            'through print job.
            xmlSetAttributeValue xmlRequestNode, "OPERATION", "ONMESSAGE"
            DoRequest xmlRequestNode, xmlResponseNode
        Else
            For Each xmlOperationNode In xmlOperationNodeList
                For Each xmlAttrib In xmlRequestNode.Attributes
                    ' copy attributes from REQUEST node to OPERATION node
                    xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
                Next
                'DR How can XMLResponseNode from multiple operations be passed back -
                'it must be a bit of a mish mash?
                DoRequest xmlOperationNode, xmlResponseNode
            Next
        End If
    End If
    ' if we get here, everything is OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    omRequest = xmlOutDoc.xml

OmRequestExit:
    
    ' clean up ================================================================
    ' release global objects
    'TK 15/11/2004 Set gobjTxBO = Nothing
    'TK 15/11/2004 Set gobjDO = Nothing
    ' release XML DOM objects
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    gobjTrace.TraceResponse omRequest
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
OmRequestVbErr:
    
    'Simple Error Logging
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    ' create ERROR response block from Err object
    omRequest = errCreateErrorResponse()
    
    gobjTrace.TraceIdentErrorResponse omRequest
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Debug.Print omRequest
    ' perform clean up
    GoTo OmRequestExit
End Function
Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, private interface, performs single REQUEST OPERATION
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

Dim Mutex As OMMUTEXLib.omMutex1

Dim lngErrorNumber As Long

Dim strErrorDescription As String
Dim strErrorSource As String
Dim strOperation As String
    
' TW 29/11/2005 CORE219
Dim xmlNode As IXMLDOMNode
' TW 29/11/2005 CORE219 End
    
    On Error GoTo vbError
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
'   IK_DEBUG
#If IK_DEBUG Then
    Dim fso As Object
    Dim tname As String
    Dim sPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tname = fso.GetTempName
    tname = Left(tname, Len(tname) - 4)
    sPath = App.Path & "\debug"
    If Not fso.FolderExists(sPath) Then
        fso.CreateFolder sPath
    End If
    vxmlRequestNode.ownerDocument.save sPath & "\omPMRequest_" & tname & ".xml"
    Set fso = Nothing
#End If
    
' TW 29/11/2005 CORE219
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.8]", "dbxmlGetCurrentParameterXML call"

    Set xmlNode = dbxmlGetCurrentParameterXML("DocumentMangementSystemType")
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.9]", "dbxmlGetCurrentParameterXML return"
    If xmlNode Is Nothing Then
        DocumentManagementSystemType = 1 'Default to DMS (uses omFVS)
    Else
        DocumentManagementSystemType = xmlGetNodeText(xmlNode, ".//AMOUNT")
    End If
    Set xmlNode = vxmlRequestNode.selectSingleNode("//DOCUMENTDETAILS")
    If xmlNode Is Nothing Then
        strHostTemplateId = ""
    Else
        On Error Resume Next
        strHostTemplateId = xmlNode.Attributes.getNamedItem("HOSTTEMPLATEID").Text
        On Error GoTo vbError
    End If
' TW 29/11/2005 CORE219 End
    
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    strOperation = UCase$(strOperation)
    ' mutex so that only a single instance of word is used

    If m_blnUseMutex Then 'CORE269 GHun
        Set Mutex = New OMMUTEXLib.omMutex1
        Mutex.AcquireMutex PRINT_MUTEX, True ' high priority mutex when interactively
    End If  'CORE269 GHun
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.0]", "OPERATION: " & strOperation
    Select Case strOperation
        Case "ONMESSAGE"
            OnMessage vxmlRequestNode, vxmlResponseNode
        
        Case "NEWDOCUMENT"
            ProcessRequest vxmlRequestNode, vxmlResponseNode
            
        ' ik_BMIDS00885 REPRINTDOCUMENT now via queue
        Case "PRINTDOCUMENT", "REPRINTDOCUMENT"
            PrintDocument vxmlRequestNode, vxmlResponseNode
        ' ik_BMIDS00885_ends
        ' ik_bm0037
        ' for edit with no save
        Case "UNLOCKDOCUMENT"
            UnlockDocument vxmlRequestNode, vxmlResponseNode
        ' ik_bm0037_ends
            
        Case "SAVEDOCUMENT"
            SaveDocument vxmlRequestNode, vxmlResponseNode
            ' Need to check if it's free format and then print it...
            If xmlGetAttributeAsLong(vxmlRequestNode, "FREEFORMAT") = 1 Then
                PrintDocument vxmlRequestNode, vxmlResponseNode
            End If
        
        Case "CREATEDOCUMENTARCHIVE"
            CreateDocumentArchive vxmlRequestNode, vxmlResponseNode, False
        
        Case "GETDOCUMENTARCHIVE", "EDITDOCUMENT"
            GetDocumentArchive vxmlRequestNode, vxmlResponseNode
        
        Case "CREATEAUDITTRAIL"
            CreateAuditTrail vxmlRequestNode, vxmlResponseNode
        
        Case "GETDOCUMENTHISTORYLIST", "GETDOCUMENTEVENTLIST"
            CallDO strOperation, vxmlRequestNode, vxmlResponseNode
        
        Case "PROCESSMERGEDOCUMENT"
            ProcessMergeDocument vxmlRequestNode, vxmlResponseNode
        
        Case "SAVEPRINTDOCUMENT"
            SavePrintDocument vxmlRequestNode, vxmlResponseNode
        ' PSC 10/11/2005 CORE196 - Start
        Case "UPDATEDOCUMENTAUDITDETAILS"
            CallDO strOperation, vxmlRequestNode, vxmlResponseNode
        ' PSC 10/11/2005 CORE196 - End
        
        Case "GEMINISENDTOFULFILLMENT"
            GeminiSendToFulfillmentXml gobjContext, vxmlRequestNode

        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, strOperation
    End Select
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    ' release mutex
    If m_blnUseMutex Then 'CORE269 GHun
        Mutex.ReleaseMutex
    End If  'CORE269 GHun
    Set Mutex = Nothing
' TW 29/11/2005 CORE219
    Set xmlNode = Nothing
' TW 29/11/2005 CORE219 End
    Exit Sub

vbError:
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & lngErrorNumber & ", Err.Source: " & strErrorSource & ", Err.Description: " & strErrorDescription, vbLogEventTypeError
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    ' release mutex
    If m_blnUseMutex Then 'CORE269 GHun
        Mutex.ReleaseMutex
    End If
    Set Mutex = Nothing
    'Raise the error
    Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
End Sub
Private Function ProcessRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As String
    
' header ----------------------------------------------------------------------------------
' description:
'   Deals with any new document requests coming in from the GUI
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

Const cstrFunctionName As String = "ProcessRequest"

Dim lngErrorNumber As Long
Dim lngFreeFormat As Long
Dim lngNumCopies As Integer

Dim strAppNum As String
Dim strDeliveryType As String
Dim strDocName As String
Dim strEditDoc As String
Dim strErrorDescription As String
Dim strErrorSource As String
Dim strNumberOfCopies As String
Dim strPrintLocation As String
Dim strPrintType As String
Dim strResponse As String
Dim strTest As String

Dim xmlArchiveNode As IXMLDOMNode
Dim xmlCheckPrintResponseNode As IXMLDOMNode
Dim xmlControlDataNode As IXMLDOMNode
Dim xmlCreateAuditTrailRequest As IXMLDOMNode
Dim xmlCreateDocumentArchiveRequest As IXMLDOMNode
Dim xmlCreateDocumentArchiveResponse As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlDocumentDetailNode As IXMLDOMNode
Dim xmlElement As IXMLDOMElement
Dim xmlElement2 As IXMLDOMElement
Dim xmlElement3 As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim xmlPrintDataNode As IXMLDOMNode
Dim xmlPrintDocumentRequestDoc As FreeThreadedDOMDocument40
Dim xmlPrintDocumentRequestNode As IXMLDOMNode
Dim xmlPrintNode As IXMLDOMNode
Dim xmlPrintResponseNode As IXMLDOMNode
Dim xmlRequestElement As IXMLDOMElement
Dim xmlTempNode As IXMLDOMNode

Dim xmlTemplateDataNode As IXMLDOMNode 'BC CORE236 01/02/2006
Dim strCreationDate As String 'BC CORE236 01/02/2006
Dim strQuoteNumber As String     'BC CORE236 01/02/2006
    
    On Error GoTo vbErrorHandler
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    strCreationDate = Format$(Now, "yyyymmdd_hhmmss") 'BC CORE236 01/02/2006
    Set xmlDoc = New FreeThreadedDOMDocument40
    ' Create the request hierarchy for DPS
    ' To start - Copy over the attributes from the vxmlRequestNode that has been passed in
    Set xmlRequestElement = vxmlRequestNode.cloneNode(False)    ' Clones the request node and attributes but nothing else...
        
    ' Create a response for the printing (DPS) call
    Set xmlCheckPrintResponseNode = vxmlResponseNode.cloneNode(False)
        
    ' Are we free formatting??
    lngFreeFormat = xmlGetAttributeAsLong(vxmlRequestNode, "FREEFORMAT")
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "PROCESSREQUEST"
    ' Only call DPS if we want to do anything other than just save....
    ' AS 06/07/04 BBG919 Fixed very slow logging to Event log (don't log document)
    
'    App.LogEvent "ProcessRequest() About to call CheckPrint()." ' vxmlRequestNode.xml: " & vxmlRequestNode.xml
    CheckPrint vxmlRequestNode, xmlCheckPrintResponseNode
'    App.LogEvent "ProcessRequest() Returned from CheckPrint()." ' xmlCheckPrintResponseNode.xml: " & xmlCheckPrintResponseNode.xml
    
    ' Should really check the response from CheckPrint....
    'CheckResponse for errors
'    Set xmlTempNode = xmlCheckPrintResponseNode
'    errCheckXMLResponse xmlTempNode.xml, True, vxmlResponseNode
    errCheckXMLResponseNode xmlCheckPrintResponseNode, vxmlResponseNode, True
    
    If xmlCheckPrintResponseNode.hasChildNodes = False Then
        If Err = 0 Then
            errThrowError cstrFunctionName, oeRecordNotFound
        Else
'            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName & " Template not found : " & strTemplateID, Err
            errThrowError cstrFunctionName & " " & Err.Source, Err.Number, Err.Description
        End If
    End If
    ' Check to see if we are a free format document i.e. one that will be edited
    ' before being printed
    ' If we are not in free format then we archive the data and create an audit
    'trail entry
    If lngFreeFormat = 0 Then
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.0]", "not FreeFormat"
        ' No error received ...
        ' Cloning a node will store the 1st element with their attributes to stop us
        ' having to set all the user ID and the rest ....
        Set xmlCreateDocumentArchiveRequest = vxmlRequestNode.cloneNode(False)
               
        Set xmlCreateDocumentArchiveResponse = vxmlResponseNode.cloneNode(False)
            
        '---------------------------------------------------
        ' Store the document info in the archive (FVS)
        '---------------------------------------------------
        Set xmlElement = xmlDoc.createElement("PRINTDOCUMENTDATA")
        'DR Made these mandatory - becuase it'll cause a runtime error later if it's not supplied
        Set xmlTempNode = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
        Set xmlNode = xmlGetMandatoryNode(xmlTempNode, "DOCUMENTDETAILS")
                
        xmlSetAttributeValue xmlElement, "CREATEDBY", xmlGetAttributeText(vxmlRequestNode, "USERID")
        
        ' Filename is a concatenation of document name and application number
        strDocName = xmlGetAttributeText(xmlNode, "DOCUMENTNAME")
        strAppNum = xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
        
        'BC CORE236 24/01/2006 Begin
        'PSC 06/02/2006 CORE240 - Start
        Set xmlTemplateDataNode = xmlGetNode(vxmlRequestNode, "TEMPLATEDATA")
        If Not xmlTemplateDataNode Is Nothing Then
            strQuoteNumber = xmlGetAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
        End If
        'PSC 06/02/2006 CORE240 - End
        
        ' AS 05/05/2006 CORE261 Only add quote number and creation date to FILENAME if using FILENET.
        If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_FILENET Then
            xmlSetAttributeValue xmlElement, "FILENAME", strDocName + "_" + strAppNum + "_" + strQuoteNumber + "_" + strCreationDate
        Else
            xmlSetAttributeValue xmlElement, "FILENAME", strDocName & "_" & strAppNum & ".DOC"
        End If
        'BC CORE236 24/01/2006 End
    
'        ' AS 29/07/04 BBG1149 As a temporary measure (until printing of PDFs is added)
'        ' all documents should be stored in DMS as Word, not PDFs.
'        ' See also BBG1043 in omDPS:PrintHandlerBO:ProcessRequest().
        strDeliveryType = xmlGetAttributeText(xmlNode, "DELIVERYTYPE")
'        strDeliveryType = "10"
'TW 3/8/2004 - Reinstated
        xmlSetAttributeValue xmlElement, "DELIVERYTYPE", strDeliveryType
        
        '---------------------------------------------------
        ' Get the info from the print response
        '---------------------------------------------------
        If xmlCheckPrintResponseNode Is Nothing Then
            Set xmlPrintResponseNode = vxmlRequestNode
        Else
            Set xmlPrintResponseNode = xmlCheckPrintResponseNode
        End If
        'DR Made this mandatory - becuase it'll cause a runtime error later if its not supplied
        Set xmlPrintNode = xmlGetMandatoryNode(xmlPrintResponseNode, "PRINTDOCUMENTDETAILS")
        
        ' AS 21/04/2005 CORE106 Pass compression attributes into omFVS.
        xmlSetAttributeValue xmlElement, "COMPRESSED", xmlGetAttributeText(xmlPrintNode, "COMPRESSED", "0")
        xmlSetAttributeValue xmlElement, "COMPRESSIONMETHOD", xmlGetAttributeText(xmlPrintNode, "COMPRESSIONMETHOD")
                
        ' AS 18/08/2005 CORE187 Pass METADATA attribute into omFVS.
        Dim strMetaData As String
        strMetaData = xmlGetAttributeText(xmlPrintNode, "METADATA")
        If strMetaData <> "" Then
            xmlSetAttributeValue xmlElement, "METADATA", strMetaData
        End If
                
        '-----------------------------------------------------------
        ' Version and GUID come back from the archive operation
        '-----------------------------------------------------------
        xmlSetAttributeValue xmlElement, "FILEGUID", xmlGetAttributeText(xmlPrintNode, "FILEGUID")
        xmlSetAttributeValue xmlElement, "FILEVERSION", xmlGetAttributeText(xmlPrintNode, "FILEVERSION")
        
        ' Filesize comes from the DPS response
        If xmlAttributeValueExists(xmlPrintNode, "PRINTDOCUMENTSIZE") Then
            xmlSetAttributeValue xmlElement, "FILESIZE", xmlGetAttributeText(xmlPrintResponseNode, "PRINTDOCUMENTSIZE")
        Else
            xmlSetAttributeValue xmlElement, "FILESIZE", "0"
        End If
        ' File contents type is going to be set at the moment
        xmlSetAttributeValue xmlElement, "FILECONTENTS_TYPE", "BIN.BASE64"
        ' File contents came back from DPS in the vxmlResponseNode
        strTest = xmlGetAttributeText(xmlPrintNode, "PRINTDOCUMENT")
        xmlSetAttributeValue xmlElement, "FILECONTENTS", strTest
        
        xmlCreateDocumentArchiveRequest.appendChild xmlElement
        
        '----------------------------------------------------------
        ' Now call the create document archive..... '
        '----------------------------------------------------------
'        App.LogEvent "ProcessRequest() About to call CreateDocumentArchive()"
        ' IK_20/04/2006_CORE261 (response not used until now)
        'CreateDocumentArchive xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse, False
        Dim xmlFVS As DOMDocument40
        Set xmlFVS = New DOMDocument40
        xmlFVS.async = False
        xmlFVS.loadXML CreateDocumentArchive(xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse, False)
        Set xmlCreateDocumentArchiveRequest = xmlFVS.documentElement
        ' IK_20/04/2006_CORE261_ends
        
        ' ik_debug
'        Dim xmlDebug As DOMDocument40
'        Set xmlDebug = New DOMDocument40
'        xmlDebug.async = False
'        xmlDebug.loadXML xmlCreateDocumentArchiveRequest.xml
'        xmlDebug.Save "c:\omiga4Trace\ProcessRequestCreateDocumentArchiveRequest.xml"
'        xmlDebug.loadXML xmlCreateDocumentArchiveResponse.xml
'        xmlDebug.Save "c:\omiga4Trace\ProcessRequestCreateDocumentArchiveResponse.xml"
        ' ik_debug_ends

'        App.LogEvent "ProcessRequest() returned from CreateDocumentArchive()"
        
' TW 29/11/2005 CORE219
        Select Case DocumentManagementSystemType
            'IK_20/04/2006_CORE261
            Case DOCUMENTMANAGEMENTSYSTEMTYPE_DMS, DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI
' TW 29/11/2005 CORE219 End
                '---------------------------------------------------
                ' Audit the operation
                '---------------------------------------------------
                Set xmlElement = xmlDoc.createElement("APPLICATIONDETAIL")
                Set xmlElement2 = xmlDoc.createElement("DOCUMENTDETAILS")
                Set xmlElement3 = xmlDoc.createElement("EVENTDETAIL")
                Set xmlControlDataNode = xmlGetNode(vxmlRequestNode, "CONTROLDATA")
                Set xmlDocumentDetailNode = xmlGetNode(xmlControlDataNode, "DOCUMENTDETAILS")
                
                Set xmlCreateAuditTrailRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlCreateAuditTrailRequest, "OPERATION", "CreateAuditTrail"
                xmlSetAttributeValue xmlElement, "APPLICATIONNUMBER", xmlGetAttributeText(xmlControlDataNode, "APPLICATIONNUMBER")
                xmlSetAttributeValue xmlElement, "TEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "TEMPLATEID")
                xmlSetAttributeValue xmlElement, "SEARCHKEY1", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY1")
                xmlSetAttributeValue xmlElement, "SEARCHKEY2", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY2")
                xmlSetAttributeValue xmlElement, "SEARCHKEY3", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY3")
                xmlCreateAuditTrailRequest.appendChild xmlElement
                
                Set xmlArchiveNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CREATEFILEDATA")
                xmlSetAttributeValue xmlElement2, "TEMPLATEID", xmlGetAttributeText(xmlDocumentDetailNode, "TEMPLATEID")
                xmlSetAttributeValue xmlElement2, "HOSTTEMPLATEID", xmlGetAttributeText(xmlDocumentDetailNode, "HOSTTEMPLATEID")
                xmlSetAttributeValue xmlElement2, "STAGEID", xmlGetAttributeText(xmlDocumentDetailNode, "STAGEID")
                xmlSetAttributeValue xmlElement2, "DOCUMENTTYPE", xmlGetAttributeText(xmlDocumentDetailNode, "DOCUMENTTYPE")
                xmlSetAttributeValue xmlElement2, "DOCUMENTGROUP", xmlGetAttributeText(xmlDocumentDetailNode, "DOCUMENTGROUP")
                xmlSetAttributeValue xmlElement2, "DOCUMENTNAME", xmlGetAttributeText(xmlDocumentDetailNode, "DOCUMENTNAME")
                xmlSetAttributeValue xmlElement2, "DOCUMENTDESCRIPTION", xmlGetAttributeText(xmlDocumentDetailNode, "DOCUMENTDESCRIPTION")
                xmlSetAttributeValue xmlElement2, "CUSTOMERNO", xmlGetAttributeText(xmlDocumentDetailNode, "CUSTOMERNAME")
                xmlSetAttributeValue xmlElement2, "RECIPIENTNAME", xmlGetAttributeText(xmlDocumentDetailNode, "RECIPIENTNAME")
                ' This comes from the host system (Does it on an edit too??)
                xmlSetAttributeValue xmlElement2, "DOCUMENTCREATETIMESTAMP", xmlGetAttributeText(xmlDocumentDetailNode, "DOCUMENTCREATETIMESTAMP")
                
                ' Date returned from FVS...
                xmlSetAttributeValue xmlElement2, "ARCHIVETIMESTAMP", xmlGetAttributeText(xmlArchiveNode, "ARCHIVETIMESTAMP")
                ' Check if there was a print timestamp
                If xmlAttributeValueExists(xmlPrintNode, "PRINTTIMESTAMP") Then
                    xmlSetAttributeValue xmlElement2, "PRINTTIMESTAMP", xmlGetAttributeText(xmlPrintNode, "PRINTTIMESTAMP")
                End If
                xmlSetAttributeValue xmlElement2, "LANGUAGE", xmlGetAttributeText(xmlDocumentDetailNode, "LANGUAGE")
                xmlSetAttributeValue xmlElement2, "SOURCESYSTEM", xmlGetAttributeText(xmlDocumentDetailNode, "SOURCESYSTEM")
                
                'BBG751
                xmlSetAttributeValue xmlElement2, "DELIVERYTYPE", xmlGetAttributeText(xmlDocumentDetailNode, "DELIVERYTYPE")
                xmlSetAttributeValue xmlElement2, "DELIVERYENGINETYPE", xmlGetAttributeText(xmlDocumentDetailNode, "DELIVERYENGINETYPE")
                
                ' AS 15/12/2006 CORE324 Add support for DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS
                xmlSetAttributeValue xmlElement2, "GEMINIPRINTSTATUS", xmlGetAttributeText(xmlDocumentDetailNode, "GEMINIPRINTSTATUS")
                
                xmlCreateAuditTrailRequest.appendChild xmlElement2
                
                xmlSetAttributeValue xmlElement3, "APPLICATIONNUMBER", xmlGetAttributeText(xmlControlDataNode, "APPLICATIONNUMBER")
                xmlSetAttributeValue xmlElement3, "DOCUMENTVERSION", xmlGetAttributeText(xmlArchiveNode, "FILEVERSION")
                xmlSetAttributeValue xmlElement3, "EVENTTIMESTAMP", Format$(Now, "dd mmm yyyy hh:nn:ss")
                xmlSetAttributeValue xmlElement3, "FILEGUID", xmlGetAttributeText(xmlArchiveNode, "FILEGUID")
                xmlSetAttributeValue xmlElement3, "EVENTKEY", CStr(EVENTKEY_CREATED)
                xmlSetAttributeValue xmlElement3, "UNITNAME", xmlGetAttributeText(vxmlRequestNode, "UNITNAME")
                xmlSetAttributeValue xmlElement3, "USERNAME", xmlGetAttributeText(vxmlRequestNode, "USERNAME")
                
        '        App.LogEvent "ProcessRequest() About to call GetPrintLocation()"
                GetPrintLocationEx xmlControlDataNode, strPrintType, strPrintLocation, strNumberOfCopies
        '        App.LogEvent "ProcessRequest() Returned from GetPrintLocation()"
                xmlSetAttributeValue xmlElement3, "PRINTLOCATION", strPrintLocation
                xmlSetAttributeValue xmlElement3, "PRINTERTYPE", strPrintType
                strNumberOfCopies = "1"
                If (strPrintType = "PRINTER") Then
                    Set xmlNode = xmlGetNode(xmlControlDataNode, "OUTPUTTYPE")
                    Set xmlNode = xmlGetNode(xmlNode, "PRINTER")
                    lngNumCopies = xmlGetNodeAsLong(xmlNode, "COPIES")
                    If lngNumCopies > 0 Then
                        strNumberOfCopies = lngNumCopies
                    End If
                End If
                xmlSetAttributeValue xmlElement3, "NUMCOPIES", strNumberOfCopies
                xmlSetAttributeValue xmlElement3, "EVENTDATE", Format$(Now, "dd mmm yyyy hh:nn:ss")
                xmlSetAttributeValue xmlElement3, "PACKFULFILLMENTGUID", xmlGetAttributeText(xmlControlDataNode, "PACKFULFILLMENTGUID")
                xmlSetAttributeValue xmlElement3, "HOSTTEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DOCUMENTID")
                xmlSetAttributeValue xmlElement3, "FILENETIMAGEREF", xmlGetAttributeText(xmlControlDataNode, "FILENETIMAGEREF")
                xmlCreateAuditTrailRequest.appendChild xmlElement3
        '        App.LogEvent "ProcessRequest() About to call CreateAuditTrail()"
                CreateAuditTrail xmlCreateAuditTrailRequest, vxmlResponseNode
' TW 29/11/2005 CORE219
        End Select
' TW 29/11/2005 CORE219 End
'        App.LogEvent "ProcessRequest() Returned from CreateAuditTrail()"

        'IK_20/04/2006_CORE261
        If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
            QueueGeminiRequest xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse
        End If
        'IK_20/04/2006_CORE261_ends
        
        ' ik_bmids00606
        ' DPS.ProcessRequest no longer prints
        ' need to create PrintDocument request to print after FVS storage etc.
        Set xmlPrintDocumentRequestDoc = New FreeThreadedDOMDocument40
        xmlPrintDocumentRequestDoc.async = False
        Set xmlPrintDocumentRequestNode = _
            xmlPrintDocumentRequestDoc.appendChild(vxmlRequestNode.cloneNode(True))
        xmlPrintDocumentRequestNode.Attributes.getNamedItem("OPERATION").Text = _
            "PRINTDOCUMENT"
        Set xmlPrintDataNode = xmlPrintDocumentRequestNode.selectSingleNode("PRINTDATA")
        For Each xmlNode In xmlPrintDataNode.childNodes
            xmlPrintDataNode.removeChild xmlNode
        Next
        ' could not have got this far without this (?)
        Set xmlNode = xmlCheckPrintResponseNode.selectSingleNode("PRINTDOCUMENTDETAILS")
        xmlCopyAttribute xmlNode, xmlPrintDataNode, "PRINTDOCUMENT"
        ' call DPS again ...
        Set xmlCheckPrintResponseNode = vxmlResponseNode.cloneNode(False)
        ' AS 06/07/04 BBG919 Fixed very slow logging to Event log (don't log document)
'        App.LogEvent "ProcessRequest() About to call CheckPrint()." ' xmlPrintDocumentRequestNode.xml: " & vxmlRequestNode.xml
        CheckPrint xmlPrintDocumentRequestNode, xmlCheckPrintResponseNode
'        App.LogEvent "ProcessRequest() Returned from CheckPrint()." ' xmlCheckPrintResponseNode.xml: " & xmlCheckPrintResponseNode.xml
        If xmlCheckPrintResponseNode.hasChildNodes = False Then
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
        ' ik_bmids00606_ends
    Else
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.0]", "FreeFormat"
        ' The Merging has been done so we can now send the document back to the host...
        ' Set up some return nodes...
        Set xmlControlDataNode = xmlGetNode(vxmlRequestNode, "CONTROLDATA")
        Set xmlDocumentDetailNode = xmlGetNode(xmlControlDataNode, "DOCUMENTDETAILS")
        Set xmlElement = xmlDoc.createElement("DOCUMENTCONTENTS")
        Set xmlPrintResponseNode = xmlCheckPrintResponseNode
        Set xmlPrintNode = xmlGetNode(xmlPrintResponseNode, "PRINTDOCUMENTDETAILS")
        ' Filesize comes from the DPS response
        If xmlAttributeValueExists(xmlPrintNode, "PRINTDOCUMENTSIZE") Then
            xmlSetAttributeValue xmlElement, "FILESIZE", xmlGetAttributeText(xmlPrintResponseNode, "PRINTDOCUMENTSIZE")
        Else
            ' If no filesize then store 0 in here just to stop the system from crashing through an invalid use of NULL!
            xmlSetAttributeValue xmlElement, "FILESIZE", "0"
        End If
        ' File contents type is going to be set at the moment
        xmlSetAttributeValue xmlElement, "FILECONTENTS_TYPE", "BIN.BASE64"
        ' Date
        xmlSetAttributeValue xmlElement, "PRINTDATE", xmlGetAttributeText(xmlPrintNode, "PRINTTIMESTAMP")
        ' File contents came back from DPS in the vxmlResponseNode
        strTest = xmlGetAttributeText(xmlPrintNode, "PRINTDOCUMENT")
        xmlSetAttributeValue xmlElement, "FILECONTENTS", strTest
        
        ' Append that to the response node
        vxmlResponseNode.appendChild xmlElement
        ' And copy the control data node over too so that it can be passed back to me for the Save operation!
        vxmlResponseNode.appendChild xmlControlDataNode
        ' And the applicationdetail
        Set xmlElement = xmlDoc.createElement("APPLICATIONDETAIL")
        xmlSetAttributeValue xmlElement, "APPLICATIONNUMBER", xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
        xmlSetAttributeValue xmlElement, "TEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "TEMPLATEID")
        xmlSetAttributeValue xmlElement, "SEARCHKEY1", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY1")
        xmlSetAttributeValue xmlElement, "SEARCHKEY2", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY2")
        xmlSetAttributeValue xmlElement, "SEARCHKEY3", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY3")
        vxmlResponseNode.appendChild xmlElement
    End If ' Free Format document
    errCheckError cstrFunctionName, TypeName(Me)
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName

ProcessRequestExit:

    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlRequestElement = Nothing
    Set xmlCheckPrintResponseNode = Nothing
    Set xmlCreateDocumentArchiveRequest = Nothing
    Set xmlCreateDocumentArchiveResponse = Nothing
    Set xmlCreateAuditTrailRequest = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlDocumentDetailNode = Nothing
    Set xmlArchiveNode = Nothing
    Set xmlPrintDocumentRequestDoc = Nothing
    Set xmlPrintDocumentRequestNode = Nothing
    Set xmlPrintResponseNode = Nothing
    Set xmlPrintNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlElement = Nothing
    Set xmlElement2 = Nothing
    Set xmlElement3 = Nothing
    Set xmlTempNode = Nothing
    
    ' IK_20/04/2006_CORE261
    Set xmlFVS = Nothing

    Exit Function

vbErrorHandler:
    
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    'Raise the error
    Err.Raise lngErrorNumber, strErrorSource, strErrorDescription

    ' go to clean-up section
    Resume ProcessRequestExit
    
End Function
Private Function CreateDocumentArchive( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal bEditDoc As Boolean) As String

Const cstrFunctionName As String = "CreateDocumentArchive"

Dim bFreeFormat As Boolean

Dim objFVSBO As Object

Dim strRequest As String
Dim strResponse As String

Dim xmlAtt As IXMLDOMAttribute
Dim xmlDoc As New FreeThreadedDOMDocument40
Dim xmlGetFileNode As IXMLDOMNode
Dim xmlNewElement As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
    
    On Error GoTo CreateDocumentArchive_Err
    
' TW 29/11/2005 CORE219
    Select Case DocumentManagementSystemType
        ' IK_20/04/2006_CORE261
        Case DOCUMENTMANAGEMENTSYSTEMTYPE_DMS, DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI
            gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
            Set xmlRequestNode = vxmlRequestNode.cloneNode(False)
            
            bFreeFormat = (xmlGetAttributeAsLong(xmlRequestNode, "FREEFORMAT") = 1)
            
            Set xmlNode = xmlGetNode(vxmlRequestNode, "PRINTDOCUMENTDATA")
            ' If we are getting a document for edit, then we need to check it out...
            If bEditDoc = True And bFreeFormat = False Then
                xmlSetAttributeValue xmlRequestNode, "OPERATION", "CHECKINFILE"
                ' Need to change the document data node to say CHECKINDATA
                Set xmlNewElement = xmlRequestNode.ownerDocument.createElement("CHECKINDATA")
            Else
                xmlSetAttributeValue xmlRequestNode, "OPERATION", "CREATEFILE"
                ' Need to change the document data node to say CREATEFILEDATA
                Set xmlNewElement = xmlRequestNode.ownerDocument.createElement("CREATEFILEDATA")
            End If
            
            ' Now copy over all the attributes
            For Each xmlAtt In xmlNode.Attributes
                xmlSetAttributeValue xmlNewElement, xmlAtt.Name, xmlAtt.Value
            Next
            
            'AS 04/04/2006 CORE261 Add support for Gemini.
            xmlSetAttributeValue xmlNewElement, "DOCUMENTMANAGEMENTSYSTEMTYPE", DocumentManagementSystemType
            
            xmlRequestNode.appendChild xmlNewElement
            If gobjContext Is Nothing Then
                Set gobjContext = GetObjectContext()
            End If
            
            ' Create an instance to the file versioning system
            If objFVSBO Is Nothing Then
                Set objFVSBO = gobjContext.CreateInstance(gstrFVS_COMPONENT & ".FileVersioningBO")
                
                gobjTrace.TraceInitialiseOffspring objFVSBO
            End If
            
            ' And call the public interface
            strRequest = xmlRequestNode.xml
            
            ' ik_debug
'            Dim xmlDebug As FreeThreadedDOMDocument40
'            Set xmlDebug = New FreeThreadedDOMDocument40
'            xmlDebug.loadXML strRequest
'            xmlDebug.Save "c:\omiga4Trace\fvsRequest.xml"
            ' ik_debug_ends
            
            'Useful debug info
            ' AS 06/07/04 BBG919 Fixed very slow logging to Event log (don't log document)
        '    App.LogEvent cstrClassName & "::" & cstrFunctionName & "()." ' Debug Information. xmlRequestNode.xml: " & xmlRequestNode.xml
            strResponse = objFVSBO.omRequest(strRequest)
            
            ' ik_debug
'            xmlDebug.loadXML strResponse
'            xmlDebug.Save "c:\omiga4Trace\fvsResponse.xml"
'            Set xmlDebug = Nothing
            ' ik_debug_ends
    
            ' IK_20/04/2006_CORE261 (response not used until now)
            CreateDocumentArchive = strRequest
            
            Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    
            'CheckResponse for errors
            ' IK_20/04/2006_CORE261
            errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
            
            ' if we get here, everything is OK
            xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
                
            Set xmlNode = xmlGetNode(xmlDoc, "RESPONSE")
            If bEditDoc = True And bFreeFormat = False Then
                Set xmlGetFileNode = xmlGetNode(xmlNode, "CHECKINFILE")
            Else
                Set xmlGetFileNode = xmlGetNode(xmlNode, "CREATEFILEDATA")
            End If
            
            ' IK_20/04/2006_CORE261
            If Not xmlGetFileNode Is Nothing Then
                vxmlResponseNode.appendChild xmlGetFileNode
            End If
        
        Case DOCUMENTMANAGEMENTSYSTEMTYPE_FILENET
            strResponse = SendDocumentToFileNet(vxmlRequestNode)
            ' IK_20/04/2006_CORE261 (response not used until now)
            CreateDocumentArchive = strResponse
        
        Case Else
    
    End Select
' TW 29/11/2005 CORE219 End
    
    
    ' IK_20/04/2006_CORE261 (response not used until now)
    ' CreateDocumentArchive = strResponse

CreateDocumentArchive_Exit:
    Set objFVSBO = Nothing
    Set xmlRequestNode = Nothing
    Set xmlNode = Nothing
    Set xmlNewElement = Nothing
    Set xmlDoc = Nothing
    Set xmlAtt = Nothing
    Set xmlGetFileNode = Nothing
    Exit Function

CreateDocumentArchive_Err:
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)

    ' go to clean-up section
    Resume CreateDocumentArchive_Exit
End Function

' ik_bm0269 drop oossINHERIT
Private Function GetDocumentArchive( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode) As String

Const cstrFunctionName As String = "GetDocumentArchive"

Dim adoCommand As New ADODB.Command
Dim adoConnection As ADODB.Connection
Dim adoRecordSet As Recordset

Dim objFVSBO As Object

Dim strDocumentVersion As String
Dim strRequest As String
Dim strResponse As String
Dim strSQL As String

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlDoc2 As FreeThreadedDOMDocument40
Dim xmlElement As IXMLDOMElement
Dim xmlGetCriteriaNode As IXMLDOMNode
Dim xmlGetFileNode As IXMLDOMNode
Dim xmlNewChildNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode

    On Error GoTo GetDocumentArchiveExit
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Debug.Print "Request to getfile = " & vxmlRequestNode.xml
    Set xmlRequestNode = vxmlRequestNode.cloneNode(True)
    ' We need to check the edit flag in the GetCriteria node to see if we need to check out the file or not...
    Set xmlGetCriteriaNode = xmlGetNode(xmlRequestNode, "GETCRITERIA")
    ' If the EDITDOCUMENT attribute is set to TRUE then we call checkoutfile instead of getfile
    
    If xmlGetAttributeText(xmlRequestNode, "EDITDOCUMENT") = "TRUE" Then
        xmlSetAttributeValue xmlRequestNode, "OPERATION", "CHECKOUTFILE"
    Else
        xmlSetAttributeValue xmlRequestNode, "OPERATION", "GETFILE"
    End If
    
    'TW Problem with gobjContext not always being set
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
    End If
    ' Create an instance to the file versioning system
    If objFVSBO Is Nothing Then
        Set objFVSBO = gobjContext.CreateInstance(gstrFVS_COMPONENT & ".FileVersioningBO")
        
        gobjTrace.TraceInitialiseOffspring objFVSBO
    End If
    ' And call the public interface
    ' Need to add the 2 attributes from the child node in the request in to the parent node of the new request
    xmlSetAttributeValue xmlRequestNode, "FILEGUID", xmlGetAttributeText(xmlGetCriteriaNode, "FILEGUID")
    xmlSetAttributeValue xmlRequestNode, "LOCKEDBY", xmlGetAttributeText(xmlRequestNode, "USERID")
    ' Need to check the version number if we're trying to edit as at the moment we can only get the latest version!!!
    
    If xmlGetAttributeText(xmlRequestNode, "EDITDOCUMENT") = "FALSE" Then
        xmlSetAttributeValue xmlRequestNode, "FILEVERSION", xmlGetAttributeText(xmlGetCriteriaNode, "FILEVERSION")
    Else
        ' Get the latest version....
        If adoGetDbProvider = omiga4DBPROVIDEROracle Then
            strSQL = "SELECT fileversion " & _
                     "FROM CURRENTVERSION " & _
                     "WHERE applicationnumber = '" & xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER") & "'" & _
                     " AND FileGUID = '" & xmlGetAttributeText(xmlGetCriteriaNode, "FILEGUID") & "'"
        Else
            strSQL = "SELECT fileversion " & _
                     "FROM CURRENTVERSION " & _
                     "WHERE applicationnumber = '" & xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER") & "'" & _
                     " AND FileGUID = 0x" & xmlGetAttributeText(xmlGetCriteriaNode, "FILEGUID")
        End If
        Debug.Print "Getting latest version from 'CURRENTVERSION' view using sql : " & strSQL
        ' Open a database connection
        Set adoConnection = New ADODB.Connection
        adoConnection.ConnectionString = adoGetDbConnectString
        adoConnection.CursorLocation = adUseClient
        adoConnection.open
            
        ' And a command to execute
        Set adoCommand.ActiveConnection = adoConnection
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
        ' And a recordset to return all our entries...
        Set adoRecordSet = adoCommand.Execute
        
        'TK 15/11/2004
        Set adoRecordSet.ActiveConnection = Nothing
        adoConnection.Close
        
        ' Now check to see if we have any records returned......
        If adoRecordSet.EOF Then
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
        ' And if we have let's start at the beginning and add them all to the XML response
        If Not adoRecordSet.EOF Then 'TK 08/09/2004 BBG879
            adoRecordSet.MoveFirst
            strDocumentVersion = adoRecordSet.fields("FileVersion").Value
            xmlSetAttributeValue xmlRequestNode, "FILEVERSION", strDocumentVersion
            
        End If

        adoRecordSet.Close 'TK 08/09/2004 BBG879
        Set adoRecordSet = Nothing
    End If
    
    'AS 28/04/2006 CORE261 Add support for Gemini.
    xmlSetAttributeValue xmlRequestNode, "DOCUMENTMANAGEMENTSYSTEMTYPE", DocumentManagementSystemType
    
    strRequest = xmlRequestNode.xml
    Debug.Print "Request to FVS = " & strRequest
    strResponse = objFVSBO.omRequest(strRequest)
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    ' ik_bm0269
'    errCheckXMLResponse strResponse, True
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    ' ik_bm0269_ends
    Set xmlNode = xmlGetNode(xmlDoc, "RESPONSE")
    If xmlGetAttributeText(xmlNode, "TYPE") = "SUCCESS" Then
    
        If xmlGetAttributeText(xmlRequestNode, "EDITDOCUMENT") = "TRUE" Then
            Set xmlGetFileNode = xmlGetNode(xmlNode, "CHECKOUTFILE")
        Else
            Set xmlGetFileNode = xmlGetNode(xmlNode, "GETFILE")
        End If
    
        Set xmlNewChildNode = vxmlResponseNode.ownerDocument.createElement("DOCUMENTCONTENTS")
        xmlSetAttributeValue xmlNewChildNode, "FILEVERSION", xmlGetAttributeText(xmlGetFileNode, "FILEVERSION")
        xmlSetAttributeValue xmlNewChildNode, "FILECONTENTSTYPE", xmlGetAttributeText(xmlGetFileNode, "FILECONTENTS_TYPE")
        xmlSetAttributeValue xmlNewChildNode, "FILECONTENTS", xmlGetAttributeText(xmlGetFileNode, "FILECONTENTS")
'TW 10/8/2004
        xmlSetAttributeValue xmlNewChildNode, "DELIVERYTYPE", xmlGetAttributeText(xmlGetFileNode, "DELIVERYTYPE")
'End TW 10/8/2004
        
        'AS 26/07/2006 CORE286 Get file from archive as a Url and not as bin base 64 string.
        Dim strFileContentsUrl As String
        strFileContentsUrl = xmlGetAttributeText(xmlGetFileNode, "FILECONTENTSURL")
        If strFileContentsUrl <> "" Then
            xmlSetAttributeValue xmlNewChildNode, "FILECONTENTSURL", strFileContentsUrl
        End If
        
        vxmlResponseNode.appendChild xmlNewChildNode
        ' Only create an audit trail entry if we are viewing a document.
        ' reprinting a document needs to be auditted after a success from DPS in the PrintDocument BO function
        If Not xmlAttributeValueExists(vxmlRequestNode, "EVENTKEY") Then
            Set xmlDoc2 = New FreeThreadedDOMDocument40
            Set xmlElement = xmlDoc2.createElement("EVENTDETAIL")
            ' Stuff from the root
            xmlSetAttributeValue xmlElement, "APPLICATIONNUMBER", xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
            xmlSetAttributeValue xmlElement, "DOCUMENTGUID", xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID")
            xmlSetAttributeValue xmlElement, "UNITNAME", xmlGetAttributeText(vxmlRequestNode, "UNITID")
            xmlSetAttributeValue xmlElement, "USERNAME", xmlGetAttributeText(vxmlRequestNode, "USERID")
            ' GUID and version are in the GETCRITERIA node
            xmlSetAttributeValue xmlElement, "DOCUMENTVERSION", xmlGetAttributeText(xmlGetCriteriaNode, "FILEVERSION")
            xmlSetAttributeValue xmlElement, "FILEGUID", xmlGetAttributeText(xmlGetCriteriaNode, "FILEGUID")
            ' And event stuff that is created now!
            xmlSetAttributeValue xmlElement, "EVENTKEY", CStr(EVENTKEY_VIEWED)
            xmlSetAttributeValue xmlElement, "EVENTDATE", xmlGetAttributeText(xmlGetCriteriaNode, "EVENTDATE")
            ' We have a print location and number of copies from the previous event
            xmlRequestNode.appendChild xmlElement
            ' Audit the data if wwe have viewed.  If we are editing we wait for the save document to audit...
    
            'AS 05/05/2005 CORE114 - Do not audit on getting archived document for workstation printing.
            If xmlGetAttributeText(xmlRequestNode, "EDITDOCUMENT", "FALSE") = "FALSE" And xmlGetAttributeText(xmlRequestNode, "REPRINTDOCUMENT", "FALSE") = "FALSE" Then
                Call CreateAuditTrail(xmlRequestNode, vxmlResponseNode)
            End If
    
        End If
        ' if we get here, everything is OK
'        xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
        
    End If
    
GetDocumentArchiveExit:
    Set objFVSBO = Nothing
    Set xmlRequestNode = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    Set xmlDoc2 = Nothing
    Set xmlNewChildNode = Nothing
    Set xmlNode = Nothing
    Set xmlGetCriteriaNode = Nothing
    Set xmlGetFileNode = Nothing
    Set adoRecordSet = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
   
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    Debug.Print Err.Description
End Function
Private Function CreateAuditTrail( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode) As String

Const cstrFunctionName As String = "CreateAuditTrail"

    On Error GoTo CreateAuditTrail_Exit
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Call the PrintManagerDO.CreateAuditTrail
    CallDO "CREATEAUDITTRAIL", vxmlRequestNode, vxmlResponseNode

CreateAuditTrail_Exit:
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Debug.Print Err.Description
    errCheckError cstrFunctionName, TypeName(Me)
End Function
Private Sub CallDO( _
    ByVal strOperation As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Calls Operation in the DO
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "CallDO"
    
    On Error GoTo CallDOErr
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim gobjDO As IPrintManagerDO

    If gobjContext Is Nothing Then
        Set gobjDO = New PrintManagerDO
    Else
        Set gobjDO = gobjContext.CreateInstance(gstrPRINTMANAGER_COMPONENT & ".PrintManagerDO")
    End If
    Select Case strOperation
        Case "GETDOCUMENTHISTORYLIST"
            gobjDO.GetDocumentHistoryList vxmlRequestNode, vxmlResponseNode
        Case "GETDOCUMENTEVENTLIST"
            gobjDO.GetDocumentEventList vxmlRequestNode, vxmlResponseNode
        Case "CREATEAUDITTRAIL", "UPDATEDOCUMENTAUDITDETAILS" ' PSC 10/11/2005 CORE196
            'gobjDO.CreateAuditTrail vxmlRequestNode, vxmlResponseNode
            gobjDO.PutAuditTrail vxmlRequestNode, vxmlResponseNode
        Case Else
            errThrowError _
                "PrintManagerRequest", _
                oeNotImplemented, _
                strOperation
    End Select

    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If

CallDOExit:
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    
    Set gobjDO = Nothing

    Exit Sub
    
CallDOErr:
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    ' go to clean-up section
    Resume CallDOExit
End Sub
Private Sub OnMessage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
            
'DR See comments at end of sub routine
'On Error GoTo OnMessage_Exit
Const cstrFunctionName As String = "OnMessage"

Dim return_value As String
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
'    App.LogEvent "omPM.OnMessage() Called", vbLogEventTypeInformation
    If CheckPrint(vxmlRequestNode, vxmlResponseNode) = "SUCCESS" Then
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "MESSQ_RESP_SUCCESS"
    Else
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "MESSQ_MOVE_QUEUE"
    End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
'DR - wheres the error handling!
'OnMessage_Exit:
'
'
'    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Function CheckPrint( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByRef vxmlResponseNode As IXMLDOMNode) As String
    
Const cstrFunctionName As String = "CheckPrint"

Dim lngErrorNumber As Long

Dim objDPS As Object

Dim strErrorDescription As String
Dim strErrorSource As String
Dim strTest As String
Dim strXML As String
    
Dim xmlInDoc As New FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
    
    On Error GoTo vbErrorHandler
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Create a Print Handler BO object to allow us to print via DPS
    If gobjContext Is Nothing Then
        Set objDPS = New PrintHandlerBO
        gobjTrace.TraceInitialiseOffspring objDPS
    Else
        Set objDPS = gobjContext.CreateInstance(gstrDPS_COMPONENT & ".PrintHandlerBO")
        gobjTrace.TraceInitialiseOffspring objDPS
    End If
    
    strXML = vxmlRequestNode.xml
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "objDPS.ProcessRequest call"
    
    strTest = objDPS.ProcessRequest(strXML)
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "objDPS.ProcessRequest return"
    
    ' AS 06/07/04 BBG919 Fixed very slow logging to Event log (don't log document)
'    App.LogEvent cstrClassName & "::" & cstrFunctionName & " Event 1." ' strXML: " & strXML
    ' Add the return XML from DPS to the response node
    
    Set xmlInDoc = xmlLoad(strTest, cstrFunctionName)
    
    ' Get the response and clone it for returning.....
    
    Set xmlNode = xmlInDoc.selectSingleNode("//PRINTDOCUMENTDETAILS")
    If xmlNode Is Nothing Then
        CheckPrint = "FAILED"
        Set vxmlResponseNode = xmlInDoc.selectSingleNode("RESPONSE")
    Else
        CheckPrint = "OK"
        vxmlResponseNode.appendChild xmlNode
    End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If
    
CheckPrintExit:
    Set objDPS = Nothing
    Set xmlInDoc = Nothing
    Set xmlNode = Nothing
    Exit Function
    
vbErrorHandler:
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    'Record error details
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & lngErrorNumber & _
                ", Err.Source: " & strErrorSource & _
                ", Err.Description: " & strErrorDescription, _
                vbLogEventTypeError
    'Simple Error Handling - Bubble up that error!
    Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    ' go to clean-up section
    Resume CheckPrintExit
End Function

Private Function PrintDocument( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode) As String
    
Const cstrFunctionName As String = "PrintDocument"

Dim adoCommand As New ADODB.Command
Dim adoConnection As ADODB.Connection
Dim adoRecordSet As Recordset

Dim objDPS As Object
    
Dim lngFreeFormat As Long

Dim strEmail As Variant
Dim strEmailAddresses As Variant
Dim strEmailAdministrator As String
Dim strEventDate As String
Dim strFileContents As String
Dim strFileID As String
Dim strGuid As String
Dim strHostTemplateDescription As String
Dim strHostTemplateName As String
Dim strNumberOfCopies As String
Dim strPrintDate As String
Dim strPrinterType As String
Dim strPrintLocation As String
Dim strResponse As String
Dim strSQL As String
Dim strVersion As String
Dim strXML As String

Dim xmlControlDataNode As IXMLDOMNode
Dim xmlControlElement As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlDoc2 As FreeThreadedDOMDocument40
Dim xmlDocumentDescriptionAttrib As IXMLDOMAttribute
Dim xmlDocumentDetailNode As IXMLDOMNode
Dim xmlDocumentDetailsNode As IXMLDOMNode
Dim xmlDocumentNameAttrib As IXMLDOMAttribute
Dim xmlDPSDoc As FreeThreadedDOMDocument40
Dim xmlDPSRequest As IXMLDOMNode
Dim xmlElement As IXMLDOMNode
Dim xmlEventNode As IXMLDOMNode
Dim xmlGetDocArchiveRequest As IXMLDOMNode
Dim xmlGetDocArchiveResponse As IXMLDOMNode
Dim xmlNewControlDataNode As IXMLDOMNode
Dim xmlNode As IXMLDOMElement
Dim xmlNode5 As IXMLDOMNode
Dim xmlNode6 As IXMLDOMNode
Dim xmlNode7 As IXMLDOMNode
Dim xmlOutputNode As IXMLDOMNode
Dim xmlOutputTypeNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode
Dim xmlTempNode2 As IXMLDOMNode
Dim xmlTempResponseNode As IXMLDOMNode
Dim xmlTmpNode As IXMLDOMNode
        
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    On Error GoTo PrintDocumentExit         ' PSC 18/11/2004 BBG1806
        
    Set xmlDoc2 = New FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    
    Set xmlDPSRequest = vxmlRequestNode.cloneNode(False)
    
    ' Open a database connection
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    adoConnection.open
        
    ' And a command to execute
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdText
    
    ' Check the freeformat flag
    lngFreeFormat = xmlGetAttributeAsLong(vxmlRequestNode, "FREEFORMAT")
    Set xmlDocumentDetailsNode = xmlGetNode(vxmlRequestNode, "DOCUMENTDETAILS")
    Set xmlControlDataNode = xmlGetNode(vxmlRequestNode, "CONTROLDATA")
    strEmailAdministrator = xmlGetAttributeText(xmlDocumentDetailsNode, "EMAILADMINISTRATOR", "")
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.0]", "lngFreeFormat = " & lngFreeFormat
    
    If lngFreeFormat = 0 Then
        
        ' We have a guid and version in the DOCUMENTDETAILS node of the request so get the file contents from the archive - aka "GetDocumentArchive"
        ' Set up some basic structure for request and response to and from the getdocarchive function
        Set xmlGetDocArchiveRequest = vxmlRequestNode.cloneNode(False)
        Set xmlGetDocArchiveResponse = vxmlResponseNode.cloneNode(False)
        
        xmlSetAttributeValue xmlGetDocArchiveRequest, "OPERATION", "GETFILE"
        Set xmlTempNode = xmlDoc.createElement("GETCRITERIA")
        xmlSetAttributeValue xmlTempNode, "FILEGUID", xmlGetAttributeText(xmlDocumentDetailsNode, "FILEGUID")
        xmlSetAttributeValue xmlTempNode, "FILEVERSION", xmlGetAttributeText(xmlDocumentDetailsNode, "FILEVERSION")
        xmlGetDocArchiveRequest.appendChild xmlTempNode
        ' Need to set an event key as we will otherwise store a view instead of a reprint!
        xmlSetAttributeValue xmlGetDocArchiveRequest, "EVENTKEY", CStr(EVENTKEY_REPRINTED)
                        
        GetDocumentArchive xmlGetDocArchiveRequest, xmlGetDocArchiveResponse
        
        Set xmlTempNode = xmlGetNode(xmlGetDocArchiveResponse, "DOCUMENTCONTENTS")
    Else
        
        ' The contents are in the documentdata node of the request already
        Set xmlTempNode = xmlGetNode(vxmlRequestNode, "DOCUMENTCONTENTS")
    End If
    
    strFileContents = xmlGetAttributeText(xmlTempNode, "FILECONTENTS")
    ' Need to check where we are printing to....  (file/fax/mail/sms/printer?)
    ' Check if the control node exists.  If it doesn't then we need to send the
    ' document to the one specified in the event audit trail
    ' This branch will be hit if it's a free format document anyway as we will have a control data node....
    If Not xmlControlDataNode Is Nothing Then
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]"
        GetPrintLocationEx xmlControlDataNode, strPrinterType, strPrintLocation, strNumberOfCopies
    Else
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]"
        ' Use the data in the event record...
        If adoGetDbProvider = omiga4DBPROVIDEROracle Then
            strSQL = "SELECT a.*, b.documentname from EventAuditDetail a, DocumentAuditDetails b WHERE a.FileGUID = '" & _
                     xmlGetAttributeText(xmlDocumentDetailsNode, "FILEGUID") & "'" & _
                     " AND a.DocumentGUID = '" & _
                     xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & "'" & _
                     " AND a.FileVersion = '" & _
                     xmlGetAttributeText(xmlDocumentDetailsNode, "FILEVERSION") & "'" & _
                     " AND a.DocumentGUID = b.DocumentGUID"
        Else
            strSQL = "SELECT a.*, b.documentname from EventAuditDetail a, DocumentAuditDetails b WHERE a.FileGUID = 0x" & _
                     xmlGetAttributeText(xmlDocumentDetailsNode, "FILEGUID") & _
                     " AND a.DocumentGUID = 0x" & _
                     xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & _
                     " AND a.FileVersion = '" & _
                     xmlGetAttributeText(xmlDocumentDetailsNode, "FILEVERSION") & "'" & _
                     " AND a.DocumentGUID = b.DocumentGUID"
        End If
                 
        Debug.Print "Reprint document - getting output data from event audit trail with SQL = " & strSQL
        adoCommand.CommandText = strSQL
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.3]", "adoCommand.Execute call"
        Set adoRecordSet = adoCommand.Execute
        
        'TK 15/11/2004
        Set adoRecordSet.ActiveConnection = Nothing
        adoConnection.Close
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.4]", "adoCommand.Execute return"
        ' Get the required values from the recordset
        strNumberOfCopies = adoRecordSet.fields("NUMBEROFCOPIES").Value
        strPrinterType = adoRecordSet.fields("PRINTERTYPE").Value
        
        ' If the printer is a remote printer then we have to find the default printer ...
        strPrintLocation = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("DOCUMENTDETAILS"), "PRINTLOCATION")
        strEventDate = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("DOCUMENTDETAILS"), "EVENTDATE")
        ' Now need to create an controldata/outputtype node for DPS to know where to print to!
        If strPrintLocation <> "" Then
            Set xmlNewControlDataNode = xmlDoc.createElement("CONTROLDATA")
            Set xmlOutputTypeNode = xmlDoc.createElement("OUTPUTTYPE")
            Set xmlOutputNode = xmlDoc.createElement(strPrinterType)
            Set xmlNode5 = xmlDoc.createElement("NUMBEROFCOPIES")
            xmlNode5.Text = strNumberOfCopies
            Set xmlNode6 = xmlDoc.createElement("OUTPUTDIRECTORY")
            xmlNode6.Text = strPrintLocation
            
            If (strPrinterType = "FILE") Or (strPrinterType = "EMAIL") Then
                'Get the document name and description from the incoming XML request
                Set xmlTmpNode = vxmlRequestNode.selectSingleNode("//REQUEST/DOCUMENTDETAILS")
                strHostTemplateName = xmlGetMandatoryAttributeText(xmlTmpNode, "DOCUMENTNAME")
                strHostTemplateDescription = xmlGetAttributeText(xmlTmpNode, "DOCUMENTDESCRIPTION")
                Set xmlTmpNode = Nothing
                
                strFileID = adoRecordSet.fields("APPLICATIONNUMBER").Value
                
            End If
            Select Case strPrinterType
                Case "FILE"
                    Set xmlDocumentDetailNode = xmlDoc.createElement("DOCUMENTDETAILS")
                    Set xmlDocumentNameAttrib = xmlDocumentDetailNode.ownerDocument.createAttribute("DOCUMENTNAME")
                    xmlDocumentNameAttrib.Text = strHostTemplateName
                    xmlDocumentDetailNode.Attributes.setNamedItem xmlDocumentNameAttrib
                    xmlNewControlDataNode.appendChild xmlDocumentDetailNode
                    Set xmlNode7 = xmlDoc.createElement("FILEID")
                    xmlNode7.Text = strFileID
                    xmlOutputNode.appendChild xmlNode7
                    xmlOutputNode.appendChild xmlNode5
                    xmlOutputNode.appendChild xmlNode6
                    xmlOutputTypeNode.appendChild xmlOutputNode
                Case "EMAIL"
                    strEmailAddresses = Split(strPrintLocation, ",", -1, 1)
                    For Each strEmail In strEmailAddresses
                        Set xmlOutputNode = xmlDoc.createElement(strPrinterType)
                        xmlSetAttributeValue xmlOutputNode, "EMAILADDRESS", strEmail
                        If strEmailAdministrator <> "" Then
                            xmlSetAttributeValue xmlOutputNode, "EMAILADMINISTRATOR", strEmailAdministrator
                        End If
                        xmlSetAttributeValue xmlOutputNode, "FILEID", strFileID
                        
                        xmlOutputTypeNode.appendChild xmlOutputNode
                    Next
                    'DR Require the document name (because its used in the file naming)
                    Set xmlDocumentDetailNode = xmlDoc.createElement("DOCUMENTDETAILS")
                    Set xmlDocumentNameAttrib = xmlDocumentDetailNode.ownerDocument.createAttribute("DOCUMENTNAME")
                    xmlDocumentNameAttrib.Text = strHostTemplateName
                    xmlDocumentDetailNode.Attributes.setNamedItem xmlDocumentNameAttrib
                    'DR Require the document description (because it's used in the file naming)
                    Set xmlDocumentDescriptionAttrib = xmlDocumentDetailNode.ownerDocument.createAttribute("DOCUMENTDESCRIPTION")
                    xmlDocumentDescriptionAttrib.Text = strHostTemplateDescription
                    xmlDocumentDetailNode.Attributes.setNamedItem xmlDocumentDescriptionAttrib
                    'DR Add the new document details to the Control node.
                    xmlNewControlDataNode.appendChild xmlDocumentDetailNode
                    xmlOutputNode.appendChild xmlNode5
                    xmlOutputNode.appendChild xmlNode6
                Case Else
                    xmlOutputNode.appendChild xmlNode5
                    xmlOutputNode.appendChild xmlNode6
                    Set xmlNode6 = xmlDoc.createElement("PRINTERNAME")
                    xmlNode6.Text = strPrintLocation
                    xmlOutputNode.appendChild xmlNode6
                    xmlOutputTypeNode.appendChild xmlOutputNode
            End Select
            ' And add them all together!!
            xmlNewControlDataNode.appendChild xmlOutputTypeNode
            xmlDPSRequest.appendChild xmlNewControlDataNode
        End If
    End If    ' Using controldata or audit trail for the DPS request XML....
    ' ik_bmids00885
    ' wasteful I know, but for re-print I am just going to replace CONTROLDATA with
    ' CONTROLDATA from REQUEST
    If vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "REPRINTDOCUMENT" Then
        If Not xmlDPSRequest.selectSingleNode("CONTROLDATA") Is Nothing Then
            xmlDPSRequest.removeChild xmlDPSRequest.selectSingleNode("CONTROLDATA")
        End If
        xmlDPSRequest.appendChild xmlControlDataNode.cloneNode(True)
    End If
    ' ik_bmids00885_ends
    If strPrintLocation <> "" Then
        ' Build the XML stream to go to DPS
        xmlSetAttributeValue xmlDPSRequest, "OPERATION", "PRINTDOCUMENT"
        ' Put the file contents into the printdata node of the new request
        Set xmlTempNode = xmlDoc.createElement("PRINTDATA")
        xmlSetAttributeValue xmlTempNode, "PRINTDOCUMENT", strFileContents
        xmlDPSRequest.appendChild xmlTempNode
        ' Convert to a string
        strXML = xmlDPSRequest.xml
        ' CheckPrint () ?????
        
        If gobjContext Is Nothing Then
            Set gobjContext = GetObjectContext()
        End If
        ' Create a Print Handler BO object to allow us to print via DPS
        If objDPS Is Nothing Then
            Set objDPS = gobjContext.CreateInstance(gstrDPS_COMPONENT & ".PrintHandlerBO")
            
            gobjTrace.TraceInitialiseOffspring objDPS
        End If
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[3.1]", "objDPS.ProcessRequest call"
        strResponse = objDPS.ProcessRequest(strXML)
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[3.2]", "objDPS.ProcessRequest return"
        
        Set xmlDPSDoc = xmlLoad(strResponse, cstrFunctionName)
        
        ' Create an audit trail entry if the outcome of the print was successful
        If xmlGetAttributeText(xmlDPSDoc.documentElement, "TYPE") = "SUCCESS" Then
            ' Create audit trail entry
            
            Set xmlRequestNode = vxmlRequestNode.cloneNode(False)
'            Set xmlDocumentDetailsNode = xmlGetNode(vxmlRequestNode, "DOCUMENTDETAILS")
            If lngFreeFormat = 1 Then
                ' If we are a free format document then we need to create a document detail node for the audit trail...
                ' But we already have the document detail node in the control data node so just clone it....
                Set xmlEventNode = xmlGetNode(vxmlRequestNode, "EVENTDETAIL")
                'xmlRequestNode.appendChild xmlTempNode2
                Set xmlTempNode = xmlGetNode(vxmlRequestNode, "APPLICATIONDETAIL")
                Set xmlTempNode2 = xmlTempNode.cloneNode(True)
                xmlRequestNode.appendChild xmlTempNode2
'                Set xmlDocumentDetailsNode = xmlGetNode(vxmlRequestNode, "DOCUMENTDETAILS")
                xmlSetAttributeValue xmlEventNode, "EVENTKEY", CStr(EVENTKEY_CREATED)   ' New doc...
                ' And get the print date from the response...
                
                Set xmlTempResponseNode = xmlGetMandatoryNode(xmlDPSDoc, "RESPONSE")
                Set xmlTmpNode = xmlGetNode(xmlTempResponseNode, "PRINTDOCUMENTDETAILS")
                strPrintDate = xmlGetAttributeText(xmlTmpNode, "PRINTTIMESTAMP")
                xmlSetAttributeValue xmlDocumentDetailsNode, "PRINTTIMESTAMP", strPrintDate
                xmlRequestNode.appendChild xmlDocumentDetailsNode
                xmlRequestNode.appendChild xmlEventNode
            Else
                Set xmlElement = xmlDoc2.createElement("EVENTDETAIL")
                ' Stuff from the root
                xmlSetAttributeValue xmlElement, "APPLICATIONNUMBER", xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
                xmlSetAttributeValue xmlElement, "DOCUMENTGUID", xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID")
                xmlSetAttributeValue xmlElement, "UNITNAME", xmlGetAttributeText(vxmlRequestNode, "UNITID")
                xmlSetAttributeValue xmlElement, "USERNAME", xmlGetAttributeText(vxmlRequestNode, "USERID")
                ' DocumentDetails node only exists if we are not a free format document...
                ' GUID and version are in the DOCUMENTDETAILS node
                xmlSetAttributeValue xmlElement, "DOCUMENTVERSION", xmlGetAttributeText(xmlDocumentDetailsNode, "FILEVERSION")
                xmlSetAttributeValue xmlElement, "FILEGUID", xmlGetAttributeText(xmlDocumentDetailsNode, "FILEGUID")
                xmlSetAttributeValue xmlElement, "EVENTKEY", CStr(EVENTKEY_REPRINTED)
                xmlSetAttributeValue xmlElement, "PRINTLOCATION", strPrintLocation
                xmlSetAttributeValue xmlElement, "EVENTDATE", strEventDate
                xmlSetAttributeValue xmlElement, "PACKFULFILLMENTGUID", xmlGetAttributeText(xmlControlDataNode, "PACKFULFILLMENTGUID")
                xmlSetAttributeValue xmlElement, "HOSTTEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DOCUMENTID")
                xmlSetAttributeValue xmlElement, "FILENETIMAGEREF", xmlGetAttributeText(xmlControlDataNode, "FILENETIMAGEREF")
                xmlRequestNode.appendChild xmlElement
            End If
            ' ik_bmids00885
            ' for re-print I am going to add CONTROLDATA from REQUEST
            If vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "REPRINTDOCUMENT" Then
                If Not xmlRequestNode.selectSingleNode("CONTROLDATA") Is Nothing Then
                    xmlRequestNode.removeChild xmlRequestNode.selectSingleNode("CONTROLDATA")
                End If
                xmlRequestNode.appendChild xmlControlDataNode.cloneNode(True)
            End If
            ' ik_bmids00885_ends
            ' Audit the data if wwe have viewed.  If we are editing we wait for the save document to audit...
            Call CreateAuditTrail(xmlRequestNode, vxmlResponseNode)
        End If ' End if response type = "sucess" or not
        ' Set up the response
        xmlSetAttributeValue vxmlResponseNode, "TYPE", xmlGetAttributeText(xmlDPSDoc.documentElement, "TYPE")
        If xmlAttributeValueExists(xmlDPSDoc.documentElement, "DESCRIPTION") Then
            xmlSetAttributeValue vxmlResponseNode, "DESCRIPTION", xmlGetAttributeText(xmlDPSDoc.documentElement, "DESCRIPTION")
        End If
    Else
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "ERROR"
        xmlSetAttributeValue vxmlResponseNode, "DESCRIPTION", "Printer Location not set up correctly"
    End If
    
PrintDocumentExit:

    Set xmlTmpNode = Nothing
    Set xmlDocumentDetailsNode = Nothing
    Set xmlGetDocArchiveRequest = Nothing
    Set xmlGetDocArchiveResponse = Nothing
    Set xmlDPSRequest = Nothing
    Set xmlDoc = Nothing
    Set xmlTempNode = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlOutputTypeNode = Nothing
    Set xmlOutputNode = Nothing
    Set xmlDoc2 = Nothing
    Set xmlDocumentDetailNode = Nothing
    Set xmlDocumentNameAttrib = Nothing
    Set xmlDocumentDescriptionAttrib = Nothing
    Set xmlNode5 = Nothing
    Set xmlNode6 = Nothing
    Set xmlNode7 = Nothing
    Set objDPS = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordSet = Nothing
    Set xmlNewControlDataNode = Nothing
    Set xmlDPSDoc = Nothing
    Set xmlElement = Nothing
    Set xmlRequestNode = Nothing
    Set xmlTempNode2 = Nothing
    Set xmlEventNode = Nothing
    Set xmlTempResponseNode = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName

    'Simple Error Logging
   ' PSC 18/11/2004 BBG1806 - Start
    If Err.Number <> 0 Then
        App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
            "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    End If

    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode
    ' PSC 18/11/2004 BBG1806 - End
End Function
Private Function SaveDocument( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode) As String

Const cstrFunctionName As String = "SaveDocument"
    
Dim adoCommand As New ADODB.Command
Dim adoConnection As ADODB.Connection
Dim adoRecordSet As Recordset

Dim strAppNum As String
Dim strDocName As String
Dim strFreeFormat As String
Dim strNumberOfCopies As String
Dim strPrinterType As String
Dim strPrintLocation As String
Dim strSQL As String

Dim xmlControlDataNode As IXMLDOMNode
Dim xmlCreateDocumentArchiveRequest As IXMLDOMNode
Dim xmlCreateDocumentArchiveResponse As IXMLDOMNode
Dim xmlDoc As New FreeThreadedDOMDocument40
Dim xmlDocDetailNode As IXMLDOMNode
Dim xmlDocumentDetailElement As IXMLDOMElement
Dim xmlEventDetailElement As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim xmlPrintDocumentDataNode As IXMLDOMNode
Dim xmlTempDocumentDetailElement As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode

Dim xmlTemplateDataNode As IXMLDOMNode 'BC CORE236 01/02/2006
Dim strCreationDate As String   'BC CORE236 01/02/2006
Dim strQuoteNumber As String     'BC CORE236 01/02/2006
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    On Error GoTo SaveDocumentExit
    
    strCreationDate = Format$(Now, "yyyymmdd_hhmmss") 'BC CORE236 01/02/2006
    'DR DMSSYS0005 if FREEFORMAT not supplied then default to 0
    strFreeFormat = xmlGetAttributeText(vxmlRequestNode, "FREEFORMAT", "0")
    ' Copy the start of the request
    Set xmlCreateDocumentArchiveRequest = vxmlRequestNode.cloneNode(False)
    Set xmlCreateDocumentArchiveResponse = vxmlResponseNode.cloneNode(False)
    
    ' Set up the appropriate nodes for the archive request
    Set xmlPrintDocumentDataNode = vxmlResponseNode.ownerDocument.createElement("PRINTDOCUMENTDATA")
    xmlSetAttributeValue xmlPrintDocumentDataNode, "CREATEDBY", xmlGetAttributeText(vxmlRequestNode, "USERID")
    
    If strFreeFormat = "1" Then
        Set xmlNode = xmlGetNode(vxmlRequestNode, "DOCUMENTCONTENTS")
        Set xmlControlDataNode = xmlGetNode(vxmlRequestNode, "CONTROLDATA")
        Set xmlDocDetailNode = xmlGetNode(xmlControlDataNode, "DOCUMENTDETAILS")
        ' Filename is a concatenation of document name and application number
        strDocName = xmlGetAttributeText(xmlDocDetailNode, "DOCUMENTNAME")
        strAppNum = xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
        
        'BC CORE236 24/01/2006 Begin
        Set xmlTemplateDataNode = xmlGetMandatoryNode(vxmlRequestNode, "TEMPLATEDATE")
        strQuoteNumber = xmlGetAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
        ' AS 05/05/2006 CORE261 Only add quote number and creation date to FILENAME if using FILENET.
        If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_FILENET Then
            xmlSetAttributeValue xmlPrintDocumentDataNode, "FILENAME", strDocName + "_" + strAppNum + "_" + strQuoteNumber + "_" + strCreationDate
        Else
            xmlSetAttributeValue xmlPrintDocumentDataNode, "FILENAME", strDocName & "_" & strAppNum & ".DOC"
        End If
        'BC CORE236 24/01/2006 End
    Else
        Set xmlNode = xmlGetNode(vxmlRequestNode, "PRINTDOCUMENTDATA")
    End If
    
    ' AS 21/04/2005 CORE106 Pass compression attributes into omFVS.
    xmlSetAttributeValue xmlPrintDocumentDataNode, "COMPRESSED", xmlGetAttributeText(xmlNode, "COMPRESSED", "0")
    xmlSetAttributeValue xmlPrintDocumentDataNode, "COMPRESSIONMETHOD", xmlGetAttributeText(xmlNode, "COMPRESSIONMETHOD")
        
    ' AS 17/05/2005 CORE129 Pass delivery type into omFVS.
    Dim strDeliveryType As String
    strDeliveryType = xmlGetAttributeText(xmlNode, "DELIVERYTYPE")
    If strDeliveryType <> "" Then
        xmlSetAttributeValue xmlPrintDocumentDataNode, "DELIVERYTYPE", strDeliveryType
    End If
    
    xmlSetAttributeValue xmlPrintDocumentDataNode, "FILEGUID", xmlGetAttributeText(xmlNode, "FILEGUID")
    xmlSetAttributeValue xmlPrintDocumentDataNode, "FILEVERSION", xmlGetAttributeText(xmlNode, "FILEVERSION")
    xmlSetAttributeValue xmlPrintDocumentDataNode, "FILESIZE", xmlGetAttributeText(xmlNode, "FILESIZE")
    xmlSetAttributeValue xmlPrintDocumentDataNode, "FILECONTENTS", xmlGetAttributeText(xmlNode, "FILECONTENTS")
    xmlSetAttributeValue xmlPrintDocumentDataNode, "FILECONTENTS_TYPE", xmlGetAttributeText(xmlNode, "FILECONTENTS_TYPE")
    
    If xmlAttributeValueExists(xmlNode, "METADATA") Then
        xmlSetAttributeValue xmlPrintDocumentDataNode, "METADATA", xmlGetAttributeText(xmlNode, "METADATA")
    End If
    xmlCreateDocumentArchiveRequest.appendChild xmlPrintDocumentDataNode
    ' Use the data in the DOCUMENTDATA node of the request to call FVS.CreateDocumentArchive (with the edit flag set to true so
    ' we know to check it in!)
    
    ' IK_20/04/2006_CORE261 (response not used until now)
    ' CreateDocumentArchive xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse, True
    Dim xmlFVS As DOMDocument40
    Set xmlFVS = New DOMDocument40
    xmlFVS.async = False
    xmlFVS.loadXML CreateDocumentArchive(xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse, True)
    Set xmlCreateDocumentArchiveRequest = xmlFVS.documentElement
    ' IK_20/04/2006_CORE261_ends
    
    ' ik_debug
'    Dim xmlDebug As DOMDocument40
'    Set xmlDebug = New DOMDocument40
'    xmlDebug.async = False
'    xmlDebug.loadXML xmlCreateDocumentArchiveRequest.xml
'    xmlDebug.Save "c:\omiga4Trace\SaveDocumentCreateDocumentArchiveRequest.xml"
'    xmlDebug.loadXML xmlCreateDocumentArchiveResponse.xml
'    xmlDebug.Save "c:\omiga4Trace\SaveDocumentCreateDocumentArchiveResponse.xml"
    ' ik_debug_ends
    
    ' And Audit the saved edit file
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlEventDetailElement = xmlDoc.createElement("EVENTDETAIL")
    If strFreeFormat = "1" Then
        Set xmlTempNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CREATEFILEDATA")
    Else
        Set xmlTempNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CHECKINFILE")
    End If
    xmlSetAttributeValue xmlEventDetailElement, "APPLICATIONNUMBER", xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
    xmlSetAttributeValue xmlEventDetailElement, "DOCUMENTGUID", xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID")
    xmlSetAttributeValue xmlEventDetailElement, "DOCUMENTVERSION", xmlGetAttributeText(xmlTempNode, "FILEVERSION")
    'DR DMSSYS0005 Fixed 24 typo
    xmlSetAttributeValue xmlEventDetailElement, "EVENTTIMESTAMP", Format$(Now, "dd mmm yyyy hh:nn:ss")
    If strFreeFormat = "1" Then
        xmlSetAttributeValue xmlEventDetailElement, "EVENTKEY", CStr(EVENTKEY_CREATED)  ' New doc...coz it's a free format one!
        xmlSetAttributeValue xmlEventDetailElement, "FILEGUID", xmlGetAttributeText(xmlTempNode, "FILEGUID")
    Else
        xmlSetAttributeValue xmlEventDetailElement, "FILEGUID", xmlGetAttributeText(xmlNode, "FILEGUID")
        xmlSetAttributeValue xmlEventDetailElement, "EVENTKEY", CStr(EVENTKEY_EDITED)
        xmlSetAttributeValue xmlEventDetailElement, "EVENTDATE", xmlGetAttributeText(xmlNode, "EVENTDATE")
    End If
    xmlSetAttributeValue xmlEventDetailElement, "UNITNAME", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlEventDetailElement, "USERNAME", xmlGetAttributeText(vxmlRequestNode, "USERID")
    ' Need to get the print locations from the previous event for this file...
    ' BUT ONLY IF WE ARE NOT IN FREE FORMAT DOCUMENT!!!
    If strFreeFormat = "0" Then
        If adoGetDbProvider = omiga4DBPROVIDEROracle Then
            strSQL = "SELECT a.*, b.documentname from EventAuditDetail a, DocumentAuditDetails b WHERE a.FileGUID = '" & _
                     xmlGetAttributeText(xmlNode, "FILEGUID") & "'" & _
                     " AND a.DocumentGUID = '" & _
                     xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & "'" & _
                     " AND a.FileVersion = '" & _
                     xmlGetAttributeText(xmlNode, "FILEVERSION") & "'" & _
                     " AND a.DocumentGUID = b.DocumentGUID"
        Else
            strSQL = "SELECT a.*, b.documentname from EventAuditDetail a, DocumentAuditDetails b WHERE a.FileGUID = 0x" & _
                     xmlGetAttributeText(xmlNode, "FILEGUID") & _
                     " AND a.DocumentGUID = 0x" & _
                     xmlGetAttributeText(vxmlRequestNode, "DOCUMENTGUID") & _
                     " AND a.FileVersion = '" & _
                     xmlGetAttributeText(xmlNode, "FILEVERSION") & "'" & _
                     " AND a.DocumentGUID = b.DocumentGUID"
        End If
                 
        Debug.Print "Save document - getting print location and info from event audit trail with SQL = " & strSQL
        Set adoConnection = New ADODB.Connection
        adoConnection.ConnectionString = adoGetDbConnectString
        adoConnection.CursorLocation = adUseClient
        adoConnection.open
            
        ' And a command to execute
        Set adoCommand.ActiveConnection = adoConnection
        adoCommand.CommandType = adCmdText
        adoCommand.CommandText = strSQL
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "adoCommand.Execute call"
        Set adoRecordSet = adoCommand.Execute
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "adoCommand.Execute return"
        
        'TK 15/11/2004
        Set adoRecordSet.ActiveConnection = Nothing
        adoConnection.Close
        
        ' Throw an error if we haven't got a valid record, i.e. end or beginning of file....
        ' Otherwise snaffle the values from the record
        If (Not adoRecordSet.EOF) And (Not adoRecordSet.BOF) Then
            ' Get the required values from the recordset
            strPrintLocation = adoRecordSet.fields("PRINTLOCATION").Value
            strNumberOfCopies = adoRecordSet.fields("NUMBEROFCOPIES").Value
            strPrinterType = adoRecordSet.fields("PRINTERTYPE").Value
        Else
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
    Else
        ' Free Format document means we have the control data node supplied - use it!
        GetPrintLocationEx xmlControlDataNode, strPrinterType, strPrintLocation, strNumberOfCopies
    End If
    xmlSetAttributeValue xmlEventDetailElement, "OUTPUTLOCATION", strPrintLocation
'    If (strPrinterType = "PRINTER") Then
'        xmlSetAttributeValue xmlEventDetailElement, "NUMCOPIES", strNumberOfCopies
'    Else
'        xmlSetAttributeValue xmlEventDetailElement, "NUMCOPIES", "1"
'    End If
    xmlSetAttributeValue xmlEventDetailElement, "NUMCOPIES", strNumberOfCopies
    xmlSetAttributeValue xmlEventDetailElement, "PRINTLOCATION", strPrintLocation
    xmlSetAttributeValue xmlEventDetailElement, "PRINTERTYPE", strPrinterType
    If strFreeFormat = "0" Then
        Set xmlNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CHECKINFILE")
        Set xmlTempDocumentDetailElement = xmlDoc.createElement("DOCUMENTDETAILS")
    Else
        Set xmlNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CREATEFILEDATA")
        Set xmlTempDocumentDetailElement = xmlGetNode(xmlControlDataNode, "DOCUMENTDETAILS")   'xmlDoc.createElement("DOCUMENTDETAILS")
    End If
    Set xmlDocumentDetailElement = xmlTempDocumentDetailElement.cloneNode(True)
    xmlSetAttributeValue xmlDocumentDetailElement, "ARCHIVETIMESTAMP", xmlGetAttributeText(xmlNode, "ARCHIVETIMESTAMP")
    vxmlRequestNode.appendChild xmlEventDetailElement
    vxmlRequestNode.appendChild xmlDocumentDetailElement
    
    If strFreeFormat = "0" Then
        CreateAuditTrail vxmlRequestNode, vxmlResponseNode
        Set xmlNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CHECKINFILE")
    Else
        Set xmlNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CREATEFILEDATA")
    End If
    ' And return success
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
        
    'IK_20/04/2006_CORE261
    If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
        QueueGeminiRequest xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse
    End If
    'IK_20/04/2006_CORE261_ends
    
    vxmlResponseNode.appendChild xmlNode
    
SaveDocumentExit:
    ' PSC 18/11/2004 BBG1806 - Start
    Dim lngErrorNumber As Long
    Dim strErrorSource As String
    Dim strErrorDescription As String
    
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    ' PSC 18/11/2004 BBG1806 - End
    
    Set xmlCreateDocumentArchiveRequest = Nothing
    Set xmlCreateDocumentArchiveResponse = Nothing
    Set xmlPrintDocumentDataNode = Nothing
    Set xmlNode = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlDoc = Nothing
    Set xmlDocDetailNode = Nothing
    Set xmlEventDetailElement = Nothing
    Set xmlDocumentDetailElement = Nothing
    Set xmlTempNode = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordSet = Nothing
    Set xmlTempDocumentDetailElement = Nothing
    
    ' IK_20/04/2006_CORE261
    Set xmlFVS = Nothing
    
    ' PSC 18/11/2004 BBG1806 - Start
    'IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    If Not gobjContext Is Nothing Then
        If lngErrorNumber = 0 Then
            gobjContext.SetComplete
        Else
            gobjContext.SetAbort
        End If
    End If
    
    If lngErrorNumber <> 0 Then
        'Simple Error Logging
        App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
    ' PSC 18/11/2004 BBG1806 - End
    ' PSC 18/11/2004 BBG1806 - End

End Function
Public Sub GetPrintLocation(ByVal xmlControlNode As IXMLDOMNode, ByRef strPrintType, ByRef strPrintLocation)
    GetPrintLocationEx xmlControlNode, strPrintType, strPrintLocation, "1"
End Sub
Private Sub GetPrintLocationEx(ByVal xmlControlNode As IXMLDOMNode, ByRef strPrintType, ByRef strPrintLocation, ByRef strNumberOfCopies As String)

Const cstrFunctionName As String = "GetPrintLocation"

Dim intNumCopies As Integer
Dim xmlNodeTemp As IXMLDOMNode
Dim xmlOutputTypeNode As IXMLDOMNode
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Set xmlOutputTypeNode = xmlGetNode(xmlControlNode, "OUTPUTTYPE")
    
    For Each xmlNodeTemp In xmlOutputTypeNode.childNodes
        strPrintType = xmlNodeTemp.nodeName
        Select Case xmlNodeTemp.nodeName
            Case "PRINTER"
                strPrintLocation = xmlGetNodeText(xmlNodeTemp, "PRINTERNAME")
                intNumCopies = xmlGetNodeAsLong(xmlNodeTemp, "COPIES")
            
            Case "FILE"
                strPrintLocation = xmlGetNodeText(xmlNodeTemp, "OUTPUTDIRECTORY")
            
            Case "EMAIL"
                If strPrintLocation <> "" Then
                    strPrintLocation = strPrintLocation & "," 'Handle multiple recipients
                End If
                strPrintLocation = strPrintLocation & xmlGetAttributeText(xmlNodeTemp, "EMAILADDRESS")
                
            Case "FAX"
                strPrintLocation = xmlNodeTemp.Text
        End Select
    Next
    
    If strPrintLocation = "" Then
        strPrintLocation = "NOTHING"
        strPrintType = "NOTHING"
    Else
        If intNumCopies < 1 Then
            intNumCopies = 1 'Default number of copies to 1
        End If
        strNumberOfCopies = intNumCopies
    End If
    
    Set xmlNodeTemp = Nothing
    Set xmlOutputTypeNode = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
End Sub

Private Sub WriteDebugInfoToFile(ByVal vstrLine As String, _
                           ByVal vstrFunctionName As String)
' Header ----------------------------------------------------------------------------------
' Description:
'   Writes the passed string to a file on the C drive.
' Pass:
'   vstrXML
'       The XML to write to the file.
' Returns:
'       N/A
' Raise Errors:
'------------------------------------------------------------------------------------------
#If WriteSQLToFile Then
    Dim hFile As Integer
    On Error Resume Next
    hFile = FreeFile
    Open "C:\temp\dpstest.log" For Append Access Write As hFile
    Print #hFile, vbCrLf & vbCrLf
    Print #hFile, "********************************************************************"
    Print #hFile, Format$(Now, "DD/MM/YYYY HH:MM:SS") & "     (" & vstrFunctionName & ")" & vbCrLf
    Print #hFile, vstrLine
    Print #hFile, "********************************************************************"
    Close hFile
#End If
End Sub
Private Function ProcessMergeDocument( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode) As String
    
Const cstrFunctionName As String = "ProcessMergeDocument"

Dim objDPS As Object

Dim xmlDPSNode As IXMLDOMNode
Dim xmlDPSResponseDoc As FreeThreadedDOMDocument40
    
    On Error GoTo ProcessMergeDocumentErr
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
    End If
    Set xmlDPSResponseDoc = New FreeThreadedDOMDocument40
       
    Set objDPS = gobjContext.CreateInstance(gstrDPS_COMPONENT & ".PrintHandlerBO")
    
    gobjTrace.TraceInitialiseOffspring objDPS
         
    ' Call the omDPS.PrintHandlerBO.MergeDocument and load response into document to check for errors
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "objDPS.MergeDocument call"
    
    Set xmlDPSResponseDoc = xmlLoad(objDPS.MergeDocument(vxmlRequestNode.xml), cstrFunctionName)
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "objDPS.MergeDocument return"
    'CheckResponse for errors
'    errCheckXMLResponse xmlDPSResponseDoc.xml, True, vxmlResponseNode
    errCheckXMLResponseNode xmlDPSResponseDoc.documentElement, vxmlResponseNode, True
       
    'Append all the nodes from the document to the response node
    For Each xmlDPSNode In xmlDPSResponseDoc.documentElement.childNodes
        vxmlResponseNode.appendChild xmlDPSNode.cloneNode(True)
    Next
    
    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If
        
ProcessMergeDocumentExit:
    Set objDPS = Nothing
    Set xmlDPSResponseDoc = Nothing
    Set xmlDPSNode = Nothing
    Exit Function
    
ProcessMergeDocumentErr:
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)

    ' go to clean-up section
    Resume ProcessMergeDocumentExit
End Function
Private Function SavePrintDocument(ByVal vxmlRequestNode As IXMLDOMNode, _
                                   ByVal vxmlResponseNode As IXMLDOMNode) As String
'****************************************************************************************
'BG 21/10/2002 BMIDS00592 - Added method.  Copied the code below largely from
'                           ProcessRequest on this class.
'IK 14/11/2002 BMIDS00831   various fixes re: task based printing
'TW 4/8/2004 Tidied up
'TW 01/12/2005 CORE222      Corrected to ensure FileNet attribues are available for CreateDocumentArchive
'****************************************************************************************
Const cstrFunctionName As String = "SavePrintDocument"

Dim lngErrorNumber As Long
Dim lngNumCopies As Long

Dim strAppNum As String
Dim strDocName As String
Dim strDocumentId As String
Dim strErrorDescription As String
Dim strErrorSource As String
Dim strDestinationType As String
Dim strTemp As String

Dim xmlArchiveNode As IXMLDOMNode
Dim xmlCreateAuditTrailRequest As IXMLDOMNode
Dim xmlCreateDocumentArchiveRequest As IXMLDOMNode
Dim xmlCreateDocumentArchiveResponse As IXMLDOMNode
Dim xmlControlDataNode As IXMLDOMElement
Dim xmlDoc As New FreeThreadedDOMDocument40
Dim xmlElement As IXMLDOMElement
Dim xmlElement2 As IXMLDOMElement
Dim xmlElement3 As IXMLDOMElement
Dim xmlPrintDocumentDataNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode
Dim xmlTemplateDataNode As IXMLDOMNode 'BC CORE236 01/02/2006

' TW 01/12/2005 CORE222
Dim xmlAttribute As IXMLDOMAttribute
' TW 01/12/2005 CORE222 End

Dim strCreationDate As String   'BC CORE236 01/02/2006
Dim strQuoteNumber As String    'BC CORE236 01/02/2006

On Error GoTo SavePrintDocumentExit     ' PSC 18/11/2004 BBG1806

    strCreationDate = Format$(Now, "yyyymmdd_hhmmss") 'BC CORE236 01/02/2006
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' Cloning a node will store the 1st element with their attributes to stop us
    ' having to set all the user ID and the rest ....
    Set xmlCreateDocumentArchiveRequest = vxmlRequestNode.cloneNode(False)
           
    Set xmlCreateDocumentArchiveResponse = vxmlResponseNode.cloneNode(False)
        
    '---------------------------------------------------
    ' Store the document info in the archive (FVS)
    '---------------------------------------------------
    Set xmlElement = xmlDoc.createElement("PRINTDOCUMENTDATA")
    
    'DR Made these mandatory - becuase it'll cause a runtime error later if its not supplied
    Set xmlTempNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    Set xmlControlDataNode = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
            
    xmlSetAttributeValue xmlElement, "CREATEDBY", xmlGetAttributeText(vxmlRequestNode, "USERID")
    ' Filename is a concatenation of document name and application number
    
    strDocName = xmlGetAttributeText(xmlControlDataNode, "HOSTTEMPLATENAME")
    strAppNum = xmlGetAttributeText(xmlTempNode, "APPLICATIONNUMBER")
    
' TW 01/12/2005 CORE222 Add necessary attribute for sending document to FileNet
    Set xmlAttribute = xmlCreateDocumentArchiveRequest.ownerDocument.createAttribute("APPLICATIONNUMBER")
    xmlAttribute.Text = strAppNum
    xmlCreateDocumentArchiveRequest.Attributes.setNamedItem xmlAttribute
    strHostTemplateId = xmlGetAttributeText(xmlControlDataNode, "DOCUMENTID")
    
    strDestinationType = xmlGetAttributeText(xmlControlDataNode, "DESTINATIONTYPE")
' TW 01/12/2005 CORE222 End
        
    'BC CORE236 24/01/2006 Begin
    ' PSC 06/02/2006 CORE240 - Start
    Set xmlTemplateDataNode = xmlGetNode(vxmlRequestNode, "TEMPLATEDATA")
    If Not xmlTemplateDataNode Is Nothing Then
        strQuoteNumber = xmlGetAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    End If
    ' PSC 06/02/2006 CORE240 - End
    ' AS 05/05/2006 CORE261 Only add quote number and creation date to FILENAME if using FILENET.
    If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_FILENET Then
        xmlSetAttributeValue xmlElement, "FILENAME", strDocName + "_" + strAppNum + "_" + strQuoteNumber + "_" + strCreationDate
    Else
        xmlSetAttributeValue xmlElement, "FILENAME", strDocName & "_" & strAppNum & ".DOC"
    End If
    'BC CORE236 24/01/2006 End
    
    '---------------------------------------------------
    ' Get the info from the print response
    '---------------------------------------------------
        
    'DR Made this mandatory - because it'll cause a runtime error later if it's not supplied
    Set xmlPrintDocumentDataNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTDOCUMENTDATA")
    
    ' AS 21/04/2005 CORE106 Pass compression attributes into omFVS.
    xmlSetAttributeValue xmlElement, "COMPRESSED", xmlGetAttributeText(xmlPrintDocumentDataNode, "COMPRESSED", "0")
    xmlSetAttributeValue xmlElement, "COMPRESSIONMETHOD", xmlGetAttributeText(xmlPrintDocumentDataNode, "COMPRESSIONMETHOD")
    
    ' AS 18/08/2005 CORE187 Pass METADATA attribute into omFVS.
    Dim strMetaData As String
    strMetaData = xmlGetAttributeText(xmlPrintDocumentDataNode, "METADATA")
    If strMetaData <> "" Then
        xmlSetAttributeValue xmlElement, "METADATA", strMetaData
    End If
    
    ' AS 17/05/2005 CORE129 Pass delivery type into omFVS.
    Dim strDeliveryType As String
    strDeliveryType = xmlGetAttributeText(xmlControlDataNode, "DELIVERYTYPE")
    If strDeliveryType <> "" Then
        xmlSetAttributeValue xmlElement, "DELIVERYTYPE", strDeliveryType
    End If
    
    '-----------------------------------------------------------
    ' Version and GUID come back from the archive operation
    '-----------------------------------------------------------
    ' BG File version - set as one as this will have been the first time it has been saved
    xmlSetAttributeValue xmlElement, "FILEGUID", ""
    xmlSetAttributeValue xmlElement, "FILEVERSION", "V1"
    ' Filesize comes from the DPS response
    xmlSetAttributeValue xmlElement, "FILESIZE", xmlGetAttributeText(xmlPrintDocumentDataNode, "PRINTDOCUMENTSIZE", "0")
    ' File contents type is going to be set at the moment
    xmlSetAttributeValue xmlElement, "FILECONTENTS_TYPE", "BIN.BASE64"
    ' File contents came back from DPS in the vxmlResponseNode
    xmlSetAttributeValue xmlElement, "FILECONTENTS", xmlGetAttributeText(xmlPrintDocumentDataNode, "FILECONTENTS")
    
    xmlCreateDocumentArchiveRequest.appendChild xmlElement
    '----------------------------------------------------------
    ' Now call the create document archive..... '
    '----------------------------------------------------------
    ' IK_20/04/2006_CORE261 (response not used until now)
    ' CreateDocumentArchive xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse, False
    Dim xmlFVS As DOMDocument40
    Set xmlFVS = New DOMDocument40
    xmlFVS.async = False
    xmlFVS.loadXML CreateDocumentArchive(xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse, False)
    Set xmlCreateDocumentArchiveRequest = xmlFVS.documentElement
    ' IK_20/04/2006_CORE261_ends
        
    ' ik_debug
'    Dim xmlDebug As DOMDocument40
'    Set xmlDebug = New DOMDocument40
'    xmlDebug.async = False
'    xmlDebug.loadXML xmlCreateDocumentArchiveRequest.xml
'    xmlDebug.Save "c:\omiga4Trace\SavePrintDocumentCreateDocumentArchiveRequest.xml"
'    xmlDebug.loadXML xmlCreateDocumentArchiveResponse.xml
'    xmlDebug.Save "c:\omiga4Trace\SavePrintDocumentCreateDocumentArchiveResponse.xml"
    ' ik_debug_ends
        
' TW 29/11/2005 CORE219
    Select Case DocumentManagementSystemType
        'IK_20/04/2006_CORE261
        Case DOCUMENTMANAGEMENTSYSTEMTYPE_DMS, DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI
' TW 29/11/2005 CORE219 End
        
            '---------------------------------------------------
            ' Audit the operation
            '---------------------------------------------------
            
            Set xmlElement = xmlDoc.createElement("APPLICATIONDETAIL")
            Set xmlElement2 = xmlDoc.createElement("DOCUMENTDETAILS")
            Set xmlElement3 = xmlDoc.createElement("EVENTDETAIL")
            
            Set xmlCreateAuditTrailRequest = vxmlRequestNode.cloneNode(False)
            xmlSetAttributeValue xmlCreateAuditTrailRequest, "OPERATION", "CreateAuditTrail"
            xmlSetAttributeValue xmlElement, "APPLICATIONNUMBER", strAppNum
            xmlSetAttributeValue xmlCreateAuditTrailRequest, "APPLICATIONNUMBER", strAppNum
            xmlSetAttributeValue xmlElement, "TEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DPSDOCUMENTID")
            xmlSetAttributeValue xmlCreateAuditTrailRequest, "TEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DPSDOCUMENTID")
            
            xmlSetAttributeValue xmlElement, "SEARCHKEY1", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY1")
            xmlSetAttributeValue xmlElement, "SEARCHKEY2", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY2")
            xmlSetAttributeValue xmlElement, "SEARCHKEY3", xmlGetAttributeText(xmlControlDataNode, "SEARCHKEY3")
            
            xmlCreateAuditTrailRequest.appendChild xmlElement
            Set xmlArchiveNode = xmlGetNode(xmlCreateDocumentArchiveResponse, "CREATEFILEDATA")
            xmlSetAttributeValue xmlElement2, "TEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DPSDOCUMENTID")
            xmlSetAttributeValue xmlElement2, "HOSTTEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DOCUMENTID")
            xmlSetAttributeValue xmlElement2, "STAGEID", xmlGetAttributeText(xmlControlDataNode, "STAGEID")
            xmlSetAttributeValue xmlElement2, "DOCUMENTTYPE", xmlGetAttributeText(xmlControlDataNode, "DOCUMENTTYPE")
            xmlSetAttributeValue xmlElement2, "DOCUMENTGROUP", xmlGetAttributeText(xmlControlDataNode, "TEMPLATEGROUPID")
            xmlSetAttributeValue xmlElement2, "DOCUMENTNAME", xmlGetAttributeText(xmlControlDataNode, "HOSTTEMPLATENAME")
            xmlSetAttributeValue xmlElement2, "DOCUMENTDESCRIPTION", xmlGetAttributeText(xmlControlDataNode, "HOSTTEMPLATEDESCRIPTION")
            xmlSetAttributeValue xmlElement2, "CUSTOMERNO", xmlGetAttributeText(xmlControlDataNode, "CUSTOMERNAME")
            xmlSetAttributeValue xmlElement2, "RECIPIENTNAME", xmlGetAttributeText(xmlControlDataNode, "RECIPIENTNAME")
            
            ' This comes from the host system (Does it on an edit too??)
            xmlSetAttributeValue xmlElement2, "DOCUMENTCREATETIMESTAMP", Format$(Now, "dd mmm yyyy hh:nn:ss")
            
            ' Date returned from FVS...
            xmlSetAttributeValue xmlElement2, "ARCHIVETIMESTAMP", xmlGetAttributeText(xmlArchiveNode, "ARCHIVETIMESTAMP")
            xmlSetAttributeValue xmlElement2, "PRINTTIMESTAMP", Format$(Now, "dd mmm yyyy hh:nn:ss")
            
            xmlSetAttributeValue xmlElement2, "LANGUAGE", "3"
            xmlSetAttributeValue xmlElement2, "SOURCESYSTEM", "Omiga4"
            
            'CORE46 Add DeliveryType and DeliveryEngineType
            xmlSetAttributeValue xmlElement2, "DELIVERYENGINETYPE", xmlGetAttributeText(xmlControlDataNode, "DELIVERYENGINETYPE")
            xmlSetAttributeValue xmlElement2, "DELIVERYTYPE", xmlGetAttributeText(xmlControlDataNode, "DELIVERYTYPE")
            
            ' AS 15/12/2006 CORE324 Add support for DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS
            xmlSetAttributeValue xmlElement2, "GEMINIPRINTSTATUS", xmlGetAttributeText(xmlControlDataNode, "GEMINIPRINTSTATUS")
            
            xmlCreateAuditTrailRequest.appendChild xmlElement2
            'xmlElement.appendChild xmlElement2
            xmlSetAttributeValue xmlElement3, "APPLICATIONNUMBER", strAppNum
            ' BG File version - set as one as this will have been the first time it has been saved
            ' IK V1
            xmlSetAttributeValue xmlElement3, "DOCUMENTVERSION", "V1"
            xmlSetAttributeValue xmlElement3, "EVENTTIMESTAMP", Format$(Now, "dd mmm yyyy hh:nn:ss")
            xmlSetAttributeValue xmlElement3, "FILEGUID", xmlGetAttributeText(xmlArchiveNode, "FILEGUID")
            xmlSetAttributeValue xmlElement3, "EVENTKEY", CStr(EVENTKEY_CREATED)
            xmlSetAttributeValue xmlElement3, "UNITNAME", xmlGetAttributeText(vxmlRequestNode, "UNITNAME")
            xmlSetAttributeValue xmlElement3, "USERNAME", xmlGetAttributeText(vxmlRequestNode, "USERNAME")
            xmlSetAttributeValue xmlElement3, "EVENTDATE", Format$(Now, "dd mmm yyyy hh:nn:ss")
            xmlSetAttributeValue xmlElement3, "PACKFULFILLMENTGUID", xmlGetAttributeText(xmlControlDataNode, "PACKFULFILLMENTGUID")
            xmlSetAttributeValue xmlElement3, "HOSTTEMPLATEID", xmlGetAttributeText(xmlControlDataNode, "DOCUMENTID")
            xmlSetAttributeValue xmlElement3, "FILENETIMAGEREF", xmlGetAttributeText(xmlControlDataNode, "FILENETIMAGEREF")
            
            ' ik_bmids00831 (bit of a frig - to say the least),
            ' see method GetPrintLocation for 'inspiration'
            xmlSetAttributeValue xmlElement3, "PRINTLOCATION", xmlGetAttributeText(xmlControlDataNode, "PRINTER")
            
            strDestinationType = xmlGetAttributeText(xmlControlDataNode, "DESTINATIONTYPE")
            Select Case strDestinationType
                Case "F"
                    strTemp = "FILE"
                Case "E"
                    strTemp = "EMAIL"
                Case "W"
                    ' AS 29/01/05 BBGRETAIL strTemp must fit into EVENTAUDITDETAIL.PRINTERTYPE,
                    ' which is nvarchar(10). WPRINTER means "workstation printer".
                    strTemp = "WPRINTER"
                Case Else
                    strTemp = "PRINTER"
            End Select
            xmlSetAttributeValue xmlElement3, "PRINTERTYPE", strTemp
            ' ik_bmids00831_ends
            
            lngNumCopies = xmlGetAttributeAsLong(xmlControlDataNode, "COPIES")
            
            If lngNumCopies > 0 Then
                strTemp = Str$(lngNumCopies)
            Else
                strTemp = "1"
            End If
            xmlSetAttributeValue xmlElement3, "NUMCOPIES", strTemp
            
            xmlCreateAuditTrailRequest.appendChild xmlElement3
                
            CreateAuditTrail xmlCreateAuditTrailRequest, vxmlResponseNode
            
            ' AS 05/06/2006 CORE277 omPM SavePrintDocument needs to return FILEGUID
            xmlSetAttributeValue vxmlResponseNode, "FILEGUID", xmlGetAttributeText(xmlArchiveNode, "FILEGUID")
            
' TW 29/11/2005 CORE219
    End Select
' TW 29/11/2005 CORE219 End

    'IK_20/04/2006_CORE261
    If DocumentManagementSystemType = DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI Then
        QueueGeminiRequest xmlCreateDocumentArchiveRequest, xmlCreateDocumentArchiveResponse
    End If
    'IK_20/04/2006_CORE261_ends
    
    Dim objDPS As Object
    If strDestinationType <> "W" Then
        ' AS 29/01/05 BBGRETAIL Document has already been printed via the workstation,
        ' so do not print again.
        ' PSC 21/10/2004 BBG1639
        
        If gobjContext Is Nothing Then
            Set gobjContext = GetObjectContext()
        End If
        
        If objDPS Is Nothing Then
            Set objDPS = gobjContext.CreateInstance(gstrDPS_COMPONENT & ".PrintHandlerBO")
            gobjTrace.TraceInitialiseOffspring objDPS
        End If
        ' PSC 21/10/2004 BBG1639 - End
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "objDPS.PrintMergedDocument call"
            
        objDPS.PrintMergedDocument vxmlRequestNode
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "objDPS.PrintMergedDocument return"
    End If
    
SavePrintDocumentExit:

    ' PSC 18/11/2004 BBG1806 - Start
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    ' PSC 18/11/2004 BBG1806 - End
    
    ' PSC 21/09/2004 BBG1639 - Start
    Set xmlCreateDocumentArchiveRequest = Nothing
    Set xmlCreateDocumentArchiveResponse = Nothing
    Set xmlElement = Nothing
    Set xmlElement2 = Nothing
    Set xmlElement3 = Nothing
    Set xmlDoc = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlPrintDocumentDataNode = Nothing
    Set xmlCreateAuditTrailRequest = Nothing
    Set xmlArchiveNode = Nothing
    Set objDPS = Nothing
    ' PSC 21/09/2004 BBG1639 - End
' TW 01/12/2005 CORE222
    Set xmlAttribute = Nothing
' TW 01/12/2005 CORE222 End
    
    ' IK_20/04/2006_CORE261
    Set xmlFVS = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
        
    ' PSC 18/11/2004 BBG1806 - Start
    If Not gobjContext Is Nothing Then
        If lngErrorNumber = 0 Then
            gobjContext.SetComplete
        Else
            gobjContext.SetAbort
        End If
    End If
    
    If lngErrorNumber <> 0 Then
        'Simple Error Logging
        App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & Err.Number & ", Err.Source: " & Err.Source & ", Err.Description: " & Err.Description, vbLogEventTypeError
    
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
    ' PSC 18/11/2004 BBG1806 - End

End Function

Private Sub UnlockDocument( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

Const cstrFunctionName As String = "UnlockDocument"

Dim objFVSBO As Object

Dim strResponse As String

Dim xmlRequestPrintDocumentDataNode As IXMLDOMNode
Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
Dim xmlThisRequestNode As IXMLDOMNode
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    On Error GoTo UnlockDocumentErr
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.async = False
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UNLOCKFILE"
    xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlThisRequestNode, "USERID", "LOCKEDBY"
    Set xmlRequestPrintDocumentDataNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTDOCUMENTDATA")
    xmlCopyMandatoryAttribute xmlRequestPrintDocumentDataNode, xmlThisRequestNode, "FILEGUID"
    xmlCopyMandatoryAttribute xmlRequestPrintDocumentDataNode, xmlThisRequestNode, "FILEVERSION"
    
    If gobjContext Is Nothing Then
        Set objFVSBO = New FileVersioningBO
    Else
        Set objFVSBO = gobjContext.CreateInstance(gstrFVS_COMPONENT & ".FileVersioningBO")
    End If
    
    gobjTrace.TraceInitialiseOffspring objFVSBO
    strResponse = objFVSBO.omRequest(xmlThisRequestDoc.xml)
    
    Set objFVSBO = Nothing
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    ' And return success
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"

    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If

UnlockDocumentExit:
    
    Set xmlRequestPrintDocumentDataNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set objFVSBO = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub        ' PSC 20/12/2004 E2EM00003127

UnlockDocumentErr:
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    GoTo UnlockDocumentExit
End Sub

Public Sub InitialiseTraceInterface( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStartElapsed As Single)
    
    If Not vblnTraceOn Then
        Exit Sub
    End If
    
    gobjTrace.TraceInitialiseFromParent vblnTraceOn, vstrFileName, vintStartElapsed
End Sub

'AS 20/12/2006 CORE325 START Refactored.
Private Sub QueueGeminiRequest( _
    ByVal vxmlFVSRequestNode As IXMLDOMNode, _
    ByVal vxmlFVSResponseNode As IXMLDOMNode)

    On Error GoTo ExitHandler
    
    Const cstrFunctionName As String = "QueueGeminiRequest"
    
    Dim xmlFileDataRequest As IXMLDOMElement
    Dim xmlFileDataResponse As IXMLDOMElement
    Dim appNumber As String
    Dim usrId As String
    Dim untId As String
    Dim deliveryTypeValidationType As String
    Dim packs() As GeminiPack
    Dim pack As GeminiPack
    ReDim pack.Documents(0 To 0)
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    pack.ApplicationNumber = vxmlFVSRequestNode.Attributes.getNamedItem("APPLICATIONNUMBER").Text
    pack.UserId = vxmlFVSRequestNode.Attributes.getNamedItem("USERID").Text
    pack.UnitId = vxmlFVSRequestNode.Attributes.getNamedItem("UNITID").Text
    
    Set xmlFileDataRequest = vxmlFVSRequestNode.selectSingleNode("CREATEFILEDATA")
    If xmlFileDataRequest Is Nothing Then
        Set xmlFileDataRequest = vxmlFVSRequestNode.selectSingleNode("CHECKINDATA")
    End If
        
    ' AS 05/12/2006 CORE312 Gemini Printing; simplified setting up parameters.
    Dim guidAttributeName As String
    If Not xmlFileDataRequest Is Nothing Then
        Set xmlFileDataResponse = vxmlFVSResponseNode.selectSingleNode("CREATEFILEDATA")
        guidAttributeName = "FILEGUID"
        If xmlFileDataResponse Is Nothing Then
            Set xmlFileDataResponse = vxmlFVSResponseNode.selectSingleNode("CHECKINFILE")
            guidAttributeName = "FILECONTENTSGUID"
        End If
        
        If Not xmlFileDataResponse Is Nothing Then
            pack.Documents(0).FileContentsGuid = xmlFileDataResponse.getAttribute(guidAttributeName)
            xmlFileDataRequest.setAttribute "FILECONTENTSGUID", pack.Documents(0).FileContentsGuid
        End If
        
        pack.Documents(0).DocumentDetails.DeliveryType = xmlFileDataRequest.getAttribute("DELIVERYTYPE")
        deliveryTypeValidationType = LCase$(GetValidationTypeForValueID("DocumentDeliveryType", pack.Documents(0).DocumentDetails.DeliveryType))
        pack.Documents(0).DocumentDetails.CompressionMethod = xmlFileDataRequest.getAttribute("COMPRESSIONMETHOD")
        pack.Documents(0).DocumentDetails.FileContentsType = xmlFileDataRequest.getAttribute("FILECONTENTS_TYPE")
        pack.Documents(0).DocumentDetails.FileContents = xmlFileDataRequest.getAttribute("FILECONTENTS")
    End If
    
    ' AS 13/11/2006 CORE309 Only put the message on the GeminiInterface queue if the document is not a
    ' TIFF (all scanned documents are TIFFs).
    If Not LCase$(deliveryTypeValidationType) = "tiff" Then
        
        GeminiSendMessageToQueue vxmlFVSRequestNode.xml, "Archive"
        
        ' AS 05/12/2006 CORE312 Gemini Printing; send message to fulfillment if appropriate.
        packs = GeminiSendToFulfillment(gobjContext, pack)
                    
    End If

ExitHandler:
    Set xmlFileDataRequest = Nothing
    Set xmlFileDataResponse = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Sub
'AS 20/12/2006 CORE325 END Refactored.
