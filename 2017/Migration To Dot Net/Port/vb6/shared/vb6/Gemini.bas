Attribute VB_Name = "Gemini"
'Workfile:      Gemini.bas
'Copyright:     Copyright © 2001 Marlborough Stirling
'Created:       13/12/2006
'Author:        Adrian Stanley
'Description:   Gemini helper functions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AS     13/12/2006  First version.
'AS     20/12/2006  CORE325 omPM: Gemini printing pack handling
'AS     04/01/2007  CORE327 Check not trying to print pack document outside of pack.
'AW     25/01/2007  EP1308 Check locked status of documents.
'------------------------------------------------------------------------------------------
Option Explicit

Public Const GEMINIPRINTSTATUS_AWAITINGAPPROVAL = 10
Public Const GEMINIPRINTSTATUS_NOTAPPROVED = 20
Public Const GEMINIPRINTSTATUS_APPROVED = 30
Public Const GEMINIPRINTSTATUS_GEMINIPRINTED = 40

Public Const GEMINIPRINTMODE_NEVER = 10
Public Const GEMINIPRINTMODE_IMMEDIATE = 20
Public Const GEMINIPRINTMODE_ONHOLD = 30

Public Const GEMINILOCATION_DMS = "DMS"
Public Const GEMINILOCATION_MOBIUSPENDING = "Mobius: pending"

Private Const m_fulfillmentRequestNamespace = "http://Omiga.Fsd.Vertex/Gemini.Fulfillment.Request"
Private Const m_geminiQueueName = "GeminiInterface"
Private Const m_geminiComponentName = "omGemini.FileVersioningBO"

Private Const MODULENAME = "Gemini"
Private Const ERRORNUMBERBASE = vbObjectError + 512 + oeRecordNotFound

Public Type GeminiDocumentDetails
    FileContentsType As String                  ' Specify on input if supplying document contents.
    FileContents As String                      ' Specify on input if supplying document contents.
    CompressionMethod As String                 ' Specify on input if supplying document contents.
    DeliveryType As Integer                     ' Specify on input if supplying document contents.
End Type

Public Type GeminiDocument
    DocumentGuid As String                      ' Specify on input or specify FileContentsGuid.
    FileGuid As String                          ' Read from database.
    FileContentsGuid As String                  ' Specify on input or specify DocumentGuid.
    DocumentVersion As String                   ' Read from database.
    DocumentDate As Date                        ' Read from database.
    DocumentName As String                      ' Read from database.
    HostTemplateId As String                    ' Read from database.
    TemplatePackMember As Boolean               ' Read from database.
    PrintStatus As Integer                      ' Read from database.
    PrintMode As Integer                        ' Read from database.
    DocumentLocation As String                  ' Read from database.
    InputDocument As Boolean                    ' True if this document was supplied in the input.
    DocumentDetails As GeminiDocumentDetails    ' Specify on input if supplying document contents.
End Type

Public Type GeminiPack
    PackFulfillmentGuid As String               ' Specify on input or specify at least one document.
    PackControlName As String                   ' Read from database.
    PackCreationDate As Date                    ' Read from database.
    ApplicationNumber As String                 ' Specify on input.
    UserId As String                            ' Specify on input.
    UnitId As String                            ' Specify on input.
    Documents() As GeminiDocument               ' If PackFulfillmentGuid is not specified then specify at least one document on input.
End Type

'------------------------------------------------------------------------------------------
' Summary:
'   Gets the fulfillment request Xml to pass to omGemini for the specific document or
'   file contents.
' Argument:
'   objContext. The COM+ ObjectContext.
' Argument:
'   pack. The input pack.
' Argument:
'   mustExist. If true then each document must already exist in DMS or on Mobius,
'   i.e., Omiga must have received the notification of the documents location in
'   Mobius.
' Argument:
'   printOnHold. If true then an Gemini print a document even if the associated template is set to
'   "On-hold" (or "Immediate"), i.e., document can be created but is only sent to Gemini printing
'   by using the Gemini Print button on DMS110.
' Argument:
'   errorIfGeminiPrintNever. If true then an error is raised if attempting to Gemini print a
'   document where the associated template is set to "Never", i.e., never Gemini print the
'   document.
'   If false then no error is raised, but the document(s) is not Gemini Printed.
'   This should be set to true if printing a pack or packs, and false if printing a single document from PM010.
' Argument:
'   errorIfGeminiPrinted. If true then an error is raised if attempting to Gemini print a
'   document that has already been Gemini Printed.
' Argument:
'   errorIfDocumentNotInPack. If true then an error is raised if attempting to Gemini print a document that is
'   not in a pack and the template for the document is part of a pack definition.
'   If false then no error is raised, but the document is not Gemini Printed.
'   This should be set to true when Gemini Printing from DMS110, and false when Gemini Printing from PM010.
' Returns:
'   An array of packs.
' Remarks:
'
'   On input, the pack parameter should be defined as follows:
'       pack.PackFulfillmentGuid - specify to include all documents in this pack, and to not
'       include any other associated packs.
'       pack.ApplicationNumber - set to the mortgage application number.
'       pack.UserId - set to the user initiating the request.
'       pack.UnitId - set to the unit id for the user.
'       Documents - set as follows:
'           It is only necessary to specify details for one document in the pack (the
'           details for the other documents will be read from the database), UNLESS
'           you are supplying the file contents for a document, in which case include its
'           details as follows.
'           For each document specify either the FileContentsGuid or the
'           DocumentGuid field.
'           If supplying the document contents, then specify all the sub-fields in the
'           DocumentDetails field. If these sub-fields are not supplied, then they will be
'           retrieved either from DMS or Mobius.
'           Do not specify the other document fields as these will be read from
'           the database.
'
'   For the specified pack, the fulfillment request will contain
'   all the associated documents:
'       If pack.PackFulfillmentGuid is specified, then includes all documents in that pack,
'       but does not include any other packs.
'       If the document is part of a pack, include all documents in the pack.
'       If the document is in more than one pack, include all the packs, and all the
'       documents in these packs, UNLESS pack.PackFulfillmentGuid is specified.
'       If any of the documents in any of the packs is not available for Gemini printing
'       (e.g., it does not exist in Mobius, or it is on-hold), then do not include any
'       of these packs.
'------------------------------------------------------------------------------------------
Public Function GeminiSendToFulfillment( _
    ByVal objContext As ObjectContext, _
    ByRef pack As GeminiPack, _
    Optional ByVal mustExist As Boolean = False, _
    Optional ByVal printOnHold As Boolean = False, _
    Optional ByVal errorIfGeminiPrintNever As Boolean = False, _
    Optional ByVal errorIfGeminiPrinted As Boolean = False, _
    Optional ByVal errorIfDocumentNotInPack As Boolean = False) As GeminiPack()
    
    Const cstrFunctionName As String = "GeminiSendToFulfillment"

    On Error GoTo ExitHandler
             
    Debug.Print "->" & cstrFunctionName & "(mustExist = " & CStr(mustExist) & ", printOnHold = " & CStr(printOnHold) & ", errorIfGeminiPrintNever = " & CStr(errorIfGeminiPrintNever) & ")"
    gobjTrace.TraceMethodEntry MODULENAME, cstrFunctionName
             
    CheckInputPack pack
    
    Dim packs() As GeminiPack
    packs = GetPacks(pack)
        
    If CheckOutputPacks(packs, mustExist, printOnHold, errorIfGeminiPrintNever, errorIfGeminiPrinted, errorIfDocumentNotInPack) Then
        ' Packs are valid for sending to Gemini Printing
        GeminiSendMessagesToQueue ToFulfillmentRequestXml(packs), "Fulfillment"
        CreateEventAuditDetails objContext, packs, EVENTKEY_GEMINIPRINTED, GEMINIPRINTSTATUS_GEMINIPRINTED
    End If
    
    GeminiSendToFulfillment = packs
    
ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "()"
    gobjTrace.TraceMethodExit MODULENAME, cstrFunctionName
    
    If Err.Number <> 0 Then
        If Not objContext Is Nothing Then
            objContext.SetAbort
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    End If
End Function

Public Sub GeminiSendToFulfillmentXml( _
    ByVal objContext As ObjectContext, _
    ByVal xmlRequestNode As IXMLDOMNode)
    
    Const cstrFunctionName As String = "GeminiSendToFulfillmentXml"

    On Error GoTo ExitHandler
             
    Debug.Print "->" & cstrFunctionName & "()"
    gobjTrace.TraceXML xmlRequestNode.xml, MODULENAME & "_" & cstrFunctionName & "_request"
    gobjTrace.TraceMethodEntry MODULENAME, cstrFunctionName
    
    Dim xmlPackNode As IXMLDOMNode
    Dim xmlDocumentNodes As IXMLDOMNodeList
    Dim xmlDocumentNode As IXMLDOMNode
    Dim xmlDocumentDetailsNode As IXMLDOMNode
    Dim packs() As GeminiPack
    Dim pack As GeminiPack
    Dim mustExist As Boolean
    Dim printOnHold As Boolean
    Dim errorIfGeminiPrintNever As Boolean
    Dim errorIfGeminiPrinted As Boolean
    Dim errorIfDocumentNotInPack As Boolean
    Dim docGuid As String
    Dim fcGuid As String
    Dim documentIndex As Integer
    
    mustExist = xmlGetAttributeAsBoolean(xmlRequestNode, "MUSTEXIST", "0")
    printOnHold = xmlGetAttributeAsBoolean(xmlRequestNode, "PRINTONHOLD", "0")
    errorIfGeminiPrintNever = xmlGetAttributeAsBoolean(xmlRequestNode, "ERRORIFGEMINIPRINTNEVER", "0")
    errorIfGeminiPrinted = xmlGetAttributeAsBoolean(xmlRequestNode, "ERRORIFGEMINIPRINTED", "0")
    errorIfDocumentNotInPack = xmlGetAttributeAsBoolean(xmlRequestNode, "ERRORIFDOCUMENTNOTINPACK", "0")
    
    Set xmlPackNode = xmlGetMandatoryNode(xmlRequestNode, "PACK")
    
    pack.PackFulfillmentGuid = xmlGetAttributeText(xmlPackNode, "PACKFULFILLMENTGUID")
    pack.ApplicationNumber = xmlGetMandatoryAttributeText(xmlPackNode, "APPLICATIONNUMBER")
    pack.UserId = xmlGetMandatoryAttributeText(xmlRequestNode, "USERID")
    pack.UnitId = xmlGetMandatoryAttributeText(xmlRequestNode, "UNITID")
    
    Set xmlDocumentNodes = xmlGetMandatoryNodeList(xmlPackNode, "DOCUMENTLIST/DOCUMENT")
    ReDim pack.Documents(0 To xmlDocumentNodes.length - 1)
    documentIndex = 0
    For Each xmlDocumentNode In xmlDocumentNodes
        pack.Documents(documentIndex).DocumentGuid = xmlGetAttributeText(xmlDocumentNode, "DOCUMENTGUID")
        pack.Documents(documentIndex).FileContentsGuid = xmlGetAttributeText(xmlDocumentNode, "FILECONTENTSGUID")
        
        Set xmlDocumentDetailsNode = xmlGetNode(xmlDocumentNode, "DOCUMENTDETAILS")
        If Not xmlDocumentDetailsNode Is Nothing Then
            pack.Documents(documentIndex).DocumentDetails.CompressionMethod = xmlGetAttributeText(xmlDocumentDetailsNode, "COMPRESSIONMETHOD")
            pack.Documents(documentIndex).DocumentDetails.DeliveryType = xmlGetAttributeText(xmlDocumentDetailsNode, "DELIVERYTYPE")
            pack.Documents(documentIndex).DocumentDetails.FileContentsType = xmlGetAttributeText(xmlDocumentDetailsNode, "FILECONTENTS_TYPE")
            pack.Documents(documentIndex).DocumentDetails.FileContents = xmlGetAttributeText(xmlDocumentDetailsNode, "FILECONTENTS")
        End If
        documentIndex = documentIndex + 1
    Next

    packs = GeminiSendToFulfillment(objContext, pack, mustExist, printOnHold, errorIfGeminiPrintNever, errorIfGeminiPrinted, errorIfDocumentNotInPack)

ExitHandler:
    Set xmlPackNode = Nothing
    Set xmlDocumentNodes = Nothing
    Set xmlDocumentNode = Nothing
    Set xmlDocumentDetailsNode = Nothing

    Debug.Print "<-" & cstrFunctionName & "()"
    gobjTrace.TraceMethodExit MODULENAME, cstrFunctionName
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub CheckInputPack(ByRef pack As GeminiPack)
    Const cstrFunctionName As String = "CheckInputPack"

    On Error GoTo ExitHandler
             
    Debug.Print "->" & cstrFunctionName & "()"
             
    Dim ref As String
    
    If Len(pack.ApplicationNumber) = 0 Then
        Err.Raise ERRORNUMBERBASE, MODULENAME & "." & cstrFunctionName, "Invalid pack.ApplicationNumber"
    Else
        ref = "Details: ApplicationNumber=" & pack.ApplicationNumber
    End If
    
    If Len(pack.PackFulfillmentGuid) = 0 Then
        On Error Resume Next
        If UBound(pack.Documents) >= 0 Then
            If Err.Number <> 0 Then
                On Error GoTo ExitHandler
                Err.Raise ERRORNUMBERBASE, MODULENAME & "." & cstrFunctionName, "Invalid pack Document" & "." & vbCrLf & ref
            Else
                On Error GoTo ExitHandler
                Dim documentIndex As Integer
                
                For documentIndex = 0 To UBound(pack.Documents)
                    Dim refDocument As String
                    refDocument = ref & ", documentIndex=" & CStr(documentIndex)
                    
                    If Len(pack.Documents(documentIndex).DocumentGuid) = 0 And Len(pack.Documents(documentIndex).FileContentsGuid) = 0 Then
                        Err.Raise ERRORNUMBERBASE, MODULENAME & "." & cstrFunctionName, "Invalid DocumentGuid or FileContentsGuid." & vbCrLf & refDocument
                    End If
                Next
            End If
        End If
    End If
   
    On Error GoTo ExitHandler

ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "()"
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Function CheckOutputPacks( _
    ByRef packs() As GeminiPack, _
    ByVal mustExist As Boolean, _
    ByVal printOnHold As Boolean, _
    ByVal errorIfGeminiPrintNever As Boolean, _
    ByVal errorIfGeminiPrinted As Boolean, _
    ByVal errorIfDocumentNotInPack) As Boolean
    
    Const cstrFunctionName As String = "CheckOutputPacks"
    
    On Error GoTo ExitHandler

    Debug.Print "->" & cstrFunctionName & "()"

    Dim success As Boolean
    Dim refPack As String
    
    success = True
    
    On Error Resume Next
    If UBound(packs) >= 0 Then
        If Err.Number <> 0 Then
            On Error GoTo ExitHandler
            Err.Raise ERRORNUMBERBASE, MODULENAME & "." & cstrFunctionName, "Invalid packs"
        Else
            ' At least one pack.
            On Error GoTo ExitHandler
            
            Dim packIndex As Integer
            For packIndex = 0 To UBound(packs)
            
                If success Then
                
                    Dim pack As GeminiPack
                    pack = packs(packIndex)
                    refPack = GetPackRef(pack)
                                                                              
                    On Error Resume Next
                    If UBound(pack.Documents) >= 0 Then
                        If Err.Number <> 0 Then
                            On Error GoTo ExitHandler
                            Err.Raise ERRORNUMBERBASE, MODULENAME & "." & cstrFunctionName, "Invalid pack Documents" & "." & vbCrLf & refPack
                        Else
                            On Error GoTo ExitHandler
                            
                            Dim documentIndex As Integer
                            For documentIndex = 0 To UBound(pack.Documents)
                                If success Then
                                
                                    Dim document As GeminiDocument
                                    document = pack.Documents(documentIndex)
                                   
                                    Dim strUserId As String
                                    Dim refDocument As String
                                    refDocument = GetDocumentRef(document, refPack)
                                                                                                                                             
                                    If IsFileVersionLocked(document, strUserId) Then
                                        Err.Raise ERRORNUMBERBASE, MODULENAME + "." + cstrFunctionName, "The following document is curently being amended by " & strUserId & " " & "." & vbCrLf & refDocument
                                    End If
                                    
                                    If _
                                        document.InputDocument And _
                                        document.PrintStatus = GEMINIPRINTSTATUS_GEMINIPRINTED And _
                                        errorIfGeminiPrinted Then
                                        ' The input document has already been Gemini Printed and it is an error
                                        ' to attempt to Gemini Print a document more than once.
                                        ' Ignore print status of other documents in the pack as these may have been
                                        ' Gemini Printed in other packs.
                                        Err.Raise ERRORNUMBERBASE, MODULENAME + "." + cstrFunctionName, "Invalid Gemini Print Status (" & GeminiPrintStatusToString(document.PrintStatus) & ")" & "." & vbCrLf & refDocument
                                    End If
                                    
                                    If document.PrintMode = GEMINIPRINTMODE_NEVER Then
                                        ' Document is Gemini Print Never.
                                        If errorIfGeminiPrintNever Then
                                            ' And this should be reported as an error.
                                            Err.Raise ERRORNUMBERBASE, MODULENAME + "." + cstrFunctionName, "Invalid Gemini Print Mode (" & GeminiPrintModeToString(document.PrintMode) & ")" & "." & vbCrLf & refDocument
                                        Else
                                            ' Otherwise do not continue with Gemini Printing.
                                            success = False
                                        End If
                                    End If
        
                                    If success And document.PrintMode = GEMINIPRINTMODE_ONHOLD Then
                                        ' Document is Gemini Print On-Hold, so only Gemini Print if printOnHold is true, i.e.,
                                        ' called from DMS110.
                                        success = printOnHold
                                    End If
        
                                    If success And _
                                        document.PrintStatus <> GEMINIPRINTSTATUS_APPROVED And _
                                        document.PrintStatus <> GEMINIPRINTSTATUS_GEMINIPRINTED Then
                                        ' Document is not Approved or not Gemini Printed.
                                        Err.Raise ERRORNUMBERBASE, MODULENAME + "." + cstrFunctionName, "Invalid Gemini Print Status (" & GeminiPrintStatusToString(document.PrintStatus) & ")" & "." & vbCrLf & refDocument
                                    End If
                                
                                    If success And _
                                        mustExist And _
                                        UCase$(document.DocumentLocation) = UCase$(GEMINILOCATION_MOBIUSPENDING) And _
                                        Len(document.DocumentDetails.FileContents) = 0 Then
                                        ' Document has been sent to Mobius but no notification has been received back as to its
                                        ' location in Mobius, and contents of the document where not included in the input pack.
                                        Err.Raise ERRORNUMBERBASE, MODULENAME + "." + cstrFunctionName, "Invalid Document Location (" & document.DocumentLocation & ")" & "." & vbCrLf & refDocument
                                    End If
                                    
                                    If success And Len(pack.PackFulfillmentGuid) = 0 And document.TemplatePackMember Then
                                        ' Document is not part of a pack and the document template is part of a pack definition.
                                        If errorIfDocumentNotInPack Then
                                            ' And this should be reported as an error.
                                            Err.Raise ERRORNUMBERBASE, MODULENAME + "." + cstrFunctionName, "Document is not part of a pack." & vbCrLf & refDocument
                                        Else
                                            ' Otherwise do not continue with Gemini Printing.
                                            success = False
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next
        End If
    End If
    On Error GoTo ExitHandler

    CheckOutputPacks = success

ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "()"
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Private Function GetPackRef(ByRef pack As GeminiPack) As String
    Dim refPack As String

    refPack = "Details: ApplicationNumber=" + pack.ApplicationNumber

    If Len(pack.PackControlName) > 0 Then
        refPack = refPack + ", PackName=" + pack.PackControlName
    End If
                                   
    If Len(pack.PackFulfillmentGuid) > 0 Then
        refPack = refPack + ", PackCreationDate=" + CStr(pack.PackCreationDate)
        refPack = refPack + ", PackFulfillmentGuid=" + pack.PackFulfillmentGuid
    End If
              
    GetPackRef = refPack

End Function

Private Function GetDocumentRef(ByRef document As GeminiDocument, ByVal refPack As String) As String

    Dim refDocument As String
    refDocument = refPack
        
    If Len(document.DocumentVersion) > 0 Then
        refDocument = refDocument & ", DocumentVersion=" & document.DocumentVersion
    End If
    
    If Len(document.DocumentName) > 0 Then
        refDocument = refDocument & ", DocumentName=" & document.DocumentName
    End If
    
    If Len(document.DocumentVersion) > 0 Then
        refDocument = refDocument & ", DocumentDate=" & CStr(document.DocumentDate)
    End If
                                
    If Len(document.DocumentGuid) > 0 Then
        refDocument = refDocument & ", DocumentGuid=" & document.DocumentGuid
    End If
    
    If Len(document.FileContentsGuid) > 0 Then
        refDocument = refDocument & ", FileContentsGuid=" & document.FileContentsGuid
    End If
    
    GetDocumentRef = refDocument
End Function

Private Function GeminiPrintModeToString(ByVal geminiPrintMode As Integer) As String
    Select Case geminiPrintMode
        Case GEMINIPRINTMODE_NEVER
            GeminiPrintModeToString = "Never"
        Case GEMINIPRINTMODE_IMMEDIATE
            GeminiPrintModeToString = "Immediate"
        Case GEMINIPRINTMODE_ONHOLD
            GeminiPrintModeToString = "On-hold"
        Case Else
            GeminiPrintModeToString = "Unknown"
    End Select
End Function

Private Function GeminiPrintStatusToString(ByVal geminiPrintStatus As Integer) As String
    Select Case geminiPrintStatus
        Case GEMINIPRINTSTATUS_AWAITINGAPPROVAL
            GeminiPrintStatusToString = "Awaiting approval"
        Case GEMINIPRINTSTATUS_NOTAPPROVED
            GeminiPrintStatusToString = "Not approved"
        Case GEMINIPRINTSTATUS_APPROVED
            GeminiPrintStatusToString = "Approved"
        Case GEMINIPRINTSTATUS_GEMINIPRINTED
            GeminiPrintStatusToString = "Gemini printed"
        Case Else
            GeminiPrintStatusToString = "Unknown"
    End Select
End Function


'------------------------------------------------------------------------------------------
' Summary:
'   Converts an array of GeminiPack items into a fulfillment request.
' Argument:
'   appNumber. The Omiga mortgage application number.
' Argument:
'   packs. The array of GeminiPack items to convert into the fulfillment request.
' Returns:
'   The fulfillment request, as an Xml string.
' Remarks
'   Each pack is a separate request. This enables each request to be put onto the
'   queue as a separate message. If one of the requests fails, it will not affect the others,
'   and it can be resubmitted to the queue.
'------------------------------------------------------------------------------------------
Private Function ToFulfillmentRequestXml(ByRef packs() As GeminiPack) As String()
    Const cstrFunctionName As String = "ToFulfillmentRequestXml"

    On Error GoTo ExitHandler
        
    Debug.Print "->" & cstrFunctionName & "()"
    gobjTrace.TraceMethodEntry MODULENAME, cstrFunctionName
        
    Dim xmlRequestDocument As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMNode
    Dim xmlPackList As IXMLDOMElement
    Dim xmlPack As IXMLDOMElement
    Dim xmlDocumentList As IXMLDOMElement
    Dim xmlDocument As IXMLDOMElement
    Dim xmlDocumentDetails As IXMLDOMElement
    
    ReDim requests(0 To UBound(packs)) As String
    
    Dim packIndex As Integer
    For packIndex = 0 To UBound(packs)
        Set xmlRequestDocument = New FreeThreadedDOMDocument40
    
        xmlRequestDocument.setProperty "NewParser", True
        xmlRequestDocument.async = False
    
        Set xmlRequest = xmlRequestDocument.createNode(NODE_ELEMENT, "FULFILLMENTREQUEST", m_fulfillmentRequestNamespace)
        xmlRequestDocument.appendChild xmlRequest

        Set xmlPackList = xmlRequestDocument.createNode(NODE_ELEMENT, "PACKLIST", m_fulfillmentRequestNamespace)
        xmlRequest.appendChild xmlPackList
    
        Set xmlPack = xmlRequestDocument.createNode(NODE_ELEMENT, "PACK", m_fulfillmentRequestNamespace)
        xmlPackList.appendChild xmlPack
        xmlPack.setAttribute "APPLICATIONNUMBER", packs(packIndex).ApplicationNumber
        
        Set xmlDocumentList = xmlRequestDocument.createNode(NODE_ELEMENT, "DOCUMENTLIST", m_fulfillmentRequestNamespace)
        xmlPack.appendChild xmlDocumentList
        
        Dim documentIndex As Integer
        For documentIndex = 0 To UBound(packs(packIndex).Documents)
            If packs(packIndex).Documents(documentIndex).PrintMode <> GEMINIPRINTMODE_NEVER Then
                Set xmlDocument = xmlRequestDocument.createNode(NODE_ELEMENT, "DOCUMENT", m_fulfillmentRequestNamespace)
                xmlDocumentList.appendChild xmlDocument
                xmlDocument.setAttribute "FILECONTENTSGUID", packs(packIndex).Documents(documentIndex).FileContentsGuid
                
                If Len(packs(packIndex).Documents(documentIndex).DocumentDetails.FileContents) > 0 Then
                    Set xmlDocumentDetails = xmlRequestDocument.createNode(NODE_ELEMENT, "DOCUMENTDETAILS", m_fulfillmentRequestNamespace)
                    xmlDocument.appendChild xmlDocumentDetails
                    xmlDocumentDetails.setAttribute "COMPRESSIONMETHOD", packs(packIndex).Documents(documentIndex).DocumentDetails.CompressionMethod
                    xmlDocumentDetails.setAttribute "DELIVERYTYPE", CStr(packs(packIndex).Documents(documentIndex).DocumentDetails.DeliveryType)
                    xmlDocumentDetails.setAttribute "FILECONTENTS_TYPE", packs(packIndex).Documents(documentIndex).DocumentDetails.FileContentsType
                    xmlDocumentDetails.setAttribute "FILECONTENTS", packs(packIndex).Documents(documentIndex).DocumentDetails.FileContents
                End If
            End If
        Next
        
        requests(packIndex) = xmlRequestDocument.xml
        
    Next

    ToFulfillmentRequestXml = requests
    
ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "()"
    gobjTrace.TraceMethodExit MODULENAME, cstrFunctionName
    
    Set xmlRequestDocument = Nothing
    Set xmlRequest = Nothing
    Set xmlPackList = Nothing
    Set xmlPack = Nothing
    Set xmlDocumentList = Nothing
    Set xmlDocument = Nothing
    Set xmlDocumentDetails = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

'------------------------------------------------------------------------------------------
' Summary:
'   Gets an array of GeminiPack items for a specified document or file contents.
' Argument:
'   pack. The input pack.
' Returns:
'   An array of packs.
' Remarks:
'   For the specified document or file contents, the fulfillment request will contain
'   all the associated documents:
'       If the document is part of a pack, include all documents in the pack.
'       If the document is in more than one pack, include all the packs, and all the
'       documents in these packs.
'       If any of the documents in any of the packs is not available for Gemini printing
'       (e.g., it does not exist in Mobius, or it is on-hold), then do not include any
'       of these packs.
'------------------------------------------------------------------------------------------
Private Function GetPacks(ByRef pack As GeminiPack) As GeminiPack()
    Const cstrFunctionName As String = "GetPacks"

    On Error GoTo ExitHandler
    
    Debug.Print "->" & cstrFunctionName & "()"
    gobjTrace.TraceMethodEntry MODULENAME, cstrFunctionName
    
    Dim adoConnection As New ADODB.Connection
    Dim adoCommand As New ADODB.Command
    Dim adoParameter As New ADODB.Parameter
    Dim adoRecordSet As Recordset
    
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    
    adoCommand.CommandType = adCmdStoredProc
    adoCommand.CommandText = "usp_GeminiGetPacks"
    
    adoCommand.Parameters.Append adoCommand.CreateParameter(, adBSTR, adParamInput, Len(pack.ApplicationNumber), pack.ApplicationNumber)
    
    Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
    adoCommand.Parameters.Append adoParameter
    adoParameter.Attributes = adParamNullable
    If Len(pack.PackFulfillmentGuid) <> 0 Then
        adoParameter.Value = GuidStringToByteArray(pack.PackFulfillmentGuid)
    Else
        adoParameter.Value = Null
    End If
    
    Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
    adoCommand.Parameters.Append adoParameter
    adoParameter.Attributes = adParamNullable
    adoParameter.Value = Null
    On Error Resume Next
    If UBound(pack.Documents) >= 0 Then
        If Err.Number = 0 Then
            On Error GoTo ExitHandler
            If Len(pack.Documents(0).DocumentGuid) <> 0 Then
                adoParameter.Value = GuidStringToByteArray(pack.Documents(0).DocumentGuid)
            End If
        End If
    End If
    On Error GoTo ExitHandler

    Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
    adoCommand.Parameters.Append adoParameter
    adoParameter.Attributes = adParamNullable
    adoParameter.Value = Null
    On Error Resume Next
    If UBound(pack.Documents) >= 0 Then
        If Err.Number = 0 Then
            On Error GoTo ExitHandler
            If Len(pack.Documents(0).FileContentsGuid) <> 0 Then
                adoParameter.Value = GuidStringToByteArray(pack.Documents(0).FileContentsGuid)
            End If
        End If
    End If
    On Error GoTo ExitHandler
        
    adoConnection.open
    adoCommand.ActiveConnection = adoConnection
    
    Set adoRecordSet = adoCommand.Execute

    Dim packs() As GeminiPack
    Dim packIndex As Integer
    packIndex = 0
    Dim documentIndex As Integer
    Dim continue As Boolean
    
    continue = True
    Do While continue And Not adoRecordSet Is Nothing
        If adoRecordSet.State = adStateOpen Then
            ReDim Preserve packs(0 To packIndex)
            packs(packIndex).ApplicationNumber = pack.ApplicationNumber
            packs(packIndex).UserId = pack.UserId
            packs(packIndex).UnitId = pack.UnitId
            
            documentIndex = 0
            
            Do While Not adoRecordSet.EOF
                ReDim Preserve packs(packIndex).Documents(0 To documentIndex)
                If adoRecordSet.fields.Count >= 14 Then
                    Dim fieldIndex As Integer
                    
                    fieldIndex = 0
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).PackFulfillmentGuid = adoGuidToString(adoRecordSet.fields.Item(fieldIndex).Value)
                    End If
                    
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).PackControlName = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                    
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).PackCreationDate = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                    
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).DocumentGuid = adoGuidToString(adoRecordSet.fields.Item(fieldIndex).Value)
                    End If
                    
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).FileGuid = adoGuidToString(adoRecordSet.fields.Item(fieldIndex).Value)
                    End If
                    
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).FileContentsGuid = adoGuidToString(adoRecordSet.fields.Item(fieldIndex).Value)
                    End If
                                       
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).DocumentVersion = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                        
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).DocumentDate = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                        
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).DocumentName = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                       
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).HostTemplateId = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                       
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).TemplatePackMember = adoRecordSet.fields.Item(fieldIndex).Value
                    End If

                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).PrintStatus = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                       
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).PrintMode = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                       
                    fieldIndex = fieldIndex + 1
                    If Not IsNull(adoRecordSet.fields.Item(fieldIndex).Value) Then
                        packs(packIndex).Documents(documentIndex).DocumentLocation = adoRecordSet.fields.Item(fieldIndex).Value
                    End If
                                       
                    If Len(packs(packIndex).Documents(documentIndex).FileContentsGuid) > 0 Then
                        Dim packDocumentIndex As Integer
                        On Error Resume Next
                        If UBound(pack.Documents) >= 0 Then
                            If Err.Number = 0 Then
                                On Error GoTo ExitHandler
                                For packDocumentIndex = 0 To UBound(pack.Documents)
                                    If _
                                        StrComp(packs(packIndex).Documents(documentIndex).DocumentGuid, pack.Documents(packDocumentIndex).DocumentGuid, vbTextCompare) = 0 Or _
                                        StrComp(packs(packIndex).Documents(documentIndex).FileContentsGuid, pack.Documents(packDocumentIndex).FileContentsGuid, vbTextCompare) = 0 Then
                                        ' This pack document is the same document that was passed in.
                                        
                                        packs(packIndex).Documents(documentIndex).InputDocument = True
                                        
                                        If Len(pack.Documents(packDocumentIndex).DocumentDetails.FileContents) > 0 Then
                                            ' Document details have been passed in so copy them in to the pack document.
                                            packs(packIndex).Documents(documentIndex).DocumentDetails = pack.Documents(packDocumentIndex).DocumentDetails
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        On Error GoTo ExitHandler
                    End If
                End If
                
                adoRecordSet.MoveNext
                documentIndex = documentIndex + 1
            Loop
            
            Set adoRecordSet = adoRecordSet.NextRecordset
            packIndex = packIndex + 1
        Else
            continue = False
        End If
    Loop
    If Not adoRecordSet Is Nothing Then
        If adoRecordSet.State = adStateOpen Then
            adoRecordSet.Close
        End If
    End If
    adoConnection.Close
    
    GetPacks = packs

ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "()"
    gobjTrace.TraceMethodExit MODULENAME, cstrFunctionName
    
    Set adoParameter = Nothing
    Set adoRecordSet = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Function GeminiSendMessagesToQueue(ByRef messages() As String, ByVal messageType As String) As String()
    Const cstrFunctionName As String = "GeminiSendMessagesToQueue"

    On Error GoTo ExitHandler
    
    ReDim responses(0 To UBound(messages)) As String
    
    Dim messageIndex As Integer
    For messageIndex = 0 To UBound(messages)
        responses(messageIndex) = GeminiSendMessageToQueue(messages(messageIndex), messageType, messageIndex)
    Next
    
    GeminiSendMessagesToQueue = responses
    
ExitHandler:

End Function

Public Function GeminiSendMessageToQueue(ByVal messageText As String, ByVal messageType As String, Optional ByVal messageIndex As Integer = 0) As String
    Const cstrFunctionName As String = "GeminiSendMessageToQueue"

    On Error GoTo ExitHandler

    Debug.Print "->" & cstrFunctionName & "(messageText=" & messageText & ")"
    gobjTrace.TraceXML messageText, MODULENAME & "_" & messageType & CStr(messageIndex) & "_request"
    gobjTrace.TraceMethodEntry MODULENAME, cstrFunctionName
    
    Dim messageQueueObject As IOmigaToMessageQueue
    
    Dim messageQueueType As Integer
    messageQueueType = GetGlobalParamAmount("MessageQueueType")
    
    Select Case messageQueueType
        Case 1 ' SQL Server MSMQ
            Set messageQueueObject = CreateObject("omToMSMQ.OmigaToMessageQueue")
        Case 2, 3 'SQL Server OMMQ Oracle OMMQ
            Set messageQueueObject = CreateObject("omToOMMQ.OmigaToMessageQueue")
        Case Else
            ' Error Message Queue type not supported.
            Err.Raise oeInvalidMessageQueueType, "Message Queue Type not supported, check global parameter MessageQueueType"
    End Select
           
    Dim request As String
    Dim response As String
    
    request = _
        "<REQUEST>" & _
            "<MESSAGEQUEUE>" & _
                "<QUEUENAME>" & m_geminiQueueName & "</QUEUENAME>" & _
                "<PROGID>" & m_geminiComponentName & "</PROGID>" & _
            "</MESSAGEQUEUE>" & _
        "</REQUEST>"
   
    Debug.Print "->messageQueueObject.AsyncSend(request=" & request & ", messageText=" & messageText & ")"
    gobjTrace.TraceMethodEntry MODULENAME, "messageQueueObject.AsyncSend", request
    response = messageQueueObject.AsyncSend(request, messageText)
    Debug.Print "<-messageQueueObject.AsyncSend() = " & response
    gobjTrace.TraceMethodExit MODULENAME, "messageQueueObject.AsyncSend", response
    
    errCheckXMLResponse response, True
    
    GeminiSendMessageToQueue = response
          
ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "() = " & response
    gobjTrace.TraceMethodExit MODULENAME, cstrFunctionName
    
    Set messageQueueObject = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Private Function CreateEventAuditDetails(ByVal objContext As ObjectContext, ByRef packs() As GeminiPack, eventKey As Integer, geminiPrintStatus As Integer) As Boolean
    Const cstrFunctionName As String = "CreateEventAuditDetails"

    On Error GoTo ExitHandler

#If Not TEST Then
    Debug.Print "->" & cstrFunctionName & "()"
    gobjTrace.TraceMethodEntry MODULENAME, cstrFunctionName
    
    Dim objPM As PrintManagerBO
    If objContext Is Nothing Then
        Set objPM = New PrintManagerBO
    Else
        Set objPM = objContext.CreateInstance(gstrPRINTMANAGER_COMPONENT & ".PrintManagerBO")
    End If
       
    Dim packIndex As Integer
    For packIndex = 0 To UBound(packs)
          
        Dim documentIndex As Integer
        For documentIndex = 0 To UBound(packs(packIndex).Documents)
            Dim request As String
            Dim response As String
            
            request = _
                "<REQUEST " & _
                    "APPLICATIONNUMBER='" & packs(packIndex).ApplicationNumber & "' " & _
                    "USERID='" & packs(packIndex).UserId & "' " & _
                    "UNITID='" & packs(packIndex).UnitId & "' " & _
                    "DOCUMENTGUID='" & packs(packIndex).Documents(documentIndex).DocumentGuid & "' " & _
                    "OPERATION='CREATEAUDITTRAIL'>" & _
                    "<EVENTDETAIL " & _
                        "EVENTKEY='" & CStr(eventKey) & "' " & _
                        "DOCUMENTVERSION='" & packs(packIndex).Documents(documentIndex).DocumentVersion & "' " & _
                        "FILEGUID='" & packs(packIndex).Documents(documentIndex).FileGuid & "' " & _
                        "PACKFULFILLMENTGUID='" & packs(packIndex).PackFulfillmentGuid & "' " & _
                        "HOSTTEMPLATEID='" & packs(packIndex).Documents(documentIndex).HostTemplateId & "'>" & _
                    "</EVENTDETAIL>" & _
                    "<DOCUMENTDETAILS " & _
                        "GEMINIPRINTSTATUS='" & CStr(geminiPrintStatus) & "'>" & _
                    "</DOCUMENTDETAILS>" & _
                "</REQUEST>"
            response = objPM.omRequest(request)
        Next
    Next
#End If

    CreateEventAuditDetails = True

ExitHandler:
#If Not TEST Then
    Debug.Print "<-" & cstrFunctionName & "()"
    gobjTrace.TraceMethodExit MODULENAME, cstrFunctionName
    Set objPM = Nothing
#End If

    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Private Function IsFileVersionLocked(ByRef document As GeminiDocument, ByRef strUserId As String) As Boolean

    Const cstrFunctionName As String = "IsFileVersionLocked"
    
    On Error GoTo ExitHandler

    Debug.Print "->" & cstrFunctionName & "()"
    
    Dim adoConnection As New ADODB.Connection
    Dim adoCommand As New ADODB.Command
    Dim adoParameter As New ADODB.Parameter
    Dim adoRecordSet As Recordset
    
    Dim blnIsLocked As Boolean
    blnIsLocked = False
    
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.CursorLocation = adUseClient
    
    adoCommand.CommandType = adCmdStoredProc
    adoCommand.CommandText = "usp_IsFvFileLocked"
    
    Set adoParameter = adoCommand.CreateParameter("@FileGuid", adBinary, adParamInput, 16)
    adoCommand.Parameters.Append adoParameter
    adoParameter.Attributes = adParamNullable
    adoParameter.Value = Null
    If Len(document.FileContentsGuid) <> 0 Then
            adoParameter.Value = GuidStringToByteArray(document.FileContentsGuid)
    End If

    Set adoParameter = adoCommand.CreateParameter("@FileVersion", adVarChar, adParamInput, 20)
    adoCommand.Parameters.Append adoParameter
    adoParameter.Attributes = adParamNullable
    adoParameter.Value = Null
    If Len(document.DocumentVersion) <> 0 Then
            adoParameter.Value = document.DocumentVersion
    End If
    
    Set adoParameter = adoCommand.CreateParameter("@IsLocked", adInteger, adParamOutput)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@UserID", adVarChar, adParamOutput, 25)
    adoCommand.Parameters.Append adoParameter
    
    adoConnection.open
    adoCommand.ActiveConnection = adoConnection
    
    Set adoRecordSet = adoCommand.Execute
    
    If Not adoRecordSet Is Nothing Then
    
        blnIsLocked = IIf(IsNull(adoCommand.Parameters("@IsLocked").Value), 0, (adoCommand.Parameters("@IsLocked").Value))
        strUserId = IIf(IsNull(adoCommand.Parameters("@UserID").Value), "", (adoCommand.Parameters("@UserID").Value))
        
        If adoRecordSet.State = adStateOpen Then
            adoRecordSet.Close
        End If
    End If
    
    adoConnection.Close
    
    IsFileVersionLocked = blnIsLocked
    
ExitHandler:
    Debug.Print "<-" & cstrFunctionName & "()"
    
    Set adoParameter = Nothing
    Set adoRecordSet = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
