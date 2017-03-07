VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PrintHandlerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PrintHandlerBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'Description:
'   Handles print requests.
'
'Dependencies:  MessageQueueListener
'               MS Word
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransactions
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    02/11/00    Created.
' RF    05/01/01    Queue listening now handled by omPM.
' RF    29/01/01    Pass in entire request, not just the data xml.
' RF    22/03/01    Ensure no Winword.exe processes are left running when errors occur.
' LD    19/06/01    SYS2386 All projects to use guidassist.bas rather than guidassist.cls
' AD    06/06/01    Pass back the merged document to the calling component.
' AD    20/06/01    Include email and fax capability.
' AD    11/07/01    DoThePrint needs to handle both new documents and documents that
'                   have been previously merged and / or edited.
' DR    10/02/02    DMSSYS0005 DoThePrint() now has a simple, straight through print
'                   operation ONMESSAGE. I've commented out the ActiverPrinter code cos it
'                   hangs some servers. I've added an on error resume next when we try to
'                   quit Word cos it sometimes errors.
' DR    10/02/02    DMSSYS0005 I've removed the mandatory flag in IsAuthorityLevelSufficient()
' DR    13/02/02    Build process overwrote the the above changes- I've reapplied them. Included
'                   the FreeThreadedDomDocument40 changes that the build process was attempting to add.
' DR    13/02/02    DMSSYS0005 MAPI session logoff and logon error bug fixed.
' DR    12/03/02    SYS4278 Adding errors to the event log.
' DR    15/04/02    Adding debug info around email usage.
' SA    15/04/02    SYS4379 Check in DoThePrint incorrect
'------------------------------------------------------------------------------------------
' Base specific History:
'
' Prog  Date        Description
' HMA   26/04/04    CORE46 Add DeliveryEngineType, ACEXML location and ACEXSL location.
'                   Add code to allow for Merging/Printing KFIs
'------------------------------------------------------------------------------------------
' BBG specific History:
'
' Prog  Date        Description
' HMA   20/05/04    BBG573 Use validation type for DeliveryEngine combo.
' HMA   24/05/04    BBG573 Speed up loading of KFI response.
' HMA   08/06/04    BBG618 Check for existance of RTFFileName before attempting to remove it.
' HMA   10/06/04    BBG72  Allow for .doc KFI templates.
' HMA   16/06/04    BBG751 Get Delivery Type and Delivery Engine type correctly.
' MV    05/07/2004  BBG452 - Amended DoThePrint() - WP20 Handle URLPostInd for Online Offers
' MV    13/07/2004  BBG875 - Amended DoThePrint() - WP20 Handle URLPostInd for Online Offers
' TW    02/08/04    Tidy up
' PSc   06/10/2004  BBG1572 - New WEBDOCUMENTTYPE attribute added to return with web documents
' PSC   19/10/2004  BBG1664 - Amend to look for print tray attributes in control data as well as document details.
'                             Also amend GetPrintTrayId to cater for blank paper ids
' PSC   21/10/2004  BBG1638 - Replace omStream within ConvertBase64ToBin, ConvertBase64ToBinViewEdit and ConvertBinToBase64
'------------------------------------------------------------------------------------------
' Core specific History:
'
' Prog  Date        Description
' AS    26/05/2005  CORE138 Edit/view before print with local printing produces garbage on print out
' AS    27/05/2005  CORE139 Edit/view before print with local printing of RTF produces blank page
' AS    31/05/2005  CORE142 PDF document not printing; due to PDFSecurity setting?
' RF    16/02/2006  CORE245 Allow printer destination of "Document Store Only"
' GHun  21/03/2006  CORE257 Changed GetKFIDocument to get PDFFontMode from a GlobalParameter
' GHun  05/03/2007  CORE00000080 Changes for EmailFulfilment
'------------------------------------------------------------------------------------------

Option Explicit
Implements ObjectControl
Public m_objWordApp As Object
'Implements IMessageQueueComponentVC1    ' for message queue listening
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
'TW 11/5/2005
'Private m_objConvertAssist As ConvertAssist 'BBG875
'TW 11/5/2005 End
'DMSSYS0005
Private Const ERR_RPC_FAILURE = -2147023170
Private Const ERR_RSM_NOTAVAILABLE = 462
Private Const ERR_PRINTER_ERROR = 1120
Private Const DIR_TEMP_EMAIL_ATTACHMENTS = "c:\temp" 'DR If a value is not supplied to the email part of
                                                     'DoThePrint() then the method will default to dumping
                                                     'attachment files into this directory.
Private Const cstrClassName = "omDPS::PrintHandlerBO"
Private Const EMAIL_PROFILE = "DMS2" 'DR This is the name of the profile from which DMS will access the email
                                     'sub-system. If you need to update email account details then do that on
                                     'the server under this profile name.
' AD 28/06/01 - as the strings for some of the GetMandatoryNode / SelectSingleNode
' calls are hard coded, this causes a problem as the dtd for the ControlData has
' changed but backwards compatibility needs to be maintained. The pattern matching is
' used more than relying on a hard path.
'------------------------------------------------------------------------------------------
' BMIDS specific History:
'
' Prog  Date        Description
' IK    28/11/2002  BMIDS00606, save / audit prior to document production
'                   operation PROCESSREQUEST no longer prints
' IK    17/02/2003  BM0200 - add TraceAssist support (remove most of OOSS inheritance)
' IK    20/02/2003  BM0231 - enhance error handling for omStream Base64ToObject calls


Private Type OUTPUT_DOCUMENT_DATA_PRINTER
    intCopies As Integer
    strPrinterName As String
    blnDuplexPrinting As Boolean
    intFirstPageTray As Integer
    intOtherPageTray As Integer
End Type

Private Type OUTPUT_DOCUMENT_DATA_FILE
    strOutputDirectory As String
    strFileID As String
    strHostTemplateName As String
End Type

Private Type OUTPUT_DOCUMENT_DATA_EMAIL
    strUserId As String
    strAdministrator As String
    strFileID As String
    strAddresses(255) As String
    lngAddresses As Long
    strHostTemplateName As String
    strHostTemplateDescription As String
    strOutputDirectory As String
End Type

Private Type OUTPUT_DOCUMENT_DATA_FAX
    strDestination As String
End Type

'CORE00000080 GHun
Private Type OUTPUT_DOCUMENT_DATA_EMAIL_FULFILMENT
    lngAddresses        As Long
    strAddresses(5)     As String
    strDocumentLocation As String
End Type
'CORE00000080 End

Private Type OUTPUT_DOCUMENT_DATA
    strOutputType As String
    
    blnPrinter As Boolean
    blnFile As Boolean
    blnEmail As Boolean
    blnFax As Boolean
    blnDMSOnly As Boolean ' RF 16/02/2006 CORE245
    blnEmailFulfilment As Boolean   'CORE00000080 GHun
    
    strCompressionMethod As String
    
    Printer As OUTPUT_DOCUMENT_DATA_PRINTER
    File As OUTPUT_DOCUMENT_DATA_FILE
    Email As OUTPUT_DOCUMENT_DATA_EMAIL
    Fax As OUTPUT_DOCUMENT_DATA_FAX
    EmailFulfilment As OUTPUT_DOCUMENT_DATA_EMAIL_FULFILMENT    'CORE00000080 GHun

End Type

Private Sub DealWithPrinterSetup(strPrinterName As String, blnDuplexPrinting As Boolean)
Const cstrFunctionName = "DealWithPrinterSetup"

Dim StrErrorEventMessage As String

'TW 3/8/2004 Code moved here to aid maintenance

    If strPrinterName = "" Then
        Exit Sub
    End If

    ' trap error 5216 - invalid printer
    On Error Resume Next
    Err.Clear
            
    'DR See Q216026 for more info on the following code:
    m_objWordApp.WordBasic.FilePrintSetup Printer:=strPrinterName, DoNotSetAsSysDefault:=1
            
    If (Err.Number = ERR_RPC_FAILURE Or Err.Number = ERR_RSM_NOTAVAILABLE) Then
        'If we get here then the WinWord process hung and was manually killed off.
        StrErrorEventMessage = "omDPS.PrintHandlerBO.DealWithPrinterSetup() Failed to set the FilePrintSetup - the WINWORD.EXE process needed to be killed before the process could continue."
        App.LogEvent StrErrorEventMessage, vbLogEventTypeError
    End If
    m_objErrAssist.SaveErr
    
    Select Case m_objErrAssist.LastErr
        Case 0
            If blnDuplexPrinting Then
                ' fixme
                ' VbPRDPSimplex     1 Single-sided printing with the current orientation setting.
                ' VbPRDPHorizontal  2 Double-sided printing using a horizontal page turn.
                ' VbPRDPVertical    3 Double-sided printing using a vertical page turn.
                'Printer.Duplex = vbPRDPVertical
            End If
        Case 5216, ERR_RPC_FAILURE, ERR_RSM_NOTAVAILABLE, ERR_PRINTER_ERROR
            ' "Invalid printer"
            errThrowError cstrFunctionName, 4906, "Printer name: '" & strPrinterName & "'"
    End Select
End Sub

Public Function GetCodeModule(ByRef rstrComponentName As String) As Object ' "As CodeModule" gave type mismatch error
' header ----------------------------------------------------------------------------------
' description:
'   get code module
'-------------------------------------------------------------------------------------------
On Error GoTo GetCodeModuleErr
    Const cstrFunctionName = "GetCodeModule"
   ' IK_BM0200 traceAssist calls
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim objComp As Object ' "As VBComponent" gave type mismatch error
    Dim objProject As Object ' "As VBProject" gave type mismatch error
    Set objProject = m_objWordApp.Application.ActiveDocument.VBProject
    If Not objProject Is Nothing Then
        Set m_objWordApp.Application.VBE.ActiveVBProject = objProject
    End If
    With m_objWordApp.Application.VBE.ActiveVBProject.VBComponents
        Dim i As Integer
        For i = 1 To .Count Step 1
            If .Item(i).Name = rstrComponentName Then
                Set objComp = .Item(i)
                Exit For
            End If
        Next i
    End With
    If Not objComp Is Nothing Then
        Set GetCodeModule = objComp.CodeModule
    Else
        Set GetCodeModule = Nothing
    End If
    ' IK_BM0200 traceAssist calls
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
GetCodeModuleErr:
    Set GetCodeModule = Nothing
    ' IK_BM0200 traceAssist calls
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Err.Raise Err.Number, cstrFunctionName, Err.Description
End Function

Public Sub DeleteAllCodeModules()

Const cstrFunctionName = "DeleteAllCodeModules"

'TW 11/5/2005
Dim VBComp As Object
Dim VBComps As Object
'TW 11/5/2005 End
    
    On Error GoTo DeleteAllCodeModulesErr:
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName

'TW 11/5/2005
    Set VBComps = m_objWordApp.Application.ActiveDocument.VBProject.VBComponents
'TW 11/5/2005 End
    
    For Each VBComp In VBComps
        If VBComp.Name = "ThisDocument" Then
            VBComp.CodeModule.DeleteLines 1, VBComp.CodeModule.CountOfLines
        Else
            VBComps.Remove VBComp
        End If
    Next VBComp
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Sub
    
DeleteAllCodeModulesErr:
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Err.Raise Err.Number, cstrFunctionName, Err.Description
End Sub

Private Function GetDeliveryEngineType(xmlTemplateNode As IXMLDOMNode) As String
Dim iEngineType As Integer

    iEngineType = xmlGetNodeAsInteger(xmlTemplateNode, "//TEMPLATE/DELIVERYENGINETYPE")
    GetDeliveryEngineType = GetValidationTypeForValueID("DeliveryEngineType", Format$(iEngineType))

End Function

Private Function GetKFIDocument(xmlIn As FreeThreadedDOMDocument40, xmlTemplateNode As IXMLDOMNode, strOutputType As String, ByVal strCompressionMethod As String) As FreeThreadedDOMDocument40
    
Dim lngErrorNumber As Long

Dim objKFI As Object

Dim strAceXML As String
Dim strAceXSL As String
Dim strAlphaCalcs As String
Dim strErrorDescription As String
Dim strErrorSource As String
Dim strOutput As String
Dim strTemplate As String
Dim strTemplateData As String

Dim xmlKFIDocIn As FreeThreadedDOMDocument40
Dim xmlKFIDocOut As FreeThreadedDOMDocument40
Dim xmlKFINode As IXMLDOMNode
Dim xmlKFIElem As IXMLDOMElement
Dim xmlTemplateDataNode As IXMLDOMNode
Dim strPDFFontMode As String    'CORE257 GHun
    
    On Error GoTo VBErr:
    Set objKFI = m_objContext.CreateInstance("eKFI.eKFIControl")
    
    strAlphaCalcs = "no"
    
    strTemplate = m_objXmlAssist.GetNodeValue(xmlTemplateNode, "//TEMPLATE/FILENAME", True)
    
    Set xmlTemplateDataNode = xmlIn.selectSingleNode("//PRINTDATA/TEMPLATEDATA")
    If xmlTemplateDataNode Is Nothing Then
        ' Default to empty template data. This node is required by eKFI.
        strTemplateData = "<TEMPLATEDATA/>"
    Else
        strTemplateData = xmlTemplateDataNode.xml
    End If
    
    'CORE257 GHun
    strPDFFontMode = CStr(GetGlobalParamAmount("PDFFontMode"))
    If Len(strPDFFontMode) = 0 Then
        strPDFFontMode = "0"    'Default the value to zero if it is blank
    End If
    'CORE257 End

    Set xmlKFIDocIn = New FreeThreadedDOMDocument40
    xmlKFIDocIn.async = False
    Set xmlKFIElem = xmlKFIDocIn.createElement("REQUEST")
    xmlKFIElem.setAttribute "Template", strTemplate
    xmlKFIElem.setAttribute "Data", strTemplateData
    xmlKFIElem.setAttribute "Output", strOutput
    xmlKFIElem.setAttribute "AlphaCalcs", strAlphaCalcs
    xmlKFIElem.setAttribute "XSL", strAceXSL
    xmlKFIElem.setAttribute "XML", strAceXML
    xmlKFIElem.setAttribute "OutputType", strOutputType
    xmlKFIElem.setAttribute "DateFormat", "dd/mm/yyyy"
    xmlKFIElem.setAttribute "IntegerFormat", "#,##0"
    xmlKFIElem.setAttribute "NumericFormat", "#,##0.00"
    xmlKFIElem.setAttribute "PDFFontMode", strPDFFontMode   'CORE257 GHun
    xmlKFIElem.setAttribute "PDFSecurity", "0"
    xmlKFIElem.setAttribute "PDFUserPassword", ""
    xmlKFIElem.setAttribute "PDFOwnerPassword", ""
    If strCompressionMethod = "ZLIB" Then
        xmlKFIElem.setAttribute "Compression", "KFI"
    ElseIf strCompressionMethod = "COMPAPI" Then
        xmlKFIElem.setAttribute "Compression", "DMS"
    End If
    Set xmlKFINode = xmlKFIDocIn.appendChild(xmlKFIElem)
 
    Set xmlKFIDocOut = objKFI.ProcessKFI(xmlKFIDocIn.documentElement.xml, True)
    If xmlKFIDocOut Is Nothing Then
        errThrowError cstrClassName, 4901, "eKFI.ProcessKFI returned empty document - Template : " & strTemplate
    End If
        
    Set GetKFIDocument = xmlKFIDocOut

GetKFIDocumentExit:
    Set xmlKFIDocIn = Nothing
    Set xmlKFIDocOut = Nothing
    Set xmlKFINode = Nothing
    Set xmlKFIElem = Nothing
    Exit Function

VBErr:
    ' OK, should check the response here and set MQResp to the appropriate value.
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within GetKFIDocument" & vbCrLf & _
                "  Err.Number: " & lngErrorNumber & _
                ", Err.Source: " & strErrorSource & _
                ", Err.Description: " & strErrorDescription, _
                vbLogEventTypeError
    m_objErrAssist.AddToErrSource "GetKFIDocument"
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), "GetKFIDocument", Err.Number, Err.Description
    End If
    Set GetKFIDocument = Nothing
    ' go to clean-up section
    Resume GetKFIDocumentExit
End Function

Private Function GetOutputFileName(strOutputDirectory As String, strFileID As String, strHostTemplateName As String, strOutputType As String) As String
Dim strSeparator As String
' Output file name is made up of a directory specified in the xml control data,
' the template name, and a file id (eg an application number) specified in the xml control data

    If (strHostTemplateName = "") And (strFileID = "") Then
        strSeparator = ""
    Else
        strSeparator = "_"
    End If
            
    GetOutputFileName = Replace(strOutputDirectory & "\" & strFileID & strSeparator & strHostTemplateName & "." & strOutputType, "\\", "\")
'    strFileId = "" 'Needed to prevent odd after-effects ??????????????????
End Function

'------------------------------------------------------------------------------------------
Public Function ProcessRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
'       XML request data stream.
'       Format:
'       <REQUEST>
'           <CONTROLDATA>
'               <TEMPLATEID>
'               <AUTHORITYLEVEL>
'               <COPIES>
'               <PRINTERLOCATION>
'           </CONTROLDATA>
'           <PRINTDATA>
'               ... (user-dependant format)
'           </PRINTDATA>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------

Const cstrFunctionName As String = "ProcessRequest"

Dim intFreeFormat As Integer
Dim intOutputType As Integer

Dim lngErrorNumber As Long


'TW 11/5/2005
'Dim objTemplateHandlerDO  As ITemplateHandlerDO
'TW 11/5/2005 End

Dim strErrorDescription As String
Dim strErrorSource As String
Dim strOperation As String
Dim strOutputType As String
Dim strResponseType As String
Dim strTemplateId As String

Dim xmlDocumentDetailNode As IXMLDOMNode
Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlKFIElem As IXMLDOMElement
Dim xmlKFINode As IXMLDOMNode
Dim xmlKFIResponseDoc As FreeThreadedDOMDocument40
Dim xmlNodeRequest As IXMLDOMNode
Dim xmlNodeTemp As IXMLDOMNode
Dim xmlNodeTemplate As IXMLDOMNode
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTemplateNode As IXMLDOMNode
Dim xmlTempResponseNode As IXMLDOMNode

    On Error GoTo ProcessRequestVbErr
    
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' ik_20030210
    m_objContext.SetComplete
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
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
    xmlIn.save sPath & "\omDPSRequest_" & tname & ".xml"
    Set fso = Nothing
#End If
    'Debug.Print vstrXMLRequest
    
    'TW Get DELIVERYTYPE from wherever it is in the xml
    Set xmlDocumentDetailNode = xmlIn.selectSingleNode("//@DELIVERYTYPE")
    If Not xmlDocumentDetailNode Is Nothing Then
        intOutputType = Val(xmlDocumentDetailNode.Text)
        strOutputType = GetValidationTypeForValueID("DocumentDeliveryType", intOutputType)
    End If
    
    If Trim$(strOutputType) = "" Then
        strOutputType = "doc"   'Default to "doc"
    End If
           
    '------------------------------------------------------------------------------------------
    ' get template details from database
    '------------------------------------------------------------------------------------------
    intFreeFormat = CSafeInt( _
        m_objXmlAssist.GetNodeValue(xmlIn, _
            "CONTROLDATA/FREEFORMAT", False))
    strOperation = UCase$(xmlIn.documentElement.Attributes.getNamedItem("OPERATION").Text)
    
'TW 11/5/2005
'    Set objTemplateHandlerDO = m_objContext.CreateInstance(App.Title & ".TemplateHandlerDO")
'TW 11/5/2005 End
    Set xmlNodeRequest = m_objXmlAssist.GetRequestNodeEx(xmlIn.documentElement)
    Set xmlNodeTemp = xmlNodeRequest.appendChild(xmlIn.createElement("TEMPLATE"))
    
    If (strOperation = "PRINTDOCUMENT") And (intFreeFormat = 0) Then
        
        'BBG751
        'Find Delivery Type otherwise default to Standard Omiga (1)
        'TemplateID is in CONTROLDATA/TEMPLATEID node or is an attribute on CONTROLDATA node
        
        Set xmlNodeTemplate = xmlNodeTemp.appendChild(xmlIn.createElement("TEMPLATEID"))
        
        strTemplateId = m_objXmlAssist.GetNodeValue(xmlIn.documentElement, "CONTROLDATA/TEMPLATEID", False, False)

        If strTemplateId = "" Then
            strTemplateId = m_objXmlAssist.GetAttributeValue(xmlIn.documentElement, "CONTROLDATA", "TEMPLATEID")
        End If
        
        If strTemplateId <> "" Then
        
            xmlNodeTemplate.Text = strTemplateId
'TW 11/5/2005
'            Set xmlTemplateNode = objTemplateHandlerDO.GetTemplate(xmlNodeRequest)
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "objDDBXML.GetTemplateXML call"
            Set xmlTemplateNode = dbxmlGetTemplateXML(strTemplateId)
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "objDDBXML.GetTemplateXML return"
'TW 11/5/2005 End
        
            'Check for eKFI processing
            If (GetDeliveryEngineType(xmlTemplateNode) = "eKFI") Then
                xmlIn.documentElement.setAttribute "OPERATION", "PRINTKFI"
            End If
        End If
                
        Set xmlTemplateNode = Nothing
    
    Else
        Set xmlNodeTemplate = m_objXmlAssist.GetMandatoryNode(xmlIn.documentElement, "CONTROLDATA/TEMPLATEID")
            
        If xmlNodeTemplate Is Nothing Then
            ' AD - should never get here if the xml is one of the above. Raise error
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 502
        End If
        xmlNodeTemp.appendChild xmlNodeTemplate
        
'TW 11/5/2005
'        Set xmlTemplateNode = objTemplateHandlerDO.GetTemplate(xmlNodeRequest)
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.3]", "objDDBXML.GetTemplateXML call"
        Set xmlTemplateNode = dbxmlGetTemplateXML(xmlNodeTemplate.Text)
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.4]", "objDDBXML.GetTemplateXML return"
'TW 11/5/2005 End
        
        '------------------------------------------------------------------------------------------
        ' check security level
        '------------------------------------------------------------------------------------------
        If IsAuthorityLevelSufficient(xmlIn.documentElement, xmlTemplateNode) = True Then
        
            'CORE46  Allow for different print engines
   
            If (GetDeliveryEngineType(xmlTemplateNode) = "eKFI") Then
                'CORE 46 Add processing for eKFI processing.
                'Delivery Engine eKFI
               
                If (strOperation = "PROCESSREQUEST") Then
                    xmlIn.documentElement.setAttribute "OPERATION", "PROCESSKFI"
                ElseIf (strOperation = "PRINTDOCUMENT") Then
                    xmlIn.documentElement.setAttribute "OPERATION", "PRINTKFI"
                End If
                   
                Dim strCompressionMethod As String
                strCompressionMethod = m_objXmlAssist.GetAttributeValue(xmlIn.documentElement, "CONTROLDATA", "COMPRESSIONMETHOD")
                Set xmlKFIResponseDoc = GetKFIDocument(xmlIn, xmlTemplateNode, strOutputType, strCompressionMethod)
                    
                strResponseType = m_objXmlAssist.GetAttributeValue(xmlKFIResponseDoc.documentElement, "RESPONSE", "TYPE")
                
                If strResponseType = "ERROR" Then
                    Set xmlTempResponseNode = xmlKFIResponseDoc.selectSingleNode("RESPONSE")
                Else
                    'Append the Document Contents response from ProcessKFI to the request to DoThePrint
                    
                    'The Request already contains a PRINTDATA node (from KFIHelper).
                    'Replace this with the new data.
                    Set xmlKFINode = xmlIn.selectSingleNode("//PRINTDATA")
                    xmlIn.documentElement.removeChild xmlKFINode
                    
                    Set xmlKFIElem = xmlIn.createElement("PRINTDATA")
                    
                    xmlKFIElem.setAttribute "PRINTDOCUMENT", xmlKFIResponseDoc.selectSingleNode("//DOCUMENTCONTENTS/@FILECONTENTS").Text
                    xmlIn.documentElement.appendChild xmlKFIElem
                
                End If
            End If
         
        Else
            '------------------------------------------------------------------------------------------
            ' failed security check
            '------------------------------------------------------------------------------------------
            ' "Insufficient authority to print selected template"
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4900
        End If
    End If
    
    If strResponseType <> "ERROR" Then
        '------------------------------------------------------------------------------------------
        ' do the print
        '------------------------------------------------------------------------------------------
        
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "DoThePrint call"
        Set xmlTempResponseNode = DoThePrint(xmlIn.documentElement, xmlTemplateNode, strOutputType)   'BBG72
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "DoThePrint return"
    End If
    
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    ' AD 06/06/01 - DoPrint returns the filestream as well as the standard O4 success response
        
    ProcessRequest = xmlTempResponseNode.xml
    
ProcessRequestExit:
#If IK_DEBUG Then
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.loadXML ProcessRequest
    xmlResponseDoc.save sPath & "\omDPSResponse_" & tname & ".xml"
    Set xmlResponseDoc = Nothing
#End If
    
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKFIElem = Nothing            'CORE46
    Set xmlKFINode = Nothing            'CORE46
    Set xmlKFIResponseDoc = Nothing     'CORE46
    Set xmlDocumentDetailNode = Nothing 'CORE46
    
    gobjTrace.TraceResponse ProcessRequest
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    'DR Raise an error if we've processed one.
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        ' App.LogEvent "Event6 encountered within " & CstrClassName & "::" & cstrFunctionName
        ' App.LogEvent "Event6.5 within " & CstrClassName & "::" & cstrFunctionName & ". About to raise error: " & strErrorDescription
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
        ' App.LogEvent "Event7 within " & CstrClassName & "::" & cstrFunctionName
    End If
    Exit Function
ProcessRequestVbErr:
    ' OK, should check the response here and set MQResp to the appropriate value.
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & lngErrorNumber & ", Err.Source: " & strErrorSource & ", Err.Description: " & strErrorDescription, vbLogEventTypeError
    ' App.LogEvent "Event1 encountered within " & CstrClassName & "::" & cstrFunctionName
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    ' App.LogEvent "Event2 encountered within " & CstrClassName & "::" & cstrFunctionName
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, lngErrorNumber, strErrorDescription
    End If
    ' App.LogEvent "Event3 encountered within " & CstrClassName & "::" & cstrFunctionName
    ProcessRequest = m_objErrAssist.CreateErrorResponse
    
    gobjTrace.TraceIdentErrorResponse ProcessRequest
    
    Resume ProcessRequestExit
End Function

Private Function IsAuthorityLevelSufficient( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal vxmlTemplateNode As IXMLDOMNode) _
    As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   check security level
' pass:
' return:
'------------------------------------------------------------------------------------------
    
'TW Simplified 2/8/2004
    
Const cstrFunctionName As String = "IsAuthorityLevelSufficient"
   
Dim intTemplateAuthorityLevel As Integer
Dim intUserAuthLevel As Integer

    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
'Get security level of user
    intUserAuthLevel = CSafeInt( _
        m_objXmlAssist.GetNodeValue(vxmlRequest, _
            "CONTROLDATA/AUTHORITYLEVEL", True))
    
'Get security level of Template (if any)
    intTemplateAuthorityLevel = CSafeInt( _
        m_objXmlAssist.GetNodeValue(vxmlTemplateNode, _
            "//TEMPLATE/SECURITYLEVEL", False))
    
    IsAuthorityLevelSufficient = (intUserAuthLevel >= intTemplateAuthorityLevel)
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Function
Private Sub CreateNewDirectory(DirName As String)
    
Const cstrFunctionName As String = "CreateNewDirectory"
    
Dim arrPathParts
Dim fs As New FileSystemObject
Dim strPath As String
Dim X As Integer
   
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    If Not fs.FolderExists(DirName) Then
        arrPathParts = Split(DirName, "\")
        strPath = arrPathParts(0)
        For X = 1 To UBound(arrPathParts, 1)
            strPath = strPath & "\" & arrPathParts(X)
            If Not fs.FolderExists(strPath) Then
                fs.CreateFolder strPath
            End If
        Next X
        App.LogEvent "Email attachments directory: " & DirName, vbLogEventTypeInformation
    End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
End Sub

Private Function DoThePrint( _
    ByVal vxmlRequest As IXMLDOMElement, _
    Optional ByVal vxmlTemplateNode As IXMLDOMNode = Nothing, _
    Optional ByVal strOutputType As String = "doc") As IXMLDOMNode    'BBG72
' header ----------------------------------------------------------------------------------
' description:
'   print template with the data in the request
' pass:
' return:
'   xml response
'TW 23/7/2004 Tidied up
'AS 08/02/2005 Refactored and fixed to support outputting RTFs via KFI engine
'------------------------------------------------------------------------------------------
    Const cstrFunctionName = "DoThePrint"
    
    Dim blnKFI As Boolean
    Dim blnMerge As Boolean
    Dim blnPrint As Boolean
    Dim blnPostToWeb As Boolean
    
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strErrorSource As String
    
    Dim strOperation As String
    
    ' Objects
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlResponse As FreeThreadedDOMDocument40
           
    ' Local objects to be freed.
    Dim xmlDocumentDetailNode As IXMLDOMNode
    Dim xmlPageTrayDetails As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempOutputType As IXMLDOMNode
    Dim xmlTempOutputTypeNode As IXMLDOMNode
    Dim xmlConvertedFileContentsElement As IXMLDOMElement
    Dim xmlFileContentsNode As IXMLDOMNode

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
    vxmlRequest.ownerDocument.save sPath & "\omDPS_DoThePrint_Request_" & tname & ".xml"
    Set fso = Nothing
#End If
    
    On Error GoTo VBErr
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
           
    ' Create default response block
    Set xmlResponse = New FreeThreadedDOMDocument40
    Set xmlResponseElement = xmlResponse.createElement("RESPONSE")
    
    xmlResponse.appendChild xmlResponseElement
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
    ' PSC 24/08/2004 BBG1233
    ' PSC 19/10/2004 BBG1664
    ' Set xmlDocumentDetailNode = xmlGetMandatoryNode(vxmlRequest, "CONTROLDATA/DOCUMENTDETAILS")

    '-----------------------------------------------------------------------------------------
    ' Check to see what type of action we are doing
    '-----------------------------------------------------------------------------------------
    ' to maintain backwards compatibility, "ProcessRequest" will give rise to version 1
    ' behaviour. "NewDocument" will merge the xml with the template and print it. "EditDocument"
    ' will print the document but won't merge it. "PrintDocument" will print the document and
    ' not merge it.
    ' However, this behaviour can be modified if the "FreeFormat" option is set. In the case of
    ' a "NewDocument" the data is merged but no print is done and in the case of "PrintDocument",
    ' the document is printed with no merge taking place.
       
    ' check operation.....
    If vxmlRequest.nodeName = "REQUEST" Then
        strOperation = vxmlRequest.Attributes.getNamedItem("OPERATION").Text
    End If
    strOperation = UCase(strOperation)
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.0]", "OPERATION: " & strOperation
    
    Select Case strOperation
        Case "PROCESSREQUEST"                   'blnKFI = false, blnMerge = conditional, blnPrint = false
            ' check free format
            Dim intFreeFormat As Integer
            intFreeFormat = xmlGetNodeAsInteger(vxmlRequest, "CONTROLDATA/FREEFORMAT")
            blnMerge = (intFreeFormat = 0)
        Case "PRINTDOCUMENT", "EDITDOCUMENT"    'blnKFI = false, blnMerge = false, blnPrint = true
            blnPrint = True
        Case "MERGEDOCUMENT"                    'blnKFI = false, blnMerge = true,  blnPrint = false
            blnMerge = True
        Case "ONMESSAGE"                        'blnKFI = false, blnMerge = true,  blnPrint = true
            blnMerge = True
            blnPrint = True
        Case "PROCESSKFI"                       'blnKFI = true,  blnMerge = false, blnPrint = false
            blnKFI = True
        Case "PRINTKFI"                         'blnKFI = true,  blnMerge = false, blnPrint = true
            blnKFI = True
            blnPrint = True
        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, strOperation
    End Select
    
    Dim OutputDocumentData As OUTPUT_DOCUMENT_DATA
    
    OutputDocumentData.blnEmail = False
    OutputDocumentData.blnFax = False
    OutputDocumentData.blnFile = False
    OutputDocumentData.blnPrinter = False
    OutputDocumentData.blnDMSOnly = False ' CORE245
    OutputDocumentData.blnEmailFulfilment = False   'CORE00000080 GHun
    OutputDocumentData.strOutputType = strOutputType
    OutputDocumentData.strCompressionMethod = ""
    
    Dim xmlControlDataNode As IXMLDOMNode
    Set xmlControlDataNode = xmlGetNode(vxmlRequest, "CONTROLDATA")
    If Not xmlControlDataNode Is Nothing Then
        OutputDocumentData.strCompressionMethod = xmlGetAttributeText(xmlControlDataNode, "COMPRESSIONMETHOD")
        Set xmlControlDataNode = Nothing
    End If
    
    'The following code is only executed when there is data to be output (ie not in merge or PROCESSKFI)
    If blnPrint Then
        'Establish the output type(s) in the request
        Set xmlTempOutputType = vxmlRequest.selectSingleNode("CONTROLDATA/OUTPUTTYPE")
        For Each xmlTempNode In xmlTempOutputType.childNodes
            Select Case xmlTempNode.nodeName
                Case "FILE"
                    OutputDocumentData.blnFile = True
                    For Each xmlTempOutputTypeNode In xmlTempNode.childNodes
                        Select Case xmlTempOutputTypeNode.nodeName
                            Case "OUTPUTDIRECTORY"
                                OutputDocumentData.File.strOutputDirectory = xmlTempOutputTypeNode.Text
                            Case "FILEID"
                                OutputDocumentData.File.strFileID = xmlTempOutputTypeNode.Text
                        End Select
                    Next
                Case "PRINTER"
                    OutputDocumentData.blnPrinter = True
                    For Each xmlTempOutputTypeNode In xmlTempNode.childNodes
                        Select Case xmlTempOutputTypeNode.nodeName
                            Case "COPIES"       ' get number of copies
                                OutputDocumentData.Printer.intCopies = CSafeInt(xmlTempOutputTypeNode.Text)
                            Case "PRINTERNAME"  ' get printer location
                                OutputDocumentData.Printer.strPrinterName = xmlTempOutputTypeNode.Text
                            Case "DUPLEX"       ' duplex required?
                                OutputDocumentData.Printer.blnDuplexPrinting = CSafeBool(xmlTempOutputTypeNode.Text)
                        End Select
                    Next
                    Set xmlPageTrayDetails = xmlGetNode(vxmlRequest, "CONTROLDATA/DOCUMENTDETAILS")
                    If xmlPageTrayDetails Is Nothing Then
                        Set xmlPageTrayDetails = xmlGetNode(vxmlRequest, "CONTROLDATA")
                    End If
                    OutputDocumentData.Printer.intFirstPageTray = xmlGetAttributeAsInteger(xmlPageTrayDetails, "FIRSTPAGEPRINTERTRAY")
                    OutputDocumentData.Printer.intOtherPageTray = xmlGetAttributeAsInteger(xmlPageTrayDetails, "OTHERPAGESPRINTERTRAY")
                Case "EMAIL"
                    OutputDocumentData.blnEmail = True
                    OutputDocumentData.Email.strUserId = xmlGetMandatoryAttributeText(vxmlRequest.selectSingleNode("//REQUEST"), "USERID")
                    OutputDocumentData.Email.strAdministrator = xmlTempNode.Attributes.getNamedItem("EMAILADMINISTRATOR").Text
                    OutputDocumentData.Email.strFileID = xmlGetMandatoryAttributeText(xmlTempNode, "FILEID")
                    OutputDocumentData.Email.lngAddresses = 0
                    Dim xmlNodeList As IXMLDOMNodeList
                    Dim xmlNodeListNode As IXMLDOMNode
                    Set xmlNodeList = xmlTempOutputType.selectNodes("EMAIL")
                    For Each xmlNodeListNode In xmlNodeList
                        OutputDocumentData.Email.strAddresses(OutputDocumentData.Email.lngAddresses) = xmlGetAttributeText(xmlNodeListNode, "EMAILADDRESS")
                        OutputDocumentData.Email.lngAddresses = OutputDocumentData.Email.lngAddresses + 1
                    Next
                    Set xmlNodeList = Nothing
                    Set xmlNodeListNode = Nothing
                Case "FAX"
                    OutputDocumentData.blnFax = True
                    OutputDocumentData.Fax.strDestination = xmlTempNode.Text
                ' RF 16/02/2006 CORE245 Start
                Case "DMS"
                    OutputDocumentData.blnDMSOnly = True
                ' RF 16/02/2006 CORE245 End
                'CORE00000080 GHun
                Case "EMAILFULFILMENT"
                    OutputDocumentData.blnEmailFulfilment = True
                    With OutputDocumentData.EmailFulfilment
                        .lngAddresses = 0
                        For Each xmlNodeListNode In xmlTempOutputType.selectNodes("EMAILFULFILMENT/RECIPIENT")
                            .strAddresses(.lngAddresses) = xmlGetMandatoryAttributeText(xmlNodeListNode, "EMAILADDRESS")
                            .lngAddresses = .lngAddresses + 1
                        Next
                        .strDocumentLocation = xmlGetNodeText(xmlTempOutputType, "EMAILFULFILMENT/DOCUMENTLOCATION")
                    End With
                'CORE00000080 End
            End Select
        Next
        
        If OutputDocumentData.blnFile Or OutputDocumentData.blnEmail Then
            'Get the Host Template Name and Description
            Set xmlDocumentDetailNode = xmlGetMandatoryNode(vxmlRequest, "//REQUEST/CONTROLDATA/DOCUMENTDETAILS")
            OutputDocumentData.File.strHostTemplateName = xmlGetMandatoryAttributeText(xmlDocumentDetailNode, "DOCUMENTNAME")
            OutputDocumentData.Email.strHostTemplateName = OutputDocumentData.File.strHostTemplateName
            OutputDocumentData.Email.strHostTemplateDescription = xmlGetAttributeText(xmlDocumentDetailNode, "DOCUMENTDESCRIPTION")
        End If

        '------------------------------------------------------------------------------------------
        ' Post to URL?
        '------------------------------------------------------------------------------------------
        
        ' PSC 24/08/2004 BBG1233 - Start
        ' PSC 19/10/2004 BBG1664 - Start
        If xmlDocumentDetailNode Is Nothing Then
            Set xmlDocumentDetailNode = xmlGetNode(vxmlRequest, "CONTROLDATA/DOCUMENTDETAILS")
        End If
        
        If Not xmlDocumentDetailNode Is Nothing Then
            Dim blnURLPost As Boolean
            blnURLPost = xmlGetAttributeAsBoolean(xmlDocumentDetailNode, "URLPOSTIND")
            If blnURLPost Then
                blnPostToWeb = xmlGetAttributeAsBoolean(xmlDocumentDetailNode, "POSTTOWEB")
            End If
        End If
        ' PSC 19/10/2004 BBG1664 - End
        ' PSC 24/08/2004 BBG1233 - End
                
    End If
            
    Set xmlFileContentsNode = vxmlRequest.selectSingleNode("//PRINTDATA/@PRINTDOCUMENT")
    
    If blnMerge Then
        Dim strMergeTemplateFileName As String
        Debug.Assert Not vxmlTemplateNode Is Nothing
        strMergeTemplateFileName = xmlGetMandatoryNodeText(vxmlTemplateNode, "//TEMPLATE/FILENAME")
        If CreateMergedDocument(blnKFI, strMergeTemplateFileName, xmlFileContentsNode, vxmlRequest, xmlResponseElement) Then
            Set DoThePrint = xmlResponseElement
        End If
    Else
        If OutputDocument( _
            blnPrint, _
            blnKFI, _
            OutputDocumentData, _
            xmlFileContentsNode, _
            xmlResponseElement) Then
            
            Set DoThePrint = xmlResponseElement
                        
            If blnPostToWeb Then
                Set xmlConvertedFileContentsElement = xmlResponseElement.selectSingleNode("PRINTDOCUMENTDETAILS")
                Call PostToWeb(vxmlRequest, xmlConvertedFileContentsElement.getAttribute("PRINTDOCUMENT"))
            End If

        End If
    End If
    
ExitPoint:
    On Error Resume Next
    
    Set xmlFileContentsNode = Nothing
    Set xmlConvertedFileContentsElement = Nothing
    Set xmlDocumentDetailNode = Nothing
    Set xmlPageTrayDetails = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempOutputType = Nothing
    Set xmlTempOutputTypeNode = Nothing
           
    'DR Raise an error if we've processed one.
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
    
#If IK_DEBUG Then
    DoThePrint.ownerDocument.save sPath & "\omDPS_DoThePrint_Response_" & tname & ".xml"
#End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    Exit Function

VBErr:
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    Set DoThePrint = ReportError(lngErrorNumber, strErrorSource, strErrorDescription, cstrFunctionName, xmlResponseElement)
    Resume ExitPoint
End Function

Public Function PrintMergedDocument(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Called when outputting a document that has been viewed or edited in Axword. The
'   document template has already been merged with the document data, and may have been
'   edited in Axword, so here we only need to output (print etc) the document.
' pass:
'   vxmlRequest - XML request.
' return:
'   xml response
'AS 08/02/2005 Refactored and fixed to support outputting RTFs via KFI engine.
'------------------------------------------------------------------------------------------
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strErrorSource As String
    
    Dim strAppNumber As String
    Dim strPrinterName As String
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlPrintDocumentDataNode As IXMLDOMNode
    Dim xmlControlDataNode As IXMLDOMNode
    Dim xmlResponse As FreeThreadedDOMDocument40
    Dim xmlResponseElement As IXMLDOMElement
    Dim xmlFileContentsNode As IXMLDOMNode
    
    Const cstrFunctionName = "PrintMergedDocument"
    On Error GoTo VBErr
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    ' Create default response block.
    Set xmlResponse = New FreeThreadedDOMDocument40
    Set xmlResponseElement = xmlResponse.createElement("RESPONSE")
    xmlResponse.appendChild xmlResponseElement
    xmlResponseElement.setAttribute "TYPE", "SUCCESS"
    
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
    vxmlRequest.ownerDocument.save sPath & "\omDPS_PrintMergedDocument_Request" & tname & ".xml"
    Set fso = Nothing
#End If
    
    Dim blnPrint As Boolean
    Dim blnKFI As Boolean
    
    Dim OutputDocumentData As OUTPUT_DOCUMENT_DATA
    
    OutputDocumentData.blnEmail = False
    OutputDocumentData.blnFax = False
    OutputDocumentData.blnFile = False
    OutputDocumentData.blnPrinter = False
    OutputDocumentData.blnEmailFulfilment = False   'CORE00000080 GHun
    
    blnPrint = True
    blnKFI = False
    
    Set xmlControlDataNode = xmlGetMandatoryNode(vxmlRequest, "CONTROLDATA")
    If Not xmlControlDataNode Is Nothing Then
        Select Case xmlGetAttributeAsInteger(xmlControlDataNode, "DELIVERYTYPE")
        Case 10
            OutputDocumentData.strOutputType = "doc"
        Case 20
            OutputDocumentData.strOutputType = "pdf"
            blnKFI = True
        Case 30
            OutputDocumentData.strOutputType = "rtf"
            blnKFI = True
        End Select
        
        OutputDocumentData.strCompressionMethod = xmlGetAttributeText(xmlControlDataNode, "COMPRESSIONMETHOD")
        
        strPrinterName = xmlGetAttributeText(xmlControlDataNode, "PRINTER")
        Set xmlPrintDataNode = vxmlRequest.selectSingleNode("PRINTDATA")
        strAppNumber = xmlGetAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
        If Not xmlControlDataNode Is Nothing Then
            Select Case xmlGetAttributeText(xmlControlDataNode, "DESTINATIONTYPE")
                Case "L", "R"
                    OutputDocumentData.blnPrinter = True
                    OutputDocumentData.Printer.strPrinterName = strPrinterName
                    OutputDocumentData.Printer.intCopies = xmlGetAttributeAsInteger(xmlControlDataNode, "COPIES")
                    OutputDocumentData.Printer.intFirstPageTray = xmlGetAttributeAsInteger(xmlControlDataNode, "FIRSTPAGEPRINTERTRAY")
                    OutputDocumentData.Printer.intOtherPageTray = xmlGetAttributeAsInteger(xmlControlDataNode, "OTHERPAGESPRINTERTRAY")
                    OutputDocumentData.Printer.blnDuplexPrinting = xmlGetAttributeAsBoolean(xmlControlDataNode, "DUPLEX")
                Case "F"
                    If strAppNumber <> "" Then
                        OutputDocumentData.blnFile = True
                        OutputDocumentData.File.strFileID = strAppNumber
                        OutputDocumentData.File.strOutputDirectory = strPrinterName
                        OutputDocumentData.File.strHostTemplateName = xmlGetAttributeText(xmlControlDataNode, "HOSTTEMPLATENAME")
                    End If
                Case "E"
                    OutputDocumentData.blnEmail = True
                    OutputDocumentData.Email.strUserId = xmlGetMandatoryAttributeText(vxmlRequest.selectSingleNode("//REQUEST"), "USERID")
                    OutputDocumentData.Email.strAdministrator = xmlGetAttributeText(xmlControlDataNode, "EMAILADMINISTRATOR")
                    OutputDocumentData.Email.strHostTemplateName = xmlGetAttributeText(xmlControlDataNode, "HOSTTEMPLATENAME")
                    OutputDocumentData.Email.strHostTemplateDescription = xmlGetAttributeText(xmlControlDataNode, "HOSTTEMPLATEDESCRIPTION")
                    OutputDocumentData.Email.lngAddresses = 1
                    OutputDocumentData.Email.strAddresses(0) = strPrinterName
                    OutputDocumentData.Email.strFileID = strAppNumber 'Used in email subject
                'CORE00000080 GHun
                Case "EF"
                    OutputDocumentData.blnEmailFulfilment = True
                    With OutputDocumentData.EmailFulfilment
                        .lngAddresses = 0
                        Dim xmlNode As IXMLDOMNode
                        Dim xmlRecipients As IXMLDOMNode
                        Set xmlRecipients = GetEmailRecipient(xmlGetMandatoryAttributeText(xmlControlDataNode, "APPLICATIONNUMBER"), xmlGetMandatoryAttributeText(xmlControlDataNode, "EMAILRECIPIENTTYPE"))
                        For Each xmlNode In xmlRecipients.selectNodes("RECIPIENT")
                            .strAddresses(.lngAddresses) = xmlGetMandatoryAttributeText(xmlNode, "EMAILADDRESS")
                            .lngAddresses = .lngAddresses + 1
                        Next
                        'TODO get doc URL from somewhere ...
                    End With
                'CORE00000080 End
            End Select
        End If
                
        Set xmlFileContentsNode = vxmlRequest.selectSingleNode("//PRINTDOCUMENTDATA/@FILECONTENTS")
        
        If OutputDocument( _
                blnPrint, _
                blnKFI, _
                OutputDocumentData, _
                xmlFileContentsNode, _
                xmlResponseElement) Then
                   
            Set PrintMergedDocument = xmlResponseElement
        End If
    End If
                     
ExitPoint:
    On Error Resume Next
    
    Set xmlFileContentsNode = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlPrintDocumentDataNode = Nothing
    
    'DR Raise an error if we've processed one.
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
    
#If IK_DEBUG Then
    PrintMergedDocument.ownerDocument.save sPath & "\omDPS_PrintMergedDocument_Response_" & tname & ".xml"
#End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    Exit Function
    
VBErr:
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    Set PrintMergedDocument = ReportError(lngErrorNumber, strErrorSource, strErrorDescription, cstrFunctionName, xmlResponseElement)
    Resume ExitPoint
End Function

Private Function OutputDocument( _
    ByVal blnPrint As Boolean, _
    ByVal blnKFI As Boolean, _
    ByRef OutputDocumentData As OUTPUT_DOCUMENT_DATA, _
    ByRef xmlFileContentsNode As IXMLDOMNode, _
    ByRef vxmlResponse As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Outputs an already merged document.
' pass:
'   blnPrint    True of document should be outputted (printed etc). If False then the
'               document is converted but not outputted.
'   blnKFI      True if the document should be outputted using the KFI engine (RTFs and PDFs).
'   OutputDocumentData
'               Controls how the document is outputted.
'   xmlFileContentsNode
'               The document to be outputted, in binary base 64 format.
'   vxmlResponse
'               The response containing the document.
' return:
'   True if successful.
'AS 08/02/2005 New function.
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName = "OutputDocument"

    ' Local objects to be freed.
    Dim kfiPrintObject As New KFIPrintProcessor.KFIPrint
    Dim xmlConvertedFileContentsElement As IXMLDOMElement
    
    Dim strOutputFileName As String
    Dim arrByte() As Byte

    Dim intFirstPageTrayId As Integer
    Dim intOtherPageTrayId As Integer

    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strErrorSource As String
    
    On Error GoTo VBErr
    
    OutputDocument = False
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
            
    Set xmlConvertedFileContentsElement = vxmlResponse.ownerDocument.createElement("PRINTDOCUMENTDETAILS")
    xmlSetAttributeValue xmlConvertedFileContentsElement, "PRINTTIMESTAMP", Format(Now, "dd/mm/yyyy hh:mm:ss")
        
    'The following code is only executed when there is data to be output (ie not in merge or PROCESSKFI)
    If blnPrint Then
    
        'CORE00000080 GHun Added EmailFulfilment
        ' RF 16/02/2006 CORE245 Start - Allow printer destination of "Document Store Only"
        'If Not OutputDocumentData.blnPrinter And Not OutputDocumentData.blnFile And Not OutputDocumentData.blnEmail And Not OutputDocumentData.blnFax Then
        If Not OutputDocumentData.blnPrinter _
            And Not OutputDocumentData.blnFile _
            And Not OutputDocumentData.blnEmail _
            And Not OutputDocumentData.blnFax _
            And Not OutputDocumentData.blnDMSOnly _
            And Not OutputDocumentData.blnEmailFulfilment Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4903 ' No action requested
        End If
        
        If OutputDocumentData.blnFile Then
            ' Create the output directory if it doesn't exist.
            Call CreateNewDirectory(OutputDocumentData.File.strOutputDirectory)
            strOutputFileName = GetOutputFileName(OutputDocumentData.File.strOutputDirectory, OutputDocumentData.File.strFileID, OutputDocumentData.File.strHostTemplateName, OutputDocumentData.strOutputType)
        End If
        
        If OutputDocumentData.blnPrinter Then
            ' Default copies to 1.
            If OutputDocumentData.Printer.intCopies = 0 Then
                OutputDocumentData.Printer.intCopies = 1
            End If
        End If
                        
        If OutputDocumentData.blnEmail Then
            OutputDocumentData.blnPrinter = False
            If OutputDocumentData.Email.strOutputDirectory = "" Then
                OutputDocumentData.Email.strOutputDirectory = GetEmailAttachmentsDirectory(OutputDocumentData.Email.strUserId)
            End If
            strOutputFileName = GetOutputFileName(OutputDocumentData.Email.strOutputDirectory, OutputDocumentData.Email.strFileID, OutputDocumentData.Email.strHostTemplateName, OutputDocumentData.strOutputType)
        End If
        
        'CORE00000080 GHun
        If OutputDocumentData.blnEmailFulfilment Then
            OutputDocumentData.blnPrinter = False
        End If
        'CORE00000080 End
    End If
    
    If xmlFileContentsNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4905
    End If

    '------------------------------------------------------------------------------------------
    ' Create the document.
    '------------------------------------------------------------------------------------------
    If blnKFI Then
        'Document has already been created ProcessRequest when ProcessKFI was called.
        
        If OutputDocumentData.strCompressionMethod <> "" Then
            xmlFileContentsNode.dataType = "bin.base64"
            arrByte = xmlFileContentsNode.nodeTypedValue
            
            If Decompress(arrByte, OutputDocumentData.strCompressionMethod) Then
                xmlFileContentsNode.nodeTypedValue = arrByte
                xmlSetAttributeValue xmlConvertedFileContentsElement, "PRINTDOCUMENT", xmlFileContentsNode.Text
                OutputDocumentData.strCompressionMethod = ""
            End If
            
        Else
            arrByte = xmlFileContentsNode.nodeTypedValue
            xmlSetAttributeValue xmlConvertedFileContentsElement, "PRINTDOCUMENT", arrByte
        End If
        
    End If
     
    '------------------------------------------------------------------------------------------
    ' Output the document.
    '------------------------------------------------------------------------------------------
    If blnPrint Then
        ' Print the document using the KFI engine
        If OutputDocumentData.blnPrinter Then
            Dim strPassword As String
            intFirstPageTrayId = GetPrintTrayId(OutputDocumentData.Printer.strPrinterName, OutputDocumentData.Printer.intFirstPageTray)
            intOtherPageTrayId = GetPrintTrayId(OutputDocumentData.Printer.strPrinterName, OutputDocumentData.Printer.intOtherPageTray)
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[5.1]", "KFIPrintProcessor call"
            If Not kfiPrintObject.PrintKFI(OutputDocumentData.strOutputType, xmlFileContentsNode, OutputDocumentData.strCompressionMethod, OutputDocumentData.Printer.strPrinterName, strPassword, OutputDocumentData.Printer.intCopies, intFirstPageTrayId, intOtherPageTrayId) Then
                errThrowError cstrFunctionName, 4817, "Failed while in KFIPrintProcessor.KFIPrint"
            End If
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[5.1]", "KFIPrintProcessor return"
        ElseIf OutputDocumentData.strCompressionMethod <> "" Then
            ' File, email or Fax - must decompress document.
            xmlFileContentsNode.dataType = "bin.base64"
            arrByte = xmlFileContentsNode.nodeTypedValue
            If Decompress(arrByte, OutputDocumentData.strCompressionMethod) Then
                xmlFileContentsNode.nodeTypedValue = arrByte
                OutputDocumentData.strCompressionMethod = ""
            End If
        End If
        
        If OutputDocumentData.blnFile Or OutputDocumentData.blnEmail Then     'File or Email
            xmlFileContentsNode.dataType = "bin.base64"
            arrByte = xmlFileContentsNode.nodeTypedValue
            Call WriteToFile(strOutputFileName, arrByte(), OutputDocumentData.strOutputType = "rtf")
        End If
        If OutputDocumentData.blnFax Then
            'TW - Not currently supported
        End If
        
        If OutputDocumentData.blnEmail Then
            lngErrorNumber = Send_Email_By_CDO(OutputDocumentData.Email.strAdministrator, OutputDocumentData.Email.strAddresses, OutputDocumentData.Email.lngAddresses, OutputDocumentData.Email.strUserId & " - " & OutputDocumentData.Email.strFileID & " - " & OutputDocumentData.Email.strHostTemplateName, OutputDocumentData.Email.strHostTemplateDescription, strOutputFileName)
            If lngErrorNumber <> 0 Then
                Err.Raise lngErrorNumber, "Send_Email_By_CDO", Error$(lngErrorNumber) & " - Unable to send email"
            End If
        End If

        'CORE00000080 GHun
        If OutputDocumentData.blnEmailFulfilment Then
            xmlFileContentsNode.dataType = "bin.base64"
            EmailFulfilment xmlFileContentsNode.nodeTypedValue, OutputDocumentData.EmailFulfilment.strAddresses, OutputDocumentData.EmailFulfilment.lngAddresses, OutputDocumentData.EmailFulfilment.strDocumentLocation
        End If
        'CORE00000080 End
    End If
        
    vxmlResponse.appendChild xmlConvertedFileContentsElement
    
    OutputDocument = True
    
ExitPoint:
    On Error Resume Next
    Set kfiPrintObject = Nothing
    Set xmlFileContentsNode = Nothing
    Set xmlConvertedFileContentsElement = Nothing
       
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
        
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
    
VBErr:
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    Call ReportError(lngErrorNumber, strErrorSource, strErrorDescription, cstrClassName, cstrFunctionName, vxmlResponse)
    ' go to clean-up section
    Resume ExitPoint
End Function

Private Function Compress(ByRef arrData() As Byte, ByVal strCompressionMethod As String) As Boolean
    
    Compress = False
    
    If Len(strCompressionMethod) > 0 Then
        Dim objCompression As dmsCompression1
        Set objCompression = New DMSCOMPRESSIONLib.dmsCompression1
        
        If Not objCompression Is Nothing Then
            objCompression.CompressionMethod = strCompressionMethod
            arrData = objCompression.SafeArrayCompressToSafeArray(arrData)
            Set objCompression = Nothing
        End If
    End If
    
    Compress = True
    
End Function

Private Function Decompress(ByRef arrData() As Byte, ByVal strCompressionMethod As String) As Boolean
    Decompress = False
    
    If Len(strCompressionMethod) > 0 Then
        Dim objCompression As dmsCompression1
        Set objCompression = New DMSCOMPRESSIONLib.dmsCompression1
        
        If Not objCompression Is Nothing Then
            objCompression.CompressionMethod = strCompressionMethod
            arrData = objCompression.SafeArrayDecompressToSafeArray(arrData)
            Set objCompression = Nothing
        End If
    End If
    
    Decompress = True
End Function

Private Function CreateMergedDocument( _
    ByVal blnKFI As Boolean, _
    ByVal strMergeTemplateFileName As String, _
    ByRef xmlFileContentsNode As IXMLDOMNode, _
    ByRef vxmlRequest As IXMLDOMElement, _
    ByRef vxmlResponse As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Generates a merged document. This should be combined with MergeDocument, as they share
'   common functionality.
' pass:
'   blnKFI      True if the document should be created using the KFI engine (RTFs and PDFs).
'   strMergeTemplateFileName
'               Template file name; only relevant for Word documents.
'   xmlFileContentsNode
'               The already merged file contents; only relevant for RTF and PDFs, which
'               will already have been generated in ProcessRequest, ProcessKFI.
'   vxmlRequest
'               Request required for DPSMerge macro in Word documents.
'   vxmlResponse
'               The response containing the document.
' return:
'   True if successful.
'AS 08/02/2005 New function.
'------------------------------------------------------------------------------------------
    
    On Error GoTo VBErr
    Const cstrFunctionName = "CreateMergedDocument"
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    CreateMergedDocument = False

    ' Local objects to be freed.
    Dim WordDoc As Object
    Dim xmlConvertedFileContentsElement As IXMLDOMElement
    
    Dim strTempFileNameBinToBase64 As String
    Dim arrByte() As Byte

    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strErrorSource As String
                        
    Set xmlConvertedFileContentsElement = vxmlResponse.ownerDocument.createElement("PRINTDOCUMENTDETAILS")
    xmlSetAttributeValue xmlConvertedFileContentsElement, "PRINTTIMESTAMP", Format(Now, "dd/mm/yyyy hh:mm:ss")

    '------------------------------------------------------------------------------------------
    ' Create the document.
    '------------------------------------------------------------------------------------------
    If blnKFI Then
        ' For the KFI engine, the document contents will be in the input request,
        ' as the document has already been merged in ProcessRequest when ProcessKFI was called.
        If xmlFileContentsNode Is Nothing Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4905
        End If
        arrByte = xmlFileContentsNode.nodeTypedValue
        xmlSetAttributeValue xmlConvertedFileContentsElement, "PRINTDOCUMENT", arrByte
    Else
        Call SetUpNewWordObject(cstrFunctionName)
        
        On Error Resume Next
        Err.Clear
        
        ' Merge the document template with the data.
        Debug.Assert Not vxmlRequest Is Nothing
        Set WordDoc = m_objWordApp.Documents.open(strMergeTemplateFileName)
        Call RunDPSMerge(vxmlRequest, cstrFunctionName)
        
        Dim xmlControlDataNode As IXMLDOMNode
        Dim strCompressionMethod As String
        Set xmlControlDataNode = xmlGetNode(vxmlRequest, "CONTROLDATA")
        If Not xmlControlDataNode Is Nothing Then
            strCompressionMethod = xmlGetAttributeText(xmlControlDataNode, "COMPRESSIONMETHOD")
        End If
        strTempFileNameBinToBase64 = ConvertBinToBase64(WordDoc, xmlConvertedFileContentsElement, strCompressionMethod)
        Set xmlControlDataNode = Nothing
        
        On Error GoTo VBErr
        If m_objErrAssist.LastErr <> 0 Then
            If m_objErrAssist.LastErr = 5174 Then
                lngErrorNumber = 4902     'Template file not found.
            Else
                lngErrorNumber = 4901     'Data merge failed.
            End If
            errThrowError cstrFunctionName, lngErrorNumber, "Template name: " & strMergeTemplateFileName
        End If
    End If

    vxmlResponse.appendChild xmlConvertedFileContentsElement

    CreateMergedDocument = True
    
ExitPoint:
    On Error Resume Next
    If Not m_objWordApp Is Nothing Then
        ' Ensure the Word object is closed and unloaded
        WordDoc.Close SaveChanges:=False
        m_objWordApp.ActiveDocument.Close
        Err.Clear
        m_objWordApp.Quit
    End If
    If strTempFileNameBinToBase64 <> "" Then
        Kill strTempFileNameBinToBase64
    End If
        
    Set WordDoc = Nothing
'TW 11/5/2005    Set m_objWordApp = Nothing
    Set xmlConvertedFileContentsElement = Nothing
       
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
        
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
    
VBErr:
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    Call ReportError(lngErrorNumber, strErrorSource, strErrorDescription, cstrClassName, cstrFunctionName, vxmlResponse)
    ' go to clean-up section
    Resume ExitPoint
End Function

'BBG452 Post To Web
Private Function PostToWeb(ByVal vxmlRequest As IXMLDOMElement, ByVal strPrintDocument As String) As Boolean
    ' Build request to be packaged
    
    Dim objMessageQueue As IomMessageQueue
    Dim objXMLFactory As FreeThreadedDOMDocument40
    Dim xmlDocumentDetailNode As IXMLDOMNode
    Dim xmlMessageQueue As IXMLDOMElement
    Dim xmlPackagedRequest As IXMLDOMElement
    Dim xmlSendQueueResponse As IXMLDOMNode
    Dim xmlSendToQueueRequest As IXMLDOMElement
    
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strErrorSource As String
    
    Const cstrFunctionName = "PostToWeb"
    
    On Error GoTo VBErr
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    PostToWeb = False
    
    Set objXMLFactory = New FreeThreadedDOMDocument40
    Set xmlPackagedRequest = objXMLFactory.createElement("REQUEST")
    xmlPackagedRequest.setAttribute "OPERATION", "PostDataToWeb"
    xmlPackagedRequest.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(vxmlRequest, "APPLICATIONNUMBER")
    
    ' PSC 06/10/2004 BBG1572
    Set xmlDocumentDetailNode = xmlGetNode(vxmlRequest, "CONTROLDATA/DOCUMENTDETAILS")
    xmlCopyAttributeValue xmlDocumentDetailNode, xmlPackagedRequest, "WEBDOCUMENTTYPE", "DOCUMENTTYPE"
    
    xmlPackagedRequest.setAttribute "DOCUMENT", strPrintDocument
    
    ' Build Package, wrapping request
    Set xmlSendToQueueRequest = objXMLFactory.createElement("REQUEST")
    Set xmlMessageQueue = objXMLFactory.createElement("MESSAGEQUEUE")
    Set xmlMessageQueue = xmlSendToQueueRequest.appendChild(xmlMessageQueue)
    xmlMessageQueue.setAttribute "QUEUENAME", GetGlobalParamString("WebPrintQueueName")
    xmlMessageQueue.setAttribute "PROGID", GetGlobalParamString("WebPrintProgId")
    xmlMessageQueue.setAttribute "XML", xmlPackagedRequest.xml
    
    'Instantiate Queue Handler
    If m_objContext Is Nothing Then
        Set objMessageQueue = New omMessageQueue
    Else
        Set objMessageQueue = m_objContext.CreateInstance("omDPS.omMessageQueue")
    End If
    
    'Send Package To Queue
    Set xmlSendQueueResponse = objMessageQueue.SendToQueue(xmlSendToQueueRequest)
    
    'Check Response
    errCheckXMLResponseNode xmlSendQueueResponse, , True

    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName

    PostToWeb = True
    
PostToWebExit:
    Set objMessageQueue = Nothing
    Set objXMLFactory = Nothing
    Set xmlDocumentDetailNode = Nothing
    Set xmlMessageQueue = Nothing
    Set xmlPackagedRequest = Nothing
    Set xmlSendQueueResponse = Nothing
    Set xmlSendToQueueRequest = Nothing
    
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
    
    Exit Function

VBErr:
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    Call ReportError(lngErrorNumber, strErrorSource, strErrorDescription, cstrClassName, cstrFunctionName)
    Resume PostToWebExit

End Function

Private Function ReportError( _
    ByVal lngErrorNumber As Long, _
    ByVal strErrorSource As String, _
    ByVal strErrorDescription As String, _
    ByVal strClassName As String, _
    ByVal strFunctionName As String, _
    Optional ByVal xmlResponse As IXMLDOMElement = Nothing) As IXMLDOMNode

    Dim blnIsWarning As Boolean
    Dim blnIsSystemError As Boolean
    
    blnIsWarning = m_objErrAssist.IsWarning
    blnIsSystemError = m_objErrAssist.IsSystemError
    
    'Simple Error Logging
    App.LogEvent _
        "Error in " & strClassName & "::" & strFunctionName & "(). " & vbCrLf & _
        "Err.Number: " & lngErrorNumber & ", " & _
        "Err.Source: " & strErrorSource & ", " & _
        "Err.Description: " & strErrorDescription, _
        vbLogEventTypeError
        
    If blnIsWarning = True And Not xmlResponse Is Nothing Then
        m_objErrAssist.AddWarning xmlResponse
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    If blnIsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, lngErrorNumber, strErrorDescription
    End If
    Set ReportError = m_objErrAssist.CreateErrorResponseEx
End Function

Private Function GetEmailAttachmentsDirectory(ByVal vstrUserId As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the name of the directory to which email attachments should be temporarily saved;
'   create the directory if it does not exist. Use DIR_TEMP_EMAIL_ATTACHMENTS as default
'   root if there is a problem getting the GlobalParameter "EmailAttachmentFilePath".
' history:
'   RF 29/04/03 LIVE00000674 Created function
'-------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "GetEmailAttachmentsDirectory"

Dim strDirName As String
Dim strErrMsg As String
Dim objGlobalParameterDO As omBase.GlobalParameterDO

    On Error GoTo GetEmailAttachmentsDirectoryVbErr
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    If Len(vstrUserId) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMissingParameter, "User Id"
    End If
    Set objGlobalParameterDO = _
        m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    On Error Resume Next
        
    strDirName = objGlobalParameterDO.GetCurrentParameterByType( _
        "EmailAttachmentFilePath", "STRING")
    m_objErrAssist.SaveErr
    Err.Clear
    On Error GoTo GetEmailAttachmentsDirectoryVbErr
    
    If m_objErrAssist.LastErr <> 0 Then
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = _
            oeRecordNotFound Then
            strErrMsg = "Failed to read GlobalParameter EmailAttachmentFilePath; " & _
                "Source: " & cstrClassName & "::" & cstrFunctionName
            App.LogEvent strErrMsg, vbLogEventTypeWarning
        
        Else
            m_objErrAssist.ReRaise
        End If
    End If
    If Len(strDirName) = 0 Then
        ' default to c:\temp for root directory
        strDirName = DIR_TEMP_EMAIL_ATTACHMENTS
    End If
    strDirName = strDirName & "\" & vstrUserId
    ' Create the output directory if it doesn't exist
    CreateNewDirectory strDirName
    
GetEmailAttachmentsDirectoryExit:
    
    Set objGlobalParameterDO = Nothing
    GetEmailAttachmentsDirectory = strDirName
    Exit Function
    
GetEmailAttachmentsDirectoryVbErr:
    Set objGlobalParameterDO = Nothing
    strErrMsg = "Error in " & cstrFunctionName & ": " & Err.Description
    
    App.LogEvent strErrMsg, vbLogEventTypeError
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub RunDPSMerge(vxmlRequest As IXMLDOMElement, cstrFunctionName)
Dim xmlNodeTemp As IXMLDOMNode
    
    On Error Resume Next
    
    Set xmlNodeTemp = vxmlRequest.selectSingleNode("PRINTDATA")
    
    'Check existence of merge data before attempting to do merge
    If Not xmlNodeTemp Is Nothing Then
        If xmlNodeTemp.childNodes.length > 0 Then
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[3.1]", "DPSMerge call"
            gobjTrace.TraceXML vxmlRequest.xml, "DPSMerge_request"
            m_objWordApp.Run "DPSMerge", vxmlRequest.xml
            
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[3.2]", "DPSMerge return"
        End If
    End If
    
    Call DeleteAllCodeModules
    m_objErrAssist.SaveErr

End Sub

Private Function Send_Email_By_CDO(strEmailAdministrator As String, strRecipients() As String, lngAddresseeCount As Long, strSubject As String, strBodyText As String, strAttachment As String) As Long
    
Const cstrFunctionName = "Send_Email_By_CDO"

Dim objConfig As New CDO.Configuration
Dim objMessage As New CDO.MESSAGE
Dim objSession As Object

Dim arrEmailDetails
Dim Index As Integer
Dim Fields

Dim strProfileName As String
Dim strSender As String
Dim strToList As String
Dim smtp_server_name As String
Dim X As Integer
'====================================================================
' Parameters :
'   smtp_server_name    - String        - naming the mail server on the installation (MSG's is MSGCHEX2K)
'   strSender           - String        - email address of the sender
'   strRecipients()     - String array  - email address(es) of the recipient(s)
'   lngAddresseeCount   - Long          - number of recipients
'   strSubject          - String        - email subject
'   strBodyText         - String        - email message
'   strAttachment       - String        - (optional) URL of attachment
'
'====================================================================
    On Error GoTo emailError:
    'Default profile name
    strProfileName = "MS Exchange Settings"
    X = InStr(1, strEmailAdministrator, ":")
    If X > 0 Then
        arrEmailDetails = Split(strEmailAdministrator, ":")
        For X = 0 To UBound(arrEmailDetails, 1)
            Select Case X
                Case 0
                    smtp_server_name = arrEmailDetails(0)
                Case 1
                    strSender = arrEmailDetails(1)
                Case 2
                    strProfileName = arrEmailDetails(2)
            End Select
        Next X
    End If
    If smtp_server_name = "" Or strSender = "" Then
        Send_Email_By_CDO = 4910 'Email Administrator format error
        Exit Function
    End If
    'App.LogEvent cstrClassName & "::" & cstrFunctionName & " Logging on to MAPI session"
    Set objSession = Session_Logon_NoDialog(strProfileName, "")
    Debug.Print "Email Server : " & smtp_server_name & " Administrator : " & strSender
    For Index = 0 To lngAddresseeCount - 1
        If Index > 0 Then
            strToList = strToList & ", "
        End If
        strToList = strToList & "<" & strRecipients(Index) & ">"
    Next Index
    Set Fields = objConfig.Fields
    'App.LogEvent CstrClassName & "::" & cstrFunctionName & " Setting up configuration"
    ' Set config fields we care about
    With Fields
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = smtp_server_name
        .Item(cdoSMTPServerPort) = 25
        .Item(cdoSMTPConnectionTimeout) = 10
        .Item(cdoSMTPAuthenticate) = cdoBasic
        .Update
    End With
    Set objMessage.Configuration = objConfig
    'App.LogEvent CstrClassName & "::" & cstrFunctionName & " Building email message"
    With objMessage
        .To = strToList
        .From = strSender
        .Subject = strSubject
        .TextBody = strBodyText
        If strAttachment <> "" Then
            .AddAttachment strAttachment
        End If
        .send
    End With
    'App.LogEvent CstrClassName & "::" & cstrFunctionName & " Email sent successfully !!"
    Set Fields = Nothing
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set objSession = Nothing
    Exit Function
    
emailError:
    Send_Email_By_CDO = Err.Number
End Function


Public Function Session_Logon_NoDialog(strProfileName As String, strPassword) As Object
    
    On Error GoTo error_olemsg
    If Session_Logon_NoDialog Is Nothing Then
        Set Session_Logon_NoDialog = CreateObject("MAPI.Session")
    End If
    If Not Session_Logon_NoDialog Is Nothing Then
        If strProfileName <> "" And strPassword <> "" Then
'            App.LogEvent "omDPS.Session_Logon_NoDialog - Profile and Password supplied"
            Session_Logon_NoDialog.Logon profileName:=strProfileName, Password:=strPassword, showDialog:=False
        Else
            If strProfileName <> "" Then
'                App.LogEvent cstrClassName & ":: Session_Logon_NoDialog - Profile supplied without Password"
                Session_Logon_NoDialog.Logon profileName:=strProfileName
            Else
                Err.Raise 1273, "Cannot logon to email system: no profile name or password"
            End If
        End If
    End If
    Exit Function
    
error_olemsg:
    Set Session_Logon_NoDialog = Nothing
End Function
Private Sub SetUpNewWordObject(cstrFunctionName)

    On Error Resume Next
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "create Word app. starts"
    
    Set m_objWordApp = CreateObject("Word.Application")
    
    ' see Technet article Q259971
    m_objWordApp.DisplayAlerts = 0 'wdAlertsNone
       
    ' AD also have to turn off the prompt that asks if you want to save normal.dot
    m_objWordApp.Options.SaveNormalPrompt = False
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName & " SetUpNewWordObject()", "[1.2]", "create Word app. ends"

End Sub

Private Sub WriteToFile(ByVal strFileName As String, ByRef theData() As Byte, ByVal blnWriteAsString As Boolean)
    Dim F As Integer

    F = FreeFile
    ' Truncates file.
    Open strFileName For Output Access Write Lock Read Write As F
    Close F
    ' Write array to the file.
    Open strFileName For Binary Access Write Lock Read Write As F
    If blnWriteAsString Then
        Dim strFile As String
        strFile = theData
        Put F, , strFile
    Else
        Put F, , theData
    End If
    Close F
End Sub

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    ' ik_bm0200 traceAssist support
'TW 11/5/2005
'    Set m_objConvertAssist = New ConvertAssist
'TW 11/5/2005
    gobjTrace.StartTrace
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
'TW 11/5/2005
'    Set m_objConvertAssist = Nothing
'TW 11/5/2005
End Sub
Private Function ConvertBinToBase64( _
    ByVal WordDoc As Object, _
    ByRef xmlRequestNode As IXMLDOMNode, _
    ByVal strCompressionMethod As String) As String

Const cstrFunctionName As String = "ConvertBinToBase64"

Dim arrData() As Byte

Dim fs As FileSystemObject

Dim F As Integer

Dim lngErrorNumber As Long

Dim strContents As String
Dim strErrorDescription As String
Dim strErrorSource As String

Dim strTempFileName As String

Dim xmlAttribute As IXMLDOMAttribute
    
    On Error GoTo vbErrorHandler
   
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    ' Convert
    Set fs = New FileSystemObject
    strTempFileName = fs.GetSpecialFolder(TemporaryFolder) & "\" & fs.GetTempName
    
    WordDoc.SaveAs strTempFileName
    ReDim arrData(FileLen(strTempFileName) - 1)
    F = FreeFile()
    Open strTempFileName For Binary Access Read As F
    Get F, , arrData
    Close F
    
    Call Compress(arrData, strCompressionMethod)
        
    Set xmlAttribute = xmlRequestNode.ownerDocument.createAttribute("PRINTDOCUMENT")
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    xmlAttribute.dataType = "bin.base64"
    xmlAttribute.nodeTypedValue = arrData
        
    ConvertBinToBase64 = strTempFileName
    
    Set xmlAttribute = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    Exit Function
        
    'DR New Error Logging+Handling
vbErrorHandler:
    
    Set xmlAttribute = Nothing
    
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
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    'Simple Error Handling - Bubble up that error!
    Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
End Function

Private Function ConvertBase64ToBin( _
    ByRef WordDocOut As Object, _
    ByRef xmlAttribute As IXMLDOMAttribute, _
    ByVal strCompressionMethod As String) As String

Const cstrFunctionName = "ConvertBase64ToBin"

Dim arrData() As Byte

Dim fs As FileSystemObject

Dim F As Integer

Dim lngErrorNumber As Long

Dim strErrorDescription As String
Dim strErrorSource As String
Dim strTempFileName As String
    
    On Error GoTo vbErrorHandler
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    xmlAttribute.dataType = "bin.base64"
    arrData = xmlAttribute.nodeTypedValue
        
    Call Decompress(arrData, strCompressionMethod)
            
    Set fs = New FileSystemObject
    strTempFileName = fs.GetSpecialFolder(TemporaryFolder) & "\" & fs.GetTempName
    F = FreeFile()
    Open strTempFileName For Binary Access Write As F
    Put F, , arrData
    Close F
    m_objWordApp.Documents.open strTempFileName
    Set WordDocOut = m_objWordApp.Documents(1)
    
    Set fs = Nothing
    
    ConvertBase64ToBin = strTempFileName
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
        
    'DR New Error Logging+Handling
vbErrorHandler:
    
    Set fs = Nothing
    
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
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    'Simple Error Handling - Bubble up that error!
    Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
End Function

Public Function MergeDocument(ByVal vstrXMLRequest As String) As String
'****************************************************************************************
'BG 17/10/2002 BMIDS00592 - Added method.  Copied the code to call the GetTemplate largely from
'                           ProcessRequest on this class.
'****************************************************************************************
Const cstrFunctionName = "MergeDocument"
    
Dim intOutputType As Integer

'TW 11/5/2005
'Dim objTemplateHandlerDO As ITemplateHandlerDO
'TW 11/5/2005 End

Dim strOutputType As String
Dim strTemplateId As String

Dim xmlDoTheMergeRespNode As IXMLDOMNode
Dim xmlIn As FreeThreadedDOMDocument40
Dim xmlKFIResponseDoc As FreeThreadedDOMDocument40
Dim xmlNodeRequest As IXMLDOMNode
Dim xmlNodeTemp As IXMLDOMNode
Dim xmlNodeTemplate As IXMLDOMNode
Dim xmlTemplateIDNode As IXMLDOMNode
Dim xmlTemplateNode As IXMLDOMNode

    On Error GoTo MergeDocumentErr
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    gobjTrace.TraceRequest vstrXMLRequest
       
    If m_objXmlAssist Is Nothing Then
        Set m_objXmlAssist = New XMLAssist
    End If
       
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
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
    xmlIn.save sPath & "\omDPS_MergeDocument_Request_" & tname & ".xml"
    Set fso = Nothing
#End If
    
'TW 11/5/2005
'    Set objTemplateHandlerDO = m_objContext.CreateInstance(App.Title & ".TemplateHandlerDO")
'TW 11/5/2005 End
    Set xmlNodeRequest = m_objXmlAssist.GetRequestNodeEx(xmlIn.documentElement)
    Set xmlNodeTemp = xmlNodeRequest.appendChild(xmlIn.createElement("TEMPLATE"))
    Set xmlNodeTemplate = m_objXmlAssist.GetMandatoryNode(xmlIn.documentElement, "CONTROLDATA")
        
    strTemplateId = xmlGetMandatoryAttributeText(xmlNodeTemplate, "DPSDOCUMENTID")
    If strTemplateId = "" Then
        ' AD - should never get here if the xml is one of the above. Raise error
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 502
    End If
    Set xmlTemplateIDNode = xmlIn.createElement("TEMPLATEID")
    xmlTemplateIDNode.Text = strTemplateId
    xmlNodeTemp.appendChild xmlTemplateIDNode
    
    xmlNodeTemplate.appendChild(xmlIn.createElement("AUTHORITYLEVEL")).Text = (xmlGetAttributeText(xmlIn.documentElement, "USERAUTHORITYLEVEL"))
'TW 11/5/2005
'    Set xmlTemplateNode = objTemplateHandlerDO.GetTemplate(xmlNodeRequest)
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.1]", "objDDBXML.GetTemplateXML call"
    Set xmlTemplateNode = dbxmlGetTemplateXML(strTemplateId)
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.2]", "objDDBXML.GetTemplateXML return"
'TW 11/5/2005 End
    If xmlTemplateNode Is Nothing Then
'        errThrowError cstrClassName, 4901, "MergeDocument Template not found : " & strTemplateID
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName & " Template not found : " & strTemplateId, 4901
    End If
    Debug.Print xmlTemplateNode.xml
    '------------------------------------------------------------------------------------------
    ' check security level
    '------------------------------------------------------------------------------------------
    If IsAuthorityLevelSufficient(xmlIn.documentElement, xmlTemplateNode) = True Then
                
        'CORE46 Allow for different print engines
                
        'Check for eKFI processing
        If (GetDeliveryEngineType(xmlTemplateNode) <> "eKFI") Then   ' Standard Omiga
            
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "DoTheMerge call"
            Set xmlDoTheMergeRespNode = DoTheMerge(xmlIn.documentElement, xmlTemplateNode)
            
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "DoTheMerge return"
        Else
            intOutputType = CSafeInt(xmlGetAttributeAsInteger(xmlNodeTemplate, "DELIVERYTYPE"))
            strOutputType = GetValidationTypeForValueID("DocumentDeliveryType", intOutputType)

            Dim strCompressionMethod As String
            strCompressionMethod = xmlGetAttributeText(xmlNodeTemplate, "COMPRESSIONMETHOD")
            Set xmlKFIResponseDoc = GetKFIDocument(xmlIn, xmlTemplateNode, strOutputType, strCompressionMethod)
            Set xmlDoTheMergeRespNode = xmlKFIResponseDoc.selectSingleNode("RESPONSE")
        End If
           
    Else
        
        '------------------------------------------------------------------------------------------
        ' failed security check
        '------------------------------------------------------------------------------------------
        ' "Insufficient authority to print selected template"
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 4900
    End If
    MergeDocument = xmlDoTheMergeRespNode.xml
'TW 11/5/2005
'    Set objTemplateHandlerDO = Nothing
'TW 11/5/2005 end
    Set xmlNodeRequest = Nothing
    Set xmlNodeTemp = Nothing
    Set xmlNodeTemplate = Nothing
    Set xmlTemplateNode = Nothing
    Set xmlIn = Nothing
    Set xmlTemplateIDNode = Nothing
    Set xmlDoTheMergeRespNode = Nothing
    Set xmlKFIResponseDoc = Nothing           'CORE46
'   IK_DEBUG
#If IK_DEBUG Then
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.loadXML MergeDocument
    xmlResponseDoc.save sPath & "\omDPS_MergeDocument_Response_" & tname & ".xml"
    Set xmlResponseDoc = Nothing
#End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    gobjTrace.TraceResponse MergeDocument
    Exit Function
MergeDocumentErr:
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
    Err.Raise Err.Number, cstrFunctionName, Err.Description
End Function
Private Function DoTheMerge( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal vxmlTemplateNode As IXMLDOMNode) _
    As IXMLDOMNode

'TW Tidied up 23/7/2004
Const cstrFunctionName = "DoTheMerge"

Dim intErrNumber As Integer

Dim lngErrorNumber As Long

Dim strErrorDescription As String
Dim strErrorSource As String
Dim strTemplateFileName As String
Dim strTempFileName As String
 
Dim WordDoc As Object

Dim xmlConvertedFileContentsElement As IXMLDOMElement
Dim xmlNodeTemp As IXMLDOMNode
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement

    On Error GoTo VBErr
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
       
    '------------------------------------------------------------------------------------------
    ' get template file name
    '------------------------------------------------------------------------------------------
'TW 11/5/2005
'    strTemplateFileName = m_objXmlAssist.GetNodeValue(vxmlTemplateNode, _
'         "//TEMPLATE/FILENAME", True)
            
    strTemplateFileName = vxmlTemplateNode.selectSingleNode("//TEMPLATE/FILENAME").Text
'TW 11/5/2005 End
            
            
    '------------------------------------------------------------------------------------------
    ' do the print merge
    '------------------------------------------------------------------------------------------
    
    Call SetUpNewWordObject(cstrFunctionName)
       
    On Error Resume Next
    Err.Clear
    ' open the template.
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "Word.Documents.Open starts"
               
    Set WordDoc = m_objWordApp.Documents.open(strTemplateFileName)
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "Word.Documents.Open ends"
    
    ' AD 06/06/01 - the xml response now includes sending back the merged document as
    ' a 64 bit text stream tagged onto the response.
    Set xmlConvertedFileContentsElement = xmlOut.createElement("PRINTDOCUMENTDETAILS")
    xmlSetAttributeValue xmlConvertedFileContentsElement, "PRINTTIMESTAMP", Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Call RunDPSMerge(vxmlRequest, cstrFunctionName)
    
    Dim xmlControlDataNode As IXMLDOMNode
    Dim strCompressionMethod As String
    Set xmlControlDataNode = xmlGetNode(vxmlRequest, "CONTROLDATA")
    If Not xmlControlDataNode Is Nothing Then
        strCompressionMethod = xmlGetAttributeText(xmlControlDataNode, "COMPRESSIONMETHOD")
    End If
    strTempFileName = ConvertBinToBase64(WordDoc, xmlConvertedFileContentsElement, strCompressionMethod)
    Set xmlControlDataNode = Nothing
    xmlResponseElem.appendChild xmlConvertedFileContentsElement
                
    On Error GoTo VBErr
    If m_objErrAssist.LastErr <> 0 Then
        If m_objErrAssist.LastErr = 5174 Then
            intErrNumber = 4902     'Template file not found
        Else
            intErrNumber = 4901     'Data merge failed
        End If
        errThrowError cstrFunctionName, intErrNumber, "Template name: " & strTemplateFileName
    End If
           
    WordDoc.Close SaveChanges:=False
        
    ' add a child element to the response ????
    Set DoTheMerge = xmlResponseElem
    
DoTheMergeExit:
    On Error Resume Next
    If Not m_objWordApp Is Nothing Then
        
        'DR DMSSYS0005 Attempt to close the doc in case its been left open due to an error.
        WordDoc.Close SaveChanges:=False
        Err.Clear
        m_objWordApp.Quit
    End If
    If strTempFileName <> "" Then
        Kill strTempFileName
    End If
    Set WordDoc = Nothing
       
    'DR Raise an error if we've processed one.
    If lngErrorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise lngErrorNumber, strErrorSource, strErrorDescription
    End If
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
    
VBErr:
    ' OK, should check the response here and set MQResp to the appropriate value.
    lngErrorNumber = Err.Number
    strErrorSource = Err.Source
    strErrorDescription = Err.Description
    'Simple Error Logging
    App.LogEvent "Error encountered within " & cstrClassName & "::" & cstrFunctionName & vbCrLf & _
                "(). Err.Number: " & lngErrorNumber & _
                ", Err.Source: " & strErrorSource & _
                ", Err.Description: " & strErrorDescription, _
                vbLogEventTypeError
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    Set DoTheMerge = m_objErrAssist.CreateErrorResponseEx
    ' go to clean-up section
    Resume DoTheMergeExit
End Function

' PSC 20/09/2004 BBG1438 - Start
Private Function GetPrintTrayId(ByVal strPrinterName As String, ByVal intTrayType As Integer) As Integer

    Dim xmlPrinter As IXMLDOMNode
    Dim strTrayId As String
    Dim strPattern As String

    If Not gxmldocPrinterAttributes Is Nothing Then
    
        ' Find the print tray attributes for the specified printer
        strPattern = "PRINTERATTRIBUTES/PRINTER[translate(@NAME, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')='" & UCase(strPrinterName) & "']"
        
        Set xmlPrinter = xmlGetNode(gxmldocPrinterAttributes, strPattern)
        
        ' Attributes found
        If Not xmlPrinter Is Nothing Then
            
            ' If headed paper required get tray code otherwise get plain paper tray code
            ' PSC 19/10/2004 BBG1664
            If GetValidationTypeForValueID("DefaultPrinterTray", intTrayType) = "H" Then
                strTrayId = xmlGetAttributeText(xmlPrinter, "HEADEDPAPERTRAYCODE")
            Else
                strTrayId = xmlGetAttributeText(xmlPrinter, "PLAINPAPERTRAYCODE")
            End If
            
        End If
    End If
    
    GetPrintTrayId = Val(strTrayId)
    
End Function
' PSC 20/09/2004 BBG1438 - End




' ik_20030206
' tracing interface
Public Sub InitialiseTraceInterface( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStart As Single)
    
    If Not vblnTraceOn Then
        Exit Sub
    End If
    
    gobjTrace.TraceInitialiseFromParent vblnTraceOn, vstrFileName, vintStart
End Sub

'CORE00000080 GHun
Private Sub EmailFulfilment(ByVal vstrRTF As String, ByRef vstrRecipients() As String, ByVal vlngNumRecipients As Long, Optional ByVal vstrDocumentLocation As String = "")
    Const cstrFunctionName As String = "EmailFulfilment"
    Dim strEmail As String
    strEmail = DeliverAsText(vstrRTF)
    SendGeminiEmail strEmail, vstrRecipients, vlngNumRecipients, vstrDocumentLocation
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Public Function DeliverAsText(ByVal strRTF As String) As String
    Const cstrFunction = "DeliverAsText.Delivery.DeliverAsText"
    Dim arrScoping As Variant
    
    Dim arrWords As Variant
    Dim thisScope As String
    Dim thisWord As String
    Dim strWork As String
    Dim intScopeLevel As Integer
    Dim X As Long
    Dim Y As Long

    'This routine converts an input rtf string to a txt string.
    
    On Error GoTo DeliverAsTextError

    ' 1 Establish the start of relevant data
    arrScoping = Split(strRTF, "{")
    
    For X = 0 To UBound(arrScoping, 1)
        intScopeLevel = intScopeLevel + 1
        thisScope = arrScoping(X)
        Y = InStr(1, thisScope, "}")
        If Y > 0 Then
            Do While Y > 0
                Y = InStr(Y + 1, thisScope, "}")
                intScopeLevel = intScopeLevel - 1
            Loop
            If InStr(1, thisScope, "\pard\plain") > 0 Then
                If intScopeLevel = 2 Then
                    Exit For
                End If
            End If
        End If
        arrScoping(X) = vbNullString
    Next X

    ' 2 Deal with 'pictures'
    DealWithBlock "\shppict", arrScoping
    DealWithBlock "\nonshppict", arrScoping

    ' 3 Deal with 'fields'
    DealWithBlock "\field", arrScoping
    
    ' 4 Deal with bullets and numbering
    DealWithBlock "\pnseclvl", arrScoping

    strWork = Join(arrScoping, "")
    
    strWork = Replace(strWork, vbCrLf, " ")

    ' 4 Remove scope terminator '}' characters
    strWork = Replace(strWork, "}", "")
    ' 5 Replace \tab
    strWork = Replace(strWork, "\tab", vbTab)
    ' 6 Replace \cell
    strWork = Replace(strWork, "\cell ", vbTab)
    ' 7 Replace \trowd
    strWork = Replace(strWork, "\trowd ", vbCrLf)
    ' 6 Replace \par by crlf
    strWork = Replace(strWork, "\par ", vbCrLf)

    ' 7 Find "Words"
    arrWords = Split(strWork, " ")
    For X = 0 To UBound(arrWords, 1)
        thisWord = arrWords(X)
        If Len(thisWord) > 0 Then
            If InStr(1, thisWord, "\sn") > 0 Or InStr(1, thisWord, "\sp") > 0 Or InStr(1, thisWord, "\sv") > 0 Or InStr(1, thisWord, "\pntxta") > 0 Or InStr(1, thisWord, "\pntxtb") > 0 Then
                thisWord = vbNullString
                arrWords(X + 1) = vbNullString
            Else
                Y = InStr(1, thisWord, "\")
                If Y > 0 Then
                    thisWord = Left$(thisWord, Y - 1)
                Else
                    thisWord = thisWord & " "
                End If
            End If
            arrWords(X) = thisWord
        End If
    Next X

    strWork = Join(arrWords, " ")
    Do While InStr(1, strWork, "  ") > 0
        strWork = Replace$(strWork, "  ", " ")
    Loop
    
    Do While InStr(1, strWork, vbCrLf & " ") > 0
        strWork = Replace$(strWork, vbCrLf & " ", vbCrLf)
    Loop
    
    DeliverAsText = Trim$(strWork)
    
DeliverAsTextError:
    
    errCheckError cstrFunction, TypeName(Me)
End Function

Private Sub DealWithBlock(ByVal strFieldType As String, ByRef arrScoping As Variant)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    
    Dim intScopeLevel As Integer
    
    Dim thisScope As String

    For X = 1 To UBound(arrScoping, 1)
        thisScope = arrScoping(X)
        If InStr(1, thisScope, strFieldType) > 0 Then
            intScopeLevel = 1
            For Y = X + 1 To UBound(arrScoping, 1)
                intScopeLevel = intScopeLevel + 1
                thisScope = arrScoping(Y)
                arrScoping(Y) = vbNullString
                Z = InStr(1, thisScope, "}")
                If Z > 0 Then
                    Do While Z > 0
                        Z = InStr(Z + 1, thisScope, "}")
                        intScopeLevel = intScopeLevel - 1
                    Loop
                    If intScopeLevel = 1 Then
                        Exit For
                    End If
                End If
            Next Y
        End If
    Next X
End Sub

Private Sub SendGeminiEmail(ByVal vstrEmail As String, ByRef rstrRecipients() As String, ByVal vlngNumRecipients As Long, Optional ByVal vstrDocumentLocation As String = "")
    Const cstrFunctionName As String = "SendGeminiEmail"
    Dim objGemini       As Object
    Dim xmlDoc          As FreeThreadedDOMDocument40
    Dim xmlRequest      As IXMLDOMElement
    Dim xmlEmail        As IXMLDOMElement
    Dim lngPos          As Long
    Dim lngEnd          As Long
    Dim strResponse     As String
    Dim strFromAddress  As String
    Dim strSubject      As String
    Dim strBody         As String

On Error GoTo SendGeminiEmailErr

    Set xmlDoc = xmlCreateDOMObject()
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlRequest.setAttribute "OPERATION", "SENDEMAIL"
    
    Set xmlEmail = xmlDoc.createElement("EMAIL")
    xmlRequest.appendChild xmlEmail
    
    strBody = vstrEmail
    
    'Check for from address in email template
    lngPos = InStr(1, strBody, "FROMADDRESS:")
    If lngPos > 0 Then
        lngEnd = InStr(1, strBody, vbCrLf)
        strFromAddress = Trim$(Mid$(strBody, lngPos + 12, lngEnd - lngPos - 12))
        strBody = Left$(strBody, lngPos - 1) & Right$(strBody, Len(strBody) - lngEnd - 1)
    End If
    
    'Check for subject in email template
    lngPos = InStr(1, strBody, "SUBJECT:")
    If lngPos > 0 Then
        lngEnd = InStr(1, strBody, vbCrLf)
        strSubject = Trim$(Mid$(strBody, lngPos + 8, lngEnd - lngPos - 8))
        strBody = Left$(strBody, lngPos - 1) & Right$(strBody, Len(strBody) - lngEnd - 1)
    End If
    
    xmlEmail.setAttribute "FROMADDRESS", strFromAddress
    xmlEmail.setAttribute "SUBJECT", strSubject
    xmlEmail.setAttribute "BODY", strBody
    If Len(vstrDocumentLocation) > 0 Then
        xmlEmail.setAttribute "DOCUMENTLOCATION", vstrDocumentLocation
    End If
    
    If m_objContext Is Nothing Then
        Set objGemini = CreateObject("omGemini.FileVersioningBO")
    Else
        Set objGemini = m_objContext.CreateInstance("omGemini.FileVersioningBO")
    End If
    
    For lngPos = 0 To vlngNumRecipients - 1
        xmlEmail.setAttribute "TOADDRESS", rstrRecipients(lngPos)
    
        strResponse = objGemini.ProcessRequest(xmlRequest.xml)
        errCheckXMLResponse strResponse, True
    Next

SendGeminiEmailErr:
    Set objGemini = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlEmail = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'CORE00000080 End

'CORE00000080 GHun
Private Function GetEmailRecipient(ByVal vstrApplicationNumber As String, ByVal vstrEmailRecipientType As String) As IXMLDOMElement
    Const cstrFunctionName  As String = "GetEmailRecipient"
    Dim objCRUD             As Object
    Dim strResponse         As String
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlEntity           As IXMLDOMElement
    
On Error GoTo GetEmailRecipientErr
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Set xmlDoc = xmlCreateDOMObject()
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlRequest.setAttribute "CRUD_OP", "READ"
    xmlRequest.setAttribute "ENTITY_REF", "GETEMAILRECIPIENT"
    xmlRequest.setAttribute "SCHEMA_NAME", "omPrintSchema"
    Set xmlEntity = xmlDoc.createElement("GETEMAILRECIPIENT")
    xmlRequest.appendChild xmlEntity
    xmlEntity.setAttribute "APPLICATIONNUMBER", vstrApplicationNumber
    xmlEntity.setAttribute "EMAILRECIPIENTTYPE", vstrEmailRecipientType

    Set objCRUD = CreateObject("omCRUD.omCRUDBO")
    strResponse = objCRUD.omRequest(xmlRequest.xml)
        
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    Set GetEmailRecipient = xmlDoc.documentElement
    
GetEmailRecipientErr:
    
    Set objCRUD = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlEntity = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'CORE00000080 End

