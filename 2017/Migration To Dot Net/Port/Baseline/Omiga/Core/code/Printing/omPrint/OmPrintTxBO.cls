VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OmPrintTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omPrintTxBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling

'Description:
'
'
'Dependencies:  MessageQueueListener
'               omPDM
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AW     08/04/01    SYS2302 Corrected ProgId
'AW     08/05/01    SYS2305 Only create one destination child, dependent on destType
'AW     14/05/01    SYS2311 Remove node FILEID
'DR     10/02/02    DMSSYS0008 Added DMS2 changes - tried to minimise the impact of DMS2 on
'                   Omiga4 by placing the changes in just one area of the code.
'DR     12/02/02    DMSSYS0008 Adding more DMS2 changes - to help support eMail/Fax and SMS
'DR     18/02/02    DMSSYS0032 Added a Customer Name Element to IOmPrintTxBO_PrintDocument()
'DR     20/02/02    DMSSYS0008 Minor bug: shouldnt assume that there are nodes in PrintDocument().
'SA     20/03/02    SYS4287 Batch Printing Rate Change Letters
'DR     25/03/02    SYS4289 Fix email resend bug
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'GD     08/10/02    BMIDS00569 - Make sure DESTINATIONTYPE gets converted from its ValueID (eg. 20) to its ValidationType (eg. 'R')
'                   ..so that the appropriate printer info gets added.
'GD     10/10/02    Back out above change. Put it into PayProc
'BG     09/10/02 BMIDS00612 Check the new indicator instead of application number to set operation correctly.  Documents with ONMESSAGE
            'will simply print.  Documents with NEWDOCUMENT will print and archive.  PRINTINDICATOR is either set
            'in the GUI (PM010 and cm100) or in omTMBO.PrintDocumentsForTask which gets fired for automatic print tasks
            'when moving to next stage.  Gui screens which do screens prints where there would be no app number - ie
            'TM020 and gn215 have not had an indicator attribute added - the code below will return an empty string
            'and the correct operation will be set, namely ONMESSAGE, as these do not need to be archive as they are not
            'application specific prints.
'BG     BMIDS00592 16/10/02 Change printdocument to handle view/edit before print
'MO     05/11/2002  BMIDS00460 - Made changes to late bind omPDM
'IK     05/11/2002  BMIDS00730 - print via BMIDS BMPrinting.BMPrintRouter component
'IK     14/11/2002  BMIDS00831 - various fixes re: integration of printing via Tasks
'IK     17/02/2003  BM0200 - add TraceAssist support
'-------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------
'Baseline Specific Changes:
'
'Prog   Date        AQR     Description
'MV     03/02/2004  Core62  Amended IOmPrintTxBO_PrintLetters;IOmPrintTxBO_RePrintDocument;IOmPrintTxBO_PrintDocument
'----------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        Defect  Description
'MV     05/07/2004  BBG452  WP20 - Amended PrintDocumentCommonCode() - Send URLPostInd to DPS
'JS     20/07/2004  BBG1057 WP06  - Changed IOmPrintTxBO_PrintDocument() to allow ACTION=DIP to return data
'AS     29/07/2004  BBG1057 Fixed XPATH - was not returning Nothing when ACTION
'                           was not 'DIP' or 'APP' (e.g., 'PrintDocument',
'                           preventing jobs getting on the print queue.
'PSC    24/08/2004  BBG1233 Amend PrintDocumentCommonCode so that DELIVERYTYPE and URLPOSTIND are got from the
'                           Host Template if present
'PSC    20/09/2004  BBG1438 Added Paper Tray Selection Capability
'INR    22/09/2004  BBG607  PrintCheque requires DeliveryType added to the ControlData in PrintDocumentCommonCode
'PSC    06/10/2004  BBG1572 Added WEBDOCUMENTTYPE to control data
'PSC    02/12/2004  E2EM00003096 Amend PrintLetters to set up delivery type
'AS     28/01/05    BBGRETAIL - added support for local workstation printing.
'MV     07/04/2005  BBG2045/CORE97  Amended PrintDocumentCommonCode() to include "Local WorkStation" - "W"
'MV     25/04/2005  CORE107  Amended IOmPrintTxBO_PrintDocument() and PrintDocumentCommonCode()
'MV     28/04/2005  CORE107  Amended IOmPrintTxBO_PrintDocument() and PrintDocumentCommonCode()
'MV     28/04/2005  CORE107  Amended PrintDocumentCommonCode()
'MV     05/05/2005  CORE116  Undo CORE107 - do not convert DESTINATIONTYPE to valid id.
'AS     10/05/2005  CORE121 omPrint.IOmPrintTxBO_PrintDocument() should include the "DeliveryType" attribute
'AS     11/05/2005  CORE122 omPrint does not merge TEMPLATEDATA when edit/view before printing.
'AS     19/05/2005  CORE132 Printing components should use MSXML4 NewParser when loading documents.
'RF     16/02/2006  CORE245 Allow printer destination of "Document Store Only"
'AS     05/12/2006  CORE312 Gemini Printing: enhance existing print components.
'AS     04/01/2007  CORE327 Ensure DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS is set.
'GHun   02/03/2007  CORE00000080 Changes for email fulfilment
'GHun   19/03/2007  CORE00000080 Raise an error if no EmailRecipient exists
'-------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IOmPrintTxBO

Private gobjContext As ObjectContext




Private Sub IOmPrintTxBO_PrintDocument(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo PrintDocumentExit
    
    Const cstrFunctionName As String = "IOmPrintTxBO_PrintDocument"
    
    gobjContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
'TW 11/5/2005
'    Dim objGlobalParameterBO As GlobalParameterBO
'TW 11/5/2005 End
    Dim objDO As IOmPrintDO
    Dim objPrintDataManagerBO As Object
    Dim objPrintManagerBO As Object
    Dim objMessageQueueing As IomMessageQueue
    
    Dim xmlPDMResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlHostTemplateDoc As FreeThreadedDOMDocument40
    Dim xmlHostTemplateNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlAxWordNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode  ' CORE46
    Dim xmlElementReq As IXMLDOMElement
    Dim xmlElementMQ As IXMLDOMElement
    Dim xmlQueueingDoc As New FreeThreadedDOMDocument40
    Dim xmlViewOrEditNode As IXMLDOMNode
    Dim xmlActionNode As IXMLDOMNode
    Dim xmlRequestContDataElem As IXMLDOMElement
    Dim xmlOutputTypeElement As IXMLDOMNode
    Dim xmlTemporaryDoc As New FreeThreadedDOMDocument40
    Dim xmlPrinterElem As IXMLDOMNode
    Dim xmlPrinterNameElem As IXMLDOMNode
    Dim xmlCopiesElem As IXMLDOMNode
    Dim xmlPMNode As IXMLDOMNode
    Dim xmlOldPrintDataNode As IXMLDOMNode
    Dim xmlNewPrintDataNode As IXMLDOMNode
    Dim xmlPDMNode As IXMLDOMNode
    Dim xmlPMResponseDoc As FreeThreadedDOMDocument40
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlControlNode As IXMLDOMNode   ' PSC 20/09/2004 BBG1438
    
    Dim strPatternViewEdit As String
    Dim strPatternAction As String
    Dim strDestinationType As String
    Dim strPrinter As String
    Dim strCopies As String
    Dim strProcessMergeDocumentResponse As String
    Dim strDestination As String
    Dim strXMLResponse As String
    Dim strQueueName As String
    Dim strMethodName As String
    Dim strReturnInd As String
    Dim strClientPrintingRouterPROGID  As String
    Dim strQuotationNumber As String     ' CORE46
    Dim strTemplateName As String        ' CORE46

    Set xmlPDMResponseDoc = New FreeThreadedDOMDocument40
    xmlPDMResponseDoc.async = False
    xmlPDMResponseDoc.validateOnParse = False
    xmlPDMResponseDoc.setProperty "NewParser", True
    
'TW 11/5/2005
'    If objGlobalParameterBO Is Nothing Then
'        Set objGlobalParameterBO = gobjContext.CreateInstance("omBase.GlobalParameterBO")
'    End If
    
'    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "objGlobalParameterBO.GetCurrentParameter call"
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "dbxmlGetCurrentParameterXML call"

'    strXMLResponse = objGlobalParameterBO.GetCurrentParameter("PrintQueueName")
    Set xmlNode = dbxmlGetCurrentParameterXML("PrintQueueName")
    
'    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "objGlobalParameterBO.GetCurrentParameter return"
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "dbxmlGetCurrentParameterXML return"
'    Set xmlDoc = xmlLoad(strXMLResponse, cstrFunctionName)
'    strQueueName = xmlGetNodeText(xmlDoc.documentElement, ".//STRING")
    strQueueName = xmlGetNodeText(xmlNode, ".//STRING")
'TW 11/5/2005 End
    
    ' IK_BMIDS00831
    ' get (static) HOSTTEMPLATE details here,
    ' then we don't have to worry about them being passed in!
    ' (so long as we have CONTROLDATA/@HOSTTEMPLATEID on REQUEST)
    
    Set xmlHostTemplateDoc = New FreeThreadedDOMDocument40
    xmlHostTemplateDoc.async = False
    
    Set xmlDoc = Nothing
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    
    Set xmlElem = xmlHostTemplateDoc.createElement("RESPONSE")
    xmlHostTemplateDoc.appendChild xmlElem
    
    ' PSC 20/09/2004 BBG1438 - Start
    Set xmlControlNode = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
    strDestination = xmlGetAttributeText(xmlControlNode, "DESTINATIONTYPE")
    ' PSC 20/09/2004 BBG1438 - End
    
    Set xmlElem = xmlDoc.createElement("FINDATTRIBUTES")
    ' PSC 20/09/2004 BBG1438
    xmlSetAttributeValue xmlElem, "HOSTTEMPLATEID", xmlGetMandatoryAttributeText(xmlControlNode, "DOCUMENTID")
    xmlDoc.appendChild xmlElem
        
    Set objDO = GetObjectContext.CreateInstance("omPrint.omPrintDO")
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.1]", "objDO.GetPrintAttributes call"
    
    objDO.GetPrintAttributes xmlElem, xmlHostTemplateDoc.firstChild
    Set objDO = Nothing
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[2.2]", "objDO.GetPrintAttributes return"
    
    errCheckXMLResponseNode xmlHostTemplateDoc.firstChild, vxmlResponseNode, True
    
    Set xmlHostTemplateNode = xmlHostTemplateDoc.selectSingleNode("RESPONSE/ATTRIBUTES")
    
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    ' IK_BMIDS00831_ends
    
#If IK_DEBUG Then
    strQueueName = ".\PrintRequestsIK"
#End If
    
    'Call Print Data manager
    'MO 5/11/2002 BMIDS00460
    'Dim objPrintDataManagerBO As PrintDataManagerBO
    Set xmlPrintDataNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    strMethodName = xmlGetAttributeText(xmlPrintDataNode, "METHODNAME")

    Set xmlAxWordNode = xmlGetNode(vxmlRequestNode, "AXWORD")
    If Not xmlAxWordNode Is Nothing Then
        strReturnInd = xmlGetAttributeText(xmlAxWordNode, "RETURNIND")
    End If
    
    'BG BMIDS00592 16/10/02 - If this is the second time through for a view/edit document then we need
    'to make sure that we go to the queue and don't try to retrieve data from PDM as we
    'already have the details.
    
    If strReturnInd = "1" Then
    
        ' PSC 20/09/2004 BBG1438 - Start
        ' Need to copy across the printer tray attributes if present in HostTemplate
        xmlCopyAttribute xmlHostTemplateNode, xmlControlNode, "FIRSTPAGEPRINTERTRAY"
        xmlCopyAttribute xmlHostTemplateNode, xmlControlNode, "OTHERPAGESPRINTERTRAY"
        ' PSC 20/09/2004 BBG1438 - End
    
        ' AS 10/05/2005 CORE121
        xmlCopyAttribute xmlHostTemplateNode, xmlControlNode, "DELIVERYTYPE"
        
        ' AS 05/12/2006 CORE312 Gemini Printing
        xmlCopyAttribute xmlHostTemplateNode, xmlControlNode, "PRECREATEFORPACK"
        xmlCopyAttribute xmlHostTemplateNode, xmlControlNode, "GEMINIPRINTMODE"
        xmlCopyAttribute xmlHostTemplateNode, xmlControlNode, "GEMINIPRINTSTATUS"

        'TW SYS5866 6/12/2002
        'TW SYS6027 31/01/2003
        'If strDestination = "E" Then
        '   'TW SYS5580 Added 21/1/2003
        '            If strMethodName <> "" Then
        '
        '                If objPrintDataManagerBO Is Nothing Then
        '                    Set objPrintDataManagerBO = gobjContext.CreateInstance("omPDM.PrintDataManagerBO")
        '                End If
        '
        '                'Call Print Data Manager and load response string into a FreeThreadedDOMDocument
        '                xmlPDMResponseDoc.loadXML objPrintDataManagerBO.omRequest(vxmlRequestNode.xml)
        '                Debug.Print xmlPDMResponseDoc.xml
        '
        '                'CheckResponse for errors
        '                errCheckXMLResponse xmlPDMResponseDoc.xml, True, vxmlResponseNode
        '
        '            End If
        '        'End TW SYS5580 21/1/2003
        '            Call PrintDocumentCommonCode(vxmlRequestNode, vxmlResponseNode, xmlPDMResponseDoc, xmlHostTemplateNode, xmlPrintDataNode, strQueueName, strMethodName, "1")
                'Else
                'Send to print queue
        
        'Preparing Queue Request Element
        Set xmlElementReq = xmlQueueingDoc.createElement("REQUEST")
                
        Set xmlElementMQ = xmlQueueingDoc.createElement("MESSAGEQUEUE")
        xmlElementReq.appendChild xmlElementMQ
        
        xmlElementMQ.setAttribute "QUEUENAME", strQueueName
        
        ' ik_bmids00703
        ' xmlElementMQ.setAttribute "PROGID", "omPM.PrintManagerBO"
        'xmlElementMQ.setAttribute "PROGID", "BMPrinting.BMPrintRouter"
        
        strClientPrintingRouterPROGID = GetGlobalParamString("ClientPrintingRouterPROGID")
        xmlElementMQ.setAttribute "PROGID", strClientPrintingRouterPROGID
        'ik_bmids00703_ends
            
        ' ik_bmids00831
        ' need SEARCHKEY1 in CONTROLDATA for document group
        If Not vxmlRequestNode.selectSingleNode("CONTROLDATA/@TEMPLATEGROUPID") Is Nothing Then
            Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("SEARCHKEY1")
            xmlAttrib.Text = _
                vxmlRequestNode.selectSingleNode("CONTROLDATA/@TEMPLATEGROUPID").Text
            vxmlRequestNode.selectSingleNode("CONTROLDATA").Attributes.setNamedItem xmlAttrib
        End If
        ' ik_bmids00831_ends
               
        xmlSetAttributeValue vxmlRequestNode.selectSingleNode("../REQUEST"), "OPERATION", "SavePrintDocument"
        xmlElementMQ.setAttribute "XML", vxmlRequestNode.xml
           
        'Call SendToQueue
        'Intialise the object
        Set objMessageQueueing = gobjContext.CreateInstance("omPrint.omMessageQueue")
    
        ' IK_BM0200 traceAssist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[3.1]", "objMessageQueueing.SendToQueue call"
        
        Set vxmlResponseNode = objMessageQueueing.SendToQueue(xmlElementReq)
    
        ' IK_BM0200 traceAssist support
        gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[3.2]", "objMessageQueueing.SendToQueue return"
        
        'CheckResponse for errors
        errCheckXMLResponse vxmlResponseNode.xml, True
    
    Else
    
'        Only call PDM if MethodName supplied
        If strMethodName <> "" Then
        
            'CORE46  Check if Quotation Number has been supplied and, if it has, add it to the document name.
            Set xmlTemplateDataNode = xmlGetNode(vxmlRequestNode, "TEMPLATEDATA")
            If Not xmlTemplateDataNode Is Nothing Then
                strQuotationNumber = xmlGetAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
                If (strQuotationNumber <> "") Then
            
                    strTemplateName = xmlGetAttributeText(xmlHostTemplateNode, "HOSTTEMPLATENAME")
                
                    If (strTemplateName <> "") Then
                        strTemplateName = strTemplateName & " Quote No " & strQuotationNumber
                
                        xmlSetAttributeValue xmlHostTemplateNode, "HOSTTEMPLATENAME", strTemplateName
                    End If
                End If
            End If

            If objPrintDataManagerBO Is Nothing Then
                Set objPrintDataManagerBO = gobjContext.CreateInstance("omPDM.PrintDataManagerBO")
                ' IK_BM0200 traceAssist support
                On Error Resume Next
                gobjTrace.TraceInitialiseOffspring objPrintDataManagerBO
                On Error GoTo PrintDocumentExit
                ' IK_20030205 traceAssist ends
            End If

            'Call Print Data Manager and load response string into a FreeThreadedDomDocument40
    
            ' IK_BM0200 traceAssist support
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[4.1]", "objPrintDataManagerBO.omRequest call"
            
            xmlPDMResponseDoc.loadXML objPrintDataManagerBO.omRequest(vxmlRequestNode.xml)
            Debug.Print xmlPDMResponseDoc.xml
    
            ' IK_BM0200 traceAssist support
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[4.2]", "objPrintDataManagerBO.omRequest return"

            'CheckResponse for errors
            errCheckXMLResponse xmlPDMResponseDoc.xml, True, vxmlResponseNode

        End If
        
        Set xmlRequestContDataElem = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
        strDestinationType = xmlGetAttributeText(xmlRequestContDataElem, "DESTINATIONTYPE")
            
        ' AS 29/01/05 BBGRETAIL Also need to produce document if printing from local workstation.
        strPatternViewEdit = "//CONTROLDATA[@EDITBEFOREPRINT = '1' or @VIEWBEFOREPRINT = '1']"
        strPatternAction = "//REQUEST[@ACTION = 'DIP' or @ACTION = 'APP']"
        
        Set xmlViewOrEditNode = xmlGetNode(vxmlRequestNode, strPatternViewEdit)
        Set xmlActionNode = xmlGetNode(vxmlRequestNode, strPatternAction)
        
        'If one of the indicators is 1 then go get the merged template document
        If (Not xmlViewOrEditNode Is Nothing) Or (Not xmlActionNode Is Nothing) Or strDestinationType = "W" Then
            
            'Replace the PRINTDATA node with what is returned from PDM call
            Set xmlOldPrintDataNode = vxmlRequestNode.selectSingleNode("PRINTDATA")
            
            BuildPrintDataNode vxmlRequestNode, xmlPDMResponseDoc, strMethodName, xmlOldPrintDataNode
                        
            'Call ProcessMergeDocument
            Set objPrintManagerBO = gobjContext.CreateInstance("omPM.PrintManagerBO")
            ' IK_BM0200 traceAssist support
            gobjTrace.TraceInitialiseOffspring objPrintManagerBO
            
            
            'Create the correct request dependent upon destinationtype
            If strDestinationType = "L" Or strDestinationType = "R" Then
            
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "PRINTER")
                strCopies = xmlGetAttributeText(xmlRequestContDataElem, "COPIES")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("PRINTER")
                
                Set xmlCopiesElem = xmlTemporaryDoc.createElement("COPIES")
                If strCopies <> "" Then
                    xmlCopiesElem.Text = strCopies
                    xmlPrinterElem.appendChild xmlCopiesElem
                End If
                
                Set xmlPrinterNameElem = xmlTemporaryDoc.createElement("PRINTERNAME")
                xmlPrinterNameElem.Text = strPrinter
                xmlPrinterElem.appendChild xmlPrinterNameElem
                
                xmlOutputTypeElement.appendChild xmlPrinterElem
                
            ElseIf strDestinationType = "F" Then
                
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "PRINTER")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("FILE")
                xmlPrinterElem.Text = strPrinter
                xmlOutputTypeElement.appendChild xmlPrinterElem
                
            ' ik_bmids00831
            ' haven't a clue whether or not EMAIL, SMS work
            ' BMIDS component will handle FAX
            ElseIf strDestinationType = "X" Then
                
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "PRINTER")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("FAX")
                xmlPrinterElem.Text = strPrinter
                xmlOutputTypeElement.appendChild xmlPrinterElem
            
            ElseIf strDestinationType = "E" Then
                
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "PRINTER")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("EMAIL")
                xmlPrinterElem.Text = strPrinter
                xmlOutputTypeElement.appendChild xmlPrinterElem
                                
            ElseIf strDestinationType = "S" Then
                    
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "PRINTER")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("SMS")
                xmlPrinterElem.Text = strPrinter
                xmlOutputTypeElement.appendChild xmlPrinterElem
            ' ik_bmids00831_ends
            
            ElseIf strDestinationType = "W" Then
                ' AS 28/01/05 BBGRETAIL - added support for local workstation printing.
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "PRINTER")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("PRINTER")
                xmlPrinterElem.Text = strPrinter
                xmlOutputTypeElement.appendChild xmlPrinterElem
            
            'CORE00000080 GHun
            ElseIf strDestinationType = "EF" Then
                strPrinter = xmlGetAttributeText(xmlRequestContDataElem, "EMAILRECIPIENTTYPE")
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("EMAILFULFILMENT")
                Set xmlNode = GetEmailRecipient(xmlGetMandatoryAttributeText(xmlRequestContDataElem, "APPLICATIONNUMBER"), strPrinter)
                For Each xmlElem In xmlNode.selectNodes("RECIPIENT")
                    xmlPrinterElem.appendChild xmlElem
                Next
                Set xmlElem = xmlGetNode(xmlRequestContDataElem, "DOCUMENTLOCATION")
                If Not xmlElem Is Nothing Then
                    xmlPrinterElem.appendChild xmlElem
                End If
                xmlOutputTypeElement.appendChild xmlPrinterElem
            'CORE00000080 End
            
            'RF 16/02/2006 CORE245 Start - Allow printer destination of "Document Store Only"
            ElseIf strDestinationType = "DMS" Then
                Set xmlOutputTypeElement = xmlTemporaryDoc.createElement("OUTPUTTYPE")
                Set xmlPrinterElem = xmlTemporaryDoc.createElement("DMS")
                xmlPrinterElem.Text = "Document Store Only"
                xmlOutputTypeElement.appendChild xmlPrinterElem
            End If
            'RF 16/02/2006 CORE245 end
                       
            'Append the newly created output element to the controldata node
            xmlRequestContDataElem.appendChild xmlOutputTypeElement
            
            'Change the OPERATION Attribute to "ProcessMergeDocument"
            xmlSetAttributeValue vxmlRequestNode.selectSingleNode("../REQUEST"), "OPERATION", "ProcessMergeDocument"
            
            Set xmlPMResponseDoc = New FreeThreadedDOMDocument40
            xmlPMResponseDoc.async = False
            xmlPMResponseDoc.validateOnParse = False
            xmlPMResponseDoc.setProperty "NewParser", True

            'call across to PrintManager and load response string into a FreeThreadedDomDocument40
    
            ' IK_BM0200 traceAssist support
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[5.1]", "objPrintManagerBO.omRequest call"
            
            xmlPMResponseDoc.loadXML objPrintManagerBO.omRequest(vxmlRequestNode.xml)
    
            ' IK_BM0200 traceAssist support
            gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[5.2]", "objPrintManagerBO.omRequest return"
            
            'CheckResponse for errors
            errCheckXMLResponse xmlPMResponseDoc.xml, True, vxmlResponseNode
            
            'Append all the nodes from the document to the response node
            For Each xmlPMNode In xmlPMResponseDoc.documentElement.childNodes
                    vxmlResponseNode.appendChild xmlPMNode.cloneNode(True)
            Next
        Else
            'all code in this block moved to a subroutine
            Call PrintDocumentCommonCode(vxmlRequestNode, vxmlResponseNode, xmlPDMResponseDoc, xmlHostTemplateNode, xmlPrintDataNode, strQueueName, strMethodName, xmlGetAttributeText(vxmlRequestNode, "PRINTINDICATOR"))
        End If
    End If

PrintDocumentExit:
    
'TW 11/5/2005
'    Set objGlobalParameterBO = Nothing
'TW 11/5/2005
    Set objPrintDataManagerBO = Nothing
    Set objMessageQueueing = Nothing
    Set objPrintManagerBO = Nothing
    Set objDO = Nothing
    Set xmlPDMResponseDoc = Nothing
    Set xmlDoc = Nothing
    Set xmlHostTemplateDoc = Nothing
    Set xmlHostTemplateNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlAxWordNode = Nothing
    Set xmlElementReq = Nothing
    Set xmlElementMQ = Nothing
    Set xmlQueueingDoc = Nothing
    Set xmlViewOrEditNode = Nothing
    Set xmlRequestContDataElem = Nothing
    Set xmlOutputTypeElement = Nothing
    Set xmlTemporaryDoc = Nothing
    Set xmlPrinterElem = Nothing
    Set xmlPrinterNameElem = Nothing
    Set xmlCopiesElem = Nothing
    Set xmlPMNode = Nothing
    Set xmlOldPrintDataNode = Nothing
    Set xmlNewPrintDataNode = Nothing
    Set xmlPDMNode = Nothing
    Set xmlPMResponseDoc = Nothing
    Set xmlAttrib = Nothing
    
    Set xmlControlNode = Nothing        ' PSC 20/09/2004 BBG1438
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
           
Private Sub PrintDocumentCommonCode( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    xmlPDMResponseDoc As FreeThreadedDOMDocument40, _
    xmlHostTemplateNode As IXMLDOMNode, _
    xmlPrintDataNode As IXMLDOMNode, _
    strQueueName As String, _
    strMethodName As String, _
    strPrintIndicator As String)
            
    Const cstrFunctionName As String = "PrintDocumentCommonCode"
    On Error GoTo PrintDocCommonCodeExit
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    'Create request to send to Queue
    Dim xmlOutRequestElement As IXMLDOMElement
    Dim xmlOutControlDataElement As IXMLDOMElement
    Dim xmlOutputElement As IXMLDOMElement
    Dim xmlFileElement As IXMLDOMElement
    Dim xmlPrinterElement As IXMLDOMElement
    Dim xmlElement1 As IXMLDOMElement
    Dim xmlElement2 As IXMLDOMElement
    Dim xmlInControlDataNode As IXMLDOMNode
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTemplateNode As IXMLDOMNode
    Dim xmlPDMResponseNode As IXMLDOMNode
    Dim xmlElement3 As IXMLDOMElement
    Dim xmlElement4 As IXMLDOMElement
    Dim xmlQueueDoc As New FreeThreadedDOMDocument40
    
    Dim objMessageQueue As IomMessageQueue
    
    Dim strAttribute As String
    Dim strDestType As String
    Dim strRemoteLocation As String
    
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.async = False
    
    Set xmlOutRequestElement = xmlOutDoc.createElement("REQUEST")
    
    ' validate REQUEST & select CONTROLATA node
    Set xmlInControlDataNode = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
    
    Set xmlOutControlDataElement = xmlOutDoc.createElement("CONTROLDATA")
    xmlOutRequestElement.appendChild xmlOutControlDataElement
    
    ' copy DESTINATIONTYPE to outbound CONTROLDATA
    xmlCopyAttribute xmlInControlDataNode, xmlOutControlDataElement, "DESTINATIONTYPE"
    
    'DR DMSSYS0008 I have changed strPrinter to strRemoteLocation becuase that more closely
    'mirrors what the data is. I daren't change the XML though so I've left as is.
    
    strAttribute = xmlGetMandatoryAttributeText(xmlInControlDataNode, "DPSDOCUMENTID")
    Set xmlElement2 = xmlOutDoc.createElement("TEMPLATEID")
    xmlElement2.Text = strAttribute
    xmlOutControlDataElement.appendChild xmlElement2
    
    strAttribute = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    Set xmlElement2 = xmlOutDoc.createElement("AUTHORITYLEVEL")
    xmlElement2.Text = strAttribute
    xmlOutControlDataElement.appendChild xmlElement2
    
    strDestType = xmlGetAttributeText(xmlInControlDataNode, "DESTINATIONTYPE")
    'RF 16/02/2006 CORE245 Start
    'strRemoteLocation = xmlGetMandatoryAttributeText(xmlInControlDataNode, "PRINTER")
    If Not (Left$(strDestType, 1) = "D" Or strDestType = "EF") Then ' D = DMS Only
        strRemoteLocation = xmlGetMandatoryAttributeText(xmlInControlDataNode, "PRINTER")
    End If
    'RF 16/02/2006 CORE245 End
    
    Set xmlOutputElement = xmlOutDoc.createElement("OUTPUTTYPE")
    xmlOutControlDataElement.appendChild xmlOutputElement
    
    'SYS2305 Only create one destination child dependent on destType
    Select Case strDestType
        Case "F"
            
            Set xmlFileElement = xmlOutDoc.createElement("FILE")
            xmlOutputElement.appendChild xmlFileElement
            
            Set xmlElement2 = xmlOutDoc.createElement("OUTPUTDIRECTORY")
            xmlElement2.Text = strRemoteLocation
            xmlFileElement.appendChild xmlElement2
            
            'DR DMSSYS0008 We require the FILEID, use the applicationId.
            Set xmlElement2 = xmlOutDoc.createElement("FILEID")
            xmlElement2.Text = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("PRINTDATA"), "APPLICATIONNUMBER")
            xmlFileElement.appendChild xmlElement2
        
        Case "X"
        
            Set xmlElement2 = xmlOutDoc.createElement("FAX")
            'DR DMSSYS0008
            xmlElement2.Text = strRemoteLocation
            xmlOutputElement.appendChild xmlElement2
            
        Case "E"
        
            Set xmlElement2 = xmlOutDoc.createElement("EMAIL")
          
            'DR DMSSYS0008 DMS2 uses attributes for this information so ....
            xmlSetAttributeValue xmlElement2, "EMAILADDRESS", strRemoteLocation
            'DR DMSSYS0008 We require the FILEID, use the applicationId.
            xmlSetAttributeValue xmlElement2, "FILEID", xmlGetAttributeText(vxmlRequestNode.selectSingleNode("PRINTDATA"), "APPLICATIONNUMBER")
            xmlSetAttributeValue xmlElement2, "EMAILADMINISTRATOR", xmlGetAttributeText(vxmlRequestNode.selectSingleNode("CONTROLDATA"), "EMAILADMINISTRATOR")
            xmlOutputElement.appendChild xmlElement2
                        
        Case "S"
            
            Set xmlElement2 = xmlOutDoc.createElement("SMS")
            'DR DMSSYS0008
            xmlElement2.Text = strRemoteLocation
            xmlOutputElement.appendChild xmlElement2

        Case "L", "R", "W"
        
            Set xmlPrinterElement = xmlOutDoc.createElement("PRINTER")
            xmlOutputElement.appendChild xmlPrinterElement
            
            Set xmlElement2 = xmlOutDoc.createElement("PRINTERNAME")
            xmlElement2.Text = strRemoteLocation
            xmlPrinterElement.appendChild xmlElement2
    
            strAttribute = xmlGetAttributeText(xmlInControlDataNode, "COPIES")
            If strAttribute = "" Then
                strAttribute = "1"
            End If
            Set xmlElement2 = xmlOutDoc.createElement("COPIES")

            xmlElement2.Text = strAttribute
            xmlPrinterElement.appendChild xmlElement2
    
        'RF 16/02/2006 CORE245 Start
        Case "DMS", "D" ' Validation type "DMS" (= DMS Only) can be truncated to "D" in some routes
        
            Set xmlElement2 = xmlOutDoc.createElement("DMS")
            xmlElement2.Text = "Document Store Only"
            xmlOutputElement.appendChild xmlElement2
            
        'CORE00000080 GHun
        Case "EF"
            strRemoteLocation = xmlGetAttributeText(xmlInControlDataNode, "EMAILRECIPIENTTYPE")
            Set xmlElement2 = xmlOutDoc.createElement("EMAILFULFILMENT")
            Set xmlNode = GetEmailRecipient(xmlGetMandatoryAttributeText(xmlInControlDataNode, "APPLICATIONNUMBER"), strRemoteLocation)
            For Each xmlElement1 In xmlNode.selectNodes("RECIPIENT")
                xmlElement2.appendChild xmlElement1
            Next
            strRemoteLocation = xmlGetAttributeText(xmlInControlDataNode, "DOCUMENTLOCATION")
            If Len(strRemoteLocation) > 0 Then
                Set xmlNode = xmlOutDoc.createElement("DOCUMENTLOCATION")
                xmlNode.Text = strRemoteLocation
                xmlElement2.appendChild xmlNode
            End If
            xmlOutputElement.appendChild xmlElement2
        'CORE00000080 End
        

        'RF 16/02/2006 CORE245 End
        
    End Select
    
    Set xmlElement1 = xmlOutDoc.createElement("PRINTDATA")
    xmlOutRequestElement.appendChild xmlElement1
    BuildPrintDataNode vxmlRequestNode, xmlPDMResponseDoc, strMethodName, xmlElement1
    
    'Debug.Print xmlOutRequestElement.xml
    
    'DR Adding Mandatory attributes for DMS2
    Dim xmlRequestControlDataNode As IXMLDOMNode
    Dim xmlDocumentDetailElement As IXMLDOMElement
    Dim strApplicationNumber As String
    Dim strOperation As String
    
    'DR Magic Numbers I'm afraid.
    Const PRINT_DEFAULT_LANGUAGE = 3
    Const PRINT_DEFAULT_SOURCESYSTEM = "OMIGA4"
    
    'Initialise
    strApplicationNumber = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("PRINTDATA"), "APPLICATIONNUMBER")
    
    'BG 09/10/02 BMIDS00612 Check the new indicator instead of application number to set operation correctly.  Documents with ONMESSAGE
    'will simply print.  Documents with NEWDOCUMENT will print and archive.  PRINTINDICATOR is either set
    'in the GUI (PM010 and cm100) or in omTMBO.PrintDocumentsForTask which gets fired for automatic task
    'when moving to next stage.  Gui screens which do screens prints where there would be no app number - ie
    'TM020 and gn215 have not had an indicator attribute added - the code below will return an empty string below
    'and the correct operation will be set, namely ONMESSAGE, as these do not need to be archive as they are not
    'application specific prints.
    
'    If (strApplicationNumber = "") Then
'        strOperation = "ONMESSAGE"
'    Else
'        strOperation = "NEWDOCUMENT"
'    End If
'
'    strPrintIndicator = xmlGetAttributeText(vxmlRequestNode, "PRINTINDICATOR")
    If (strPrintIndicator = "1") Then
        strOperation = "NEWDOCUMENT"
    Else
        strOperation = "ONMESSAGE"
    End If
    
    'BG 09/10/02 BMIDS00612 END
    
    'DR Set up the standard request attributes
    xmlSetAttributeValue xmlOutRequestElement, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlOutRequestElement, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlOutRequestElement, "MACHINENAME", xmlGetAttributeText(vxmlRequestNode, "MACHINENAME")
    xmlSetAttributeValue xmlOutRequestElement, "CHANNELID", xmlGetAttributeText(vxmlRequestNode, "CHANNELID")
    xmlSetAttributeValue xmlOutRequestElement, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlOutRequestElement, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlOutRequestElement, "OPERATION", strOperation
    
    'DR Dont know what this is? It doesnt have a value and I dont understand why its on this node anyway!?!
    xmlSetAttributeValue xmlOutRequestElement, "DOCUMENTID", xmlGetAttributeText(vxmlRequestNode, "DOCUMENTID")
    
    Set xmlRequestControlDataNode = vxmlRequestNode.selectSingleNode("CONTROLDATA")
    
    If Not (xmlRequestControlDataNode Is Nothing Or strOperation = "ONMESSAGE") Then
        Set xmlDocumentDetailElement = xmlOutDoc.createElement("DOCUMENTDETAILS")
        
        'DR The naming conventions around here and the GUI are *@!#ing incrediable!
        'IK too true!
        
        ' IK_bmids00831, use (appropriate) values direct from HOSTTEMPLATE record
        If Not xmlHostTemplateNode Is Nothing Then
        
            xmlSetAttributeValue xmlDocumentDetailElement, "TEMPLATEID", xmlGetMandatoryAttributeText(xmlHostTemplateNode, "DPSTEMPLATEID")
            xmlSetAttributeValue xmlDocumentDetailElement, "HOSTTEMPLATEID", xmlGetMandatoryAttributeText(xmlHostTemplateNode, "HOSTTEMPLATEID")
            xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTNAME", xmlGetAttributeText(xmlHostTemplateNode, "HOSTTEMPLATENAME")
            xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTDESCRIPTION", xmlGetAttributeText(xmlHostTemplateNode, "HOSTTEMPLATEDESCRIPTION")
            xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTGROUP", xmlGetAttributeText(xmlHostTemplateNode, "TEMPLATEGROUPID")
            xmlSetAttributeValue xmlOutRequestElement.selectSingleNode("CONTROLDATA"), "SEARCHKEY1", xmlGetAttributeText(xmlHostTemplateNode, "TEMPLATEGROUPID")
         
            ' PSC 24/08/2004 BBG1233 - Start
            xmlSetAttributeValue xmlDocumentDetailElement, "DELIVERYTYPE", xmlGetAttributeText(xmlHostTemplateNode, "DELIVERYTYPE")
            xmlSetAttributeValue xmlDocumentDetailElement, "URLPOSTIND", xmlGetAttributeText(xmlHostTemplateNode, "URLPOSTIND")
            ' PSC 24/08/2004 BBG1233 - End
            
            ' PSC 20/09/2004 BBG1438 - Start
            xmlSetAttributeValue xmlDocumentDetailElement, "FIRSTPAGEPRINTERTRAY", xmlGetAttributeText(xmlHostTemplateNode, "FIRSTPAGEPRINTERTRAY")
            xmlSetAttributeValue xmlDocumentDetailElement, "OTHERPAGESPRINTERTRAY", xmlGetAttributeText(xmlHostTemplateNode, "OTHERPAGESPRINTERTRAY")
            ' PSC 20/09/2004 BBG1438 - End
                           
            'AS 04/01/2007 CORE327 Ensure GEMINIPRINTMODE is set.
            xmlSetAttributeValue xmlDocumentDetailElement, "PRECREATEFORPACK", xmlGetAttributeText(xmlHostTemplateNode, "PRECREATEFORPACK")
            xmlSetAttributeValue xmlDocumentDetailElement, "GEMINIPRINTMODE", xmlGetAttributeText(xmlHostTemplateNode, "GEMINIPRINTMODE")
        Else
        
            'Mandatory Info
            xmlSetAttributeValue xmlDocumentDetailElement, "TEMPLATEID", xmlGetMandatoryAttributeText(xmlRequestControlDataNode, "DPSDOCUMENTID")
            xmlSetAttributeValue xmlDocumentDetailElement, "HOSTTEMPLATEID", xmlGetMandatoryAttributeText(xmlRequestControlDataNode, "DOCUMENTID")
            xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTNAME", xmlGetAttributeText(xmlRequestControlDataNode, "HOSTTEMPLATENAME")
            xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTDESCRIPTION", xmlGetAttributeText(xmlRequestControlDataNode, "HOSTTEMPLATEDESCRIPTION")
            xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTGROUP", xmlGetAttributeText(xmlRequestControlDataNode, "TEMPLATEGROUPID")
    
            'DR - I may need the searchkey1 value placed into the controldata node
            If (xmlGetAttributeText(xmlRequestControlDataNode, "TEMPLATEGROUPID") <> "") Then
                'If it has a value then use it otherwise return everything
                xmlSetAttributeValue xmlOutRequestElement.selectSingleNode("CONTROLDATA"), "SEARCHKEY1", xmlGetAttributeText(xmlRequestControlDataNode, "TEMPLATEGROUPID")

            End If
            
            ' PSC 24/08/2004 BBG1233 - Start
            'CORE46  Add Delivery Type
            xmlSetAttributeValue xmlDocumentDetailElement, "DELIVERYTYPE", xmlGetAttributeText(xmlRequestControlDataNode, "DELIVERYTYPE")
            xmlSetAttributeValue xmlDocumentDetailElement, "URLPOSTIND", xmlGetAttributeText(xmlRequestControlDataNode, "URLPOSTIND") 'BBG452
            ' PSC 24/08/2004 BBG1233 - End
            
            ' PSC 20/09/2004 BBG1438 - Start
            xmlSetAttributeValue xmlDocumentDetailElement, "FIRSTPAGEPRINTERTRAY", xmlGetAttributeText(xmlRequestControlDataNode, "FIRSTPAGEPRINTERTRAY")
            xmlSetAttributeValue xmlDocumentDetailElement, "OTHERPAGESPRINTERTRAY", xmlGetAttributeText(xmlRequestControlDataNode, "OTHERPAGESPRINTERTRAY")
            ' PSC 20/09/2004 BBG1438 - End
            
            'AS 04/01/2007 CORE327 Ensure GEMINIPRINTMODE is set.
            xmlSetAttributeValue xmlDocumentDetailElement, "PRECREATEFORPACK", xmlGetAttributeText(xmlRequestControlDataNode, "PRECREATEFORPACK")
            xmlSetAttributeValue xmlDocumentDetailElement, "GEMINIPRINTMODE", xmlGetAttributeText(xmlRequestControlDataNode, "GEMINIPRINTMODE")
        End If
        
        'AS 04/01/2007 CORE327 Ensure GEMINIPRINTSTATUS is set.
        xmlSetAttributeValue xmlDocumentDetailElement, "GEMINIPRINTSTATUS", xmlGetAttributeText(xmlRequestControlDataNode, "GEMINIPRINTSTATUS")
        
        'Optional Info (?)
        xmlSetAttributeValue xmlDocumentDetailElement, "STAGEID", xmlGetAttributeText(xmlRequestControlDataNode, "STAGEID")
        xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTTYPE", xmlGetAttributeText(xmlRequestControlDataNode, "DOCUMENTTYPE")
        xmlSetAttributeValue xmlDocumentDetailElement, "CUSTOMERNO", xmlGetAttributeText(xmlRequestControlDataNode, "CUSTOMERNO")

        'DR DMSSYS0032
        xmlSetAttributeValue xmlDocumentDetailElement, "CUSTOMERNAME", xmlGetAttributeText(xmlRequestControlDataNode, "CUSTOMERNAME")
        xmlSetAttributeValue xmlDocumentDetailElement, "RECIPIENTNAME", xmlGetAttributeText(xmlRequestControlDataNode, "RECIPIENTNAME")
        
        'DR I copied the timestamp from omPM.PrintManagerBO.SaveDocument() So at least its consistent.
        xmlSetAttributeValue xmlDocumentDetailElement, "DOCUMENTCREATETIMESTAMP", Format$(Now(), "dd mmm yyyy hh:mm:ss")
        xmlSetAttributeValue xmlDocumentDetailElement, "LANGUAGE", PRINT_DEFAULT_LANGUAGE
        xmlSetAttributeValue xmlDocumentDetailElement, "SOURCESYSTEM", PRINT_DEFAULT_SOURCESYSTEM
        
        ' PSC 24/08/2004 BBG1233
        xmlCopyAttribute xmlRequestControlDataNode, xmlDocumentDetailElement, "POSTTOWEB"
          
        ' PSC 24/08/2004 BBG1572
        xmlCopyAttribute xmlRequestControlDataNode, xmlDocumentDetailElement, "WEBDOCUMENTTYPE"
          
        xmlOutRequestElement.selectSingleNode("CONTROLDATA").appendChild xmlDocumentDetailElement
        
        ' IK_BMIDS00831
        ' get additional DMS attributes
        
        If Len(xmlDocumentDetailElement.Attributes.getNamedItem("CUSTOMERNO").Text) = 0 Then
            If Not xmlPrintDataNode.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
                If Len(xmlPrintDataNode.Attributes.getNamedItem("CUSTOMERNUMBER").Text) > 0 Then
                    xmlSetAttributeValue _
                        xmlDocumentDetailElement, _
                        "CUSTOMERNO", _
                        xmlPrintDataNode.Attributes.getNamedItem("CUSTOMERNUMBER").Text
                End If

            End If
        End If
        
        If Len(xmlDocumentDetailElement.Attributes.getNamedItem("CUSTOMERNAME").Text) = 0 Then
            If xmlOutRequestElement.selectNodes("PRINTDATA/CUSTOMER/@CUSTOMERNAME").length = 1 Then
                xmlDocumentDetailElement.Attributes.getNamedItem("CUSTOMERNAME").Text = _
                    xmlOutRequestElement.selectSingleNode("PRINTDATA/CUSTOMER/@CUSTOMERNAME").Text

            End If
        End If
        
        If Len(xmlDocumentDetailElement.Attributes.getNamedItem("RECIPIENTNAME").Text) = 0 Then
            If xmlOutRequestElement.selectNodes("PRINTDATA/RECIPIENT/@RECIPIENTCOMPANYNAME").length = 1 Then
                xmlDocumentDetailElement.Attributes.getNamedItem("RECIPIENTNAME").Text = _
                    xmlOutRequestElement.selectSingleNode("PRINTDATA/RECIPIENT/@RECIPIENTCOMPANYNAME").Text
            End If
        End If
                  
    Else
        'BBG607 PrintCheque requires DeliveryType
        Dim strDeliveryType As String
        
        strDeliveryType = xmlGetAttributeText(xmlRequestControlDataNode, "DELIVERYTYPE")
        If Len(strDeliveryType) > 0 Then
            Set xmlDocumentDetailElement = xmlOutDoc.createElement("DOCUMENTDETAILS")
            xmlSetAttributeValue xmlDocumentDetailElement, "DELIVERYTYPE", strDeliveryType
            xmlOutRequestElement.selectSingleNode("CONTROLDATA").appendChild xmlDocumentDetailElement
        End If

    End If

    'DR End of DMS2 Changes
    '~~~~~~~~~~~~~~~~~~~~~~~~
   
    'Send to print queue
    'Preparing Queue Request Element
    Set xmlElement3 = xmlQueueDoc.createElement("REQUEST")
    
    Set xmlElement4 = xmlQueueDoc.createElement("MESSAGEQUEUE")
    xmlElement3.appendChild xmlElement4
    
    xmlElement4.setAttribute "QUEUENAME", strQueueName

    xmlElement4.setAttribute "PROGID", "omPM.PrintManagerBO"
    
    '   Set XML which has been built up from CONTROLDATA & PRINTDATA
    xmlElement4.setAttribute "XML", xmlOutRequestElement.xml
       
    'Call SendToQueue
    'Intialise the object
    Set objMessageQueue = gobjContext.CreateInstance("omPrint.omMessageQueue")
        
    Set vxmlResponseNode = objMessageQueue.SendToQueue(xmlElement3)
    'CheckResponse for errors
    errCheckXMLResponse vxmlResponseNode.xml, True

PrintDocCommonCodeExit:

    Set objMessageQueue = Nothing
    Set xmlOutControlDataElement = Nothing
    Set xmlOutputElement = Nothing
    Set xmlFileElement = Nothing
    Set xmlPrinterElement = Nothing
    Set xmlElement2 = Nothing
    Set xmlElement3 = Nothing
    Set xmlElement4 = Nothing
    Set xmlOutDoc = Nothing
    Set xmlInControlDataNode = Nothing
    Set xmlOutRequestElement = Nothing
    Set xmlElement1 = Nothing
    Set xmlQueueDoc = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

' AS 11/05/2005 CORE122 omPrint does not merge TEMPLATEDATA when edit/view before printing.
Private Function BuildPrintDataNode( _
    ByVal xmlRequestNode As IXMLDOMNode, _
    ByVal xmlPDMResponseDoc As FreeThreadedDOMDocument40, _
    ByVal strMethodName As String, _
    ByRef xmlPrintDataNode As IXMLDOMNode) As Boolean
    
On Error GoTo ExitPoint

    Dim xmlDataNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    BuildPrintDataNode = False
    
    If strMethodName = "" Then
        Set xmlDataNode = xmlRequestNode.selectSingleNode("TEMPLATEDATA")
    Else
        Set xmlDataNode = xmlPDMResponseDoc.selectSingleNode("RESPONSE")
    End If
    
    If Not xmlDataNode Is Nothing Then
        For Each xmlNode In xmlDataNode.childNodes
            xmlPrintDataNode.appendChild xmlNode.cloneNode(True)
        Next
    End If
    
    BuildPrintDataNode = True
    
ExitPoint:
    Set xmlDataNode = Nothing
    Set xmlNode = Nothing
End Function

' ik_BMIDS00885
Private Sub IOmPrintTxBO_RePrintDocument(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName As String = "RePrintDocument"
    
    ' ik_20030213
    gobjContext.SetComplete
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Dim strQueueName As String
    Dim strClientPrintingRouterPROGID  As String
    
    Dim objMessageQueue As IomMessageQueue

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
'TW 11/5/2005
    'Get Queue parameter
'    Dim objGlobalParameterBO As GlobalParameterBO
'    Set objGlobalParameterBO = gobjContext.CreateInstance("omBase.GlobalParameterBO")
'
'    Set xmlDoc = New FreeThreadedDOMDocument40
'    xmlDoc.async = False
'    xmlDoc.loadXML objGlobalParameterBO.GetCurrentParameter("PrintQueueName")
'    Set objGlobalParameterBO = Nothing
'
'    strQueueName = xmlGetNodeText(xmlDoc.documentElement, ".//STRING")
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.1]", "dbxmlGetCurrentParameterXML call"

    Set xmlNode = dbxmlGetCurrentParameterXML("PrintQueueName")
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.2]", "dbxmlGetCurrentParameterXML return"
    strQueueName = xmlGetNodeText(xmlNode, ".//STRING")
'TW 11/5/2005 End
    
#If IK_DEBUG_Q Then
    strQueueName = ".\PrintRequestsIK"
#End If
    
    Set xmlDoc = Nothing
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    Set xmlElem = xmlDoc.createElement("REQUEST")
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("MESSAGEQUEUE")
    xmlElem.setAttribute "QUEUENAME", strQueueName
    'xmlElem.setAttribute "PROGID", "BMPrinting.BMPrintRouter"
    strClientPrintingRouterPROGID = GetGlobalParamString("ClientPrintingRouterPROGID")
    xmlElem.setAttribute "PROGID", strClientPrintingRouterPROGID
    xmlElem.setAttribute "XML", vxmlRequestParentNode.xml
    xmlNode.appendChild xmlElem
    
    Set objMessageQueue = gobjContext.CreateInstance("omPrint.omMessageQueue")
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "objMessageQueue.SendToQueue call"
    
    Set xmlNode = objMessageQueue.SendToQueue(xmlDoc.firstChild)
    Set objMessageQueue = Nothing
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "objMessageQueue.SendToQueue return"
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    errCheckXMLResponseNode xmlNode, vxmlResponseNode, True

    Set xmlDoc = Nothing

End Sub
' ik_BMIDS00885_ends

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub IOmPrintTxBO_CreateCaseTask( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CASETASK entities.
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

    On Error GoTo CreateCaseTaskVbErr
    
    Const cstrFunctionName As String = "CreateCaseTask"
    
    ' ik_20030213
    gobjContext.SetComplete
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    

    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim objDO As IOmPrintDO
    
    Set xmlCaseTaskList = vxmlRequestParentNode.selectNodes("CASETASK")
        
    If xmlCaseTaskList.length > 0 Then
        ' call CreateCaseTask if CASETASK node in request
        Set objDO = gobjContext.CreateInstance("MsgTm.MsgTmDO")
        objDO.CreateCaseTask xmlCaseTaskList, vxmlResponseNode
    End If
    
    ' ik_20030213
    ' gobjContext.SetComplete
    
CreateCaseTaskExit:
    
    Set objDO = Nothing
    Set xmlCaseTaskList = Nothing
    
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateCaseTaskVbErr:

    gobjContext.SetAbort
    
    GoTo CreateCaseTaskExit

End Sub

Private Sub IOmPrintTxBO_PrintLetters(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo PrintLettersExit
    
    Const cstrFunctionName As String = "IOmPrintTxBO_PrintLetters"
    
    ' ik_20030213
    gobjContext.SetComplete
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    
    Dim xmlPDMResponseDoc As FreeThreadedDOMDocument40
    Set xmlPDMResponseDoc = New FreeThreadedDOMDocument40
    xmlPDMResponseDoc.async = False
    xmlPDMResponseDoc.validateOnParse = False
    xmlPDMResponseDoc.setProperty "NewParser", True
    
'TW 11/5/2005
    'Get Queue parameter
'    Dim objGlobalParameterBO As GlobalParameterBO
'    Dim strXMLResponse As String, strQueueName As String
'    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strClientPrintingRouterPROGID  As String
    Dim xmlNode As IXMLDOMNode
    
'    If objGlobalParameterBO Is Nothing Then
'        Set objGlobalParameterBO = gobjContext.CreateInstance("omBase.GlobalParameterBO")
'    End If
'
'    strXMLResponse = objGlobalParameterBO.GetCurrentParameter("PrintQueueName")
'    Set xmlDoc = xmlLoad(strXMLResponse, cstrFunctionName)
'    strQueueName = xmlGetNodeText(xmlDoc.documentElement, ".//STRING")
    
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.1]", "dbxmlGetCurrentParameterXML call"

    Dim strQueueName As String
    Set xmlNode = dbxmlGetCurrentParameterXML("PrintQueueName")
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[0.2]", "dbxmlGetCurrentParameterXML return"
    strQueueName = xmlGetNodeText(xmlNode, ".//STRING")
'TW 11/5/2005 End
    
#If IK_DEBUG Then
    strQueueName = ".\PrintRequestsIK"
#End If
    
    'Call Print Data manager
    'MO 5/11/2002 BMIDS00460
    'Dim objPrintDataManagerBO As PrintDataManagerBO
    Dim objPrintDataManagerBO As Object
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim strMethodName As String
    
    Set xmlPrintDataNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    strMethodName = xmlGetAttributeText(xmlPrintDataNode, "METHODNAME")
    
    If objPrintDataManagerBO Is Nothing Then
        Set objPrintDataManagerBO = gobjContext.CreateInstance("omPDM.PrintDataManagerBO")
        ' IK_BM0200 traceAssist support
        On Error Resume Next
        gobjTrace.TraceInitialiseOffspring objPrintDataManagerBO
        On Error GoTo PrintLettersExit
        ' IK_20030205 traceAssist ends
    End If

    'Call Print Data Manager and load response string into a FreeThreadedDomDocument40
    xmlPDMResponseDoc.loadXML objPrintDataManagerBO.omRequest(vxmlRequestNode.xml)
    Debug.Print xmlPDMResponseDoc.xml

    'CheckResponse for errors
    errCheckXMLResponse xmlPDMResponseDoc.xml, True, vxmlResponseNode
    
    'Create request to send to Queue
    Dim xmlElement As IXMLDOMElement
    Dim xmlControlDataElement As IXMLDOMElement
    Dim xmlOutputElement As IXMLDOMElement
    Dim xmlFileElement As IXMLDOMElement
    Dim xmlPrinterElement As IXMLDOMElement
    Dim xmlElement1 As IXMLDOMElement
    Dim xmlElement2 As IXMLDOMElement
    Dim xmlTempDoc As New FreeThreadedDOMDocument40
    Dim xmlControlDataNode As IXMLDOMNode
    
    ' PSC 02/12/2004 E2EM00003096 - Start
    Dim xmlDocDetails As IXMLDOMNode
    Dim xmlFileIdNode As IXMLDOMNode
    Dim xmlFileIdAttr As IXMLDOMAttribute
    Dim xmlAppNoAttr As IXMLDOMAttribute
    Dim strAppNo As String
    ' PSC 02/12/2004 E2EM00003096 - End

    
    Set xmlElement = xmlTempDoc.createElement("REQUEST")
    
    Set xmlElement1 = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
    If Not xmlElement1 Is Nothing Then
        Set xmlControlDataElement = xmlTempDoc.createElement("CONTROLDATA")
        xmlElement.appendChild xmlControlDataElement
        ' validate REQUEST & select CONTROLATA node
        Set xmlControlDataNode = xmlGetMandatoryNode(vxmlRequestNode, "CONTROLDATA")
        
        'DR DMSSYS0008 I have changed strPrinter to strRemoteLocation becuase that more closely
        'mirrors what the data is. I daren't change the XML though so I've left as is.
        Dim strAttribute, strDestType, strRemoteLocation As String
        strAttribute = xmlGetMandatoryAttributeText(xmlControlDataNode, "DPSDOCUMENTID")
        Set xmlElement2 = xmlTempDoc.createElement("TEMPLATEID")
        xmlElement2.Text = strAttribute
        xmlControlDataElement.appendChild xmlElement2
        
        strAttribute = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
        Set xmlElement2 = xmlTempDoc.createElement("AUTHORITYLEVEL")
        xmlElement2.Text = strAttribute
        xmlControlDataElement.appendChild xmlElement2
                
        strDestType = xmlGetMandatoryAttributeText(xmlControlDataNode, "DESTINATIONTYPE")
        strRemoteLocation = xmlGetMandatoryAttributeText(xmlControlDataNode, "PRINTER")
        
        ' PSC 02/12/2004 E2EM00003096 - Start
        Set xmlDocDetails = xmlControlDataNode.selectSingleNode("DOCUMENTDETAILS")
        If Not xmlDocDetails Is Nothing Then
            xmlControlDataElement.appendChild xmlDocDetails.cloneNode(True)
        End If
        ' PSC 02/12/2004 E2EM00003096 - End
        
        Set xmlOutputElement = xmlTempDoc.createElement("OUTPUTTYPE")
        xmlControlDataElement.appendChild xmlOutputElement
        
        'SYS2305 Only create one destination child dependent on destType
        If strDestType = "F" Then
                
                Set xmlFileElement = xmlTempDoc.createElement("FILE")
                xmlOutputElement.appendChild xmlFileElement
                
                Set xmlElement2 = xmlTempDoc.createElement("OUTPUTDIRECTORY")
                xmlElement2.Text = strRemoteLocation
                xmlFileElement.appendChild xmlElement2
                
                'DR DMSSYS0008 We require the FILEID, use the applicationId.
                ' PSC 02/12/2004 E2EM00003096 - Start
                Set xmlFileIdNode = xmlTempDoc.createElement("FILEID")
                xmlFileElement.appendChild xmlFileIdNode
                ' PSC 02/12/2004 E2EM00003096 - End
            
        ElseIf strDestType = "X" Then
            
                Set xmlElement2 = xmlTempDoc.createElement("FAX")
                'DR DMSSYS0008
                xmlElement2.Text = strRemoteLocation
                xmlOutputElement.appendChild xmlElement2
                
        ElseIf strDestType = "E" Then
            
                Set xmlElement2 = xmlTempDoc.createElement("EMAIL")
              
                'DR DMSSYS0008 DMS2 uses attributes for this information so ....
                xmlSetAttributeValue xmlElement2, "EMAILADDRESS", strRemoteLocation
                'DR DMSSYS0008 We require the FILEID, use the applicationId.
                
                ' PSC 02/12/2004 E2EM00003096 - Start
                Set xmlFileIdAttr = xmlTempDoc.createAttribute("FILEID")
                xmlElement2.Attributes.setNamedItem (xmlFileIdAttr)
                ' PSC 02/12/2004 E2EM00003096 - End
                xmlOutputElement.appendChild xmlElement2
                            
        ElseIf strDestType = "S" Then
                
                Set xmlElement2 = xmlTempDoc.createElement("SMS")
                'DR DMSSYS0008
                xmlElement2.Text = strRemoteLocation
                xmlOutputElement.appendChild xmlElement2

        ElseIf strDestType = "L" Or strDestType = "R" Then
            
                Set xmlPrinterElement = xmlTempDoc.createElement("PRINTER")
                xmlOutputElement.appendChild xmlPrinterElement
                
                Set xmlElement2 = xmlTempDoc.createElement("PRINTERNAME")
                xmlElement2.Text = strRemoteLocation
                xmlPrinterElement.appendChild xmlElement2
        
                strAttribute = xmlGetAttributeText(xmlControlDataNode, "COPIES")
                If strAttribute = "" Then
                    strAttribute = "1"
                End If
                Set xmlElement2 = xmlTempDoc.createElement("COPIES")
                xmlElement2.Text = strAttribute
                xmlPrinterElement.appendChild xmlElement2
        
        'CORE00000080 GHun
        ElseIf strDestType = "EF" Then
            strRemoteLocation = xmlGetAttributeText(xmlControlDataNode, "EMAILRECIPIENTTYPE")
            Set xmlElement2 = xmlTempDoc.createElement("EMAILFULFILMENT")
            Set xmlNode = GetEmailRecipient(xmlGetMandatoryAttributeText(xmlControlDataNode, "APPLICATIONNUMBER"), strRemoteLocation)
            For Each xmlElement1 In xmlNode.selectNodes("RECIPIENT")
                xmlElement2.appendChild xmlElement1
            Next
            strRemoteLocation = xmlGetAttributeText(xmlControlDataNode, "DOCUMENTLOCATION")
            If Len(strRemoteLocation) > 0 Then
                Set xmlNode = xmlTempDoc.createElement("DOCUMENTLOCATION")
                xmlNode.Text = strRemoteLocation
                xmlElement2.appendChild xmlNode
            End If
            xmlOutputElement.appendChild xmlElement2
        'CORE00000080 End
        End If
                             
    End If
    
    'This now has all the <CONTROLDATA> it needs
    Debug.Print xmlElement.xml
    
    'Send each application to the print queue
    Dim xmlQueueElement As IXMLDOMElement
    Dim xmlElement4 As IXMLDOMElement
    Dim xmlQueueDoc As New FreeThreadedDOMDocument40
    
    'Preparing Queue Request Element - this will be the same for every application
    'so we only need to do it once.
    
    Set xmlQueueElement = xmlQueueDoc.createElement("REQUEST")
    Set xmlElement4 = xmlQueueDoc.createElement("MESSAGEQUEUE")
    xmlQueueElement.appendChild xmlElement4
    
    xmlElement4.setAttribute "QUEUENAME", strQueueName
    xmlElement4.setAttribute "PROGID", "omPM.PrintManagerBO"
        
    ' ik_bmids00703
    ' xmlElement4.setAttribute "PROGID", "omPM.PrintManagerBO"
    'xmlElement4.setAttribute "PROGID", "BMPrinting.BMPrintRouter"
    strClientPrintingRouterPROGID = GetGlobalParamString("ClientPrintingRouterPROGID")
    xmlElement4.setAttribute "PROGID", strClientPrintingRouterPROGID
    ' ik_bmids00703_ends
    
    'DR Start of DMS2 Changes
    '~~~~~~~~~~~~~~~~~~~~~~~~
    Dim strApplicationNumber As String
    Dim strOperation As String

    'Initialise
    strApplicationNumber = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("PRINTDATA"), "APPLICATIONNUMBER")
    If (strApplicationNumber = "") Then
        strOperation = "ONMESSAGE"
    Else
        strOperation = "NEWDOCUMENT"
    End If

    'DR Set up the standard request attributes
    xmlSetAttributeValue xmlElement, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlElement, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlElement, "MACHINENAME", xmlGetAttributeText(vxmlRequestNode, "MACHINENAME")
    xmlSetAttributeValue xmlElement, "CHANNELID", xmlGetAttributeText(vxmlRequestNode, "CHANNELID")
    xmlSetAttributeValue xmlElement, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlElement, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlElement, "OPERATION", strOperation

    'DR Dont know what this is? It doesnt have a value and I dont understand why its on this node anyway!?!
    xmlSetAttributeValue xmlElement, "DOCUMENTID", xmlGetAttributeText(vxmlRequestNode, "DOCUMENTID")

    'DR End of DMS2 Changes
    '~~~~~~~~~~~~~~~~~~~~~~~~
    
    'Now - get list of all applications and loop through
'    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTemplateNode As IXMLDOMNode
    Dim xmlPDMResponseNode As IXMLDOMNode
    Dim objMessageQueue As IomMessageQueue

    'Intialise the object ready to send to queue
    Set objMessageQueue = gobjContext.CreateInstance("omPrint.omMessageQueue")
    
    Set xmlElement1 = xmlGetMandatoryNode(xmlPDMResponseDoc, "RESPONSE")
    Debug.Print xmlElement1.xml
    
    Set xmlElement2 = xmlTempDoc.createElement("PRINTDATA")
    xmlElement.appendChild xmlElement2
        
    Set xmlPDMResponseNode = xmlPDMResponseDoc.selectSingleNode("RESPONSE")
    If Not xmlPDMResponseNode Is Nothing Then
        'For each RateChangeBatchAudit node...
        For Each xmlNode In xmlPDMResponseNode.childNodes
        
            Set xmlTempNode = xmlNode.cloneNode(True)
            
            ' PSC 02/12/2004 E2EM00003096 - Start
            Set xmlAppNoAttr = xmlTempNode.selectSingleNode("descendant-or-self::*[@APPLICATIONNUMBER]/@APPLICATIONNUMBER")
            
            If Not xmlAppNoAttr Is Nothing Then
                If Not xmlFileIdNode Is Nothing Then
                    xmlFileIdNode.Text = xmlAppNoAttr.Value
                End If
                
                If Not xmlFileIdAttr Is Nothing Then
                    xmlFileIdAttr.Value = xmlAppNoAttr.Value
                End If
            End If
            ' PSC 02/12/2004 E2EM00003096 - End
              
            xmlElement2.appendChild xmlTempNode
            
            'Set XML which has been built up from CONTROLDATA & PRINTDATA for application
            xmlElement4.setAttribute "XML", xmlElement.xml
            
            Debug.Print xmlQueueElement.xml
            
            Set vxmlResponseNode = objMessageQueue.SendToQueue(xmlQueueElement)
            
            xmlElement2.removeChild xmlTempNode
            Set xmlTempNode = Nothing
            xmlElement4.setAttribute "XML", ""
                      
            
            'What are we going to do if this errors???
            'MORE ANALYSIS REQUIRED!!
            
        Next
    End If
       
    errCheckXMLResponse vxmlResponseNode.xml, True, vxmlResponseNode
            
PrintLettersExit:
    
'TW 11/5/2005
'    Set objGlobalParameterBO = Nothing
'TW 11/5/2005
    Set objPrintDataManagerBO = Nothing
    Set objMessageQueue = Nothing
    
    Set xmlControlDataElement = Nothing
    Set xmlOutputElement = Nothing
    Set xmlFileElement = Nothing
    Set xmlPrinterElement = Nothing
    Set xmlElement2 = Nothing
    Set xmlQueueElement = Nothing
    Set xmlElement4 = Nothing
    Set xmlTempDoc = Nothing
    Set xmlControlDataNode = Nothing
    
    Set xmlElement = Nothing
    Set xmlElement1 = Nothing
    Set xmlQueueDoc = Nothing
    
    ' PSC 02/12/2004 E2EM00003096 - Start
    Set xmlDocDetails = Nothing
    Set xmlFileIdNode = Nothing
    Set xmlFileIdAttr = Nothing
    Set xmlAppNoAttr = Nothing
    ' PSC 02/12/2004 E2EM00003096 - End

    
    ' ik_20030213
    ' gobjContext.SetComplete
    
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'CORE00000080 GHun
Private Function GetEmailRecipient(ByVal vstrApplicationNumber As String, ByVal vstrEmailRecipientType As String) As IXMLDOMElement
    Const cstrFunctionName  As String = "GetEmailRecipient"
    Dim objCRUD             As omCRUDBO
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

    If Not gobjContext Is Nothing Then
        Set objCRUD = gobjContext.CreateInstance("omCRUD.omCRUDBO")
    Else
        Set objCRUD = CreateObject("omCRUD.omCRUDBO")
    End If
    strResponse = objCRUD.omRequest(xmlRequest.xml)
        
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    If xmlDoc.selectSingleNode("RESPONSE/RECIPIENT") Is Nothing Then
        errThrowError cstrFunctionName, oeRecordNotFound, "No email address can be found for the recipient."
    End If
    
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
