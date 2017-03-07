VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PrintDataManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PrintDataManagerBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling
'Description:   Print Data Manager Business Object which 'supports transactions' only
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SA    26/02/02    SYS4176 Added functionality for new templates
'                          PaymentProcessingReportTemplate
'                          RateChangeReportTemplate
'                          RateChangeLetterTemplate
'                          ValuationReportTemplate
'SA    25/03/02    SYS4287 Rate Change Letters added
'SA    04/04/02    SYS4340 RateChangeReportTemplate method changed.
'SA    05/04/02    SYS4340 RateChangeLetterTemplate method changed.
'SA    05/04/02    SYS4340 ValuationReportTemplate and PayProcReportTemplate changed to deal with no records
'SA    15/04/02    SYS4379 Valuation & PayProc Reports not sending Total figures correctly.
'SA    16/04/02    SYS4412 RateChangeLetter not displaying New Rate correctly
'DRC   22/05/02    SYS471  call to checkxmlerror in OfferTemplate
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        AQR         Description
'MDC    14/06/2002  BMIDS00074  CMWP6 BM016 - Move BaseRateSet from MortgageProduct to InterestRateType
'AW     30/08/2002  BM029       CWP1    Added PrintChequeTemplate
'MV     27/03/2003  BM0469      Amended PipelineTemplate to include Declined and Cancelled APPS
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        AQR        Description
'MV     24/02/2004  BBG76      Chagned GetElementsByTagName to SelectNodes (MSXML 4.0 Parser Error)
'DM     29/04/2004             Changed the Transactioning level back to MTSTransactionMode  = 3  'UsesTransaction
'                              as this got change by accident. Not under any specific AQR get picked up by the next change.
'SJha   16/08/2004  BBG1228:   New Procedure SolsPanInstructTemplate.
'SJha   17/08/2004  BBG1237:   Updated procedure for  Employers Reference Letter Template.
'SJha   20/08/2004  BBG1253:   Updated procedure  "LenderReferenceTemplate"  for  New Lenders Reference Letter Template.
'SJha   23/08/2004  BBG1255:   Updated procedure  "LandlordReferenceTemplate"  for Landlord Reference Letter Template.
'SJha   26/08/2004  BBG1268:   New Procedure for Comparables Refer Letter Template
'IK     26/08/2004  BBG1288:   New Procedures for:
'                              DelayedCompletionFaxTemplate
'                              COTConfirmDetailsFaxTemplate
'                              SolicitorFaxEditableTemplate
'SDS    26/08/2004  BBG1232:   Offer Documentation
'IK     26/08/2004  BBG1293:   New Procedure for: DownVal Advise Letter
'IK     31/08/2004  BBG1312:   New Procedure for: Sols Panel Cancel Letter
'IK     31/08/2004  BBG1314:   Common Procedure for: 'DownVal Advise Letter' & 'Retention Advise Letter'
'SJha   01/09/2004  BBG1296:   New Procedure for Comparables Request Letter Template
'SJha   06/09/2004  BBG1327:   New Procedure for Reports Letter Template.
'SJha   06/09/2004  BBG1265:   Updated Format of Fax Number and also
'                              Combo value has been changed for the function isTaskPresent for Comparables Request Letter Template.
'SJha   08/09/2004  BBG1345:   Procedure for Employers Reference Chase Letter Template [Using existing procedure "EmployersRefTemplate"]
'BT     08/09/2004  BBG1358:   Direct Dip Accept and Post App Direct Decline templates
'BT     10/09/2004  BBG1363:   Pre Underwriting Direct template
'PSC    13/09/2004  BBG1405:   Amend STKFIQuotationTemplate and LTKFIQuotationTemplate to include OVERTURN_KFI
'SJha   13/09/2004  BBG1357:   Procedure for Landlord Reference Chase Letter Template [Using existing procedure "LandlordReferenceTemplate"]
'BT     14/09/2004  BBG1419:   New Packaged tasks need to be catered for.
'BT     14/09/2004  BBG1422:   New procedure for DIP Accept ST
'BT     14/09/2004  BBG1423:   New templates App form unsigned (broker & direct)
'BT     15/09/2004  BBG1428:   New Offer Expiry Fax template and Change of Details
'SJha   15/09/2004  BBG1414:   Procedure for Lender Reference Chase Letter Template [Using existing procedure "LenderReferenceTemplate"]
'IK     01/09/2004  BBG1318:   New Procedure for: Direct Decline Letter
'IK     15/09/2004  BBG1425:   New Procedure for: COT Chase Fax
'INR    17/09/2004  BBG1437    Add Std & LT Offer template method calls
'TK     21/09/2004  BBG1159    Add PrintChequeTemplate method call
'SR     23/09/2004  CORE82     Modified methods STKFIQuotationTemplate, STOfferTemplate, LTOfferTemplate, LTKFIQuotationTemplate
'BT     27/09/2004  BBG1472    AddCoverLetterDetails is not formatting the single line address correctly
'INR    28/09/2004  BBG1489    Non Regulated KFI template method calls.
'BT     07/10/2004  E2E2105    Various changes to allow wildcards to stored procs
'IK     07/10/2004  E2EM00002106, various fixes to PreUndBrokerLetterTemplate & PreUnderwriterTasks
'IK     18/10/2004  E2EM00002257, various fixes to ComparablesRequestTemplate
'IK     18/10/2004  E2EM00002310, various fixes to ComparablesReferTemplate
'IK     18/10/2004  E2EM00002311, various fixes to FaxDealNotAvailableTemplate
'IK     18/10/2004  E2EM00002312, various fixes to DipResultFaxTemplate
'BT     21/10/2004  BBG1668    Bank of England rate change template
'BT     26/10/2004  BBG1707    Multiple applicant cover letters require major change to KFI TEMPLATEDATA
'BT     02/04/2004  BBG1752    Handle autonumbering sections manually via UpdateKFISectionNos routine
'IK     08/11/2004  E2EM00002892 - NewMaxLoanLetterTemplate, INCENTIVETYPE, INCENTIVETYPE from NEWPROPERTYDEPOSIT
'IK     17/11/2004  E2EM00002105 - get full APPLICATIONINTERMEDIARY details
'IK     17/11/2004  E2EM00002311 - FaxDealNotAvailableTemplate, add TEAMNAME etc.
'IK     18/11/2004  E2EM00002257 - ComparablesRequestTemplate, Packager Contact Name missing
'IK     19/11/2004  E2EM00002312 -
'                   DipResultFaxTemplate, Packager Contact Name missing
'                   Ltd Company name required
'IK     23/11/2004  E2EM00002759 - KFILetterTemplate
'                   title not picked up for packager contact
'IK     24/11/2004  BBG1786 - SolsPanAckLetterTemplate
'                   solicitor acting for lender required
'IK     24/11/2004  E2EM00002257 - ComparablesRequestTemplate, Packager Firm Name missing
'IK     24/11/2004  E2EM00002105 - SolsPanInstructTemplate
'                   not returning existing lender from NAMEANDADDRESSDIRECTORY
'IK     29/11/2004  E2EM00002105 - SolsPanInstructTemplate
'                   existing account details when no MORTGAGELOAN
'                   make sure we get first applicant
'IK     30/11/2004  E2EM00002105 - SolsPanInstructTemplate
'                   not solicitor acting for lender
'PSC    02/12/2004  E2EM00003096 Rewrite BOERateChangeTemplate
'IK     06/12/2004  E2EM00003089 - PreUndBrokerLetterTemplate, PreUndDirectLetterTemplate
'                   need completed tasks as well as incomplete tasks
'IK     06/12/2004  E2EM00002312 - Limited Company risk assessment code translation
'IK     06/12/2004  E2EM00003117 - Pre Underwriter letters need latest TEAMNAME
'                   (this may turn out to be a more generalised problem)
'IK     14/12/2004  E2EM00002921 - ADDRESSEENAME & SALUTATION on offer cover letters
'                   to include multiple applicants at same address
'IK     23/12/2004  E2EM00002312 - risk assessment paragraphs only when rule failed
'IK     05/01/2005  E2EM00002921 -
'                   ADDRESSEENAME & SALUTATION on offer cover letters refined
'                   return of SECTIONNO attributes
'                   add RISK attributes to Lifetime Section 7 child nodes
'IK     06/01/2005  BBG1852 - single cover letter for Ltd. Co. applications
'                   (incorporated in 05/01/2005 E2EM00002921 changes)
'IK     06/12/2004  E2EM00003201 - E2EM00003089 missed PreUndDirectLetterTemplate
'JD     12/01/2005  E2EM00003200 Changed task spellings where marked (PostAppDirDelineLetter)
'HMA    13/01/2005  E2EM00002106 Changes to PreUndBrokerLetterTemplate
'JD     17/01/2005  E2EM00003199 Changes to DirectDecline Letter
'TK     18/01/2005  BBG1898 Changes to STOfferTemplate
'JD     18/01/2005  E2EM00003200 PostAppDirDeclineLetter change to BrokerOrPackagerName for Individual
'TK     19/01/2005  BBG1901 Changes to LTOfferTemplate
'JD     19/01/2005  E2EM00002982 should use customers solicitor for lifetime products
'JD     19/01/2005  E2EM00003228 SolsPanAckTemplate : If there is only one legalrep don't delete it even if @SEPARATELEGALREPRESENTATIVE != '1'
'HMA    20/01/2005  E2EM00002893 Change to Pre Underwriting template processing.
'HMA    20/01/2005  BBG1889      Change task names to agree with BBG Live configuration.
'IK     21/01/2005  E2EM00003198 - Direct DIP Accept, incorrect RISKASSESSMENT node search
'TK     21/01/2005  E2EM00002982 - Always use customers solicitor for Lifetime products
'TK     28/01/2005  BBG1921 - Roll back changes done for E2EM00002982
'HMA    24/01/2005  E2EM00003199 - Update DirectDeclineLetterTemplate for new version.
'HMA    31/01/2005  BBG1889      - Update templates for applicant tasks.
'HMA    03/02/2005  BBG1925      - Update offer templates.
'TK     03/02/2005  E2EM00002312 - Amend AddRACodes - populate LC29 and LC06 for Accept RaRuleResult - commented out unused objects/variables
'MV     09/02/2005  BBG1923      - Added BanksReftemplate method calls
'HMA    08/02/2005  E2EM00003322 - Display name of Limited Company on Offer letters.
'HMA    08/02/2005  BBG1852      - Display Names of Directors in Offer Letter for Limited Companies.
'HMA    28/02/2005  BBG1993      - Update SolsPanelInstructTemplate.
'HMA    02/03/2005  E2EM00003288 - Update DipResultFaxTemplate
'TK     02/03/2005  E2EM00002312 - Roll back changes for populating LC29 and LC06 for Accept RaRuleResult
'TK     03/03/2005  BBG2003      - Amended ReportLetterTemplate()
'HMA    07/03/2005  E2EM00003189 - Changes to ComparableRequestTemplate
'TK     09/03/2005  BBG2020      - Amended SolsPanInstructTemplate()
'AS     21/06/2005  BBGR403 Added tracing.
'TW     11/07/2005  BBGR832      - Extract USERID and UNITID from the request block and pass these to PrepareTemplateCRUD
'TW     18/10/2005  MAR223       - Changes for MARS project to support Pack processing
'TW                                Also removed unused methods
'BC     25/01/2006  MAR1128/35/36  Remove AddCoverLetterDetails processing
'BC     23/02/2006  MAR1320        Ensure that Offer document is added to the response xml at the end of STOfferTemplate
'
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Baseline History:
'
'Prog   Date        AQR        Description
'HMA    26/04/2004  CORE46     Added functions to deal with KFI documents.
'------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'EPSOM History:
'
'Prog   Date        AQR        Description
'INR    08/02/2007  EP2_583    Added to Request for KFI.
'INR    05/04/2007  EP2_2203 Check the response for the correct error rather than making up an unspecified error to return

Option Explicit
Implements ObjectControl
Private gobjContext As ObjectContext

Private Function PDMReplacement(ByVal strMethod As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As Boolean
' Header ----------------------------------------------------------------------------------
' Description:
'   Performs the equivalent of PrintDocument for Print Templates which use the
'   KFIEngine for processing.
'
' Request broker, private interface, performs single REQUEST OPERATION
'   Any results of operation (data retrieved, error info etc.) appended to vxmlResponseNode
' Pass:
'   strMethod
'       PDM Method
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' Return:
'       True if Template has been processed
'       False if Template must be processed by PrintDocument
' Exceptions:
'       None.
' History:
'
' TW    6/4/2005 Created for BBG Retail
' TW    17/6/2005 BBGR366 Modified for use of omCRUD
' TW    11/7/2005 BBGR832 Extract USERID and UNITID from the request block and pass these to PrepareTemplate
' TW    18/10/2005 MAR223 Changes for MARS project to support Pack processing
'------------------------------------------------------------------------------------------
Const strFunctionName As String = "PDMReplacement"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim fs As FileSystemObject

Dim intApplicationFactFindNumber As Integer
Dim intCustomerVersionNumber As Integer

Dim PseudoPDM As PrepareTemplate.PrepareTemplateData

Dim strApplicationNumber As String
Dim strCustomerNumber As String
Dim strEnginePath As String
Dim strGlobalParamName As String
Dim strPath As String
Dim strSuccess As String
Dim strTemplateName As String
Dim strTemplatePath As String
Dim strXMLPath As String
Dim strXMLResponse As String
Dim strXSLPath As String
Dim strUserId As String
' TW 11/7/2005 BBGR832
Dim strUnitId As String
' TW 11/7/2005 BBGR832 End

Dim strUseCRUD As String

Dim xmlAttribute As IXMLDOMAttribute
Dim xmlElem As IXMLDOMElement
Dim xmlEngineData As FreeThreadedDOMDocument40
Dim xmlMethodNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlPrintDataNode As IXMLDOMNode
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlTempResponseDoc As FreeThreadedDOMDocument40

Dim X As Integer
    
    On Error GoTo PDMReplacementExit
    
    Set fs = New FileSystemObject
    
    strPath = App.Path
    X = InStrRev(strPath, "\")
    strPath = Left$(strPath, X)
    strXMLPath = strPath & "XML\"
    strEnginePath = strXMLPath & "NewEngine.xml"

    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.0]", "XML Path: " & strXMLPath
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "EnginePath: " & strEnginePath

' Check for the existence of the New Engine xml
    If fs.FileExists(strEnginePath) Then
        Set xmlEngineData = New FreeThreadedDOMDocument40
        xmlEngineData.Load strEnginePath
    
' Check if there is an entry in the New Engine xml for the PDM Method
        Set xmlMethodNode = xmlEngineData.selectSingleNode("//TEMPLATE[@PDMMethod='" & UCase$(strMethod) & "']")
        If Not xmlMethodNode Is Nothing Then
            strTemplateName = GetAttributeValue(xmlMethodNode, "Name")
' TW    17/6/2005 BBGR366
'            strXSLPath = strXMLPath & strTemplateName & ".xsl"
            strXSLPath = strXMLPath & strMethod & ".xsl"
' TW    17/6/2005 BBGR366 End
            strTemplatePath = strPath & "Printing\Templates\" & strTemplateName & ".rtf"

            gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.2]", "XSL Path: " & strXSLPath
            gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.3]", "Template Path: " & strTemplatePath

' Check that the template and associated xsl exist
            If fs.FileExists(strXSLPath) And fs.FileExists(strTemplatePath) Then
            
                Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
                
' Get the parameters
                strApplicationNumber = GetAttributeValue(xmlPrintDataNode, "APPLICATIONNUMBER")
                intApplicationFactFindNumber = Val(Format$(GetAttributeValue(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")))
                strCustomerNumber = GetAttributeValue(xmlPrintDataNode, "CUSTOMERNUMBER")
                intCustomerVersionNumber = Val(Format$(GetAttributeValue(xmlPrintDataNode, "CUSTOMERVERSIONNUMBER")))
' TW 11/7/2005 BBGR832
'                strUserId = ""
                strUserId = GetAttributeValue(vxmlRequestNode, "USERID")
                strUnitId = GetAttributeValue(vxmlRequestNode, "UNITID")
'                strGlobalParamName = ""
' TW 11/7/2005 BBGR832 End
                
                If Not gobjContext Is Nothing Then
                    Set PseudoPDM = gobjContext.CreateInstance("PrepareTemplate.PrepareTemplateData")
                Else
                    Set PseudoPDM = New PrepareTemplate.PrepareTemplateData
                End If
                
' Process the template
' TW    17/6/2005 BBGR366
'                strXMLResponse = PseudoPDM.PopulateTemplate(strMethod, strTemplatePath, strXSLPath, strApplicationNumber, intApplicationFactFindNumber, strCustomerNumber, intCustomerVersionNumber, strPrimaryXML, strUserId, strGlobalParamName)
                strUseCRUD = UCase$(GetAttributeValue(xmlMethodNode, "UseCRUD"))
                gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.4]", "PseudoPDM.PopulateTemplate(" & strMethod & ", " & strTemplatePath & ", " & strXSLPath & ", " & strApplicationNumber & ", " & CStr(intApplicationFactFindNumber) & ", " & strCustomerNumber & ", " & CStr(intCustomerVersionNumber) & ", " & vxmlRequestNode.xml & ", " & strUserId & ", " & strUserId & ", " & strUseCRUD & ")"
                On Error Resume Next
                gobjTrace.TraceInitialiseOffspring PseudoPDM
                On Error GoTo PDMReplacementExit
' TW 11/7/2005 BBGR832
'                strXMLResponse = PseudoPDM.PopulateTemplate(strMethod, strTemplatePath, strXSLPath, strApplicationNumber, intApplicationFactFindNumber, strCustomerNumber, intCustomerVersionNumber, strPrimaryXML, strUserId, strGlobalParamName, (strUseCRUD = "Y"))
' TW 18/10/2005 MAR223
                strXMLResponse = PseudoPDM.PopulateTemplate(strMethod, strTemplatePath, strXSLPath, strApplicationNumber, intApplicationFactFindNumber, strCustomerNumber, intCustomerVersionNumber, vxmlRequestNode.xml, strUserId, strUnitId, (strUseCRUD = "Y"))
' TW 18/10/2005
                gobjTrace.TraceXML strXMLResponse, "PseudoPDM_response"
' TW 11/7/2005 BBGR832 End
' TW    17/6/2005 BBGR366
                
                Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
                xmlTempResponseDoc.setProperty "NewParser", True
                xmlTempResponseDoc.LoadXML strXMLResponse
    
                If xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
                    strSuccess = "ERROR"
                Else
                    strSuccess = "SUCCESS"
                    PDMReplacement = True
                End If
                gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.5]", "PseudoPDM.PopulateTemplate = " & strSuccess

                Set xmlResponseDoc = New FreeThreadedDOMDocument40
                Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
                xmlElem.setAttribute "TYPE", strSuccess
                Set xmlNode = xmlResponseDoc.appendChild(xmlElem)

                CheckXMLResponse xmlResponseDoc.xml, True, vxmlResponseNode
    
                vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")
            End If
        End If
    End If
   
PDMReplacementExit:
    Set fs = Nothing

    Set PseudoPDM = Nothing
    
    Set xmlAttribute = Nothing
    Set xmlElem = Nothing
    Set xmlEngineData = Nothing
    Set xmlMethodNode = Nothing
    Set xmlNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
End Function

Private Function GetAttributeValue(xmlNode As IXMLDOMNode, strAttributeName As String) As String
'TW 5/4/2005
Dim xmlAttribute As IXMLDOMAttribute
    Set xmlAttribute = xmlNode.Attributes.getNamedItem(strAttributeName)
    If xmlAttribute Is Nothing Then
        GetAttributeValue = ""
    Else
        GetAttributeValue = xmlAttribute.Text
    End If
    Set xmlAttribute = Nothing
'End TW 5/4/2005
End Function
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    gobjTrace.StartTrace
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Public Function OmRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Request, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
' history: BG 14/03/01 Created for sys2096
'------------------------------------------------------------------------------------------
    On Error GoTo OmRequestVbErr
    Const strFunctionName As String = "OmRequest"
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    ' will be superceded by any subsequent SetAbort
    GetObjectContext.SetComplete
    ' create REQUEST & RESPONSE FreeThreadedDOMDocument40s
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlInDoc.async = False
    xmlOutDoc.async = False
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.LoadXML vstrXMLRequest
    
    ' check for request node
    Set xmlRequestNode = GetMandatoryNode(xmlInDoc, "REQUEST")
    CheckMandatoryAttribute xmlRequestNode, "OPERATION"
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        ' multiple operations
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                ' copy attributes from REQUEST node to OPERATION node
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If
    ' if we get here, everything is OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    OmRequest = xmlOutDoc.xml
    Debug.Print OmRequest

OmRequestExit:
    
    ' clean up ================================================================
    ' release XML DOM objects
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    gobjTrace.TraceResponse OmRequest
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function
OmRequestVbErr:
    
    ' create ERROR response block from Err object
    OmRequest = CreateErrorResponse()
    gobjTrace.TraceIdentErrorResponse OmRequest
    Debug.Print OmRequest
    ' perform clean up
    GoTo OmRequestExit
End Function
Private Sub DoRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
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
' history: BG 14/03/01 Created for sys2096
'------------------------------------------------------------------------------------------
    On Error GoTo DoRequestExit
    Const strFunctionName As String = "DoRequest"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    Dim strOperation As String
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    strOperation = UCase$(strOperation)
    Select Case strOperation
        Case "PRINTDOCUMENT"
            PrintDocument vxmlRequestNode, vxmlResponseNode
        'SYS4287 For Ratechange letters, we being called from PrintLEtters
        Case "PRINTLETTERS"
            PrintDocument vxmlRequestNode, vxmlResponseNode
                
        Case Else
            Err.Raise _
                eNOTMIPLEMENTED, _
                strFunctionName, _
                vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text & _
                " Not implemented"
    End Select
DoRequestExit:
    
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
    CheckError strFunctionName, TypeName(Me)
End Sub
Private Sub PrintDocument(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
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
' history: BG 14/03/01 Created for sys2096
'------------------------------------------------------------------------------------------
    On Error GoTo PrintDocumentExit
    Const strFunctionName As String = "PrintDocument"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim strMethod As String
    Dim xmlPrintDataNode As IXMLDOMNode
    
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    strMethod = GetMandatoryAttributeText(xmlPrintDataNode, "METHODNAME")
    
    strMethod = UCase$(strMethod)
    
'TW 6/4/2005
    If Not PDMReplacement(strMethod, vxmlRequestNode, vxmlResponseNode) Then
'End TW 6/4/2005
    
        Select Case strMethod
            Case "STKFIQUOTATIONTEMPLATE"
                STKFIQuotationTemplate vxmlRequestNode, vxmlResponseNode
            Case "LTKFIQUOTATIONTEMPLATE"
                LTKFIQuotationTemplate vxmlRequestNode, vxmlResponseNode
            Case "STOFFERTEMPLATE"
                STOfferTemplate vxmlRequestNode, vxmlResponseNode
            Case "LTOFFERTEMPLATE"
                LTOfferTemplate vxmlRequestNode, vxmlResponseNode
            Case "TEOFFERTEMPLATE"
                TEOfferTemplate vxmlRequestNode, vxmlResponseNode
            Case "NRKFIQUOTATIONTEMPLATE"
                NRKFIQuotationTemplate vxmlRequestNode, vxmlResponseNode
            Case Else
                Err.Raise _
                    eNOTMIPLEMENTED, _
                    strFunctionName, _
                    vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text & _
                    " Not implemented"
        End Select
       
    '#If BT_DEBUG Then
    '    Dim xmltmpDoc As DOMDocument
    '    Set xmltmpDoc = New DOMDocument
    '    xmltmpDoc.async = False
    '    xmltmpDoc.LoadXML vxmlResponseNode.xml
    '
    '    Dim fso As Object
    '    Dim tname As String
    '    Dim sPath As String
    '    Set fso = CreateObject("Scripting.FileSystemObject")
    '    tname = fso.GetTempName
    '    tname = Left(tname, Len(tname) - 4)
    '    sPath = App.Path & "\debug"
    '    If Not fso.FolderExists(sPath) Then
    '        fso.CreateFolder sPath
    '    End If
    '    xmltmpDoc.save sPath & "\omPDM_" & tname & ".xml"
    '    Set fso = Nothing
    '#End If
'TW 6/4/2005
    End If
'End TW 6/4/2005
    Set xmlPrintDataNode = Nothing
    Exit Sub

PrintDocumentExit:
    Set xmlPrintDataNode = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    CheckError strFunctionName, TypeName(Me)

End Sub



Private Sub GetUniqueCustomerAddress(ByVal strAppNo As String, _
                                     ByVal strApplicationFactFindNumber As String, _
                                     ByRef strCustomerArray() As String, _
                                     ByRef intNumAddresses As Integer)
    '
    ' Subroutine GetUniqueCustomerAddress is a routine to cater for the KFI cover letter
    ' requirement to print a cover letter for each unique address in the applicants list
    '
    ' This will return an array of up to 4 customer numbers which need to be passed into
    ' the build cover letter routines
    '
    On Error GoTo GetUniqueCustomerAddressError
    
    Dim xmlAPPRequestDoc As FreeThreadedDOMDocument40
    Dim xmlAPPElem As IXMLDOMElement
    Dim xmlAppNode As IXMLDOMNode

    'Create New Request
    Set xmlAPPRequestDoc = New FreeThreadedDOMDocument40
    xmlAPPRequestDoc.async = False

    Set xmlAPPElem = xmlAPPRequestDoc.createElement("REQUEST")
    xmlAPPElem.setAttribute "COMBOLOOKUP", "Y"

'   Use named omPDM SCHEMA file
    xmlAPPElem.setAttribute "RB_SCHEMA", "omPDM"
    Set xmlAppNode = xmlAPPRequestDoc.appendChild(xmlAPPElem)

    'Create element to search ALLAPPLICATIONDETAILS scheme
    'but only go down as far as customer address as that's all we need.
    Dim xmlallappdetsnode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmltemp2node As IXMLDOMNode

    Set xmlAPPElem = xmlAPPRequestDoc.createElement("ALLAPPLICATIONDETAILS")
    xmlAPPElem.setAttribute "_SCHEMA_", "ALLAPPLICATIONDETAILS"
    xmlAPPElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlAPPElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    Set xmlallappdetsnode = xmlAppNode.appendChild(xmlAPPElem)

    Set xmlAPPElem = xmlAPPRequestDoc.createElement("APPLICATIONCUSTOMERROLE")
    xmlAPPElem.setAttribute "CUSTOMERROLETYPE", "1"
    Set xmlTempNode = xmlallappdetsnode.appendChild(xmlAPPElem)

    Set xmlAPPElem = xmlAPPRequestDoc.createElement("CUSTOMER")
    Set xmltemp2node = xmlTempNode.appendChild(xmlAPPElem)

    Set xmlAPPElem = xmlAPPRequestDoc.createElement("CUSTOMERADDRESS")
    xmlAPPElem.setAttribute "_EXTRACTTYPE_", "NODE"
    xmltemp2node.appendChild xmlAPPElem

    'Create New Response
    Dim objRBDO As Object
    Dim xmlAPPResponseDoc As FreeThreadedDOMDocument40

    Set xmlAPPResponseDoc = New FreeThreadedDOMDocument40

    'Call Request Broker and load response string into a FreeThreadedDOMDocument40
    xmlAPPResponseDoc.setProperty "NewParser", True
    xmlAPPResponseDoc.validateOnParse = False

    ' Request Data via Request Broker
    Set objRBDO = GetObjectContext.CreateInstance("omRB.OmRequestDO")
    xmlAPPResponseDoc.LoadXML objRBDO.OmDataRequest(xmlAPPRequestDoc.xml)
    Set objRBDO = Nothing

    Set xmltemp2node = Nothing
    Set xmlTempNode = Nothing
    Set xmlallappdetsnode = Nothing

    Dim xmlCustNodes As IXMLDOMNodeList
    Set xmlCustNodes = xmlAPPResponseDoc.selectNodes("RESPONSE/ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1']/CUSTOMER")

    Dim xmlUniqueDoc As FreeThreadedDOMDocument40
    Dim xmlUniqueElem As IXMLDOMElement
    Dim xmlUniqueNode As IXMLDOMNode

    'Create New Request
    Set xmlUniqueDoc = New FreeThreadedDOMDocument40
    xmlUniqueDoc.async = False

    Set xmlUniqueElem = xmlUniqueDoc.createElement("ADDRESS")
    Set xmlUniqueNode = xmlUniqueDoc.appendChild(xmlUniqueElem)

    Dim xmlCust As IXMLDOMNode
    Dim xmlAddress As IXMLDOMNode
    For Each xmlCust In xmlCustNodes
        ' Only select correspondence or if not then current addresses
        If Not xmlCust.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='2']") Is Nothing Then
            Set xmlAddress = xmlCust.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='2']")
        Else
            Set xmlAddress = xmlCust.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='1']")
        End If
        If Not xmlAddress Is Nothing Then
            xmlUniqueNode.appendChild xmlAddress.cloneNode(True)
        End If
    Next

    Set xmlAddress = Nothing
    Set xmlCust = Nothing
    Set xmlCustNodes = Nothing

    ' So we should have a list of unique addresses, one per applicant. See if any are
    ' the same
    Dim blnDuplicate As Boolean
    blnDuplicate = True
    ' Loop until we have a complete loop with no duplicates found
    While blnDuplicate
        blnDuplicate = False
        Dim intNumApplicants As Integer, intOuter As Integer, intInner As Integer
        intNumApplicants = xmlUniqueNode.childNodes.length
        For intOuter = 0 To intNumApplicants - 1
            ' Loop round all other nodes to look for differences
            For intInner = intOuter + 1 To intNumApplicants - 1
                If (GetAttributeText(xmlUniqueNode.childNodes(intOuter), "POSTCODE") = _
                        GetAttributeText(xmlUniqueNode.childNodes(intInner), "POSTCODE")) And _
                   (GetAttributeText(xmlUniqueNode.childNodes(intOuter), "BUILDINGORHOUSENUMBER") = _
                        GetAttributeText(xmlUniqueNode.childNodes(intInner), "BUILDINGORHOUSENUMBER")) And _
                   (GetAttributeText(xmlUniqueNode.childNodes(intOuter), "BUILDINGORHOUSENAME") = _
                        GetAttributeText(xmlUniqueNode.childNodes(intInner), "BUILDINGORHOUSENAME")) And _
                   (GetAttributeText(xmlUniqueNode.childNodes(intOuter), "FLATNUMBER") = _
                        GetAttributeText(xmlUniqueNode.childNodes(intInner), "FLATNUMBER")) And _
                   (GetAttributeText(xmlUniqueNode.childNodes(intOuter), "STREET") = _
                        GetAttributeText(xmlUniqueNode.childNodes(intInner), "STREET")) Then
                    'Flag as duplicate and delete the intInner node & then start again
                    blnDuplicate = True
                    xmlUniqueNode.removeChild xmlUniqueNode.childNodes(intInner)
                    Exit For
                End If
            Next

            'If a duplicate found then exit the loop to reset the whole process
            'because we will have removed a node.
            If blnDuplicate Then
                Exit For
            End If
        Next
    Wend

    'Move the unique addresses into the return array
    For intOuter = 0 To xmlUniqueNode.childNodes.length - 1
        strCustomerArray(intOuter + 1) = GetAttributeText(xmlUniqueNode.childNodes.Item(intOuter), "CUSTOMERNUMBER")
    Next
    intNumAddresses = xmlUniqueNode.childNodes.length

    Set xmlUniqueNode = Nothing
    Set xmlUniqueElem = Nothing
    Set xmlUniqueDoc = Nothing
    Set xmlAPPResponseDoc = Nothing
    Set xmlAPPElem = Nothing
    Set xmlAPPRequestDoc = Nothing
    Set xmlAppNode = Nothing
    Set objRBDO = Nothing
    Set xmltemp2node = Nothing
    Set xmlTempNode = Nothing
    Set xmlallappdetsnode = Nothing
    Set xmlAddress = Nothing
    Set xmlCust = Nothing
    Set xmlCustNodes = Nothing

GetUniqueCustomerAddressError:
    Set xmlUniqueNode = Nothing
    Set xmlUniqueElem = Nothing
    Set xmlUniqueDoc = Nothing
    Set xmlAPPResponseDoc = Nothing
    Set xmlAPPElem = Nothing
    Set xmlAPPRequestDoc = Nothing
    Set xmlAppNode = Nothing
    Set objRBDO = Nothing
    Set xmltemp2node = Nothing
    Set xmlTempNode = Nothing
    Set xmlallappdetsnode = Nothing
    Set xmlAddress = Nothing
    Set xmlCust = Nothing
    Set xmlCustNodes = Nothing
End Sub

                    

Private Sub STKFIQuotationTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve data for the standard KFI Quotation document and format it.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
    On Error GoTo STKFIQuotationTemplateExit
    Const strFunctionName As String = "STKFIQuotationTemplate"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim objKFIHelp As Object
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlNewResponseDoc As New FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    'TK 21/09/2004 Dim xmlTempElement As IXMLDOMElement
    'TK 21/09/2004 Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempDoc As New FreeThreadedDOMDocument40
    
    Dim strAppNo As String
    Dim strApplicationFactFindNumber As String
    Dim strQuotationNumber As String
    Dim strKFIDocType As String
    Dim strKFIOut As String
    Dim strSuccess As String
    Dim strLetterType As String
    
    ' Select PRINTDATA node
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    
    strAppNo = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")
    
    'MKC BBG717 - Get RefType attribute which holds letter type Info
    '"OFFER_APPLICANT_STD", "OFFER_BROKER_STD", "OFFER_PACKAGER_STD", "OFFER_SOLICITOR_STD"
    strLetterType = GetAttributeText(xmlPrintDataNode, "REFTYPE")
    
    ' Select TEMPLATEDATA node
    Set xmlTemplateDataNode = GetMandatoryNode(vxmlRequestNode, "TEMPLATEDATA")
    
    strQuotationNumber = GetMandatoryAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    
    'Set the Document Type to "Standard"
    strKFIDocType = "Standard"
    
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    xmlElem.setAttribute "RB_TEMPLATE", "KFITEMPLATE"  ' SR 23/09/2004 : CORE82
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlNode.appendChild xmlElem

    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.selectSingleNode("APPLICATION").appendChild xmlElem
  
    Set xmlElem = xmlRequestDoc.createElement("QUOTATION")
    xmlElem.setAttribute "QUOTATIONNUMBER", strQuotationNumber
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("OTHERINSURANCECOMPANY")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("VALUERINSTRUCTION")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("CUSTOMERROLE")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    'EP2_583
    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONINTRODUCER")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    'EP2_583
    Set xmlElem = xmlRequestDoc.createElement("NEWPROPERTY")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "KFIDaysValidFor"
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "FSARefNoLandG"
    xmlNode.appendChild xmlElem

    ' Call OmKFIHelp.GetFormattedData to extract the data and format it into the format required
    ' by the KFI template document.
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.validateOnParse = False
    
    Set objKFIHelp = GetObjectContext.CreateInstance("omKFIHelper.OmKFIHelp")
    strKFIOut = objKFIHelp.GetFormattedData(xmlRequestDoc.xml, strKFIDocType)
    Set objKFIHelp = Nothing
    
    xmlTempResponseDoc.LoadXML strKFIOut
    
    'EP2_2203 Check the response for the correct error rather than making up an unspecified error to return
    If Not xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
        strSuccess = "SUCCESS"
    
        'Form a RESPONSE document
        Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
        xmlElem.setAttribute "TYPE", strSuccess
        Set xmlNode = xmlResponseDoc.appendChild(xmlElem)
    Else
        'CheckResponse for errors
        CheckXMLResponse strKFIOut, True, vxmlResponseNode
'
'        strSuccess = "ERROR"
    End If

'    'Form a RESPONSE document
'    Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
'    xmlElem.setAttribute "TYPE", strSuccess
'    Set xmlNode = xmlResponseDoc.appendChild(xmlElem)

'    'CheckResponse for errors
'    CheckXMLResponse xmlResponseDoc.xml, True, vxmlResponseNode
    'INR BBG1437 Cover Letter not required for Std KFI.
    'STD Offer is now printed with the STOfferTemplate method.
    'MKC BBG717 - Add BBG Cover Letter Details If required
    'MKC 28/07/04 - Start
    'SDS  26/08/2004 BBG1232__START
    'If strLetterType <> "" Then
'    If Len(strLetterType) > 0 Then
'        strLetterType = UCase(strLetterType)
'        If strLetterType = "OFFER_APP_RMC" Or strLetterType = "OFFER_APP_RLMC" Or _
'            strLetterType = "OFFER_APP_NONREG" Or strLetterType = "OFFER_SOL_RMC" Or _
'            strLetterType = "OFFER_SOL_RLMC" Or strLetterType = "OFFER_SOL_NONREG" Or _
'            strLetterType = "OFFER_BROKER_RMC" Or strLetterType = "OFFER_BROKER_RLMC" Or _
'            strLetterType = "OFFER_BROKER_NONREG" Or strLetterType = "OFFER_PACK_RMC" Or _
'            strLetterType = "OFFER_PACK_RLMC" Or strLetterType = "OFFER_PACK_NONREG" Then
'    'SDS  26/08/2004 BBG1232__END
'
'                ' Build Cover Letter Request
'                xmlTempDoc.setProperty "NewParser", True
'                xmlTempDoc.validateOnParse = False
'                xmlTempDoc.LoadXML BuildCoverLetterRequest(strAppNo, strApplicationFactFindNumber)
'
'                ' Send Built Request To Request broker to get data
'                xmlNewResponseDoc.setProperty "NewParser", True
'                xmlNewResponseDoc.validateOnParse = False
'                xmlNewResponseDoc.LoadXML AddCoverLetterDetails(xmlTempDoc, xmlTempResponseDoc, _
'                                                        strLetterType, strAppNo, vxmlResponseNode)
'            ' BT 11/08/04 BBG1197 Restructured the response
'            vxmlResponseNode.appendChild xmlNewResponseDoc.selectSingleNode("TEMPLATEDATA")
'        ' PSC 13/09/2004 BBG1405 - Start
'        Else
'            vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")
'        ' PSC 13/09/2004 BBG1405 - End
'        End If
'    Else
'        vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")
'    End If
    
    'INR BBG1437 Add the template data to our response
    vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")
   
   
STKFIQuotationTemplateExit:
    Set xmlPrintDataNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    'TK 21/09/2004 Set xmlTempElement = Nothing
    'TK 21/09/2004 Set xmlTempNode = Nothing
    Set xmlTempDoc = Nothing
    Set objKFIHelp = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlNewResponseDoc = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    CheckError strFunctionName, TypeName(Me)
End Sub


Private Sub LTKFIQuotationTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve data for the lifetime KFI Quotation document and format it.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
    On Error GoTo LTKFIQuotationTemplateExit
    Const strFunctionName As String = "LTKFIQuotationTemplate"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim objKFIHelp As Object
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    'TK 21/09/2004 Dim xmlNewResponseDoc As New FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    'TK 21/09/2004 Dim xmlTempElement As IXMLDOMElement
    'TK 21/09/2004 Dim xmlTempNode As IXMLDOMNode
    'TK 21/09/2004 Dim xmlTempDoc As New FreeThreadedDOMDocument40
    
    Dim strAppNo As String
    Dim strApplicationFactFindNumber As String
    Dim strQuotationNumber As String
    Dim strKFIDocType As String
    Dim strKFIOut As String
    Dim strSuccess As String
    Dim strLetterType As String
    
    ' Select PRINTDATA node
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    
    strAppNo = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")
    
    'MKC BBG717 - Get RefType attribute which holds letter type Info
    '"OFFER_APPLICANT_STD", "OFFER_BROKER_STD", "OFFER_PACKAGER_STD", "OFFER_SOLICITOR_STD"
    strLetterType = GetAttributeText(xmlPrintDataNode, "REFTYPE")
    
    ' Select TEMPLATEDATA node
    Set xmlTemplateDataNode = GetMandatoryNode(vxmlRequestNode, "TEMPLATEDATA")
    
    strQuotationNumber = GetMandatoryAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    
    'Set the Document Type to "Lifetime"
    strKFIDocType = "Lifetime"
    
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    xmlElem.setAttribute "RB_TEMPLATE", "KFITEMPLATE"  ' SR 23/09/2004 : CORE82
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlNode.appendChild xmlElem

    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.selectSingleNode("APPLICATION").appendChild xmlElem
  
    Set xmlElem = xmlRequestDoc.createElement("QUOTATION")
    xmlElem.setAttribute "QUOTATIONNUMBER", strQuotationNumber
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("OTHERINSURANCECOMPANY")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("VALUERINSTRUCTION")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("CUSTOMERROLE")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "KFIDaysValidFor"
    xmlNode.appendChild xmlElem

    ' Call OmKFIHelp.GetFormattedData to extract the data and format it into the format required
    ' by the KFI template document.
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.validateOnParse = False
    
    Set objKFIHelp = GetObjectContext.CreateInstance("omKFIHelper.OmKFIHelp")
    strKFIOut = objKFIHelp.GetFormattedData(xmlRequestDoc.xml, strKFIDocType)
    Set objKFIHelp = Nothing
    
    xmlTempResponseDoc.LoadXML strKFIOut

    If Not xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
        strSuccess = "SUCCESS"
    Else
        strSuccess = "ERROR"
    End If

    'Form a RESPONSE document
    Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
    xmlElem.setAttribute "TYPE", strSuccess
    Set xmlNode = xmlResponseDoc.appendChild(xmlElem)

    'CheckResponse for errors
    CheckXMLResponse xmlResponseDoc.xml, True, vxmlResponseNode
    
    'INR BBG1437 Cover Letter not required for std KFI.
    'Lifetime Offer is now printed withe the LTOfferTemplate method.
    'MKC BBG717 - Add BBG Cover Letter Details If required
    'MKC 28/07/04 - Start
    'SDS  26/08/2004 BBG1232__START
    'If strLetterType <> "" Then
'    If Len(strLetterType) > 0 Then
'        strLetterType = UCase(strLetterType)
'        If strLetterType = "OFFER_APP_RMC" Or strLetterType = "OFFER_APP_RLMC" Or _
'            strLetterType = "OFFER_APP_NONREG" Or strLetterType = "OFFER_SOL_RMC" Or _
'            strLetterType = "OFFER_SOL_RLMC" Or strLetterType = "OFFER_SOL_NONREG" Or _
'            strLetterType = "OFFER_BROKER_RMC" Or strLetterType = "OFFER_BROKER_RLMC" Or _
'            strLetterType = "OFFER_BROKER_NONREG" Or strLetterType = "OFFER_PACK_RMC" Or _
'            strLetterType = "OFFER_PACK_RLMC" Or strLetterType = "OFFER_PACK_NONREG" Then
'    'SDS  26/08/2004 BBG1232__END
'
'            ' Build Cover Letter Request
'            xmlTempDoc.setProperty "NewParser", True
'            xmlTempDoc.validateOnParse = False
'            xmlTempDoc.LoadXML BuildCoverLetterRequest(strAppNo, strApplicationFactFindNumber)
'
'            ' Send Built Request To Request broker to get data
'            xmlNewResponseDoc.setProperty "NewParser", True
'            xmlNewResponseDoc.validateOnParse = False
'            xmlNewResponseDoc.LoadXML AddCoverLetterDetails(xmlTempDoc, xmlTempResponseDoc, _
'                                                            strLetterType, strAppNo, vxmlResponseNode)
'            ' BT 11/08/04 BBG1197 Restructured the response
'            vxmlResponseNode.appendChild xmlNewResponseDoc.selectSingleNode("TEMPLATEDATA")
'        ' PSC 13/09/2004 BBG1405 - Start
'        Else
'            vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")
'        ' PSC 13/09/2004 BBG1405 - End
'        End If
'    Else
'        vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")
'    End If
    
    'INR BBG1437 Add the template data to our response
    vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")

LTKFIQuotationTemplateExit:
    Set xmlPrintDataNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    'TK 21/09/2004 Set xmlTempElement = Nothing
    'TK 21/09/2004 Set xmlTempNode = Nothing
    'TK 21/09/2004 Set xmlTempDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set objKFIHelp = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    CheckError strFunctionName, TypeName(Me)
End Sub


'INR BBG1437  Add functions to print Standard, Lifetime & Transfer of Equity Offers

Private Sub STOfferTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve data for the standard Offer document and format it.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
' BBG1707 Put a unique node underneath TEMPLATEDATA for each unique applicant. This will
' force a cover letter per unique applicant address. The template will need to have
' a UNIQUEAPPLICANT tag around the whole document
' IK 14/12/2004 E2EM00002921 -
'               resolve addressee names for customers at same address
'               major re-structuring of method, original code not preseved
'------------------------------------------------------------------------------------------
    On Error GoTo STOfferTemplateExit
    Const strFunctionName As String = "STOfferTemplate"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

    Dim objKFIHelp As Object
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlKFIResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRbRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppFFNode As IXMLDOMNode 'TK 17/01/2005 BBG1898 ApplicationFactFind node

    Dim strAppNo As String
    Dim strApplicationFactFindNumber As String
    Dim strQuotationNumber As String
    Dim strKFIDocType As String
    Dim strKFIOut As String
'    Dim strSuccess As String
    Dim strLetterType As String
    
    ' Select PRINTDATA node
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    
    strAppNo = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")
    
    'Get RefType attribute which holds letter type Info
    '"OFFER_APPLICANT_STD", "OFFER_BROKER_STD", "OFFER_PACKAGER_STD", "OFFER_SOLICITOR_STD"
    strLetterType = UCase$(GetAttributeText(xmlPrintDataNode, "REFTYPE"))
    
    ' Select TEMPLATEDATA node
    Set xmlTemplateDataNode = GetMandatoryNode(vxmlRequestNode, "TEMPLATEDATA")
    
    strQuotationNumber = GetMandatoryAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    
    'Set the Document Type to "OFFER"
    strKFIDocType = "OFFER"
    
    Set xmlKFIResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    xmlElem.setAttribute "RB_TEMPLATE", "KFITEMPLATE"  ' SR 23/09/2004 : CORE82
    Set xmlRequestNode = xmlRequestDoc.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    Set xmlNode = xmlRequestNode.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    'TK 17/01/2005 BBG1898 ApplicationFactFind node
    Set xmlAppFFNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlRequestDoc.createElement("QUOTATION")
'    xmlElem.setAttribute "QUOTATIONNUMBER", strQuotationNumber
    xmlElem.setAttribute "QUOTATIONNUMBER_REF", "ACTIVEQUOTENUMBER"
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("OTHERINSURANCECOMPANY")
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("VALUERINSTRUCTION")
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("CUSTOMERROLE")
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "KFIDaysValidFor"
    xmlRequestNode.appendChild xmlElem

    ' Call OmKFIHelp.GetFormattedData to extract the data and format it into the format required
    ' by the KFI template document.
    xmlKFIResponseDoc.setProperty "NewParser", True
    xmlKFIResponseDoc.validateOnParse = False
    
    Set objKFIHelp = GetObjectContext.CreateInstance("omKFIHelper.OmKFIHelp")
    strKFIOut = objKFIHelp.GetFormattedData(xmlRequestDoc.xml, strKFIDocType)
    Set objKFIHelp = Nothing
    
    xmlKFIResponseDoc.LoadXML strKFIOut
    
    'CheckResponse for errors
    If xmlKFIResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
        CheckXMLResponse strKFIOut, True, vxmlResponseNode
    End If
    
    vxmlResponseNode.appendChild xmlKFIResponseDoc.selectSingleNode("TEMPLATEDATA") 'BC MAR1320 23Feb06
    
'BC MAR1128 25 Jan 2006 Begin
'    Call UpdateKFISectionNos(xmlKFIResponseDoc)
    

    ' Build Cover Letter Request
'    Set xmlRbRequestDoc = New FreeThreadedDOMDocument40
'    xmlRbRequestDoc.async = False
'    xmlRbRequestDoc.setProperty "NewParser", True
'    xmlRbRequestDoc.validateOnParse = False
'    BuildCoverLetterRequest xmlRbRequestDoc, strAppNo, strApplicationFactFindNumber
'
'    AddCoverLetterDetails _
'        xmlRbRequestDoc, _
'        xmlKFIResponseDoc, _
'        strLetterType, _
'        strAppNo, _
'        vxmlResponseNode
    'BC MAR1128 25 Jan 2006 Begin
   
STOfferTemplateExit:
    
    Set objKFIHelp = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAppFFNode = Nothing 'TK 17/01/2005 BBG1898 ApplicationFactFind node
    Set xmlRequestDoc = Nothing
    Set xmlKFIResponseDoc = Nothing
    Set xmlRbRequestDoc = Nothing
    
    Set xmlPrintDataNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlRequestNode = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
    CheckError strFunctionName, TypeName(Me)

End Sub


Private Sub LTOfferTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve data for the lifetime Offer document and format it.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
' BBG1707 Put a unique node underneath TEMPLATEDATA for each unique applicant. This will
' force a cover letter per unique applicant address. The template will need to have
' a UNIQUEAPPLICANT tag around the whole document
' IK 14/12/2004 E2EM00002921 -
'               resolve addressee names for customers at same address
'               major re-structuring of method, original code not preseved
'------------------------------------------------------------------------------------------
    On Error GoTo LTOfferTemplateExit
    Const strFunctionName As String = "LTOfferTemplate"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim objKFIHelp As Object
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlKFIResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRbRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppFFNode As IXMLDOMNode 'TK 18/01/2005 BBG1901 ApplicationFactFind node

    Dim strAppNo As String
    Dim strApplicationFactFindNumber As String
    Dim strQuotationNumber As String
    Dim strKFIDocType As String
    Dim strKFIOut As String
'    Dim strSuccess As String
    Dim strLetterType As String
    
    ' Select PRINTDATA node
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    
    strAppNo = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")
    
    'Get RefType attribute which holds letter type Info
    '"OFFER_APPLICANT_STD", "OFFER_BROKER_STD", "OFFER_PACKAGER_STD", "OFFER_SOLICITOR_STD"
    strLetterType = UCase$(GetAttributeText(xmlPrintDataNode, "REFTYPE"))
    
    ' Select TEMPLATEDATA node
    Set xmlTemplateDataNode = GetMandatoryNode(vxmlRequestNode, "TEMPLATEDATA")
    
    strQuotationNumber = GetMandatoryAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    
    'Set the Document Type to "LIFETIMEOFFER"
    strKFIDocType = "LIFETIMEOFFER"
    
    Set xmlKFIResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    xmlElem.setAttribute "RB_TEMPLATE", "KFITEMPLATE"  ' SR 23/09/2004 : CORE82
    Set xmlRequestNode = xmlRequestDoc.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    Set xmlNode = xmlRequestNode.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    'TK 18/01/2005 BBG1901 ApplicationFactFind nodeSet xmlNode = xmlNode.appendChild(xmlElem)
    Set xmlAppFFNode = xmlNode.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("QUOTATION")
    'TK 18/01/2005 BBG1901 xmlElem.setAttribute "QUOTATIONNUMBER", strQuotationNumber
    xmlElem.setAttribute "QUOTATIONNUMBER_REF", "ACTIVEQUOTENUMBER"
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("OTHERINSURANCECOMPANY")
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("VALUERINSTRUCTION")
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("CUSTOMERROLE")
    xmlAppFFNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "KFIDaysValidFor"
    xmlRequestNode.appendChild xmlElem

    ' Call OmKFIHelp.GetFormattedData to extract the data and format it into the format required
    ' by the KFI template document.
    xmlKFIResponseDoc.setProperty "NewParser", True
    xmlKFIResponseDoc.validateOnParse = False
    
    Set objKFIHelp = GetObjectContext.CreateInstance("omKFIHelper.OmKFIHelp")
    strKFIOut = objKFIHelp.GetFormattedData(xmlRequestDoc.xml, strKFIDocType)
    Set objKFIHelp = Nothing
    
    xmlKFIResponseDoc.LoadXML strKFIOut
    
    'CheckResponse for errors
    If xmlKFIResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
        CheckXMLResponse strKFIOut, True, vxmlResponseNode
    End If

' IK 05/01/2005 E2EM00002921
    Call UpdateKFISectionNos(xmlKFIResponseDoc, True)
' IK 05/01/2005 E2EM00002921 ends
    
    ' Build Cover Letter Request
    Set xmlRbRequestDoc = New FreeThreadedDOMDocument40
    xmlRbRequestDoc.async = False
    xmlRbRequestDoc.setProperty "NewParser", True
    xmlRbRequestDoc.validateOnParse = False
    BuildCoverLetterRequest xmlRbRequestDoc, strAppNo, strApplicationFactFindNumber

    AddCoverLetterDetails _
        xmlRbRequestDoc, _
        xmlKFIResponseDoc, _
        strLetterType, _
        strAppNo, _
        vxmlResponseNode
    
LTOfferTemplateExit:
    
    Set objKFIHelp = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAppFFNode = Nothing 'TK 18/01/2005 BBG1901 ApplicationFactFind node
    Set xmlRequestDoc = Nothing
    Set xmlKFIResponseDoc = Nothing
    Set xmlRbRequestDoc = Nothing
    
    Set xmlPrintDataNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlRequestNode = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
    CheckError strFunctionName, TypeName(Me)

End Sub

Private Sub TEOfferTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve data for the Transfer of Equity Offer document and format it.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
' BBG1707 Put a unique node underneath TEMPLATEDATA for each unique applicant. This will
' force a cover letter per unique applicant address. The template will need to have
' a UNIQUEAPPLICANT tag around the whole document
' IK 14/12/2004 E2EM00002921 -
'               resolve addressee names for customers at same address
'               major re-structuring of method, original code not preseved
'------------------------------------------------------------------------------------------
    On Error GoTo TEOfferTemplateExit
    Const strFunctionName As String = "TEOfferTemplate"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim objKFIHelp As Object
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode
    
    Dim xmlRbRequestDoc As FreeThreadedDOMDocument40
    Dim xmlKFIResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strAppNo As String
    Dim strApplicationFactFindNumber As String
    Dim strQuotationNumber As String
    Dim strKFIDocType As String
    Dim strKFIOut As String
'    Dim strSuccess As String
    Dim strLetterType As String
    
    ' Select PRINTDATA node
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    
    strAppNo = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")
    
    'Get RefType attribute which holds letter type Info
    '"OFFER_APPLICANT_STD", "OFFER_BROKER_STD", "OFFER_PACKAGER_STD", "OFFER_SOLICITOR_STD"
    strLetterType = GetAttributeText(xmlPrintDataNode, "REFTYPE")
    
    ' Select TEMPLATEDATA node
    Set xmlTemplateDataNode = GetMandatoryNode(vxmlRequestNode, "TEMPLATEDATA")
    
    strQuotationNumber = GetMandatoryAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    
    'Set the Document Type to "TRANSFEROFFER"
    strKFIDocType = "TRANSFEROFFER"
    
    Set xmlKFIResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRbRequestDoc = New FreeThreadedDOMDocument40
    xmlRbRequestDoc.async = False
    
    Set xmlElem = xmlRbRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    ' ik_todo - use KFITEMPLATE?
    Set xmlRequestNode = xmlRbRequestDoc.appendChild(xmlElem)

    Set xmlElem = xmlRbRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    Set xmlNode = xmlRequestNode.appendChild(xmlElem)

    Set xmlElem = xmlRbRequestDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    Set xmlNode = xmlRequestNode.appendChild(xmlElem)
  
    Set xmlElem = xmlRbRequestDoc.createElement("QUOTATION")
    xmlElem.setAttribute "QUOTATIONNUMBER", strQuotationNumber
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRbRequestDoc.createElement("OTHERINSURANCECOMPANY")
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRbRequestDoc.createElement("VALUERINSTRUCTION")
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRbRequestDoc.createElement("CUSTOMERROLE")
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlRbRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "KFIDaysValidFor"
    xmlRequestNode.appendChild xmlElem

    ' Call OmKFIHelp.GetFormattedData to extract the data and format it into the format required
    ' by the KFI template document.
    xmlKFIResponseDoc.setProperty "NewParser", True
    xmlKFIResponseDoc.validateOnParse = False
    
    Set objKFIHelp = GetObjectContext.CreateInstance("omKFIHelper.OmKFIHelp")
    strKFIOut = objKFIHelp.GetFormattedData(xmlRbRequestDoc.xml, strKFIDocType)
    Set objKFIHelp = Nothing
    
    xmlKFIResponseDoc.LoadXML strKFIOut

    'CheckResponse for errors
    If xmlKFIResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
        CheckXMLResponse strKFIOut, True, vxmlResponseNode
    End If

    Call UpdateKFISectionNos(xmlKFIResponseDoc)
    
    ' Build Cover Letter Request
    Set xmlRbRequestDoc = New FreeThreadedDOMDocument40
    xmlRbRequestDoc.async = False
    xmlRbRequestDoc.setProperty "NewParser", True
    xmlRbRequestDoc.validateOnParse = False
    BuildCoverLetterRequest xmlRbRequestDoc, strAppNo, strApplicationFactFindNumber

    AddCoverLetterDetails _
        xmlRbRequestDoc, _
        xmlKFIResponseDoc, _
        strLetterType, _
        strAppNo, _
        vxmlResponseNode
    
TEOfferTemplateExit:
    Set objKFIHelp = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlPrintDataNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRbRequestDoc = Nothing
    Set xmlKFIResponseDoc = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
    CheckError strFunctionName, TypeName(Me)

End Sub

Private Function xmlGetAttributeAsInteger( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Integer
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsInteger = CInt(strValue)
    End If

End Function

Private Function xmlGetAttributeText( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As String
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            xmlGetAttributeText = vxmlNode.Attributes.getNamedItem(vstrAttribName).Text
        End If
    End If

End Function


Private Function xmlGetAttributeAsDate( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Date
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsDate(strValue) Then
        xmlGetAttributeAsDate = CDate(strValue)
    End If

End Function


Private Function xmlAttributeExists( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            If Len(vxmlNode.Attributes.getNamedItem(vstrAttribName).Text) > 0 Then
                xmlAttributeExists = True
            End If
        End If
    End If

End Function

Public Function xmlGetAttributeAsLong( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Long
    xmlGetAttributeAsLong = CSafeLng(GetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function



Private Function xmlGetAttributeAsBoolean( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    
    Dim strValue As String
    
    strValue = UCase$(xmlGetAttributeText(vxmlNode, vstrAttribName))
    
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetAttributeAsBoolean = True
    End If
    
End Function
'BBG1489 Non Regulated KFI Quotation document method
Private Sub NRKFIQuotationTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve data for the Non Regulated KFI Quotation document and format it.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
    On Error GoTo NRKFIQuotationTemplateExit
    Const strFunctionName As String = "NRKFIQuotationTemplate"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim objKFIHelp As Object
    
    Dim xmlPrintDataNode As IXMLDOMNode
    Dim xmlTemplateDataNode As IXMLDOMNode
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strAppNo As String
    Dim strApplicationFactFindNumber As String
    Dim strQuotationNumber As String
    Dim strKFIDocType As String
    Dim strKFIOut As String
    Dim strSuccess As String
    Dim strLetterType As String
    
    ' Select PRINTDATA node
    Set xmlPrintDataNode = GetMandatoryNode(vxmlRequestNode, "PRINTDATA")
    
    strAppNo = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = GetMandatoryAttributeText(xmlPrintDataNode, "APPLICATIONFACTFINDNUMBER")
    
    'Get RefType attribute which holds letter type Info
    '"OFFER_APPLICANT_STD", "OFFER_BROKER_STD", "OFFER_PACKAGER_STD", "OFFER_SOLICITOR_STD"
    strLetterType = GetAttributeText(xmlPrintDataNode, "REFTYPE")
    
    ' Select TEMPLATEDATA node
    Set xmlTemplateDataNode = GetMandatoryNode(vxmlRequestNode, "TEMPLATEDATA")
    
    strQuotationNumber = GetMandatoryAttributeText(xmlTemplateDataNode, "QUOTATIONNUMBER")
    
    'Set the Document Type to "NONREG"
    strKFIDocType = "NONREG"
    
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    xmlElem.setAttribute "RB_TEMPLATE", "KFITEMPLATE"
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)

    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlNode.appendChild xmlElem

    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONFACTFIND")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.selectSingleNode("APPLICATION").appendChild xmlElem
  
    Set xmlElem = xmlRequestDoc.createElement("QUOTATION")
    xmlElem.setAttribute "QUOTATIONNUMBER", strQuotationNumber
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("OTHERINSURANCECOMPANY")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("VALUERINSTRUCTION")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("CUSTOMERROLE")
    xmlNode.selectSingleNode("//APPLICATIONFACTFIND").appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("GLOBALDATAITEM")
    xmlElem.setAttribute "_SCHEMA_", "KFIDATA"
    xmlElem.setAttribute "NAME", "KFIDaysValidFor"
    xmlNode.appendChild xmlElem

    ' Call OmKFIHelp.GetFormattedData to extract the data and format it into the format required
    ' by the KFI template document.
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.validateOnParse = False
    
    Set objKFIHelp = GetObjectContext.CreateInstance("omKFIHelper.OmKFIHelp")
    strKFIOut = objKFIHelp.GetFormattedData(xmlRequestDoc.xml, strKFIDocType)
    Set objKFIHelp = Nothing
    
    xmlTempResponseDoc.LoadXML strKFIOut

    If Not xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA") Is Nothing Then
        strSuccess = "SUCCESS"
    Else
        strSuccess = "ERROR"
    End If

    'Form a RESPONSE document
    Set xmlElem = xmlResponseDoc.createElement("RESPONSE")
    xmlElem.setAttribute "TYPE", strSuccess
    Set xmlNode = xmlResponseDoc.appendChild(xmlElem)

    'CheckResponse for errors
    CheckXMLResponse xmlResponseDoc.xml, True, vxmlResponseNode
    
    'Add the template data to our response
    vxmlResponseNode.appendChild xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA")

NRKFIQuotationTemplateExit:
    Set xmlPrintDataNode = Nothing
    Set xmlTemplateDataNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDoc = Nothing
    Set objKFIHelp = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    CheckError strFunctionName, TypeName(Me)
End Sub

' ====================================================================================
' KFI Cover Letter functions
' ====================================================================================

Private Sub BuildCoverLetterRequest( _
    ByVal vxmlRequestDoc As FreeThreadedDOMDocument40, _
    ByVal strAppNo As String, _
    ByVal strApplicationFactFindNumber As String)
' Header ----------------------------------------------------------------------------------
' Ddescription: Build up XML Request Schema for BBG KFI Cover Letters
'
' pass:
'   strAppNo
'       Application Number
'   strApplicationFactFindNumber
'       Application Fact Find Number
' return:
'   XML Request Schema
'------------------------------------------------------------------------------------------
'BT BBG1707    Added CustomerNumber argument to allow multiple KFI cover letter addresses
'              which is passed into the CUSTCORRADDRESS xml request
'------------------------------------------------------------------------------------------
    On Error GoTo BuildCoverLetterRequestExit
    Const strFunctionName As String = "BuildCoverLetterRequest"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlElem = vxmlRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    
'   Use named omPDM SCHEMA file
    xmlElem.setAttribute "RB_SCHEMA", "omPDM"
    Set xmlNode = vxmlRequestDoc.appendChild(xmlElem)
    
    Set xmlElem = vxmlRequestDoc.createElement("UNITNAME")
    xmlElem.setAttribute "_SCHEMA_", "UNITNAME"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlNode.appendChild xmlElem
    
    Set xmlElem = vxmlRequestDoc.createElement("PRODUCTTYPE")
    xmlElem.setAttribute "_SCHEMA_", "PRODUCTTYPE"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem
    
    Set xmlElem = vxmlRequestDoc.createElement("SECURITYADDRESS")
    xmlElem.setAttribute "_SCHEMA_", "SECURITYADDRESS"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem
    
    Set xmlElem = vxmlRequestDoc.createElement("BROKERORAPPLICANTADDRESS")
    xmlElem.setAttribute "_SCHEMA_", "BROKERORAPPLICANTADDRESS"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem

'IK 14/12/2004 E2EM00002921
'    Set xmlElem = vxmlRequestDoc.createElement("CUSTCORRADDRESS")
'    xmlElem.setAttribute "_SCHEMA_", "CUSTCORRADDRESS"
'    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
'    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
'    'BBG1707
'    xmlElem.setAttribute "CUSTOMERNUMBER", strCustomerNumber
'    xmlNode.appendChild xmlElem
'
'    Set xmlElem = vxmlRequestDoc.createElement("LTDCOMPANYADDRESS")
'    xmlElem.setAttribute "_SCHEMA_", "LTDCOMPANYADDRESS"
'    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
'    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
'    xmlNode.appendChild xmlElem
'IK 14/12/2004 E2EM00002921 ends

    Set xmlElem = vxmlRequestDoc.createElement("APPLNADDITIONALQUESTIONS")
    xmlElem.setAttribute "_SCHEMA_", "APPLNADDITIONALQUESTIONS"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem

    Set xmlElem = vxmlRequestDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "_SCHEMA_", "CASEACTIVITY"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlNode.appendChild xmlElem

    Set xmlElem = vxmlRequestDoc.createElement("NEWPROPERTY")
    xmlElem.setAttribute "_SCHEMA_", "NEWPROPERTY"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    xmlNode.appendChild xmlElem
    
    'If steLetterType = "NETWORK" Then
        Set xmlElem = vxmlRequestDoc.createElement("INTERMEDIARIES")
        xmlElem.setAttribute "_SCHEMA_", "INTERMEDIARIES"
        xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
        xmlNode.appendChild xmlElem
    'End If

    'If strLetterType = "SOLICITOR" Then
        Set xmlElem = vxmlRequestDoc.createElement("LEGALREP")
        xmlElem.setAttribute "_SCHEMA_", "LEGALREP"
        xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
        xmlNode.appendChild xmlElem
    'End If
    
    Set xmlElem = vxmlRequestDoc.createElement("GLOBALPARAMETER")
    xmlElem.setAttribute "_SCHEMA_", "GLOBALPARAMETER"
    xmlElem.setAttribute "NAME", "KFIMDay"
    xmlNode.appendChild xmlElem

    'Create element to search ALLAPPLICATIONDETAILS scheme
    'but only go down as far as customer as that's all we need.
    Dim xmlallappdetsnode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    
    Set xmlElem = vxmlRequestDoc.createElement("ALLAPPLICATIONDETAILS")
    xmlElem.setAttribute "_SCHEMA_", "ALLAPPLICATIONDETAILS"
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
    Set xmlallappdetsnode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = vxmlRequestDoc.createElement("APPLICATIONCUSTOMERROLE")
    Set xmlTempNode = xmlallappdetsnode.appendChild(xmlElem)
    
    Set xmlElem = vxmlRequestDoc.createElement("CUSTOMER")
'IK 14/12/2004 E2EM00002921
    Set xmlTempNode = xmlTempNode.appendChild(xmlElem)
    Set xmlElem = vxmlRequestDoc.createElement("CUSTOMERADDRESS")
'IK 14/12/2004 E2EM00002921 ends
    xmlElem.setAttribute "_EXTRACTTYPE_", "NODE"
    xmlTempNode.appendChild xmlElem

Exit Sub
        
BuildCoverLetterRequestExit:
    Set vxmlRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlallappdetsnode = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    CheckError strFunctionName, TypeName(Me)

End Sub

Private Sub AddCoverLetterDetails( _
    ByVal vxmlRbRequestDoc As FreeThreadedDOMDocument40, _
    ByVal vxmlKFIResponseDoc As FreeThreadedDOMDocument40, _
    ByVal strLetterID As String, ByVal strAppNo As String, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description: Retrieves data using xml schema returned from BuildCoverLetterRequest function
'               and appends info to the Main KFI XML file based on certain criteria
'
' pass:
'   vxmlRequestDoc
'       Cover Letter REQUEST XML
'   vxmlKFIResponseDoc
'       Main KFI
'   strLetterID
'
'   strAppNo
' return:
'
' exceptions:
'       oeRecordNotFound
' history:
' IK 14/12/2004 E2EM00002921 -
'               resolve addressee names for customers at same address
'               major re-structuring of method, original code not preseved
'------------------------------------------------------------------------------------------
    On Error GoTo AddCoverLetterDetailsExit
    Const strFunctionName As String = "AddCoverLetterDetails"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim objRBDO As Object
    
    Dim xmlRbResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRbResponseNode As IXMLDOMNode
    Dim xmlRbCustomerRoleNode As IXMLDOMNode
    
    Dim xmlKFIResponseNode As IXMLDOMNode
    
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseUniqueApplicantNode As IXMLDOMNode
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim blnLtdCompBToL As Boolean
'    Dim blnUseSolicitorActingForLender As Boolean
'    Dim blnLifetimeEnclosure As Boolean
    Dim blnBuyToLet As Boolean
    Dim blnMarsMpsOffer As Boolean
    
'    Dim strAddress As String
'    Dim strAddresseeName As String
    Dim strProdCode As String
'    Dim strSuccess As String
'    Dim strPath As String
'    Dim strTemp As String

    Dim dteMortgageDay As Date
    Dim dteDIPDataCapture As Date
    
    'E2EM00003332
    Dim strCompanyName As String
    Dim intCount As Integer
    Dim intTotal As Integer
    Dim strSalutation As String

    'Call Request Broker and load response string into a FreeThreadedDOMDocument40
    Set xmlRbResponseDoc = New FreeThreadedDOMDocument40
    xmlRbResponseDoc.async = False
    xmlRbResponseDoc.setProperty "NewParser", True
    xmlRbResponseDoc.validateOnParse = False
    
    ' Request Data via Request Broker
    Set objRBDO = GetObjectContext.CreateInstance("omRB.OmRequestDO")
    xmlRbResponseDoc.LoadXML objRBDO.OmDataRequest(vxmlRbRequestDoc.xml)
    Set objRBDO = Nothing
    
    'CheckResponse for errors
    CheckXMLResponse xmlRbResponseDoc.xml, True, vxmlResponseNode
        
    ' drop <COPYTO> nodes from KFIHelper data, not valid for multiple customers
    For Each xmlNode In vxmlKFIResponseDoc.selectNodes("TEMPLATEDATA/COPYTO")
        vxmlKFIResponseDoc.selectSingleNode("TEMPLATEDATA").removeChild xmlNode
    Next
    
    'Get Product Code
    Set xmlNode = xmlRbResponseDoc.selectSingleNode("RESPONSE/PRODUCTTYPE")
    If Not xmlNode Is Nothing Then
        strProdCode = GetAttributeText(xmlNode, "SPECIALGROUP")
        If strProdCode = "20" Or strProdCode = "25" Then
            blnLtdCompBToL = True
        End If
        If strProdCode = "10" Or strProdCode = "15" Or strProdCode = "20" Or strProdCode = "25" Then
            blnBuyToLet = True
        End If
    End If

    ' Specific additional question will be used for whether an Mars/MPS offer has been issued
    Set xmlNode = xmlRbResponseDoc.selectSingleNode("RESPONSE/APPLNADDITIONALQUESTIONS[@QUESTIONREFERENCE='2']")
    If Not xmlNode Is Nothing Then
        If xmlGetAttributeAsBoolean(xmlNode, "RESPONSE") Then
            blnMarsMpsOffer = True
        End If
    End If

    ' Global parameter showing when mortgage day is....
    Set xmlNode = xmlRbResponseDoc.selectSingleNode("RESPONSE/GLOBALPARAMETER[@NAME='KFIMDay']")
    If Not xmlNode Is Nothing Then
        dteMortgageDay = xmlGetAttributeAsDate(xmlNode, "STRING")
    Else
        dteMortgageDay = CDate("31/October/2004")
    End If

    ' Find when the application went into DIP data capture stage
    Set xmlNode = xmlRbResponseDoc.selectSingleNode("RESPONSE/CASEACTIVITY/CASESTAGE[@STAGEID='50']")
    If xmlNode Is Nothing Then
        Set xmlNode = xmlRbResponseDoc.selectSingleNode("RESPONSE/CASEACTIVITY/CASESTAGEARCHIVE[@STAGEID='50']")
    End If
    If Not xmlNode Is Nothing Then
        dteDIPDataCapture = xmlGetAttributeAsDate(xmlNode, "STAGEORIGINATINGDATETIME")
    End If
    
    If blnLtdCompBToL Then
    
       'BBG1852 Look at all the Directors and construct a Salutation from the names.
        
        intTotal = xmlRbResponseDoc.selectNodes("RESPONSE/ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1']/CUSTOMER").length
    
        strSalutation = ""
        intCount = 0
        For Each xmlNode In xmlRbResponseDoc.selectNodes("RESPONSE/ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1']/CUSTOMER")
            If Len(strSalutation) <> 0 Then
                intCount = intCount + 1
                If (intTotal - intCount) > 1 Then
                    strSalutation = strSalutation & ", "
                Else
                    strSalutation = strSalutation & " & "
                End If
            End If
            strSalutation = strSalutation & GetSalutation(xmlNode, 5) ' BBG1852
        Next
        
        ' Ltd. Co. - drop applicant nodes
        For Each xmlNode In xmlRbResponseDoc.selectNodes("RESPONSE/ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1']")
            xmlRbResponseDoc.selectSingleNode("RESPONSE/ALLAPPLICATIONDETAILS").removeChild xmlNode
        Next
        
 ' IK 05/01/2005 E2EM00002921
' E2EM00003332 If the Company Name is not held in the Correspondence Salutation, use the Surname attribute
        For Each xmlNode In xmlRbResponseDoc.selectNodes("RESPONSE/ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='3']/CUSTOMER")
            strCompanyName = xmlGetAttributeText(xmlNode, "CORRESPONDENCESALUTATION")
            If strCompanyName = "" Then
                strCompanyName = xmlGetAttributeText(xmlNode, "SURNAME")
            End If
        
            SetAttributeValue xmlNode, "ADDRESSEENAME", strCompanyName
            SetAttributeValue xmlNode, "SALUTATION", strSalutation
            
        Next
' IK 05/01/2005 E2EM00002921 ends
 
    Else
        ResolveClientsAtSameAddress xmlRbResponseDoc
    End If
    
    Set xmlKFIResponseNode = vxmlKFIResponseDoc.selectSingleNode("TEMPLATEDATA")
    Set xmlRbResponseNode = xmlRbResponseDoc.selectSingleNode("RESPONSE")
    
    AddStandardCoverLetterData strAppNo, xmlKFIResponseNode, xmlRbResponseNode
    
    'Add Retrieved Cover Letter Data to Main KFI XML
    Set xmlResponseNode = vxmlResponseNode.appendChild(xmlKFIResponseNode.cloneNode(False))

    If strLetterID = "OFFER_APP_RMC" _
    Or strLetterID = "OFFER_APP_RLMC" _
    Or strLetterID = "OFFER_APP_NONREG" _
    Then
    
        For Each xmlRbCustomerRoleNode In xmlRbResponseNode.selectNodes("ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE")
            
            Set xmlElem = vxmlResponseNode.ownerDocument.createElement("UNIQUEAPPLICANT")
            Set xmlResponseUniqueApplicantNode = xmlResponseNode.appendChild(xmlElem)
            For Each xmlNode In xmlKFIResponseNode.childNodes
                xmlResponseUniqueApplicantNode.appendChild xmlNode.cloneNode(True)
            Next
            
            UpdateAddresseeDetails xmlResponseUniqueApplicantNode, xmlRbCustomerRoleNode
            CreateCopyToNodes xmlResponseUniqueApplicantNode, xmlRbCustomerRoleNode
        
            Call AddTypeSpecificCoverLetterData( _
                strLetterID, strAppNo, strProdCode, _
                blnLtdCompBToL, blnMarsMpsOffer, blnBuyToLet, _
                dteMortgageDay, dteDIPDataCapture, _
                xmlRbResponseNode, xmlRbCustomerRoleNode, xmlResponseUniqueApplicantNode)
        
        Next
    
    Else
    
        Set xmlRbCustomerRoleNode = xmlRbResponseNode.selectSingleNode("ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE")
            
        Set xmlElem = vxmlResponseNode.ownerDocument.createElement("UNIQUEAPPLICANT")
        Set xmlResponseUniqueApplicantNode = xmlResponseNode.appendChild(xmlElem)
        For Each xmlNode In xmlKFIResponseNode.childNodes
            xmlResponseUniqueApplicantNode.appendChild xmlNode.cloneNode(True)
        Next
        
        UpdateAddresseeDetails xmlResponseUniqueApplicantNode, xmlRbCustomerRoleNode
'        CreateCopyToNodes xmlResponseUniqueApplicantNode, xmlRbCustomerRoleNode
            
        Call AddTypeSpecificCoverLetterData( _
            strLetterID, strAppNo, strProdCode, _
            blnLtdCompBToL, blnMarsMpsOffer, blnBuyToLet, _
            dteMortgageDay, dteDIPDataCapture, _
            xmlRbResponseNode, xmlRbCustomerRoleNode, xmlResponseUniqueApplicantNode)
    
    End If
        
AddCoverLetterDetailsExit:
    
    Set objRBDO = Nothing
    Set xmlRbResponseDoc = Nothing
    Set xmlRbResponseNode = Nothing
    Set xmlResponseUniqueApplicantNode = Nothing
    Set xmlRbCustomerRoleNode = Nothing
    Set xmlKFIResponseNode = Nothing
    Set xmlKFIResponseNode = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlAttrib = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
    CheckError strFunctionName, TypeName(Me)
 
End Sub

Private Sub AddTypeSpecificCoverLetterData( _
    ByVal vstrLetterID As String, _
    ByVal vstrAppNo As String, _
    ByVal vstrProdCode As String, _
    ByVal vblnLtdCompBToL As Boolean, _
    ByVal vblnMarsMpsOffer As Boolean, _
    ByVal vblnBuyToLet As Boolean, _
    ByVal vdteMortgageDay As Date, _
    ByVal vdteDIPDataCapture As Date, _
    ByVal vxmlRbResponseNode As IXMLDOMNode, _
    ByVal vxmlRbCustomerRoleNode As IXMLDOMNode, _
    ByVal vxmlResponseUniqueApplicantNode As IXMLDOMElement)
    
    On Error GoTo AddTypeSpecificCoverLetterDataExit
    
    Const strFunctionName As String = "AddTypeSpecificCoverLetterData"
    gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName
    
    Dim xmlKFILetterNode As IXMLDOMNode
    Dim xmlKFIAddresseeNode As IXMLDOMNode
    Dim xmlKFICustomerNode As IXMLDOMNode
    Dim xmlKFINotCustomerNode As IXMLDOMNode
    Dim xmlKFINode As IXMLDOMNode
    
    Dim xmlRbNode As IXMLDOMNode
    Dim xmlRbAddresseeNode As IXMLDOMNode

    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim blnUseSolicitorActingForLender As Boolean
    Dim blnLifetimeEnclosure As Boolean
    Dim blnSentToAdditional As Boolean
    
    Dim strSalutation As String
    Dim strSalutationHeading As String ' BBG1852
    Dim strAddresseeName As String
    Dim strAddress As String
    Dim strTemp As String
    Dim strLocation As String
    
    Dim strCompany As String   'BBG1852
    
    'Append <LETTER> tag to temp document
    Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("LETTER")
    
    Select Case vstrLetterID
        Case "OFFER_APP_RMC", "OFFER_APP_RLMC", "OFFER_APP_RMC", "OFFER_APP_NONREG"
            strTemp = "OL001"
        Case "OFFER_SOL_RMC", "OFFER_SOL_RLMC", "OFFER_SOL_RMC", "OFFER_SOL_NONREG"
            strTemp = "OL002"
        Case "OFFER_BROKER_RMC", "OFFER_BROKER_RLMC", "OFFER_BROKER_RMC", "OFFER_BROKER_NONREG"
            strTemp = "OL003"
        Case "OFFER_PACK_RMC", "OFFER_PACK_RLMC", "OFFER_PACK_RMC", "OFFER_PACK_NONREG"
            strTemp = "OL004"
    End Select
    xmlElem.setAttribute "REFNUMBER", strTemp
    
    ' Put in the values of ACCOUNTNUMBER,APPLICATIONNUMBER and ISSUEDATE again as
    ' the originals at the top level will be formatted however the KFI engine feels like
    ' it. Putting the curly brackets around them forces it to display as is.
    
    xmlElem.setAttribute "ACCOUNTNUMBER", "{" & vstrAppNo & "}"
    xmlElem.setAttribute "APPLICATIONNUMBER", "{" & vstrAppNo & "}"
    xmlElem.setAttribute "ISSUEDATE", "{" & Format$(Date, "d mmmm yyyy") & "}"
    
    'Add Team Name
    Set xmlNode = vxmlRbResponseNode.selectSingleNode("UNITNAME")
    If Not xmlNode Is Nothing Then
        CopyAttribute xmlNode, xmlElem, "UNITNAME"
    End If
    
    ' Override if it is a customer offer for a limited company
    If vstrLetterID = "OFFER_APP_RMC" Or vstrLetterID = "OFFER_APP_RLMC" Or vstrLetterID = "OFFER_APP_NONREG" Then
        Set xmlNode = vxmlRbCustomerRoleNode.selectSingleNode("CUSTOMER")
        If Not xmlNode Is Nothing Then
            If vblnLtdCompBToL Then
                'E2EM00003332
                'Get the company name from ADDRESSEENAME already set up.
                strCompany = GetAttributeText(xmlNode, "ADDRESSEENAME")

                strSalutationHeading = GetAttributeText(xmlNode, "SALUTATION")   ' BBG1852
                strSalutation = GetAttributeText(xmlNode, "SALUTATION") & ", " & strCompany
            Else
                strSalutation = GetAttributeText(xmlNode, "SALUTATION")
            End If
        End If
    Else
        strSalutation = "Sirs"
    End If
    
    xmlElem.setAttribute "SALUTATION", strSalutation
    
    Set xmlKFILetterNode = vxmlResponseUniqueApplicantNode.appendChild(xmlElem)
  
    'Create Addressee Element
    Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("ADDRESSEE")
    
    'Select Create XML Request Based On Letter Type Passed
    Select Case vstrLetterID
        Case "OFFER_APP_RMC", "OFFER_APP_RLMC", "OFFER_APP_NONREG"
            Set xmlRbAddresseeNode = vxmlRbCustomerRoleNode.selectSingleNode("CUSTOMER")

        Case "OFFER_BROKER_RMC", "OFFER_BROKER_RLMC", "OFFER_BROKER_NONREG"
            Set xmlRbAddresseeNode = vxmlRbResponseNode.selectSingleNode("INTERMEDIARIES[@INTERMEDIARYTYPE=50]")

        Case "OFFER_PACK_RMC", "OFFER_PACK_RLMC", "OFFER_PACK_NONREG"
            Set xmlRbAddresseeNode = vxmlRbResponseNode.selectSingleNode("INTERMEDIARIES[@INTERMEDIARYTYPE=20 or @INTERMEDIARYTYPE=30]")

        Case "OFFER_SOL_RMC", "OFFER_SOL_RLMC", "OFFER_SOL_NONREG"
            If vblnLtdCompBToL Or vstrProdCode = "120" Then
                'TK 21/01/2005 E2EM00002982 always use customers solicitor for Lifetime products
                'TK 28/01/2004 BBG1921
                'If vstrProdCode = "120" Then
                '    Set xmlRbAddresseeNode = vxmlRbResponseNode.selectSingleNode("LEGALREP[@SEPARATELEGALREPRESENTATIVE=0]")
                'Else
                Set xmlRbAddresseeNode = vxmlRbResponseNode.selectSingleNode("LEGALREP[@SEPARATELEGALREPRESENTATIVE=1]")
                'End If
            Else
                Set xmlRbAddresseeNode = vxmlRbResponseNode.selectSingleNode("LEGALREP[@SEPARATELEGALREPRESENTATIVE=0]")
            End If

    End Select
    
    If Not xmlRbAddresseeNode Is Nothing Then
        'Addressee Name Attribute
        Select Case vstrLetterID
            Case "OFFER_APP_RMC", "OFFER_APP_RLMC", "OFFER_APP_NONREG"
                
                ' BBG1852 Add the Directors names for Limited Companies
                If vblnLtdCompBToL Then
                    strAddresseeName = strSalutationHeading   'BBG1852
                Else
                    strAddresseeName = GetAttributeText(xmlRbAddresseeNode, "ADDRESSEENAME")
                End If
                
                If Not xmlRbAddresseeNode.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='2']") Is Nothing Then
                    Set xmlRbAddresseeNode = xmlRbAddresseeNode.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='2']")
                Else
                    Set xmlRbAddresseeNode = xmlRbAddresseeNode.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='1']")
                End If
                
            Case "OFFER_BROKER_RMC", "OFFER_BROKER_RLMC", "OFFER_BROKER_NONREG"
                strAddresseeName = GetAttributeText(xmlRbAddresseeNode, "TITLE") & " "
                strAddresseeName = LTrim$(strAddresseeName) & Left$(GetAttributeText(xmlRbAddresseeNode, "FORENAME"), 1) & " "
                strAddresseeName = LTrim$(strAddresseeName) & GetAttributeText(xmlRbAddresseeNode, "SURNAME")
            
            Case "OFFER_SOL_RMC", "OFFER_SOL_RLMC", "OFFER_SOL_NONREG"
                strAddresseeName = GetAttributeText(xmlRbAddresseeNode, "CONTACTTITLE_TEXT") & " "
                strAddresseeName = LTrim$(strAddresseeName) & Left$(GetAttributeText(xmlRbAddresseeNode, "CONTACTFORENAME"), 1) & " "
                strAddresseeName = LTrim$(strAddresseeName) & GetAttributeText(xmlRbAddresseeNode, "CONTACTSURNAME")
    
            Case "OFFER_PACK_RMC", "OFFER_PACK_RLMC", "OFFER_PACK_NONREG"
                strAddresseeName = GetAttributeText(xmlRbAddresseeNode, "NAME")
    
        End Select
        
        If Len(strAddresseeName) > 0 Then
            xmlElem.setAttribute "ADDRESSEENAME", strAddresseeName
        End If
    
    End If
    
    Set xmlKFIAddresseeNode = xmlKFILetterNode.appendChild(xmlElem)
    
    If Not xmlRbAddresseeNode Is Nothing Then
        
        Select Case vstrLetterID
            Case "OFFER_BROKER_RMC", "OFFER_BROKER_RLMC", "OFFER_BROKER_NONREG"
                AddCoverLetterAddressLine vxmlRbResponseNode.selectSingleNode("INTERMEDIARIES[@INTERMEDIARYTYPE=40]"), xmlKFIAddresseeNode, "NAME"
        End Select

        'Postcode Attribute
        CopyAttribute xmlRbAddresseeNode, xmlKFIAddresseeNode, "POSTCODE"
            
        Select Case vstrLetterID
            Case "OFFER_APP_RMC", "OFFER_APP_RLMC", "OFFER_APP_NONREG"
                'BBG1852 Add Company name as an ADDRESSLINE
                If vblnLtdCompBToL Then
                    AddCoverLetterAddressLine xmlRbAddresseeNode.parentNode, xmlKFIAddresseeNode, "CORRESPONDENCESALUTATION"
                End If
            Case "OFFER_SOL_RMC", "OFFER_SOL_RLMC", "OFFER_SOL_NONREG"
                AddCoverLetterAddressLine xmlRbAddresseeNode, xmlKFIAddresseeNode, "COMPANYNAME"
        End Select
        
        strTemp = GetAttributeText(xmlRbAddresseeNode, "FLATNUMBER")
        If Len(strTemp) > 0 Then
            Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("ADDRESS")
            xmlElem.setAttribute "ADDRESSLINE", "Flat " & strTemp
            xmlKFIAddresseeNode.appendChild xmlElem
        End If
        
        AddCoverLetterAddressLine xmlRbAddresseeNode, xmlKFIAddresseeNode, "BUILDINGORHOUSENAME"
        
        strTemp = GetAttributeText(xmlRbAddresseeNode, "BUILDINGORHOUSENUMBER") & " "
        strTemp = LTrim$(strTemp) & GetAttributeText(xmlRbAddresseeNode, "STREET")
        If Len(strTemp) > 0 Then
            Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("ADDRESS")
            xmlElem.setAttribute "ADDRESSLINE", strTemp
            xmlKFIAddresseeNode.appendChild xmlElem
        End If
        
        AddCoverLetterAddressLine xmlRbAddresseeNode, xmlKFIAddresseeNode, "DISTRICT"
        AddCoverLetterAddressLine xmlRbAddresseeNode, xmlKFIAddresseeNode, "TOWN"
        AddCoverLetterAddressLine xmlRbAddresseeNode, xmlKFIAddresseeNode, "COUNTY"
    End If
    
    If vstrLetterID = "OFFER_APP_RMC" Or vstrLetterID = "OFFER_APP_RLMC" Or vstrLetterID = "OFFER_APP_NONREG" Then
        Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("CUSTOMER")
        
        Set xmlKFICustomerNode = xmlKFILetterNode.appendChild(xmlElem)
    
        ' Formatting differences for the customer letter
        If vstrProdCode = "120" Then
            Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("LIFETIME")
            Set xmlNode = xmlKFICustomerNode.appendChild(xmlElem)
            
            If vblnMarsMpsOffer And vdteDIPDataCapture < vdteMortgageDay Then
                Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("LIFEBEFOREYES")
                xmlNode.appendChild xmlElem
            End If
            'JD E2EM00002982 should use customers solicitor
            '28/01/2005 BBG1921 - Always use lenders solicitor
            blnUseSolicitorActingForLender = True
            blnLifetimeEnclosure = True
        
        Else
            Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("NOTLIFETIME")
            Set xmlNode = xmlKFICustomerNode.appendChild(xmlElem)
        
            ' There is no regulation on Buy to Let cases so don't set either of these flags
            If Not vblnBuyToLet Then
'                If vdteDIPDataCapture < CDate("31/October/2004") Then
                If vdteDIPDataCapture < vdteMortgageDay Then
                    If Not vblnMarsMpsOffer Then
                        Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("NOTLIFEBEFORENO")
                        xmlNode.appendChild xmlElem
    
                    Else
                        Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("NOTLIFEBEFOREYES")
                        xmlNode.appendChild xmlElem
                    End If
                End If
            End If
            
            If vblnLtdCompBToL Then
                blnUseSolicitorActingForLender = True
            End If
        End If
    
        ' Next tag the solicitor, imtermediary or packager onto the sub nodes as appropriate
        If blnUseSolicitorActingForLender Then
            Set xmlNode = vxmlRbResponseNode.selectSingleNode("LEGALREP[@SEPARATELEGALREPRESENTATIVE=1]")
        Else
            Set xmlNode = vxmlRbResponseNode.selectSingleNode("LEGALREP[@SEPARATELEGALREPRESENTATIVE=0]")
        End If
        
        If Not xmlNode Is Nothing Then
            blnSentToAdditional = True
            strAddress = xmlGetAttributeText(xmlNode, "CONTACTTITLE_TEXT") & " "
            strAddress = LTrim$(strAddress) & Left$(xmlGetAttributeText(xmlNode, "CONTACTFORENAME"), 1)
            ConcatAttrib strAddress, xmlNode, "CONTACTSURNAME"
            ConcatAttrib strAddress, xmlNode, "COMPANYNAME", ", "
            If Len(strAddress) > 0 Then
                strAddress = strAddress & ", "
            End If
            ConcatAddress strAddress, xmlNode
            Call AddToCustomerElements(xmlKFICustomerNode, "CUSTSOLICITOR", _
                                       "SOLICITORNAMEANDADDRESS", strAddress)
        End If
        
        Set xmlNode = vxmlRbResponseNode.selectSingleNode("INTERMEDIARIES[@INTERMEDIARYTYPE=50]")
        If Not xmlNode Is Nothing Then
            blnSentToAdditional = True
            strAddress = xmlGetAttributeText(xmlNode, "TITLE") & " "
            strAddress = LTrim$(strAddress) & Left$(xmlGetAttributeText(xmlNode, "FORENAME"), 1)
            ConcatAttrib strAddress, xmlNode, "SURNAME"
            ConcatAttrib strAddress, xmlNode, "NAME", ", "
            If Len(strAddress) > 0 Then
                strAddress = strAddress & ", "
            End If
            ConcatAddress strAddress, xmlNode
            Call AddToCustomerElements(xmlKFICustomerNode, "CUSTINTERMEDIARY", _
                                       "INTERMEDIARYNAMEANDADDRESS", strAddress)
        End If
        
        Set xmlNode = vxmlRbResponseNode.selectSingleNode("INTERMEDIARIES[@INTERMEDIARYTYPE=30]")
        If Not xmlNode Is Nothing Then
            blnSentToAdditional = True
            strAddress = xmlGetAttributeText(xmlNode, "TITLE") & " "
            strAddress = LTrim$(strAddress) & Left$(xmlGetAttributeText(xmlNode, "FORENAME"), 1)
            ConcatAttrib strAddress, xmlNode, "SURNAME"
            ConcatAttrib strAddress, xmlNode, "NAME", ", "
            If Len(strAddress) > 0 Then
                strAddress = strAddress & ", "
            End If
            ConcatAddress strAddress, xmlNode
            Call AddToCustomerElements(xmlKFICustomerNode, "CUSTPACKAGER", _
                                       "PACKAGERNAMEANDADDRESS", strAddress)
        End If
                
        If blnSentToAdditional Then
            Call AddToCustomerElements(xmlKFICustomerNode, "SENTTOADDITIONAL", "DUMMY", "1")
        End If
        
'-----------NOT CUSTOMER---------------
    Else
        
        Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("NOTCUSTOMER")
        
        'Copy the Security address
        strTemp = ""
        ConcatAddress strTemp, vxmlRbResponseNode.selectSingleNode("SECURITYADDRESS")
        If Len(strTemp) > 0 Then
            xmlElem.setAttribute "SECURITYADDRESS", strTemp
        End If
        
        Set xmlKFINotCustomerNode = xmlKFILetterNode.appendChild(xmlElem)
        
        If vblnLtdCompBToL Then   'Ltd Company
            Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("LTDCOMPANY")
            xmlElem.setAttribute "LIMITEDCOMPANY", GetAttributeText(vxmlResponseUniqueApplicantNode, "APPLICANTNAME")
        Else
            Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("NOTLTDCOMPANY")
            ' TEMPLATEDATA node will have concatenated names of multiple applicants
            xmlElem.setAttribute "APPLICANTNAME", GetAttributeText(vxmlResponseUniqueApplicantNode.parentNode, "APPLICANTNAME")
        End If
        xmlKFINotCustomerNode.appendChild xmlElem
        
        Select Case vstrLetterID
            
            Case "OFFER_BROKER_RMC", "OFFER_BROKER_RLMC", "OFFER_BROKER_NONREG", _
                "OFFER_PACK_RMC", "OFFER_PACK_RLMC", "OFFER_PACK_NONREG"
                Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("BROKERORNETWORK")
                xmlKFINotCustomerNode.appendChild xmlElem
    
            Case "OFFER_SOL_RMC", "OFFER_SOL_RLMC", "OFFER_SOL_NONREG"
            
                Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("ENCLOSURES")
                Set xmlKFINode = xmlKFILetterNode.appendChild(xmlElem)
                Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("SOLICITOR")
                Set xmlKFINode = xmlKFINode.appendChild(xmlElem)
                Set xmlRbNode = vxmlRbResponseNode.selectSingleNode("NEWPROPERTY")
                If Not xmlRbNode Is Nothing Then
                    Select Case GetAttributeText(xmlRbNode, "PROPERTYLOCATION")
                        Case "10"
                            strLocation = "ENGLANDANDWALES"
                        Case "20"
                            strLocation = "SCOTLAND"
                        Case "30"
                            strLocation = "NORTHERNIRELAND"
                    End Select
                    If Len(strLocation) > 0 Then
                        Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement(strLocation)
                        xmlKFINode.appendChild xmlElem
                    End If
                    Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("BUILDINGINSURANCEDECLARATION")
                    xmlKFINode.appendChild xmlElem
                End If
                
                Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("SOLICITOR")
                Set xmlKFINode = xmlKFINotCustomerNode.appendChild(xmlElem)
                
                If vblnLtdCompBToL Then   'Ltd Company
                    Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("LIMITEDCOMPANY")
                    xmlKFINode.appendChild xmlElem
                Else
                    Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement("NOTLIMITEDCOMPANY")
                    Set xmlKFINode = xmlKFINode.appendChild(xmlElem)
                    If Len(strLocation) > 0 Then
                        Set xmlElem = vxmlResponseUniqueApplicantNode.ownerDocument.createElement(strLocation)
                        xmlKFINode.appendChild xmlElem
                    End If
                End If
    
        End Select
        
    End If
    
AddTypeSpecificCoverLetterDataExit:

    Set xmlKFILetterNode = Nothing
    Set xmlKFIAddresseeNode = Nothing
    Set xmlKFICustomerNode = Nothing
    Set xmlKFINotCustomerNode = Nothing
    Set xmlKFINode = Nothing
    
    Set xmlRbNode = Nothing
    Set xmlRbAddresseeNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    
    CheckError strFunctionName, TypeName(Me)

End Sub

Private Sub AddCoverLetterAddressLine( _
    ByVal vxmlSourceNode As IXMLDOMNode, _
    ByVal vxmlTargetNode As IXMLDOMNode, _
    ByVal vstrAttribName As String)
    
    Dim xmlElem As IXMLDOMElement

    If Not vxmlSourceNode Is Nothing Then
        If AttributeValueExists(vxmlSourceNode, vstrAttribName) Then
            Set xmlElem = vxmlTargetNode.ownerDocument.createElement("ADDRESS")
            xmlElem.setAttribute "ADDRESSLINE", GetAttributeText(vxmlSourceNode, vstrAttribName)
            vxmlTargetNode.appendChild xmlElem
        End If
    End If
    
    Set xmlElem = Nothing

End Sub

Private Sub AddStandardCoverLetterData( _
    ByVal vstrAppNo As String, _
    ByVal vxmlRbResponseNode As IXMLDOMElement, _
    ByVal vxmlKFIResponseNode As IXMLDOMElement)
    
    Dim xmlNode As IXMLDOMNode
    Dim strAddress As String
    
    'Construct the Response xml from data
    '**** ADD STANDARD Cover Letter Elements to Main KFI Template XML ****
'   Add Account Number
    vxmlKFIResponseNode.setAttribute "ACCOUNTNUMBER", vstrAppNo

'   Add Application Number
    vxmlKFIResponseNode.setAttribute "APPLICATIONNUMBER", vstrAppNo
    
    'Add Issue Date
    vxmlKFIResponseNode.setAttribute "ISSUEDATE", Format$(Date, "d mmmm yyyy")
    
    'Add Team Name
    Set xmlNode = vxmlRbResponseNode.selectSingleNode("UNITNAME")
    If Not xmlNode Is Nothing Then
        CopyAttribute xmlNode, vxmlKFIResponseNode, "UNITNAME"
    End If

    'Add Security Address
    Set xmlNode = vxmlRbResponseNode.selectSingleNode("SECURITYADDRESS")
    If Not xmlNode Is Nothing Then
        ConcatAddress strAddress, xmlNode
        vxmlKFIResponseNode.setAttribute "SECURITYADDRESS", strAddress
        Set xmlNode = Nothing
    End If
    
    Set xmlNode = Nothing

End Sub

Private Sub ConcatAddress( _
    ByRef rvstrTarget As String, _
    ByVal vxmlAddressNode As IXMLDOMNode, _
    Optional ByVal vblnAddFlat As Boolean = True)
    
    Dim strAddress As String
    
    If Not vxmlAddressNode Is Nothing Then
        If vblnAddFlat Then
            ConcatAttrib strAddress, vxmlAddressNode, "FLATNUMBER"
            If Len(strAddress) > 0 Then strAddress = "Flat " & strAddress
        End If
        ConcatAttrib strAddress, vxmlAddressNode, "BUILDINGORHOUSENAME", ", "
        ConcatAttrib strAddress, vxmlAddressNode, "BUILDINGORHOUSENUMBER", ", ", ","
        If strAddress = "," Then
            strAddress = ""
        End If
        ConcatAttrib strAddress, vxmlAddressNode, "STREET"
        ConcatAttrib strAddress, vxmlAddressNode, "DISTRICT", ", "
        ConcatAttrib strAddress, vxmlAddressNode, "TOWN", ", "
        ConcatAttrib strAddress, vxmlAddressNode, "COUNTY", ", "
        ConcatAttrib strAddress, vxmlAddressNode, "POSTCODE", ", "
        rvstrTarget = rvstrTarget & strAddress
    End If
    
End Sub
    
Private Sub ConcatAttrib( _
    ByRef rvstrTarget As String, _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrPrefix As String = " ", _
    Optional ByVal vstrDefault As String = "")
    
    If xmlAttributeExists(vxmlNode, vstrAttribName) Then
        If Len(rvstrTarget) > 0 Then rvstrTarget = rvstrTarget & vstrPrefix
        rvstrTarget = rvstrTarget & xmlGetAttributeText(vxmlNode, vstrAttribName)
    ElseIf Len(vstrDefault) > 0 Then
        rvstrTarget = rvstrTarget & vstrDefault
    End If
    
End Sub

Private Sub UpdateAddresseeDetails( _
    ByVal vxmlResponseUniqueApplicantNode As IXMLDOMNode, _
    ByVal vxmlRbCustomerRoleNode As IXMLDOMNode)

    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlCustomerAddressNode As IXMLDOMNode
    Dim xmlTemplateNode As IXMLDOMNode   ' BBG1925

    ' also fix incorrect APPLICANTNAME
    Set xmlCustomerNode = vxmlRbCustomerRoleNode.selectSingleNode("CUSTOMER")
    If Not xmlCustomerNode Is Nothing Then
        CopyAttributeValue xmlCustomerNode, vxmlResponseUniqueApplicantNode, "ADDRESSEENAME", "APPLICANTNAME"
    End If
    
    'BBG1925 In the TEMPLATEDATA node rename the APPLICANT attributes to OFFERAPPLICANT
    '        so that each copy of the offer DOCUMENT is addressed to all applicants at the first applicant's address.
    '        Note that a cover letter is sent with the document to the applicants living at each unique address.
    Set xmlTemplateNode = vxmlResponseUniqueApplicantNode.parentNode

    CopyAttributeValue xmlTemplateNode, xmlTemplateNode, "APPLICANTNAME", "OFFERAPPLICANTNAME"
    CopyAttributeValue xmlTemplateNode, xmlTemplateNode, "APPLICANTADDRESSLINE1", "OFFERAPPLICANTADDRESSLINE1"
    CopyAttributeValue xmlTemplateNode, xmlTemplateNode, "APPLICANTADDRESSLINE2", "OFFERAPPLICANTADDRESSLINE2"
    CopyAttributeValue xmlTemplateNode, xmlTemplateNode, "APPLICANTADDRESSLINE3", "OFFERAPPLICANTADDRESSLINE3"
    CopyAttributeValue xmlTemplateNode, xmlTemplateNode, "APPLICANTPOSTCODE", "OFFERAPPLICANTPOSTCODE"
    
    ' remove existing values from TEMPLATEDATA node
    vxmlResponseUniqueApplicantNode.parentNode.Attributes.removeNamedItem "APPLICANTADDRESSLINE1"
    vxmlResponseUniqueApplicantNode.parentNode.Attributes.removeNamedItem "APPLICANTADDRESSLINE2"
    vxmlResponseUniqueApplicantNode.parentNode.Attributes.removeNamedItem "APPLICANTADDRESSLINE3"
    vxmlResponseUniqueApplicantNode.parentNode.Attributes.removeNamedItem "APPLICANTPOSTCODE"
    
    UpdateKFIAddress vxmlResponseUniqueApplicantNode, xmlCustomerNode, "APPLICANT"

    'BBG1925
    Set xmlCustomerNode = Nothing
    Set xmlCustomerAddressNode = Nothing
    Set xmlTemplateNode = Nothing

End Sub

Private Sub CreateCopyToNodes( _
    ByVal vxmlKFIResponseNode As IXMLDOMNode, _
    ByVal vxmlRbCustomerRoleNode As IXMLDOMNode)
    
    Dim xmlSiblingNode As IXMLDOMNode
    
    Set xmlSiblingNode = vxmlRbCustomerRoleNode.previousSibling
    Do While Not (xmlSiblingNode Is Nothing)
        CreateSingleCopyTo vxmlKFIResponseNode, xmlSiblingNode
        Set xmlSiblingNode = xmlSiblingNode.previousSibling
    Loop
    Set xmlSiblingNode = vxmlRbCustomerRoleNode.nextSibling
    Do While Not (xmlSiblingNode Is Nothing)
        CreateSingleCopyTo vxmlKFIResponseNode, xmlSiblingNode
        Set xmlSiblingNode = xmlSiblingNode.nextSibling
    Loop

    Set xmlSiblingNode = Nothing

End Sub

Private Sub CreateSingleCopyTo( _
    ByVal vxmlTargetNode As IXMLDOMNode, _
    ByVal vxmlSourceNode As IXMLDOMNode)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    
    Set xmlCustomerNode = vxmlSourceNode.selectSingleNode("CUSTOMER")
    If Not xmlCustomerNode Is Nothing Then
        
        Set xmlElem = vxmlTargetNode.ownerDocument.createElement("COPYTO")
        Set xmlNode = vxmlTargetNode.appendChild(xmlElem)
    
        Set xmlElem = vxmlTargetNode.ownerDocument.createElement("COPYADDRESSEE")
        
        CopyAttributeValue xmlCustomerNode, xmlElem, "ADDRESSEENAME", "COPYNAME"
        
        UpdateKFIAddress xmlElem, xmlCustomerNode, "COPY"
        
        xmlNode.appendChild xmlElem
        
    End If

    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlCustomerNode = Nothing

End Sub

Private Sub UpdateKFIAddress( _
    ByVal vxmlKFIResponseNode As IXMLDOMElement, _
    ByVal vxmlRbCustomerNode As IXMLDOMNode, _
    ByVal vstrPrefix As String)
    
    Dim intLines As Integer
    Dim strLine As String
    
    Dim xmlAddressNode As IXMLDOMNode
    Set xmlAddressNode = vxmlRbCustomerNode.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='2']")
    If xmlAddressNode Is Nothing Then
        Set xmlAddressNode = vxmlRbCustomerNode.selectSingleNode("CUSTOMERADDRESS[@ADDRESSTYPE='1']")
    End If
    
    If Not xmlAddressNode Is Nothing Then
            
        ' remove existing values
        vxmlKFIResponseNode.Attributes.removeNamedItem vstrPrefix & "ADDRESSLINE1"
        vxmlKFIResponseNode.Attributes.removeNamedItem vstrPrefix & "ADDRESSLINE2"
        vxmlKFIResponseNode.Attributes.removeNamedItem vstrPrefix & "ADDRESSLINE3"
        vxmlKFIResponseNode.Attributes.removeNamedItem vstrPrefix & "POSTCODE"
    
        intLines = 1
        If xmlAttributeExists(xmlAddressNode, "FLATNUMBER") Then
            strLine = "Flat " & xmlGetAttributeText(xmlAddressNode, "FLATNUMBER")
        End If
        ConcatAttrib strLine, xmlAddressNode, "BUILDINGORHOUSENAME", ", "
        ConcatAttrib strLine, xmlAddressNode, "BUILDINGORHOUSENUMBER", ", "
        ConcatAttrib strLine, xmlAddressNode, "STREET"
        If Len(strLine) > 0 Then
            vxmlKFIResponseNode.setAttribute vstrPrefix & "ADDRESSLINE1", strLine
            intLines = 2
        End If
        
        strLine = xmlGetAttributeText(xmlAddressNode, "DISTRICT")
        ConcatAttrib strLine, xmlAddressNode, "TOWN", ", "
        If Len(strLine) > 0 Then
            vxmlKFIResponseNode.setAttribute vstrPrefix & "ADDRESSLINE" & CStr(intLines), strLine
            intLines = intLines + 1
        End If
        
        strLine = xmlGetAttributeText(xmlAddressNode, "COUNTY")
        If Len(strLine) > 0 Then
            vxmlKFIResponseNode.setAttribute vstrPrefix & "ADDRESSLINE" & CStr(intLines), strLine
            intLines = intLines + 1
        End If
        
        strLine = xmlGetAttributeText(xmlAddressNode, "POSTCODE")
        If Len(strLine) > 0 Then
            If intLines = 4 Then
                vxmlKFIResponseNode.setAttribute vstrPrefix & "POSTCODE", strLine
            Else
                vxmlKFIResponseNode.setAttribute vstrPrefix & "ADDRESSLINE" & CStr(intLines), strLine
            End If
        End If
    
    End If
    
    Set xmlAddressNode = Nothing

End Sub

Private Sub ResolveClientsAtSameAddress(ByVal vxmlRbResponseDoc As FreeThreadedDOMDocument40)
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomerRoleNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlSiblingNode As IXMLDOMNode
    Dim xmlAddressMasterNode As IXMLDOMNode
    Dim xmlAddressTargetNode As IXMLDOMNode
    
    Dim strPath As String
    
    For Each xmlCustomerRoleNode In vxmlRbResponseDoc.selectNodes("RESPONSE/ALLAPPLICATIONDETAILS/APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1'][CUSTOMER]")
        ' save salutation etc. for later concat. operation
        Set xmlCustomerNode = xmlCustomerRoleNode.selectSingleNode("CUSTOMER")
        AddCustomerIdNode xmlCustomerNode, xmlCustomerNode
        
        strPath = "CUSTOMER/CUSTOMERADDRESS[@ADDRESSTYPE='2']"
        Set xmlAddressMasterNode = xmlCustomerRoleNode.selectSingleNode(strPath)
        If xmlAddressMasterNode Is Nothing Then
            strPath = "CUSTOMER/CUSTOMERADDRESS[@ADDRESSTYPE='1']"
            Set xmlAddressMasterNode = xmlCustomerRoleNode.selectSingleNode(strPath)
        End If
        If Not xmlAddressMasterNode Is Nothing Then
            ' check siblings for duplicate addresses
            Set xmlSiblingNode = xmlCustomerRoleNode.nextSibling
            Do While Not (xmlSiblingNode Is Nothing)
                Set xmlAddressTargetNode = xmlSiblingNode.selectSingleNode(strPath)
' IK 05/01/2005 E2EM00002921
                Set xmlSiblingNode = xmlSiblingNode.nextSibling
' IK 05/01/2005 E2EM00002921 ends
                If Not xmlAddressTargetNode Is Nothing Then
                    If AddressMatch(xmlAddressMasterNode, xmlAddressTargetNode) Then
                        ' save salutation etc. for later concat. operation
                        AddCustomerIdNode xmlCustomerNode, xmlAddressTargetNode.parentNode
' IK 05/01/2005 E2EM00002921
                        ' remove redundant APPLICATIONCUSTOMERROLE node
                        xmlCustomerRoleNode.parentNode.removeChild xmlAddressTargetNode.parentNode.parentNode
' IK 05/01/2005 E2EM00002921 ends
                    End If
                End If
            Loop
        End If
        If xmlCustomerNode.selectNodes("tempWho").length = 1 Then
            SetAttributeValue xmlCustomerNode, "ADDRESSEENAME", GetSalutation(xmlCustomerNode, 5)   ' BBG1852  Display Initials.
            SetAttributeValue xmlCustomerNode, "SALUTATION", GetAttributeText(xmlCustomerNode, "CORRESPONDENCESALUTATION")
        Else
            ConcatNames xmlCustomerNode
        End If
        
        For Each xmlNode In xmlCustomerNode.selectNodes("tempWho")
            xmlCustomerNode.removeChild xmlNode
        Next
        
    Next

    Set xmlNode = Nothing
    Set xmlCustomerRoleNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlSiblingNode = Nothing
    Set xmlAddressMasterNode = Nothing
    Set xmlAddressTargetNode = Nothing

End Sub

Private Sub AddCustomerIdNode( _
    ByVal vxmlTargetNode As IXMLDOMNode, _
    ByVal vxmlSourceNode As IXMLDOMNode)
                            
    Dim xmlElem As IXMLDOMElement
    Set xmlElem = vxmlTargetNode.ownerDocument.createElement("tempWho")
    
    CopyAttribute vxmlSourceNode, xmlElem, "TITLE_TEXT"
    CopyAttribute vxmlSourceNode, xmlElem, "TITLEOTHER"
    CopyAttribute vxmlSourceNode, xmlElem, "FIRSTFORENAME"
    CopyAttribute vxmlSourceNode, xmlElem, "SECONDFORENAME"
    CopyAttribute vxmlSourceNode, xmlElem, "OTHERFORENAMES"
    CopyAttribute vxmlSourceNode, xmlElem, "SURNAME"
    CopyAttribute vxmlSourceNode, xmlElem, "CORRESPONDENCESALUTATION"
    CopyAttribute vxmlSourceNode, xmlElem, "MARITALSTATUS"
    vxmlTargetNode.appendChild xmlElem
    
    Set xmlElem = Nothing

End Sub

Private Function AddressMatch( _
    ByVal vxmlMaster As IXMLDOMElement, _
    ByVal vxmlTarget As IXMLDOMElement) _
    As Boolean
    
    If Len(GetAttributeText(vxmlMaster, "POSTCODE")) > 0 Then
        If GetAttributeText(vxmlMaster, "POSTCODE") = GetAttributeText(vxmlTarget, "POSTCODE") Then
            If Len(GetAttributeText(vxmlMaster, "BUILDINGORHOUSENUMBER")) > 0 Then
                If GetAttributeText(vxmlMaster, "BUILDINGORHOUSENUMBER") = GetAttributeText(vxmlTarget, "BUILDINGORHOUSENUMBER") Then
                    AddressMatch = True
                End If
            ElseIf Len(GetAttributeText(vxmlMaster, "BUILDINGORHOUSENAME")) > 0 Then
                If GetAttributeText(vxmlMaster, "BUILDINGORHOUSENAME") = GetAttributeText(vxmlTarget, "BUILDINGORHOUSENAME") Then
                    AddressMatch = True
                End If
            End If
        End If
    Else
        If (GetAttributeText(vxmlMaster, "BUILDINGORHOUSENAME") = GetAttributeText(vxmlTarget, "BUILDINGORHOUSENAME")) _
        And (GetAttributeText(vxmlMaster, "BUILDINGORHOUSENUMBER") = GetAttributeText(vxmlTarget, "BUILDINGORHOUSENUMBER")) _
        And (GetAttributeText(vxmlMaster, "FLATNUMBER") = GetAttributeText(vxmlTarget, "FLATNUMBER")) _
        And (GetAttributeText(vxmlMaster, "STREET") = GetAttributeText(vxmlTarget, "STREET")) _
        And (GetAttributeText(vxmlMaster, "DISTRICT") = GetAttributeText(vxmlTarget, "DISTRICT")) _
        And (GetAttributeText(vxmlMaster, "TOWN") = GetAttributeText(vxmlTarget, "TOWN")) _
        And (GetAttributeText(vxmlMaster, "COUNTY") = GetAttributeText(vxmlTarget, "COUNTY")) _
        Then
            AddressMatch = True
        End If
    End If

End Function

Private Sub ConcatNames(ByVal vxmlMaster As IXMLDOMElement)
' (master) CUSTOMERROLE node has (temporary) tempWho nodes added
' these nodes contain salutation etc. attributes for applicants at same address
    
    Dim strSalutation As String
        
    Dim intLoop As Integer
' IK 05/01/2005 E2EM00002921
    Dim intTotal As Integer
    Dim intCount As Integer
' IK 05/01/2005 E2EM00002921 ends
        
    Dim xmlNode As IXMLDOMNode
    Dim xmlSiblingNode As IXMLDOMNode
    
    Dim blnDupFound As Boolean

' IK 05/01/2005 E2EM00002921
    intTotal = vxmlMaster.selectNodes("tempWho").length
' IK 05/01/2005 E2EM00002921 ends
    
    ' any duplicate addressee names?
    For Each xmlNode In vxmlMaster.selectNodes("tempWho")
        SetAttributeValue xmlNode, "ADDRESSEENAME", GetSalutation(xmlNode, 5)  ' BBG1852  Use Initials
    Next
    
    For intLoop = 1 To 3
        blnDupFound = False
        For Each xmlNode In vxmlMaster.selectNodes("tempWho")
            Set xmlSiblingNode = xmlNode.nextSibling
            Do While Not (xmlSiblingNode Is Nothing)
                If GetAttributeText(xmlNode, "ADDRESSEENAME") = GetAttributeText(xmlSiblingNode, "ADDRESSEENAME") Then
                    SetAttributeValue xmlNode, "addresseeNameEx", GetSalutation(xmlNode, intLoop)
                    SetAttributeValue xmlSiblingNode, "addresseeNameEx", GetSalutation(xmlSiblingNode, intLoop)
                    blnDupFound = True
                End If
                Set xmlSiblingNode = xmlSiblingNode.nextSibling
            Loop
        Next
        For Each xmlNode In vxmlMaster.selectNodes("tempWho")
            If Not xmlNode.Attributes.getNamedItem("addresseeNameEx") Is Nothing Then
                CopyAttributeValue xmlNode, xmlNode, "addresseeNameEx", "ADDRESSEENAME"
            End If
        Next
        If Not blnDupFound Then
            Exit For
        End If
    Next
        
' IK 05/01/2005 E2EM00002921
    intCount = 0
    For Each xmlNode In vxmlMaster.selectNodes("tempWho")
        If Len(strSalutation) <> 0 Then
            intCount = intCount + 1
            If (intTotal - intCount) > 1 Then
                strSalutation = strSalutation & ", "
            Else
                strSalutation = strSalutation & " & "
            End If
        End If
' IK 05/01/2005 E2EM00002921 ends
        strSalutation = strSalutation & GetAttributeText(xmlNode, "ADDRESSEENAME")
    Next
    
    SetAttributeValue vxmlMaster, "ADDRESSEENAME", strSalutation
        
    ' any duplicate CORRESPONDENCESALUTATION?
    For Each xmlNode In vxmlMaster.selectNodes("tempWho")
' IK 05/01/2005 E2EM00002921
'        SetAttributeValue xmlNode, "SALUTATION", GetAttributeText(vxmlMaster, "CORRESPONDENCESALUTATION")
        SetAttributeValue xmlNode, "SALUTATION", GetSalutation(xmlNode, 5)   ' BBG1852  Title, Initials and surname only
' IK 05/01/2005 E2EM00002921 ends
    Next
    
    For intLoop = 1 To 3
        blnDupFound = False
        For Each xmlNode In vxmlMaster.selectNodes("tempWho")
            Set xmlSiblingNode = xmlNode.nextSibling
            Do While Not (xmlSiblingNode Is Nothing)
                If GetAttributeText(xmlNode, "SALUTATION") = GetAttributeText(xmlSiblingNode, "SALUTATION") Then
                    SetAttributeValue xmlNode, "SALUTATIONEx", GetSalutation(xmlNode, intLoop + 3)
                    SetAttributeValue xmlSiblingNode, "SALUTATIONEx", GetSalutation(xmlSiblingNode, intLoop + 3)
                    blnDupFound = True
                End If
                Set xmlSiblingNode = xmlSiblingNode.nextSibling
            Loop
        Next
        For Each xmlNode In vxmlMaster.selectNodes("tempWho")
            If Not xmlNode.Attributes.getNamedItem("SALUTATIONEx") Is Nothing Then
                CopyAttributeValue xmlNode, xmlNode, "SALUTATIONEx", "SALUTATION"
            End If
        Next
        If Not blnDupFound Then
            Exit For
        End If
    Next
        
    strSalutation = ""
' IK 05/01/2005 E2EM00002921
    intCount = 0
    For Each xmlNode In vxmlMaster.selectNodes("tempWho")
        If Len(strSalutation) <> 0 Then
            intCount = intCount + 1
            If (intTotal - intCount) > 1 Then
                strSalutation = strSalutation & ", "
            Else
                strSalutation = strSalutation & " & "
            End If
' IK 05/01/2005 E2EM00002921 ends
        End If
        strSalutation = strSalutation & GetAttributeText(xmlNode, "SALUTATION")
    Next
    
    SetAttributeValue vxmlMaster, "SALUTATION", strSalutation
    
    Set xmlNode = Nothing
    Set xmlSiblingNode = Nothing

End Sub

' IK 05/01/2005 E2EM00002921
'Private Function GetSalutation( _
'    ByVal vxmlCustomerNode As IXMLDOMNode, _
'    ByVal vintOption As Integer) _
'    As String

Private Function GetSalutation( _
    ByVal vxmlCustomerNode As IXMLDOMNode, _
    Optional ByVal vintOption As Integer = 0) _
    As String
    
    ' option 0: Title, Surname
' IK 05/01/2005 E2EM00002921 ends
    ' option 1: Title, First Forename, Surname
    ' option 2: Title, First Forename, Second Forename, Surname
    ' option 3: Title, First Forename, Second Forename, OtherForenames, Surname
    ' option 4: Title, First initial, Surname
    ' option 5: Title, First initial, second initial, Surname
    ' option 6: Title, First initial, second initial, other initials, Surname
    
    'BBG1852 Remove unnecessary spaces
    
    Dim strName As String
    Dim strAdd As String
    
    strName = GetAttributeText(vxmlCustomerNode, "TITLEOTHER") & " "
    If Len(strName) = 1 Then
        strName = GetAttributeText(vxmlCustomerNode, "TITLE_TEXT") & " "
    End If
    If vintOption = 1 Or vintOption = 2 Or vintOption = 3 Then
        strAdd = GetAttributeText(vxmlCustomerNode, "FIRSTFORENAME")
        If strAdd <> "" Then
            strName = LTrim$(strName) & strAdd & " "
        End If
    End If
    If vintOption = 2 Or vintOption = 3 Then
        strAdd = GetAttributeText(vxmlCustomerNode, "SECONDFORENAME")
        If strAdd <> "" Then
            strName = LTrim$(strName) & strAdd & " "
        End If
    End If
    If vintOption = 3 Then
        strAdd = GetAttributeText(vxmlCustomerNode, "OTHERFORENAMES")
        If strAdd <> "" Then
            strName = LTrim$(strName) & strAdd & " "
        End If
    End If
    If vintOption = 4 Or vintOption = 5 Or vintOption = 6 Then
        strAdd = GetAttributeText(vxmlCustomerNode, "FIRSTFORENAME")
        If strAdd <> "" Then
            strName = LTrim$(strName) & Left$(strAdd, 1) & " "
        End If
    End If
    If vintOption = 5 Or vintOption = 6 Then
        strAdd = GetAttributeText(vxmlCustomerNode, "SECONDFORENAME")
        If strAdd <> "" Then
            strName = LTrim$(strName) & Left$(strAdd, 1) & " "
        End If
    End If
    If vintOption = 6 Then
        strAdd = GetAttributeText(vxmlCustomerNode, "OTHERFORENAMES")
        If strAdd <> "" Then
            strName = LTrim$(strName) & Left$(strAdd, 1) & " "
        End If
    End If
    GetSalutation = LTrim$(strName) & GetAttributeText(vxmlCustomerNode, "SURNAME")

End Function

' IK 05/01/2005 E2EM00002921
Private Sub UpdateKFISectionNos( _
    ByRef xmlTempResponseDoc As IXMLDOMNode, _
    Optional ByVal vblnIsLifeTime As Boolean = False)
' IK 05/01/2005 E2EM00002921 ends
    '
    ' Subroutine UpdateKFISectionNos is a routine which will run down the sections
    ' in the node and add in an attribute which will be used for the consecutive
    ' section numbering.
    '
    Dim intSection As Integer, intCondition As Integer, intLoop As Integer
    intSection = 1
    intCondition = 1
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlChild As IXMLDOMElement
    Dim xmlChild2 As IXMLDOMElement
    For intLoop = 1 To 20
' IK 05/01/2005 E2EM00002921
        Set xmlElem = xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA/SECTION" & intLoop)
' IK 05/01/2005 E2EM00002921 ends
        If Not xmlElem Is Nothing Then
            xmlElem.setAttribute "SECTIONNO", intSection
            ' Additionally the offer templates sometimes have a sub node on them
            ' so add the attribute to any child nodes also
            For Each xmlChild In xmlElem.childNodes
                xmlChild.setAttribute "SECTIONNO", intSection
                'Go down an extra level just in case
                For Each xmlChild2 In xmlChild.childNodes
                    xmlChild2.setAttribute "SECTIONNO", intSection
                Next
            Next
            intSection = intSection + 1
        End If
    Next
    ' Another special (because of the autonumbering issue). In the conditionsection
    ' also have any CONDITION nodes numbered manually in the general or specific sections
' IK 05/01/2005 E2EM00002921
    Set xmlElem = xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA/CONDITIONSECTION/GENERAL")
' IK 05/01/2005 E2EM00002921 ends
    If Not xmlElem Is Nothing Then
        For Each xmlChild In xmlElem.childNodes
            If xmlChild.baseName = "CONDITION" Then
                xmlChild.setAttribute "CONDITIONNO", intCondition
                intCondition = intCondition + 1
            End If
        Next
    End If
' IK 05/01/2005 E2EM00002921
    Set xmlElem = xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA/CONDITIONSECTION/SPECIFIC")
' IK 05/01/2005 E2EM00002921 ends
    If Not xmlElem Is Nothing Then
        For Each xmlChild In xmlElem.childNodes
            If xmlChild.baseName = "CONDITION" Then
                xmlChild.setAttribute "CONDITIONNO", intCondition
                intCondition = intCondition + 1
            End If
        Next
    End If
    
' IK 05/01/2005 E2EM00002921 - add RISK attributes for Lifetime Section 7
    If vblnIsLifeTime Then
        intCondition = 1
        For Each xmlElem In xmlTempResponseDoc.selectSingleNode("TEMPLATEDATA/SECTION7").childNodes
            xmlElem.setAttribute "RISK", intCondition
            intCondition = intCondition + 1
        Next
    End If
' IK 05/01/2005 E2EM00002921 - ends
    
    Set xmlChild2 = Nothing
    Set xmlChild = Nothing
    Set xmlElem = Nothing
End Sub

Private Sub AddToCustomerElements(ByVal xmlKFICustomerNode As IXMLDOMNode, _
                                  ByVal strElement As String, _
                                  ByVal strAttribute As String, _
                                  ByVal strValue As String)
' Routine to add solicitors etc address into the five sub nodes of the KFI
' response doc if they are present
            
    Dim xmlNewElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement
    Set xmlTempElem = xmlKFICustomerNode.selectSingleNode("LIFETIME")
    If Not xmlTempElem Is Nothing Then
        Set xmlNewElem = xmlTempElem.ownerDocument.createElement(strElement)
        xmlNewElem.setAttribute strAttribute, strValue
        xmlTempElem.appendChild xmlNewElem
    End If

    Set xmlTempElem = xmlKFICustomerNode.selectSingleNode("NOTLIFETIME")
    If Not xmlTempElem Is Nothing Then
        Set xmlNewElem = xmlTempElem.ownerDocument.createElement(strElement)
        xmlNewElem.setAttribute strAttribute, strValue
        xmlTempElem.appendChild xmlNewElem
    End If

    Set xmlNewElem = Nothing
    Set xmlTempElem = Nothing
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
