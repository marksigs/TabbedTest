VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AdminRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AdminRulesBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:

'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: Uses Transactions

'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MDC    05/06/01    Created
'JR     28/06/01    Amended Rules: 7.5.7, 7.5.9 and 7.5.11
'MDC    09/08/01    SYS2481 - Refunds & Retentions
'MDC    16/08/01    SYS2602 - Amend REQUEST to CBREQUEST
'MDC    18/09/01    SYS2719 - OmiPlus 20
'INR    26/10/01    SYS2719 - Changes following revue
'KN     31/10/01    SYS2719 - OmiPlus 20 Contd (01.13)
'KN     07/11/01    SYS2719 - OmiPlus 20 Contd (01.14)
'KN     21/11/01    SYS2719 - OmiPlus 20 Contd (01.14)
'                   Various changes to ensure attributes are populated
'                   Even if data does not exist on database
'MDC    06/12/01    SYS3412 - Error finding accepted mortgagesubquote
'KN     19/12/01    SYS2719 - OmiPlus 20 Contd (01.15)
'PSC    21/01/02    SYS3838 - Tidy up code to use only specific helper methods
'DM     07/02/02    SYS4018 - Changed EMPTYVALUE to EMPTYNUMBER and altered typo
'DM 8/2/02  SYS4045 Check element is being returned from request broker
' DM 15-02-2002 SYS4098 removed collateral number form security
' DM 15/02/2002 SYS4100 Was using the wrong xml node.
' DM 19/02/200 SYS4112 Using wrong xml node.
' DM 28/02/2002  SYS4024
' DM 08/03/02 SYS2317
' DM    11/03/02    SYS4256 Set firstadvance date for the component.
' DM 13/03/02 SYS4272 Set up userid and unitid
' DM 27/03/02 SYS4313 Set up USERAUTHORITYLEVEL
' DM 08/04/02 SYS4371 Remove until the CBA can handle this
' DM 25/04/02 SYS4495 Removed Rule 7.5.102 from SetFirstProductDetails.
' DM 01/05/02 SYS4525 Problem with getting CREATIONDATE when payment type is not initial advance
' DM 08/05/02 SYS4545 was using the wrong attribute name of APPROVALDATE, changed to APPLICATIONAPPROVALDATE
'DM 17/05/02 Temp fix SYS4641
' DM 21/05/02 SYS4628 RELATEDCIFNUMBER should be set to the Application number
' DM 21/05/2002 SYS4667
'STB 27/05/2002 SYS4649 Report On Title element is optional.
'STB 29/05/2002 SYS4629 As above and CompletionDate attribute may be missing.
'STB 30/05/2002 SYS4494 Payment Source information is being set into a long integer (loosing the decimal portion).
'STB 31/05/2002 SYS4654 BASERATECODE is now read from the correct XML element.
'STB 07/06/2002 SYS4640 FEETRANSACTIONDATE for transaction type 99 set to issue date.
'STB 13/06/2002 SYS4868 HOLDBACKTYPE should be prefixed with a zero for advances.
'STB 13/06/2002 SYS4867 HOLDBACKINDICATOR set to 'Y' without a HOLDBACK block being created.
'------------------------------------------------------------------------------------------
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MARS Specific History

'Prog    Date        AQR     Description
'SD      03/10/2005  MAR47   Completion Rules
'SD      06/10/2005  MAR47   Removing unused functions
'SD      17/10/2005  MAR47   Removing Middlename from Customer
'                            Renaming <Customer> as <CustomerDetails>
'                            Renaming <CustomerNumber> in <Customer> to <PersonId>
'                            Adding DateToPreviousLegalAddress to <PreviousAddressDetails>
'SD      01/11/2005  MAR326  Adding the elements in GenericMessages.xsd to the response
'SD      24/11/2005  MAR673  Added AcceptedQuoteNumber to searchCriteris for FiedRateProductCode
'                            Changed Tenure to TenureOfProperty
'SD      24/11/2005  MAR673  Restricting gross ext floor area to 5 digits max, temporary work around
'SD      25/11/2005  MAR694  Cloned PaymentRecord & BatchAudit as SolicitorFundingMethod was not being populated
'SD      25/11/2005  MAR694  Adding Tenure, restricting PanelNumber to 8
'SD      14/12/2005          Changing the mapped values for the ApplicationType and ValidationType for Employment Status
'SD      05/01/2006          Picking up Income from IncomeSummary, not CurrentEmployersRef
'SD      13/01/2006  1842    Changing SetupComponentRecord function
'SD      17/01/2006  1994    Changing SetupCustomerCreation as incorrect CustomerVersion is being picked up
'SD      19/01/2006  1811    Setting CountryCode as ValidationType in AddCustoemrAddress
'SD      20/01/2006          Checking for CustomerVersion
'SD      26/01/2006          Removed NEWLOAN and put MORTGAGESUBQUOTE under APPLICATIONFACTFIND
'SD      26/01/2006          In AddLoanComponent, Picking up ReportOnTitle from ApplicationFactFind and not from Application
'SD      26/01/2006          ManualPortedLoanInd set to 0 if it is null
'JJ      09/02/2006  2396    If ValuationReport values not received, pick them up from NewProperty
'                    2364    If LastValuationDate not populated, pick it up from HomeTrackValuationDetails.VALUATIONRESULTSDATERECEIVED
'SD      09/02/2006  2396    Removing PropertyDescription from the check when populating PropertyType
'SD      13/02/2006          1)Updating comments
'                            2)in AccountCreation, removing the check - where RepaymentBankAccountIndicator = 1,as we have only one payment account
'                            3)SourceChannel - check for ValidationType of 'WF' (Post), in place of 'FA' (Additional Borrowing)
'SD      16/02/2006          Implementing tracing
'SD      21/02/2006          Using PropertyDesc to populate PropertyType
'SC      22/02/2006          Ensure that the log file is closed properly on exit
'SC      06/03/2006          UAT Defect 586: Present condition valuation not being populated for Hometrack valuations
'SC      06/03/2006          UAT Defect 1343: Income multiplier type H added
'SC      17/03/2006          UAT Defect 1728: Take the delphi score instead of the detect score for the credit score
'SC      27/03/2006          UAT Defect 1754: Look for the validation type of 'C' instead of 'CH' for properties
'                            with a tenure of Commonhold
'SC      27/03/2006          UAT Defect 1489: Rewrite of valuation fees logic
'SC      03/04/2006          UAT Defect 2170: If the income basis can not be derived, then do not append the xmlIncomeBasis node
'SC      18/04/2006          UAT Defect 2542: Changes to LTV logic to cater for a LTV which is a decimal OR integer
'SC      02/05/2006          UAT Defect 2950: Changes to SetupFees procedure to cater for the 'Already Refunded' payment event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'EPSOM Specific History
'
'Prog   Date        AQR     Description
'SAB    10/05/2006  EP574   Create initial version for Completion Handover document
'SAB    10/05/2006  EP574   Create SetupBankAccountDetails function
'SAB    10/05/2006  EP574   Create SetupSolicitorDetails function
'SAB    10/05/2006  EP574   Create SetupValuerDetails function
'SAB    10/05/2006  EP574   Create SetupPeripheralSecurityDetails function
'SAB    10/05/2006  EP574   Create SetupInvestorDetails function
'SAB    10/05/2006  EP574   Create SetupRetentionDetails function
'SAB    10/05/2006  EP574   Create SetupLoanDetails function
'SAB    10/05/2006  EP574   Create SetupChargeDetails function
'SAB    10/05/2006  EP574   Create SetupBrokerDetails function
'SAB    10/05/2006  EP574   Amended SetupPropertyDetails function for EPSOM
'SAB    10/05/2006  EP574   Amended SetupCustomerCreation function for EPSOM
'SAB    10/05/2006  EP574   Amended AddCustomer function for EPSOM
'SAB    10/05/2006  EP574   Amended AddCustomerAddress function for EPSOM
'SAB    10/05/2006  EP574   Amended SetupFees function for EPSOM
'SAB    19/05/2006  EP583   Amended SetupCompletionBlock - removed redundant code and added new code
'SAB    19/05/2006  EP583   Amended SetupPeripheralSecurityDetails - additional code for identified data
'SAB    19/05/2006  EP583   Create SetupMortgageFundingDetails function
'SAB    19/05/2006  EP583   Amended SetupRetentionDetails - Amended, so that only displayed if there is a retention mount
'SAB    19/05/2006  EP583   Amended SetupLoanDetails function for addtional data, including new Rate calculation process
'SAB    19/05/2006  EP583   Amended SetupChargeDetails function for addtional data, including new CCJ data
'SAB    19/05/2006  EP583   Amended SetupPropertyDetails - added new code for missing data
'SAB    19/05/2006  EP583   Amended SetupCustomerCreation - added new code for missing data
'SAB    19/05/2006  EP583   Amended AddCustomer - added new code for missing data
'SAB    19/05/2006  EP583   Amended SetupFees - added new code for missing data
'SAB    19/05/2006  EP583   Amended SetupChargeDetails to allow for not existence of BESPOKEBUREAUDATA data
'SAB    24/05/2006  EP607   Create SetupHeaderDetails function
'SAB    24/05/2006  EP607   Amend SetupCompletionBlock to call SetupHeaderDetails
'SAB    24/05/2006  EP607   Amend SetupBankAccountDetails to use REPAYMENTBANKACCOUNTINDICATOR to identify correct Bank Account
'                           Also amended to check for NameAndAddressDirectory and PanelBankAccount if ThirdParty does not exist
'SAB    24/05/2006  EP607   Amend SetupInvestorDetails to default to "Deutsche Bank" not "Barclays"
'SAB    24/05/2006  EP607   Amendments to SetupLoanDetails for TermOfLoan description, xmlComponentPaymentOptionsNode node name,
'                           ProductScheme fix and LTV Banding combo changed
'PE     08/06/2006  EP694   Populated missing fields
'PE     12/06/2006  EP729   Handover Doc - Further Enhancements
'PE     14/06/2006  EP768   Added Broker Number to Handover document
'SW     20/06/2006  EP771   Added Repayment Vehicle to SetupLoanDetails
'PE     30/06/2006  EP928   Empty attribute causing conversion errors. Updated all conversion to "safe" conversions.
'DC     31/07/2006  EP1055  Defensive code in case second customer doesn't have a seperate address
'PE     22/08/2006  EP1093  Zero gross allowable income causing divide by zero problem
'PE     23/08/2006  EP1090  Retrieve valuer company name from nameandaddressdirectory
'LH     24/08/2006  EP1090  Product scheme & Omiga Product Code are now retrieved
'LH     29/08/2006  EP1090  Product scheme text is now retrieved
'AW     13/10/2006  EP1211  Corrected date processing around bankruptcy checks
'DRC    05/03/2007  EP2_956 Changes for Epsom 2 - Mainly Introducer, Multi Loan Cmponent
'AW     28/10/2007  EP2_2125 DBM178  Amended AddCustomer()
'                                    Added SetupCashbackDetails()
'                                    Re-write SetupRetentionDetails(), SetupFees()
'AW     30/03/2007  EP2_956 Minor formatting issues
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Implements ObjectControl

Private gstrRuleNumber As String

Private gobjContext As ObjectContext
Private m_xmlApplication As IXMLDOMNode
Private m_xmlApplicationFactFind As IXMLDOMNode
Private m_xmlPayment As IXMLDOMNode
Private m_xmlMortgageSubQuote As IXMLDOMNode
Private m_xmlFirstLoanComp As IXMLDOMNode
Private m_xmlCustomer As IXMLDOMNode
Private m_xmlNumbers As IXMLDOMNode
Private m_xmlDisbPayment As IXMLDOMNode
Private m_xmlRefundNode As IXMLDOMNode
Private m_xmlIncentiveNode As IXMLDOMNode

Private m_strCustomerEntityNumber As String
Private m_strSolicitorCIFNumber As String
Private m_strHdrMortgagor As String          ' KN - SYS2719 - 31/10/01

'SYS4867 - Use a more accurate datatype (variant/decimal).
'Private m_lngRetentionAmount As Long         ' KN - SYS2719 - 31/10/01
Private m_vdecRetentionAmount As Variant

Private m_intMaxTermInMonths As Integer      ' KN - SYS2719 - 31/10/01
Private m_strDDReference As String           ' KN - SYS2719 - 31/10/01
Private m_strInterestRateEndDate As String   ' KN - SYS2719 - 08/11/01
Private m_strInterestRateStartDate As String ' KN - SYS2719 - 08/11/01
Private m_blnJointApplication As Boolean
Private m_blnRetentionExists As Boolean

Private m_strAddressPostCode As String              ' SAB - EPSOM - 16/05/2006
Private m_strMaxValInstructionSeqNo As String       ' SAB - EPSOM - 17/05/2006
Private m_dblJointGrossAllowableIncome As Double    ' SAB - EPSOM - 18/05/2006
Private m_dblApplicationFeeAmount As Double         ' SAB - EPSOM - 18/05/2006

Private m_xmlLendingBranchCombo As IXMLDOMNode      'DRC - EP2_956 - 27/02/2007
Private m_xmlNatureofBusinessCombo As IXMLDOMNode   'DRC - EP2_956 - 27/02/2007

Private Const EMPTYVALUE = "EMPTY"
Private Const EMPTYNUMBER = "0"
Private Const RESERVEDVALUE = "RESERVED"
Private Const DOGUARANTOR = False            ' KN - SYS2719 - 21/11/01
Private Const DOPRODUCTINCENTIVE = False     ' KN - SYS2719 - 21/11/01
Private Const DEFAULTNAMESPACE = "" 'SD - MAR47 - 13/10/05
'When set to TRUE, the two flags above will result in nodes for
'Guarantor & ProductIncentive ALWAYS being created
'If the required data is not present in the database the
'attributes will be populated with "EMPTY" OR "0"

Private Enum AdhocType
    adhSolicitor
    adhPayee
End Enum

Private Enum RefundType
    rftFeeRefund
    rftIncentiveRefund
End Enum

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
    Set m_xmlApplication = Nothing
    Set m_xmlApplicationFactFind = Nothing
    Set m_xmlPayment = Nothing
    Set m_xmlMortgageSubQuote = Nothing
    Set m_xmlFirstLoanComp = Nothing
    Set m_xmlCustomer = Nothing
    Set m_xmlNumbers = Nothing
    Set m_xmlDisbPayment = Nothing
    Set m_xmlRefundNode = Nothing
    Set m_xmlIncentiveNode = Nothing
End Sub

Public Function OmRequest(ByVal vstrXMLRequest As String) As String
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

    On Error GoTo OmRequestVbErr
    
    Const strFunctionName As String = "OmRequest"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    'Dim xmlInDoc As New MSXML2.DOMDocument40
    'Dim xmlOutDoc As New MSXML2.DOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    'Dim xmlResponseNode As MSXML2.IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
        
    ' will be superceded by any subsequent SetAbort
    'SD 06/10/2005 gobjContext is empty for some unknown reason
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
    End If
    
    gobjContext.SetComplete
    
    ' create REQUEST & RESPONSE DOMDocuments
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    xmlInDoc.loadXML vstrXMLRequest
    
    Call Trace.TraceMain
        
    LogDetails 1, "------------------------------------------------------------------------------------------------------------"
    LogDetails 1, "OmRequest : Start "
       
    ' check for request node
    Set xmlRequestNode = GetMandatoryNode(xmlInDoc, "REQUEST")
    
    LogDetails 1, "Application Number : " & GetAttributeText(xmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONNUMBER")
    LogDetails 1, "Saving the Request to REQ_RES_Files folder", xmlInDoc, "CompletionsRequest"
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        'xmlInDoc.Save ("C:\Mars\Request.xml")
        
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
    'xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    'xmlOutDoc.appendChild xmlResponseNode
    'xmlOutDoc.Save ("C:\Mars\outRequest.xml")
    OmRequest = xmlOutDoc.xml
    Debug.Print OmRequest
    
    LogDetails 1, "Saving the Response to REQ_RES_Files folder", xmlOutDoc, "CreateMortgageAccountRequest"
    
    LogDetails 1, "OmRequest : End "
    
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
            
    'SC 22/02/2006: close file
    LogDetails 10, ""
    
    Exit Function

OmRequestVbErr:
    
    ' create ERROR response block from Err object
    OmRequest = CreateErrorResponse()
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
'------------------------------------------------------------------------------------------

    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
        
    Dim strOperation As String
    
    LogDetails 3, "DoRequest : Start "
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
    
        Case "RUNCOMPLETIONRULES"
            RunCompletionRules vxmlRequestNode, vxmlResponseNode
        
        Case Else
            Err.Raise eNOTMIPLEMENTED, _
                      strFunctionName, _
                      "OPERATION " & _
                      vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text & _
                      " Not implemented"
    End Select
    
    LogDetails 3, "DoRequest : End "
    
DoRequestExit:
    
    CheckError strFunctionName, TypeName(Me)

End Sub

Private Sub RunCompletionRules(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Run business rules to create a Completions Request block to be passed to OSG.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'------------------------------------------------------------------------------------------

On Error GoTo RunCompletionRulesExit
    
Const strFunctionName As String = "RunCompletionRules"

    Dim xmlNode As IXMLDOMNode
    Dim intFailCount As Integer
    
    Dim xmlBatchAuditNode As IXMLDOMNode
    Dim xmlPaymentRecordNode As IXMLDOMNode
    
    LogDetails 3, "RunCompletionRules : Start "
    
    Set xmlBatchAuditNode = vxmlRequestNode.selectSingleNode("BATCHAUDIT")
    Set xmlPaymentRecordNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    
    'EP2_956
    Set m_xmlLendingBranchCombo = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/COMBOGROUP[@GROUPNAME='LendingBranch']")
    Set m_xmlNatureofBusinessCombo = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/COMBOGROUP[@GROUPNAME='NatureOfBusiness']")
    If Not xmlBatchAuditNode Is Nothing Then
        vxmlResponseNode.appendChild xmlBatchAuditNode.cloneNode(True)
    End If
    If Not xmlPaymentRecordNode Is Nothing Then
        vxmlResponseNode.appendChild xmlPaymentRecordNode.cloneNode(True)
    End If
    
    CopyAttribute vxmlRequestNode, vxmlResponseNode, "USERID"
    CopyAttribute vxmlRequestNode, vxmlResponseNode, "UNITID"
    
    'vxmlResponseNode.appendChild vxmlRequestNode.cloneNode(True)
    'For Each xmlNode In vxmlRequestNode.childNodes
    '    vxmlResponseNode.appendChild xmlNode.cloneNode(True)
    'Next

    'intFailCount = intFailCount + Rule3005(xmlApplicationNode, xmlResponseNode)
    'intFailCount = intFailCount + Rule3801(xmlApplicationNode, xmlResponseNode)

    Call SetupCompletionBlock(vxmlRequestNode, vxmlResponseNode)
    
    Set xmlNode = vxmlResponseNode.ownerDocument.createAttribute("TYPE")
    xmlNode.Text = "SUCCESS"
    vxmlResponseNode.Attributes.setNamedItem xmlNode
    
    Set xmlNode = vxmlResponseNode.ownerDocument.createAttribute("FAILCOUNT")
    xmlNode.Text = intFailCount
    vxmlResponseNode.Attributes.setNamedItem xmlNode
    
    Set xmlNode = Nothing
    
    LogDetails 3, "RunCompletionRules : End"

RunCompletionRulesExit:
     
    Set xmlNode = Nothing
    Set xmlBatchAuditNode = Nothing
    Set xmlPaymentRecordNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)

End Sub

Private Sub SetupCompletionBlock(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupCompletionBlock
' Author:       Sam Dali
' Date:         11/08/05
' Purpose:      Create Completion block, the outer most block in validation response
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to
'------------------------------------------------------------------------

On Error GoTo SetupCompletionBlockExit

Const strFunctionName As String = "SetupCompletionBlock"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlCompletionBlock As IXMLDOMNode
Dim xmlPropertyDetailsBlock As IXMLDOMNode

LogDetails 3, "SetupCompletionBlock : Start"

Set xmlDoc = New FreeThreadedDOMDocument40
xmlDoc.validateOnParse = False
xmlDoc.setProperty "NewParser", True


'SAB 18/05/2006 EP574 - Amended for EPSOM to create initial version for Completion Handover document

Set xmlCompletionBlock = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CompletionsHandover", "")


'set up Header
Call SetupHeaderDetails(vxmlRequestNode, xmlCompletionBlock)


'set up CustomerCreation
Call SetupCustomerCreation(vxmlRequestNode, xmlCompletionBlock)

'set up data for BankAccount
Call SetupBankAccountDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Solicitor Details
Call SetupSolicitorDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Broker Details
Call SetupBrokerDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Valuer Details
Call SetupValuerDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for PropertyDetails
Call SetupPropertyDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Charges
Call SetupChargeDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Component
Call SetupLoanDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Investor
'For Epsom 2 this will be handled in SetupLoanDetails as it's part of
' multi loan component
'Call SetupInvestorDetails(vxmlRequestNode, xmlCompletionBlock)

'set up data for Peripheral Security
Call SetupPeripheralSecurityDetails(vxmlRequestNode, xmlCompletionBlock)

'AW 28/03/07    EP2_2125
'set up data for Cashback details
Call SetupCashbackDetails(vxmlRequestNode, xmlCompletionBlock)

'set up Retention details (if nothing then do not display)
Call SetupRetentionDetails(vxmlRequestNode, xmlCompletionBlock)

'set up Fees
Call SetupFees(vxmlRequestNode, xmlCompletionBlock)

'set up Mortgage Funding
Call SetupMortgageFundingDetails(vxmlRequestNode, xmlCompletionBlock)



vxmlResponseNode.appendChild xmlCompletionBlock

LogDetails 3, "SetupCompletionBlock : End"

SetupCompletionBlockExit:
    Set xmlCompletionBlock = Nothing
    Set xmlDoc = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub
Private Sub SetupGenericMessageElements(ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo SetupGenericMessageElementsExit
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim xmlClientDeviceNode As IXMLDOMNode
    Dim xmlTellerIDNode As IXMLDOMNode
    Dim xmlTellerPwdNode As IXMLDOMNode
    Dim xmlProxyID As IXMLDOMNode
    Dim xmlProxyPwdNode As IXMLDOMNode
    Dim xmlOperatorNode As IXMLDOMNode
    Dim xmlProductTypeNode As IXMLDOMNode
    Dim xmlSessionIDNode As IXMLDOMNode
    Dim xmlCustomerNumberNode As IXMLDOMNode
    Dim xmlCommunicationChannelNode As IXMLDOMNode
    Dim xmlCommunicationDirectionNode As IXMLDOMNode
    Dim xmlServiceNameNode As IXMLDOMNode
    
    LogDetails 3, "SetupGenericMessageElements : Start "
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlClientDeviceNode = xmlDoc.createElement("ClientDevice")
    Set xmlTellerIDNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TellerID", "")
    Set xmlTellerPwdNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TellerPwd", "")
    Set xmlProxyID = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProxyID", "")
    Set xmlProxyPwdNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProxyPwd", "")
    Set xmlOperatorNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Operator", "")
    Set xmlProductTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProductType", "")
    Set xmlSessionIDNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SessionID", "")
    Set xmlCommunicationChannelNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommunicationChannel", "")
    Set xmlCommunicationDirectionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommunicationDirection", "")
    Set xmlServiceNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ServiceName", "")
    Set xmlCustomerNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CustomerNumber", "")
    
    vxmlResponseNode.appendChild xmlClientDeviceNode
    vxmlResponseNode.appendChild xmlTellerIDNode
    vxmlResponseNode.appendChild xmlTellerPwdNode
    vxmlResponseNode.appendChild xmlProxyID
    vxmlResponseNode.appendChild xmlProxyPwdNode
    vxmlResponseNode.appendChild xmlOperatorNode
    vxmlResponseNode.appendChild xmlProductTypeNode
    vxmlResponseNode.appendChild xmlSessionIDNode
    vxmlResponseNode.appendChild xmlCommunicationChannelNode
    vxmlResponseNode.appendChild xmlCommunicationDirectionNode
    vxmlResponseNode.appendChild xmlServiceNameNode
    vxmlResponseNode.appendChild xmlCustomerNumberNode
    
    LogDetails 3, "SetupGenericMessageElements : End "
    
SetupGenericMessageElementsExit:
    Set xmlClientDeviceNode = Nothing
    Set xmlTellerIDNode = Nothing
    Set xmlTellerPwdNode = Nothing
    Set xmlProxyID = Nothing
    Set xmlProxyPwdNode = Nothing
    Set xmlOperatorNode = Nothing
    Set xmlProductTypeNode = Nothing
    Set xmlSessionIDNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlCommunicationChannelNode = Nothing
    Set xmlCommunicationDirectionNode = Nothing
    Set xmlServiceNameNode = Nothing
    
End Sub


Private Sub SetupBankAccountDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupBankAccountDetails
' Author:       Steve Badman
' Date:         10/05/2006
' Purpose:      Create Bank account details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupBankAccountDetailsExit
    
    Const strFunctionName As String = "SetupBankAccountDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlApplicationBankBuildingSOCNode As IXMLDOMNode
    Dim strSortCode As String
    Dim xmlApplicantBankAccountDetailsNode As IXMLDOMNode
    Dim xmlBankAccountNameNode As IXMLDOMNode
    Dim xmlBankAccountSortCodeNode As IXMLDOMNode
    Dim xmlBankAccountNumberNode As IXMLDOMNode
    
    LogDetails 3, "SetupBankAccountDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlApplicantBankAccountDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ApplicantBankAccountDetails", "")
    
    Dim strRuleName As String
    strRuleName = "SetupBankAccountDetails"
    
    'SAB 22/05/2006 - EPSOM - Amended to filter in case the Applicant has more than one Bank Account
    'Set xmlApplicationBankBuildingSOCNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC")
    Set xmlApplicationBankBuildingSOCNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC[@REPAYMENTBANKACCOUNTINDICATOR='1']")
    
    Set xmlBankAccountNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AccountName", "")
    Set xmlBankAccountSortCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SortCode", "")
    Set xmlBankAccountNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AccountNumber", "")

    If Not xmlApplicationBankBuildingSOCNode Is Nothing Then
        'AccountName
        gstrRuleNumber = strRuleName & ".1"
        xmlBankAccountNameNode.Text = GetAttributeText(xmlApplicationBankBuildingSOCNode, "ACCOUNTNAME")
        xmlApplicantBankAccountDetailsNode.appendChild xmlBankAccountNameNode
        
        
        'SortCode
        gstrRuleNumber = strRuleName & ".2"
        strSortCode = GetAttributeText(xmlApplicationBankBuildingSOCNode.selectSingleNode("THIRDPARTY"), "THIRDPARTYBANKSORTCODE")
        
        'SAB 22/05/2006 - EPSOM - Amended to check for NameAndAddressDirectory and PanelBankAccount if ThirdParty does not exist
        If strSortCode = "" Then
            strSortCode = GetAttributeText(xmlApplicationBankBuildingSOCNode.selectSingleNode("NAMEANDADDRESSDIRECTORY"), "NAMEANDADDRESSBANKSORTCODE")
            If strSortCode = "" Then
                strSortCode = GetAttributeText(xmlApplicationBankBuildingSOCNode.selectSingleNode("PANEL/PANELBANKACCOUNT"), "BANKSORTCODE")
            End If
        End If
        
        'If Not strSortCode = "" Then
        '    strSortCode = Replace(strSortCode, "-", "")
        'End If
        
        xmlBankAccountSortCodeNode.Text = strSortCode
        xmlApplicantBankAccountDetailsNode.appendChild xmlBankAccountSortCodeNode
        
        
        'AccountNumber
        gstrRuleNumber = strRuleName & ".3"
        xmlBankAccountNumberNode.Text = GetAttributeText(xmlApplicationBankBuildingSOCNode, "ACCOUNTNUMBER")
        xmlApplicantBankAccountDetailsNode.appendChild xmlBankAccountNumberNode
    End If

    vxmlResponseNode.appendChild xmlApplicantBankAccountDetailsNode
    
    LogDetails 3, "SetupBankAccountDetails : End "
    
    
SetupBankAccountDetailsExit:
    Set xmlDoc = Nothing
    Set xmlApplicationBankBuildingSOCNode = Nothing
    Set xmlApplicantBankAccountDetailsNode = Nothing
    Set xmlBankAccountNameNode = Nothing
    Set xmlBankAccountSortCodeNode = Nothing
    Set xmlBankAccountNumberNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetupSolicitorDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupSolicitorDetails
' Author:       Steve Badman
' Date:         10/05/2006
' Purpose:      Create Solicitor Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupSolicitorDetailsExit
    
    Const strFunctionName As String = "SetupSolicitorDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlSolicitorDetailsNode As IXMLDOMNode
    Dim xmlApplicationLegalRepNode As IXMLDOMNode
    Dim xmlSearchNode As IXMLDOMNode 'temp node used for searching
    Dim xmlSolicitorNumberNode As IXMLDOMNode
    Dim xmlSolicitorCompanyNameNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
    Dim xmlSolicitorAddressNode As IXMLDOMNode
    Dim xmlFlatNameOrNumberNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNameNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNumberNode As IXMLDOMNode
    Dim xmlStreetNode As IXMLDOMNode
    Dim xmlDistrictNode As IXMLDOMNode
    Dim xmlTownOrCityNode As IXMLDOMNode
    Dim xmlCountyNode As IXMLDOMNode
    Dim xmlPostCodeNode As IXMLDOMNode
    Dim strRuleName As String
    Dim strAddressGUID As String
    Dim xmlContactTelephoneDetailsNodeList As IXMLDOMNodeList
    Dim xmlContactTelephoneDetailsNode As IXMLDOMNode
    Dim strUsage As String
    Dim xmlSolWorkPhoneNode As IXMLDOMNode
    Dim xmlSolOtherPhoneNode As IXMLDOMNode
    Dim xmlAreaCodeNode As IXMLDOMNode
    Dim xmlLocalNumberNode As IXMLDOMNode
    Dim xmlExtensionNode As IXMLDOMNode
    Dim xmlDXNumberNode As IXMLDOMNode
    Dim xmlDXLocationNode As IXMLDOMNode

    
    LogDetails 3, "SetupSolicitorDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlSolicitorDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorDetails", "")
    
    strRuleName = "SetupSolicitorDetails"
    gstrRuleNumber = strRuleName
    
        
    Set xmlApplicationLegalRepNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP")
    

    If Not xmlApplicationLegalRepNode Is Nothing Then
        
        'get the NAMEANDADDRESSDIRECTORY node or THIRDPARTY node
        Set xmlSearchNode = xmlApplicationLegalRepNode.selectSingleNode("NAMEANDADDRESSDIRECTORY")
        If xmlSearchNode Is Nothing Then
            Set xmlSearchNode = xmlApplicationLegalRepNode.selectSingleNode("THIRDPARTY")
        End If
        
   
        'Solicitor Number
        gstrRuleNumber = strRuleName & ".1"
        Set xmlSolicitorNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorNumber", "")
        xmlSolicitorNumberNode.Text = "N/A"
        xmlSolicitorDetailsNode.appendChild xmlSolicitorNumberNode
   
   
        'Solicitor Company Name
        gstrRuleNumber = strRuleName & ".2"
        Set xmlSolicitorCompanyNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CompanyName", "")
        xmlSolicitorCompanyNameNode.Text = GetAttributeText(xmlSearchNode, "COMPANYNAME")
        xmlSolicitorDetailsNode.appendChild xmlSolicitorCompanyNameNode
   
        
        'get AddressGUID and Address node
        strAddressGUID = GetAttributeText(xmlSearchNode, "ADDRESSGUID")
        
        If strAddressGUID <> "" Then
            Set xmlAddressNode = xmlSearchNode.selectSingleNode("ADDRESS")
        End If
        
        'set all nodes, dependant on address
        Set xmlSolicitorAddressNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorAddress", "")
        Set xmlFlatNameOrNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FlatNameOrNumber", "")
        Set xmlHouseOrBuildingNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingName", "")
        Set xmlHouseOrBuildingNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingNumber", "")
        Set xmlStreetNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Street", "")
        Set xmlDistrictNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "District", "")
        Set xmlTownOrCityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TownOrCity", "")
        Set xmlCountyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "County", "")
        Set xmlPostCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PostCode", "")
        
        
        If Not xmlAddressNode Is Nothing Then
            'FlatNameOrNumber
            gstrRuleNumber = strRuleName & ".3"
            xmlFlatNameOrNumberNode.Text = GetAttributeText(xmlAddressNode, "FLATNUMBER")
            
            'HouseOrBuildingName
            gstrRuleNumber = strRuleName & ".4"
            xmlHouseOrBuildingNameNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENAME")
            
            'HouseOrBuildingNumber
            gstrRuleNumber = strRuleName & ".5"
            xmlHouseOrBuildingNumberNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENUMBER")
            
            'Street
            gstrRuleNumber = strRuleName & ".6"
            xmlStreetNode.Text = GetAttributeText(xmlAddressNode, "STREET")
            
            'District
            gstrRuleNumber = strRuleName & ".7"
            xmlDistrictNode.Text = GetAttributeText(xmlAddressNode, "DISTRICT")
            
            'TownOrCity
            gstrRuleNumber = strRuleName & ".8"
            xmlTownOrCityNode.Text = GetAttributeText(xmlAddressNode, "TOWN")
            
            'County
            gstrRuleNumber = strRuleName & ".9"
            xmlCountyNode.Text = GetAttributeText(xmlAddressNode, "COUNTY")
            
            'PostCode
            gstrRuleNumber = strRuleName & ".10"
            xmlPostCodeNode.Text = GetAttributeText(xmlAddressNode, "POSTCODE")
        End If
            
        xmlSolicitorAddressNode.appendChild xmlFlatNameOrNumberNode
        xmlSolicitorAddressNode.appendChild xmlHouseOrBuildingNameNode
        xmlSolicitorAddressNode.appendChild xmlHouseOrBuildingNumberNode
        xmlSolicitorAddressNode.appendChild xmlStreetNode
        xmlSolicitorAddressNode.appendChild xmlDistrictNode
        xmlSolicitorAddressNode.appendChild xmlTownOrCityNode
        xmlSolicitorAddressNode.appendChild xmlCountyNode
        xmlSolicitorAddressNode.appendChild xmlPostCodeNode
        
        xmlSolicitorDetailsNode.appendChild xmlSolicitorAddressNode
    
    
    
    
        'get the solicitor contact telephone details
        
        Set xmlContactTelephoneDetailsNodeList = xmlSearchNode.selectNodes("CONTACTTELEPHONEDETAILS")
        
        For Each xmlContactTelephoneDetailsNode In xmlContactTelephoneDetailsNodeList
            strUsage = GetAttributeText(xmlContactTelephoneDetailsNode, "USAGE")
            If Not strUsage = "" Then
                If IsValidationType("ContactTelephoneUsage", CSafeInt(strUsage), "W") Then   '"W"=Work
                    'work phone number
                    gstrRuleNumber = strRuleName & ".12"
                    Set xmlSolWorkPhoneNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorWorkPhoneNumber", "")
                    
                    Set xmlAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
                    Set xmlLocalNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
                    Set xmlExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
                    
                    xmlAreaCodeNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "AREACODE")
                    xmlLocalNumberNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "TELENUMBER")
                    xmlExtensionNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "EXTENSIONNUMBER")
                    
                    xmlSolWorkPhoneNode.appendChild xmlAreaCodeNode
                    xmlSolWorkPhoneNode.appendChild xmlLocalNumberNode
                    xmlSolWorkPhoneNode.appendChild xmlExtensionNode
                
                    xmlSolicitorDetailsNode.appendChild xmlSolWorkPhoneNode
                Else
                    'Other Phone number
                    gstrRuleNumber = strRuleName & ".11"
                
                    Set xmlSolOtherPhoneNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorOtherPhoneNumber", "")
                    
                    Set xmlAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
                    Set xmlLocalNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
                    Set xmlExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
                    
                    xmlAreaCodeNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "AREACODE")
                    xmlLocalNumberNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "TELENUMBER")
                    xmlExtensionNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "EXTENSIONNUMBER")
                    
                    xmlSolOtherPhoneNode.appendChild xmlAreaCodeNode
                    xmlSolOtherPhoneNode.appendChild xmlLocalNumberNode
                    xmlSolOtherPhoneNode.appendChild xmlExtensionNode
                
                    xmlSolicitorDetailsNode.appendChild xmlSolOtherPhoneNode
                End If
            End If
        Next


        'Exchange service/number
        gstrRuleNumber = strRuleName & ".13"
        Set xmlDXNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DXNumber", "")
        xmlDXNumberNode.Text = GetAttributeText(xmlApplicationLegalRepNode, "DXID")
        xmlSolicitorDetailsNode.appendChild xmlDXNumberNode
    
    
        'Exchange city
        gstrRuleNumber = strRuleName & ".14"
        Set xmlDXLocationNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DXLocation", "")
        xmlDXLocationNode.Text = GetAttributeText(xmlApplicationLegalRepNode, "DXLOCATION")
        xmlSolicitorDetailsNode.appendChild xmlDXLocationNode
    End If


    vxmlResponseNode.appendChild xmlSolicitorDetailsNode
    
    LogDetails 3, "SetupSolicitorDetails : End "
    
    
SetupSolicitorDetailsExit:
    Set xmlDoc = Nothing
    Set xmlSolicitorDetailsNode = Nothing
    Set xmlApplicationLegalRepNode = Nothing
    Set xmlSearchNode = Nothing
    Set xmlSolicitorNumberNode = Nothing
    Set xmlSolicitorCompanyNameNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlSolicitorAddressNode = Nothing
    Set xmlFlatNameOrNumberNode = Nothing
    Set xmlHouseOrBuildingNameNode = Nothing
    Set xmlHouseOrBuildingNumberNode = Nothing
    Set xmlStreetNode = Nothing
    Set xmlDistrictNode = Nothing
    Set xmlTownOrCityNode = Nothing
    Set xmlCountyNode = Nothing
    Set xmlPostCodeNode = Nothing
    Set xmlContactTelephoneDetailsNodeList = Nothing
    Set xmlContactTelephoneDetailsNode = Nothing
    Set xmlSolWorkPhoneNode = Nothing
    Set xmlSolOtherPhoneNode = Nothing
    Set xmlAreaCodeNode = Nothing
    Set xmlLocalNumberNode = Nothing
    Set xmlExtensionNode = Nothing
    Set xmlDXNumberNode = Nothing
    Set xmlDXLocationNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetupValuerDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupValuerDetails
' Author:       Steve Badman
' Date:         12/05/2006
' Purpose:      Create Valuer Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupValuerDetailsExit
    
    Const strFunctionName As String = "SetupValuerDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlValuerDetailsNode As IXMLDOMNode
    Dim xmlValuerInstructionNode As IXMLDOMNode
    Dim xmlValuerNumberNode As IXMLDOMNode
    Dim xmlValuerCompanyNameNode As IXMLDOMNode
    Dim xmlDirectoryNode As IXMLDOMNode
    
    LogDetails 3, "SetupValuerDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlValuerDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValuerDetails", "")
    
    strRuleName = "SetupValuerDetails"
    
        
    Set xmlValuerInstructionNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION")

    If Not xmlValuerInstructionNode Is Nothing Then
   
        'Valuer Number
        gstrRuleNumber = strRuleName & ".1"
        Set xmlValuerNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValuerNumber", "")
        xmlValuerNumberNode.Text = "N/A"
        xmlValuerDetailsNode.appendChild xmlValuerNumberNode
   
   
        'Valuer Company Name
        gstrRuleNumber = strRuleName & ".2"
        Set xmlValuerCompanyNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CompanyName", "")
        'EP1090 - 23/08/2006 - Peter Edney
        Set xmlDirectoryNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION/NAMEANDADDRESSDIRECTORY")
        If Not (xmlDirectoryNode Is Nothing) Then
            xmlValuerCompanyNameNode.Text = GetAttributeText(xmlDirectoryNode, "COMPANYNAME")
        Else
            xmlValuerCompanyNameNode.Text = GetAttributeText(xmlValuerInstructionNode, "VALUERNAME")
        End If
        xmlValuerDetailsNode.appendChild xmlValuerCompanyNameNode
    End If


    vxmlResponseNode.appendChild xmlValuerDetailsNode
    
    LogDetails 3, "SetupValuerDetails : End "
    
    
SetupValuerDetailsExit:
    Set xmlDoc = Nothing
    Set xmlValuerDetailsNode = Nothing
    Set xmlValuerInstructionNode = Nothing
    Set xmlValuerNumberNode = Nothing
    Set xmlValuerCompanyNameNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub


Private Sub SetupHeaderDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupHeaderDetails
' Author:       Steve Badman
' Date:         22/05/2006
' Purpose:      Create Header Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupHeaderDetailsExit
    
    Const strFunctionName As String = "SetupHeaderDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlHeaderDetailsNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlFeesNodeList As IXMLDOMNodeList
    Dim xmlFeesNode As IXMLDOMNode
    Dim dblFeeAmount As Double
    
    LogDetails 3, "SetupHeaderDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlHeaderDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HeaderDetails", "")
    strRuleName = "SetupHeaderDetails"
    
    
    
    'Loan No.
    gstrRuleNumber = strRuleName & ".1"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LoanNo", "")
    xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONNUMBER")
    xmlHeaderDetailsNode.appendChild xmlTempNode
    
        
    'Submitted
    gstrRuleNumber = strRuleName & ".2"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Submitted", "")
    'EP2_956 DRC
    xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONDATE")
    xmlHeaderDetailsNode.appendChild xmlTempNode
        

    'Introducer
    gstrRuleNumber = strRuleName & ".3"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Introducer", "")
    'EP2_956 - use the new FSA tables
    xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode(".//ARFIRM"), "ARFIRMNAME")
    If Len(xmlTempNode.Text) = 0 Then
      xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode(".//PRINCIPALFIRMBROKER"), "PRINCIPALFIRMNAME")
    End If
    
    xmlHeaderDetailsNode.appendChild xmlTempNode


    'Introducer Fee
    gstrRuleNumber = strRuleName & ".4"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IntroducerFee", "")
    'EP2_956 - Total all the MortgageIntroducerFees
    Set xmlFeesNodeList = vxmlRequestNode.selectNodes(".//MORTGAGEINTRODUCERFEE")
    dblFeeAmount = 0
    For Each xmlFeesNode In xmlFeesNodeList
       dblFeeAmount = dblFeeAmount + GetAttributeAsDouble(xmlFeesNode, "FEEAMOUNT")
    Next
    If dblFeeAmount > 0.1 Then
       xmlTempNode.Text = CStr(dblFeeAmount)
    End If
    'EP2_956 - End
    xmlHeaderDetailsNode.appendChild xmlTempNode




    vxmlResponseNode.appendChild xmlHeaderDetailsNode
    
    LogDetails 3, "SetupHeaderDetails : End "
    
    
SetupHeaderDetailsExit:
    Set xmlDoc = Nothing
    Set xmlHeaderDetailsNode = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetupPeripheralSecurityDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupPeripheralSecurityDetails
' Author:       Steve Badman
' Date:         12/05/2006
' Purpose:      Create PeripheralSecurity Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupPeripheralSecurityDetailsExit
    
    Const strFunctionName As String = "SetupPeripheralSecurityDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlPeripheralSecurityDetailsNode As IXMLDOMNode
    Dim xmlForeignInsuranceNode As IXMLDOMNode
    Dim xmlCreditRatingNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim strGrossLoan As String
    Dim strNatureOfLoan As String
    Dim strNumberOfApplicants As String
    Dim intNumberOfApplicants As Integer
    Dim xmlIncomeSummary As IXMLDOMNodeList
    Dim xmlIncome As IXMLDOMNode
    Dim blnJointIncome As Boolean
    Dim strALLOWABLEANNUALINCOME As String
    Dim strALLOWABLEMONTHLYINCOME As String
    Dim strGROSSCONFIRMEDALLOWABLEINCOME As String
    Dim xmlBureauNode As IXMLDOMNode
    Dim xmlApplicationCreditCheckNodeList As IXMLDOMNodeList
    Dim xmlApplicationCreditCheckNode As IXMLDOMNode
    Dim strSequenceNumber As String
    Dim strCreditCheckGUID As String
    
    Dim xmlBureau2List As IXMLDOMNodeList
    Dim xmlBureau2Node As IXMLDOMNode
    
    LogDetails 3, "SetupPeripheralSecurityDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlPeripheralSecurityDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PeripheralSecurity", "")
    strRuleName = "SetupPeripheralSecurityDetails"
    
    
    ' ******************************************
    ' *** INSURANCE DATA NOT HELD FOR PHASE1 ***
    ' ******************************************
        
        
        
    ' *** Foreign Insurance : START ***
    Set xmlForeignInsuranceNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ForeignInsurance", "")
        
   
    'Foreign Insurance Type
    gstrRuleNumber = strRuleName & ".1"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ForeignInsuranceType", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode
    
    
    'Start/Received Date
    gstrRuleNumber = strRuleName & ".2"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "StartOrReceivedDate", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode
    
    
    'ExpiryDate
    gstrRuleNumber = strRuleName & ".3"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ExpiryDate", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode
    
    
    'ReferenceNumber
    gstrRuleNumber = strRuleName & ".4"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ReferenceNumber", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode
    
    
    'DescriptiveTextLine1
    gstrRuleNumber = strRuleName & ".5"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DescriptiveTextLine1", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode
    

    'DescriptiveTextLine2
    gstrRuleNumber = strRuleName & ".6"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DescriptiveTextLine2", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode


    'Policy/Security Value
    gstrRuleNumber = strRuleName & ".7"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PolicyOrSecurityValue", "")
    xmlTempNode.Text = "*** NOT IN PHASE 1 ***"
    xmlForeignInsuranceNode.appendChild xmlTempNode


    xmlPeripheralSecurityDetailsNode.appendChild xmlForeignInsuranceNode
    ' *** Foreign Insurance : END ***

    
    
    
    
    ' *** Credit Rating : START ***
    Set xmlCreditRatingNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CreditRating", "")
        
   
    'IncomeMethodUsed
    gstrRuleNumber = strRuleName & ".1"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IncomeMethodUsed", "")
    
    'EP694 - 07/06/2006 - Peter Edney
    strNatureOfLoan = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "NATUREOFLOAN")
    If IsValidationType("NatureOfLoan", strNatureOfLoan, "BR") Then
        xmlTempNode.Text = "Rental Income"
    Else
        strNumberOfApplicants = GetMandatoryAttributeText(GetMandatoryNode(vxmlRequestNode, "APPLICATION/APPLICATIONFACTFIND"), "NUMBEROFAPPLICANTS")
        If strNumberOfApplicants <> "" Then
            intNumberOfApplicants = CSafeInt(strNumberOfApplicants)
        Else
            intNumberOfApplicants = 0
        End If
        
        If intNumberOfApplicants > 1 Then
            blnJointIncome = True
            Set xmlIncomeSummary = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/INCOMESUMMARY")
            For Each xmlIncome In xmlIncomeSummary
                strALLOWABLEANNUALINCOME = GetAttributeText(xmlIncome, "ALLOWABLEANNUALINCOME")
                strALLOWABLEMONTHLYINCOME = GetAttributeText(xmlIncome, "ALLOWABLEMONTHLYINCOME")
                strGROSSCONFIRMEDALLOWABLEINCOME = GetAttributeText(xmlIncome, "GROSSCONFIRMEDALLOWABLEINCOME")
                If (strALLOWABLEANNUALINCOME = "" Or strALLOWABLEANNUALINCOME = "0") _
                    And (strALLOWABLEMONTHLYINCOME = "" Or strALLOWABLEMONTHLYINCOME = "0") _
                    And (strGROSSCONFIRMEDALLOWABLEINCOME = "" Or strGROSSCONFIRMEDALLOWABLEINCOME = "0") Then
                    blnJointIncome = False
                End If
            Next
            If blnJointIncome Then
                xmlTempNode.Text = "Joint Income"
            Else
                xmlTempNode.Text = "Single Income"
            End If
        Else
            xmlTempNode.Text = "Single Income"
        End If
    End If
    
    xmlCreditRatingNode.appendChild xmlTempNode
    
    'Bankruptcy/IVA
    gstrRuleNumber = strRuleName & ".2"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BankruptcyOrIVA", "")
        
    'EP694 - 07/06/2006 - Peter Edney
    'get the correct APPLICATIONCREDITCHECK data
    Set xmlApplicationCreditCheckNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK")

    For Each xmlApplicationCreditCheckNode In xmlApplicationCreditCheckNodeList
        'Ordered descending, so the 1st record will be the correct one
        strSequenceNumber = GetAttributeText(xmlApplicationCreditCheckNode, "SEQUENCENUMBER")
        strCreditCheckGUID = GetAttributeText(xmlApplicationCreditCheckNode, "CREDITCHECKGUID")
        Exit For
    Next
    
    Set xmlBureauNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/CREDITCHECK/BESPOKEBUREAUDATA[@CREDITCHECKGUID='" & strCreditCheckGUID & "']")
    If Not xmlBureauNode Is Nothing Then
        
        If GetAttributeText(xmlBureauNode, "ACTIVEBANKRUPTCY") <> "N" Then
           xmlTempNode.Text = "Current CB"
        End If
        
        If xmlTempNode.Text = "" Then
            'AW  13/10/06  EP1211
            If GetAttributeText(xmlBureauNode, "DATELATESTBANKRUPTCY") <> "000000" Then
                'Date stored as varchar(6). Need to create a valid date for comparison
                Dim dtBankruptcyDate As Date
                Dim strBankruptcyDate As String
                
                strBankruptcyDate = GetAttributeText(xmlBureauNode, "DATELATESTBANKRUPTCY")
                strBankruptcyDate = Mid(strBankruptcyDate, 1, 2) & "/" & Mid(strBankruptcyDate, 3, 2) & "/" & Mid(strBankruptcyDate, 5, 2)
                dtBankruptcyDate = CSafeDate(strBankruptcyDate)

                If GetAge(dtBankruptcyDate) < 3 Then
                        xmlTempNode.Text = "Prior CB"
                End If
            End If
            'AW  13/10/06  EP1211 - End
        End If
            
        If xmlTempNode.Text = "" Then
            If GetAttributeText(xmlBureauNode, "IVAWORSTSTATUS") = "M" Then
                xmlTempNode.Text = "Current IVA"
            End If
        End If
        
        If xmlTempNode.Text = "" Then
        
            If GetAttributeText(xmlBureauNode, "IVAWORSTSTATUS") = "S" Then
                'AW  13/10/06  EP1211
                If GetAttributeText(xmlBureauNode, "IVADATE") <> "000000" Then
                    'Date stored as varchar(6). Need to create a valid date for comparison
                    Dim dtIVADate As Date
                    Dim strIVADate As String
                    
                    strIVADate = GetAttributeText(xmlBureauNode, "IVADATE")
                    strIVADate = Mid(strIVADate, 1, 2) & "/" & Mid(strIVADate, 3, 2) & "/" & Mid(strIVADate, 5, 2)
                    dtIVADate = CSafeDate(strIVADate)
    
                    If GetAge(dtIVADate) < 3 Then
                            xmlTempNode.Text = "Prior IVA"
                    End If
                End If
                'AW  13/10/06  EP1211 - End
            End If

        End If
        
        If xmlTempNode.Text = "" Then
            xmlTempNode.Text = "No"
        End If
    Else
        xmlTempNode.Text = "No"
    End If
    
    xmlCreditRatingNode.appendChild xmlTempNode
    
    'PreviousArrears
    gstrRuleNumber = strRuleName & ".3"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PreviousArrears", "")
    
    'EP694 - 07/06/2006 - Peter Edney
    xmlTempNode.Text = "N"
    Set xmlBureau2List = vxmlRequestNode.selectNodes("//APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/CREDITCHECK/BXCUSTOMERBUREAU2[@PERSONTYPE='1']")
    For Each xmlBureau2Node In xmlBureau2List
        If GetAttributeText(xmlBureau2Node, "BX18_SecuredArrears") = "A" Or GetAttributeText(xmlBureau2Node, "BX18_UnsecuredArrears") = "A" Then
            xmlTempNode.Text = "Y"
        End If
    Next

    xmlCreditRatingNode.appendChild xmlTempNode
    
    
    'Income For Affordability
    gstrRuleNumber = strRuleName & ".4"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IncomeForAffordability", "")
    
    'Use the Joint Income calculated from all Customer(s)
    xmlTempNode.Text = m_dblJointGrossAllowableIncome
    xmlCreditRatingNode.appendChild xmlTempNode
    
    
    'Debt to Income Ratio
    gstrRuleNumber = strRuleName & ".5"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DebtToIncomeRatio", "")
    
    strGrossLoan = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "TOTALLOANAMOUNT")
    'EP1093 - 22/08/2006 - Peter Edney
    'AW EP2_956 Restrict decimal places
    If (Not (strGrossLoan = "")) And (m_dblJointGrossAllowableIncome > 0) Then
        xmlTempNode.Text = Format((CSafeDbl(strGrossLoan) / m_dblJointGrossAllowableIncome), ".00")
    End If
    xmlCreditRatingNode.appendChild xmlTempNode
    
    
    xmlPeripheralSecurityDetailsNode.appendChild xmlCreditRatingNode
    ' *** Credit Rating : END ***
    
    
    
    
    vxmlResponseNode.appendChild xmlPeripheralSecurityDetailsNode
    
    LogDetails 3, "SetupPeripheralSecurityDetails : End "
    
    
SetupPeripheralSecurityDetailsExit:
    Set xmlDoc = Nothing
    Set xmlPeripheralSecurityDetailsNode = Nothing
    Set xmlForeignInsuranceNode = Nothing
    Set xmlCreditRatingNode = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

'EP694 - 07/06/2006 - Peter Edney
Function GetAge(dteDOB As Date) As Integer

Dim intYears As Integer
Dim dteNow As Date
    
    dteNow = Now
    intYears = Year(dteNow) - Year(dteDOB)
    
    If DateAdd("yyyy", intYears, dteDOB) > dteNow Then
        intYears = intYears - 1
    End If
    
    GetAge = intYears
    
End Function

Private Sub SetupInvestorDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupInvestorDetails
' Author:       Steve Badman
' Date:         17/05/2006
' Purpose:      Create Investor Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupInvestorDetailsExit
    
    Const strFunctionName As String = "SetupInvestorDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlInvestorDetailsNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    
    LogDetails 3, "SetupInvestorDetails : Start "
    
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlInvestorDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InvestorDetails", "")
    strRuleName = "SetupInvestorDetails"
    

   
    'Investor - LineOfCredit
    gstrRuleNumber = strRuleName & ".1"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Investor", "")
    'No combo for this - For initial Phase there will only be a single line of credit "Barclays"
    'xmlTempNode.Text = "Barclays"
    'SAB 24/05/2006 - EP607 - Marcus Bennett confirmed this hsould in fact be "Deutsche Bank" not "Barclays"
    xmlTempNode.Text = "Deutsche Bank"
    xmlInvestorDetailsNode.appendChild xmlTempNode

    
    'Committed amount - MORTGAGESUBQUOTE.TOTALLOANAMOUNT ??
    gstrRuleNumber = strRuleName & ".2"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommittedAmount", "")
    'xmlTempNode.Text = "??? To be identified ???"
    xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "TOTALLOANAMOUNT")
    xmlInvestorDetailsNode.appendChild xmlTempNode
    


    vxmlResponseNode.appendChild xmlInvestorDetailsNode
    
    LogDetails 3, "SetupInvestorDetails : End "
    
    
SetupInvestorDetailsExit:
    Set xmlDoc = Nothing
    Set xmlInvestorDetailsNode = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub



Private Sub SetupMortgageFundingDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupMortgageFundingDetails
' Author:       Steve Badman
' Date:         18/05/2006
' Purpose:      Create Mortgage Funding Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupMortgageFundingDetailsExit
    
    Const strFunctionName As String = "SetupMortgageFundingDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlMortgageFundingNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim strTotalLoamAmount As String
    Dim strRetention As String
    Dim strGrossAdvanceAmount As String
    Dim strNetAdvanceAmount As String
    
    LogDetails 3, "SetupMortgageFundingDetails : Start "
    
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlMortgageFundingNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageFunding", "")
    strRuleName = "SetupMortgageFundingDetails"
    

   
    'Gross advance amount
    gstrRuleNumber = strRuleName & ".1"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "GrossAdvanceAmount", "")
    strTotalLoamAmount = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "TOTALLOANAMOUNT")
    strRetention = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "RETENTIONAMOUNT")
    If Not strTotalLoamAmount = "" Then
        If Not strTotalLoamAmount = "" Then
            strGrossAdvanceAmount = CSafeDbl(strTotalLoamAmount) - CSafeDbl(strRetention)
            xmlTempNode.Text = strGrossAdvanceAmount
        End If
    End If
    xmlMortgageFundingNode.appendChild xmlTempNode

    
    'Net advance amount
    gstrRuleNumber = strRuleName & ".2"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NetAdvanceAmount", "")
    strNetAdvanceAmount = CSafeDbl(strGrossAdvanceAmount) - m_dblApplicationFeeAmount
    If Not strNetAdvanceAmount = "" Then
        xmlTempNode.Text = strNetAdvanceAmount
    End If
    xmlMortgageFundingNode.appendChild xmlTempNode
    
    


    vxmlResponseNode.appendChild xmlMortgageFundingNode
    
    LogDetails 3, "SetupMortgageFundingDetails : End "
    
    
SetupMortgageFundingDetailsExit:
    Set xmlDoc = Nothing
    Set xmlMortgageFundingNode = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

'AW 28/03/07    EP2_2125    -   Re-write to use valuation report
Private Sub SetupRetentionDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupRetentionDetails
' Author:       Steve Badman
' Date:         17/05/2006
' Purpose:      Create Retention Details data in completion block
'               Only to be displayed if there is a Retention amount
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupRetentionDetailsExit
    
    Const strFunctionName As String = "SetupRetentionDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlValnRepValuationNode As IXMLDOMNode
    Dim xmlRetentionDetailsNode As IXMLDOMNode
    Dim xmlValnInstructionNodeList As IXMLDOMNodeList
    Dim xmlTempNode As IXMLDOMNode
    Dim strRetention As String
    Dim intMaxSequenceNo As Integer
    Dim intSequenceNo As Integer
    Dim strMaxSequenceNo As String
    Dim lngRetentionAmount As Long
    
    LogDetails 3, "SetupRetentionDetails : Start "
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strRuleName = "SetupRetentionDetails"
    
    'Retention amount
    gstrRuleNumber = strRuleName & ".2"
    
    Set xmlValnInstructionNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION")
    
    intMaxSequenceNo = 0
    lngRetentionAmount = 0
        
    'Find the one with the highest INSTRUCTIONSEQUENCENO
    For Each xmlTempNode In xmlValnInstructionNodeList
         intSequenceNo = GetAttributeAsInteger(xmlTempNode, "INSTRUCTIONSEQUENCENO")
         If intSequenceNo > intMaxSequenceNo Then
             intMaxSequenceNo = intSequenceNo
         End If
    Next

    strMaxSequenceNo = CStr(intMaxSequenceNo)
    
    Set xmlValnRepValuationNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION/VALNREPVALUATION[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")
    
    If Not xmlValnRepValuationNode Is Nothing Then
        lngRetentionAmount = CSafeLng(GetAttributeText(xmlValnRepValuationNode, "RETENTIONWORKS"))
    End If
    
    If lngRetentionAmount > 0 Then
        Set xmlRetentionDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RetentionDetails", "")
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RetentionAmount", "")
        xmlTempNode.Text = CStr(lngRetentionAmount)
        xmlRetentionDetailsNode.appendChild xmlTempNode
        vxmlResponseNode.appendChild xmlRetentionDetailsNode
    End If
 
    LogDetails 3, "SetupRetentionDetails : End "
    
    
SetupRetentionDetailsExit:
    Set xmlDoc = Nothing
    Set xmlRetentionDetailsNode = Nothing
    Set xmlValnRepValuationNode = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub
'AW 28/03/07    EP2_2125    -   End

Private Sub SetupLoanDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupLoanDetails
' Author:       Steve Badman
' Date:         12/05/2006
' Purpose:      Create Loan Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupLoanDetailsExit
    
    Const strFunctionName As String = "SetupLoanDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlLoanDetailsNode As IXMLDOMNode
    Dim xmlInvestorDetailsNode As IXMLDOMNode
    
    Dim xmlComponentOriginationNode As IXMLDOMNode
    Dim xmlApplicationOfferNode As IXMLDOMNode
    Dim xmlApplicationOfferNodeList As IXMLDOMNodeList
    Dim intHighestAppOfferSeqNo As Integer
    Dim strHighestAppOfferSeqNo As String
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlDisbursementPaymentNode As IXMLDOMNode
    Dim strCommitmentIssueDate As String
    Dim strCommitmentExpiryDate As String
    Dim strCommitmentAcceptedDate As String
    
    Dim xmlComponentRateAndTermNode As IXMLDOMNode
    Dim strPackagerName As String
    Dim xmlLoanComponentNode As IXMLDOMNode
    Dim xmlMortgageProductNode As IXMLDOMNode
    Dim xmlMortgageProductLanguageNode As IXMLDOMNode
    Dim strLTV As String
    Dim strLoanComponentTermInYears As String
    Dim strLoanComponentTermInMonths As String
    Dim strMortgageProductDetail As String
    Dim xmlInterestRateList As IXMLDOMNodeList
    Dim xmlInterestRateNode As IXMLDOMNode
    Dim strInterestRateTypeSequenceNumber As String
    Dim strInterestRatePeriod As String
    Dim strBaseInterestRate As String
    Dim strBaseRateDescription As String
    Dim xmlBaseRateBandList As IXMLDOMNodeList
    Dim xmlBaseRateBandNode As IXMLDOMNode
    Dim intBaseRateBandLTV As Integer
    Dim strRateDifference As String
    Dim strResolvedRate As String
    Dim strRateType As String
    Dim strInterestRateCode As String
    Dim strRateChangeFrequency As String
    Dim strRateOrRateIncrement As String
    Dim strCompletionDate As String
    Dim strDateToSetInterestRate As String
    Dim intInterestRatePeriod As Integer
    Dim strLoanTerm As String
    Dim intTotalLoanCompTermInMonths As Integer
    Dim intTotalTermMonthsAfterDiscountPeriod As Integer
    Dim intTermYearsAfterDiscountPeriod As Integer
    Dim intTermMonthsAfterDiscountPeriod As Integer
    Dim intGroupSequenceNumber As Integer
    
    Dim xmlComponentPaymentOptionsNode As IXMLDOMNode
    Dim strRepaymentMethod As String
    Dim strPrimaryPaymentAmount As String
    
    Dim xmlComponentMiscellaneousNode As IXMLDOMNode
    Dim strProductScheme As String
    Dim strLTVBanding As String
    Dim strRegulationStatus As String
    Dim strMortgageProductCode As String

    
    Dim xmlOrgNode As IXMLDOMNode
    
'SW 21/06/2006 EP771
    Dim strRepaymentVehicle As String
  'EP2_956
    Dim xmlLoanComponentNodeList As IXMLDOMNodeList
    
    LogDetails 3, "SetupLoanDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strRuleName = "SetupLoanDetails"
  'EP2_956 - Need to loop for each loan component and add on Investor Details
  
    Set xmlLoanComponentNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT")
    For Each xmlLoanComponentNode In xmlLoanComponentNodeList
       
        Set xmlLoanDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LoanDetails", "")
        gstrRuleNumber = strRuleName
        
         
         
        '** Get Component Origination details : START **
        Set xmlComponentOriginationNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentOrigination", "")
        
        
        
        
        
        'Date of loan application
        gstrRuleNumber = strRuleName & ".1"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ApplicationDate", "")
        xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONDATE")
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'Date Commitment issued
        gstrRuleNumber = strRuleName & ".2"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommitmentIssuedDate", "")
        'Use APPLICATIONOFFER.OFFERISSUEDATE
        'get the APPLICATIONOFFER data
        Set xmlApplicationOfferNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/APPLICATIONOFFER")
        'EP2_956 need to get the one with the highest sequence number
        intHighestAppOfferSeqNo = 0
        'get the APPLICATIONOFFER data for the latest Offer
        For Each xmlApplicationOfferNode In xmlApplicationOfferNodeList
            If (GetAttributeAsInteger(xmlApplicationOfferNode, "OFFERSEQUENCENO") > intHighestAppOfferSeqNo) Then
                intHighestAppOfferSeqNo = GetAttributeAsInteger(xmlApplicationOfferNode, "OFFERSEQUENCENO")
            End If
        Next
        strHighestAppOfferSeqNo = intHighestAppOfferSeqNo
        Set xmlApplicationOfferNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONOFFER[@OFFERSEQUENCENO='" & strHighestAppOfferSeqNo & "']")
        strCommitmentIssueDate = GetAttributeText(xmlApplicationOfferNode, "OFFERISSUEDATE")
        If Not strCommitmentIssueDate = "" Then
            xmlTempNode.Text = strCommitmentIssueDate
        End If
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'Commitment expiry date
        gstrRuleNumber = strRuleName & ".3"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommitmentExpiryDate", "")
        'Use APPLICATIONOFFER.OFFEREXPIRYDATE or if NULL use COMPLETIONDATE
        
        If xmlApplicationOfferNode Is Nothing Then
            strCommitmentExpiryDate = ""
        Else
            strCommitmentExpiryDate = Format(GetAttributeText(xmlApplicationOfferNode, "OFFEREXPIRYDATE"), "DD/MM/YYYY")
        End If
        If Not strCommitmentExpiryDate = "" Then
            xmlTempNode.Text = strCommitmentExpiryDate
        Else
            xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "COMPLETIONDATE")
        End If
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'Commitment accepted date
        gstrRuleNumber = strRuleName & ".4"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommitmentAcceptedDate", "")
        'Use APPLICATIONUNDERWRITING.UNDERWRITERDECISIONDATETIME or if NULL use COMPLETIONDATE
        strCommitmentAcceptedDate = Format(GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONUNDERWRITING"), "UNDERWRITERDECISIONDATETIME"), "DD/MM/YYYY")
        If Not strCommitmentAcceptedDate = "" Then
            xmlTempNode.Text = strCommitmentAcceptedDate
        Else
            xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "COMPLETIONDATE")
        End If
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'Committed amount - MORTGAGESUBQUOTE.TOTALLOANAMOUNT ??
        gstrRuleNumber = strRuleName & ".5"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommittedAmount", "")
        'xmlTempNode.Text = "??? To be identified ???"
        xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "TOTALLOANAMOUNT")
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        'LoanComponentAmount
        gstrRuleNumber = strRuleName & ".5a"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LoanComponentAmount", "")
        xmlTempNode.Text = GetAttributeText(xmlLoanComponentNode, "TOTALLOANCOMPONENTAMOUNT")
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        'OriginalLTV
        gstrRuleNumber = strRuleName & ".5b"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "OriginalLTV", "")
        xmlTempNode.Text = GetAttributeText(xmlLoanComponentNode, "ORIGINALLTV")
        If xmlTempNode.Text = "" Then
          xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "LTV")
        End If
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        'Initial Advance date
        gstrRuleNumber = strRuleName & ".6"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InitialAdvanceDate", "")
        'EP2_956
        '
        ' Use DisbursementPayment.IssueDate where the PaymentType = 'I' (initial advance)
                                                                                                                                                  
        Set xmlDisbursementPaymentNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/DISBURSEMENTPAYMENT[@PAYMENTTYPE_TYPE_I]")
        If Not xmlDisbursementPaymentNode Is Nothing Then
          xmlTempNode.Text = GetAttributeText(xmlDisbursementPaymentNode, "ISSUEDATE")
        Else
          xmlTempNode.Text = ""
        End If
        ' xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "COMPLETIONDATE")
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'Originator Branch
        gstrRuleNumber = strRuleName & ".7"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "OriginatorBranch", "")
        'Set to Postcode mapping for Postcode of new propertry
        xmlTempNode.Text = m_strAddressPostCode
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'Name of Packager - APPLICATION/PACKAGERINTERMEDIARYORGANISATION.NAME
        gstrRuleNumber = strRuleName & ".8"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PackagerName", "")
        
        'EP694 - 07/06/2006 - Peter Edney
        'EP2_956 Use the PacKager Principal firm name
        Set xmlOrgNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONINTRODUCER/PRINCIPALFIRMPACKAGER")
        If xmlOrgNode Is Nothing Then
            strPackagerName = ""
        Else
            strPackagerName = GetAttributeText(xmlOrgNode, "PRINCIPALFIRMNAME")
        End If
        If Not strPackagerName = "" Then
            xmlTempNode.Text = strPackagerName
        End If
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        
        'RepaymentVehicle
        gstrRuleNumber = strRuleName & ".9"
            
    'SW 21/06/2006 EP771 Start
     '   Set xmlLoanComponentNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT")
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RepaymentVehicle", "")
        
        strRepaymentVehicle = GetAttributeText(xmlLoanComponentNode, "REPAYMENTVEHICLE")
        
        If Not strRepaymentVehicle = "" Then
            xmlTempNode.Text = GetComboText("RepaymentVehicle", CSafeInt(strRepaymentVehicle))
        Else
            'EP694 - 07/06/2006 - Peter Edney
            xmlTempNode.Text = "Unknown"
        End If
            
        xmlComponentOriginationNode.appendChild xmlTempNode
    'SW 21/06/2006 EP771 End
        
        
        'External Account Number
        gstrRuleNumber = strRuleName & ".10"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ExternalAccountNumber", "")
        xmlTempNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONNUMBER")
        xmlComponentOriginationNode.appendChild xmlTempNode
        
        xmlLoanDetailsNode.appendChild xmlComponentOriginationNode
        '** Get Component Origination details : START **
            
            
        
        
        
        
        
        '** Get Component Rate and Term : START **
        'Should be only one LOANCOMPONENT, but multiple INTERESTRATETYPE
        
    'SW 21/06/2006 EP771    Set xmlLoanComponentNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT")
        Set xmlMortgageProductNode = xmlLoanComponentNode.selectSingleNode("MORTGAGEPRODUCT")
        Set xmlMortgageProductLanguageNode = xmlMortgageProductNode.selectSingleNode("MORTGAGEPRODUCTLANGUAGE")
        
        
        'Get the LTV
        strLTV = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "LTV")
        
        'Get the LoanComponent Data
        strLoanComponentTermInYears = GetAttributeText(xmlLoanComponentNode, "TERMINYEARS")
        strLoanComponentTermInMonths = GetAttributeText(xmlLoanComponentNode, "TERMINMONTHS")
        
        
        'Get the MortgageProductLanguage Data
        strMortgageProductDetail = GetAttributeText(xmlMortgageProductLanguageNode, "PRODUCTTEXTDETAILS")
        
        
        
        
        Set xmlInterestRateList = xmlLoanComponentNode.selectNodes("MORTGAGEPRODUCT/INTERESTRATETYPE")
        For Each xmlInterestRateNode In xmlInterestRateList
            Set xmlComponentRateAndTermNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentRateAndTerm", "")
            
            strInterestRateTypeSequenceNumber = GetAttributeText(xmlInterestRateNode, "INTERESTRATETYPESEQUENCENUMBER")
                    
            'Get the Interest Rate Period
            strInterestRatePeriod = GetAttributeText(xmlInterestRateNode, "INTERESTRATEPERIOD")
                    
            'Get the Base Interest Rate
            strBaseInterestRate = GetAttributeText(xmlInterestRateNode.selectSingleNode("BASERATESET/BASERATE"), "BASEINTERESTRATE")
            
            'Get the Base Rate Description
            strBaseRateDescription = GetAttributeText(xmlInterestRateNode.selectSingleNode("BASERATESET/BASERATE"), "RATEDESCRIPTION")
            
            'get the BaseRateBand ratedifference
            'List ordered by MAXIMUMLTV ascedning, so once a match is found then exit the loop
            Set xmlBaseRateBandList = xmlInterestRateNode.selectNodes("BASERATESET/BASERATEBAND")
            For Each xmlBaseRateBandNode In xmlBaseRateBandList
                intBaseRateBandLTV = CSafeInt(GetAttributeText(xmlBaseRateBandNode, "MAXIMUMLTV"))
                If intBaseRateBandLTV > CSafeInt(strLTV) Then
                    strRateDifference = GetAttributeText(xmlBaseRateBandNode, "RATEDIFFERENCE")
                    strResolvedRate = CSafeDbl(strRateDifference) + CSafeDbl(strBaseInterestRate)
                    Exit For
                End If
            Next
            
            
            
            
            strRateType = GetAttributeText(xmlInterestRateNode, "RATETYPE")
                    
            Select Case strRateType
                Case "B":
                    strInterestRateCode = strBaseRateDescription
                    strRateChangeFrequency = "99a - Floating rate"
                    strRateOrRateIncrement = strRateDifference
                Case "C":
                    strInterestRateCode = strBaseRateDescription
                    strRateChangeFrequency = "99a - Floating rate"
                    strRateOrRateIncrement = strRateDifference
                Case "D":
                    strInterestRateCode = strBaseRateDescription
                    strRateChangeFrequency = "99a - Floating rate"
                    strRateOrRateIncrement = strRateDifference
                Case "F":
                    strInterestRateCode = "Fixed"
                    strRateChangeFrequency = "00 - Fixed Rate"
                    strRateOrRateIncrement = strResolvedRate
                Case Else:
            End Select
            
            
            
            'Interest Rate Code
            gstrRuleNumber = strRuleName & strInterestRateTypeSequenceNumber & ".11"
            Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InterestRateCode", "")
            xmlTempNode.Text = strInterestRateCode
            xmlComponentRateAndTermNode.appendChild xmlTempNode
            
            
            'Rate or RateIncrement
            gstrRuleNumber = strRuleName & strInterestRateTypeSequenceNumber & ".12"
            Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RateOrRateIncrement", "")
            xmlTempNode.Text = strRateOrRateIncrement
            xmlComponentRateAndTermNode.appendChild xmlTempNode
                    
            
            'Date to set interest rate - Completion Date
            gstrRuleNumber = strRuleName & strInterestRateTypeSequenceNumber & ".13"
            Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DateToSetInterestRate", "")
            strCompletionDate = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "COMPLETIONDATE")
            
            If Not strInterestRatePeriod = "-1" Then
                'completion date
                strDateToSetInterestRate = strCompletionDate
                intInterestRatePeriod = CSafeInt(strInterestRatePeriod)
            Else
                'completion date + strInterestRatePeriod
                strDateToSetInterestRate = CStr(DateAdd("m", intInterestRatePeriod, CSafeDate(strCompletionDate)))
            End If
            xmlTempNode.Text = strDateToSetInterestRate
            xmlComponentRateAndTermNode.appendChild xmlTempNode
            
            
            'Rate Change Frequency
            gstrRuleNumber = strRuleName & strInterestRateTypeSequenceNumber & ".14"
            Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RateChangeFrequency", "")
            xmlTempNode.Text = strRateChangeFrequency
            xmlComponentRateAndTermNode.appendChild xmlTempNode
            
            
            'Term of Loan
            gstrRuleNumber = strRuleName & strInterestRateTypeSequenceNumber & ".15"
            Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TermOfLoan", "")
            If Not strInterestRatePeriod = "-1" Then
                'SAB 23/05/2006 - EP607 - Amended to display the full term rather than just the discount period.
                strLoanTerm = strLoanComponentTermInYears & " Year(s) and " & strLoanComponentTermInMonths & " Month(s)"
                'strLoanTerm = strInterestRatePeriod & " Month(s)"
                
                'calculate Term for the rest of the loan
                intTotalLoanCompTermInMonths = (CSafeInt(strLoanComponentTermInYears) * 12) + CSafeInt(strLoanComponentTermInMonths)
                intTotalTermMonthsAfterDiscountPeriod = intTotalLoanCompTermInMonths - CSafeInt(strInterestRatePeriod)
                intTermYearsAfterDiscountPeriod = Round(intTotalTermMonthsAfterDiscountPeriod / 12, 0)
                intTermMonthsAfterDiscountPeriod = intTotalTermMonthsAfterDiscountPeriod - (intTermYearsAfterDiscountPeriod * 12)
            Else
                strLoanTerm = intTermYearsAfterDiscountPeriod & " Year(s) and " & intTermMonthsAfterDiscountPeriod & " Month(s)"
            End If
            xmlTempNode.Text = strLoanTerm
            xmlComponentRateAndTermNode.appendChild xmlTempNode
            
            
            'Product Code
            gstrRuleNumber = strRuleName & strInterestRateTypeSequenceNumber & ".16"
            Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProductCode", "")
            xmlTempNode.Text = strMortgageProductDetail
            xmlComponentRateAndTermNode.appendChild xmlTempNode
            
            
            xmlLoanDetailsNode.appendChild xmlComponentRateAndTermNode
            
            Set xmlComponentRateAndTermNode = Nothing
        Next
        '** Get Component Rate and Term : END **
        
    
    
    
    
        '** Get Component Payment Options : START **
        'SAB 23/05/2006 - EP607 - amended to use different Node name
        'Set xmlComponentPaymentOptionsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentOrigination", "")
        Set xmlComponentPaymentOptionsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentPaymentOptions", "")
        
        'Repayment Method
        gstrRuleNumber = strRuleName & ".18"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RepaymentMethod", "")
        
        strRepaymentMethod = GetAttributeText(xmlLoanComponentNode, "REPAYMENTMETHOD")
        If Not strRepaymentMethod = "" Then
            xmlTempNode.Text = GetComboText("RepaymentType", CSafeInt(strRepaymentMethod))
        End If
        xmlComponentPaymentOptionsNode.appendChild xmlTempNode
    
    
        'Primary Payment Amount
        gstrRuleNumber = strRuleName & ".19"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PrimaryPaymentAmount", "")
        strPrimaryPaymentAmount = GetAttributeText(xmlLoanComponentNode, "NETMONTHLYCOST")
        If Not strPrimaryPaymentAmount = "" Then
            xmlTempNode.Text = Format(strPrimaryPaymentAmount, ".00")
        End If
        xmlComponentPaymentOptionsNode.appendChild xmlTempNode
    
    
        xmlLoanDetailsNode.appendChild xmlComponentPaymentOptionsNode
        '** Get Component Payment Options : END **
    
    
        '** Get Component Miscellaneous Fields : START **
        Set xmlComponentMiscellaneousNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentMiscellaneous", "")
        
        'ProductScheme
        gstrRuleNumber = strRuleName & ".20"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProductScheme", "")
        
        'SAB 23/05/2006 - EP607 - Amended as this is held on ApplicationFactFind.
        'strProductScheme = GetAttributeText(xmlLoanComponentNode, "PRODUCTSCHEME")
        'strProductScheme = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "PRODUCTSCHEME")
        
        'LH 29/08/3006 - EP1090 - Amended this to retrieve product scheme correctly, product scheme was not populating in APPLICATIONFACTFIND table in all instances
        ' Retrieve the product scheme
        'Multi Loan Component
        'Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT")
        Set xmlNode = xmlLoanComponentNode
        If Not xmlNode Is Nothing Then
            Set xmlNode = xmlNode.selectSingleNode("SPECIALGROUP")
            If Not xmlNode Is Nothing Then
                'Get the ValidationType from combo "SpecialGroup" where combo ValueId = SpecialGroup.GroupTypeSequenceNumber
                If Not xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER") Is Nothing Then
                    intGroupSequenceNumber = xmlNode.Attributes.getNamedItem("GROUPTYPESEQUENCENUMBER").nodeValue
                    Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/COMBOGROUP[@GROUPNAME='SpecialGroup']/COMBOVALUE[@VALUEID='" & intGroupSequenceNumber & "']")
                    If Not xmlNode Is Nothing Then
                        If Not xmlNode.Attributes.getNamedItem("VALUENAME") Is Nothing Then
                            strProductScheme = xmlNode.Attributes.getNamedItem("VALUENAME").nodeValue
                        End If
                    End If
                End If
            End If
        End If
        
        If Not strProductScheme = "" Then
            xmlTempNode.Text = strProductScheme
        End If
        xmlComponentMiscellaneousNode.appendChild xmlTempNode
    
    
        'LTV Banding
        gstrRuleNumber = strRuleName & ".21"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LTVBanding", "")
        '** Could not identify an LTV Banding combo **
        'SAB 23/05/2006 - EP607 - Amended to use Hard-code Optimus values for LTV Banding
        strLTVBanding = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "LTV")
        If Not strLTVBanding = "" Then
            'xmlTempNode.Text = strLTVBanding
            If CSafeDbl(strLTVBanding) <= 75 Then
                xmlTempNode.Text = "Up to 75%"
            ElseIf CSafeDbl(strLTVBanding) <= 80 Then
                xmlTempNode.Text = "75.01% to 80%"
            ElseIf CSafeDbl(strLTVBanding) <= 85 Then
                xmlTempNode.Text = "80.01% to 85%"
            ElseIf CSafeDbl(strLTVBanding) <= 90 Then
                xmlTempNode.Text = "85.01% to 90%"
            ElseIf CSafeDbl(strLTVBanding) <= 95 Then
                xmlTempNode.Text = "90.01% to 95%"
            End If
        End If
        xmlComponentMiscellaneousNode.appendChild xmlTempNode
    
    
        'Regulation Status
        gstrRuleNumber = strRuleName & ".22"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RegulationStatus", "")
        strRegulationStatus = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "REGULATIONINDICATOR")
        If Not strRegulationStatus = "" Then
            xmlTempNode.Text = GetComboText("RegulationIndicator", CSafeInt(strRegulationStatus))
        End If
        xmlComponentMiscellaneousNode.appendChild xmlTempNode
    
    
        'Omiga Product Code
        gstrRuleNumber = strRuleName & ".23"
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "OmigaProductCode", "")
        strMortgageProductCode = GetAttributeText(xmlMortgageProductNode, "MORTGAGEPRODUCTCODE")
        If Not strMortgageProductCode = "" Then
            xmlTempNode.Text = strMortgageProductCode
        End If
        xmlComponentMiscellaneousNode.appendChild xmlTempNode
    
        xmlLoanDetailsNode.appendChild xmlComponentMiscellaneousNode
        '** Get Component Miscellaneous Fields : START **

' now do the Investor Details
        Set xmlInvestorDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InvestorDetails", "")
        strRuleName = "SetupInvestorDetails"
    
     'This bit cloned from SetupInvestorDetails
   
    'Investor - LineOfCredit
    gstrRuleNumber = strRuleName & ".1"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Investor", "")
    'No combo for this - For initial Phase there will only be a single line of credit "Barclays"
    'xmlTempNode.Text = "Barclays"
    'SAB 24/05/2006 - EP607 - Marcus Bennett confirmed this hsould in fact be "Deutsche Bank" not "Barclays"
    xmlTempNode.Text = "Deutsche Bank"
    xmlInvestorDetailsNode.appendChild xmlTempNode

    
    'Committed amount - MORTGAGESUBQUOTE.TOTALLOANAMOUNT ??
    gstrRuleNumber = strRuleName & ".2"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CommittedAmount", "")
    
    xmlTempNode.Text = GetAttributeText(xmlLoanComponentNode, "TOTALLOANCOMPONENTAMOUNT")
    xmlInvestorDetailsNode.appendChild xmlTempNode
    
    xmlLoanDetailsNode.appendChild xmlInvestorDetailsNode

    vxmlResponseNode.appendChild xmlLoanDetailsNode
    
    Next 'Loan Component
    
    LogDetails 3, "SetupLoanDetails : End "
    
    
SetupLoanDetailsExit:
    Set xmlDoc = Nothing
    Set xmlLoanDetailsNode = Nothing
    Set xmlComponentOriginationNode = Nothing
    Set xmlApplicationOfferNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlComponentRateAndTermNode = Nothing
    Set xmlLoanComponentNode = Nothing
    Set xmlMortgageProductNode = Nothing
    Set xmlMortgageProductLanguageNode = Nothing
    Set xmlInterestRateList = Nothing
    Set xmlInterestRateNode = Nothing
    Set xmlBaseRateBandList = Nothing
    Set xmlBaseRateBandNode = Nothing
    Set xmlComponentPaymentOptionsNode = Nothing
    Set xmlComponentMiscellaneousNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetupChargeDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupChargeDetails
' Author:       Steve Badman
' Date:         12/05/2006
' Purpose:      Create Charge Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupChargeDetailsExit
    
    Const strFunctionName As String = "SetupChargeDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlChargeDetailsNode As IXMLDOMNode

    Dim xmlProductInformationNode As IXMLDOMNode
    Dim xmlLineOfBusinessNode As IXMLDOMNode
    Dim strLineOfBusiness As String
    Dim xmlApplicationTypeNode As IXMLDOMNode
    Dim strApplicationType As String
    Dim xmlMortgagePurposeTypeNode As IXMLDOMNode
    Dim strPurposeOfLoan As String

    Dim xmlPaymentTermsNode As IXMLDOMNode
    Dim xmlFirstPaymentDateNode As IXMLDOMNode
    Dim strFirstPaymentDate As String
    Dim dtCompletionDate As Date

    Dim xmlMiscellaneousChargeDataNode As IXMLDOMNode
    Dim xmlCCJDataNode As IXMLDOMNode
    Dim xmlApplicationCreditCheckNodeList As IXMLDOMNodeList
    Dim xmlApplicationCreditCheckNode As IXMLDOMNode
    Dim xmlBespokeBureauData As IXMLDOMNode
    Dim strSequenceNumber As String
    Dim strCreditCheckGUID As String
    
    Dim xmlTempNode As IXMLDOMNode
    Dim strTemp As String

    Dim strSelfCertInd As String
    Dim strTypeOfBuyer As String
    Dim strUnsatisfiedCCJ As String
    Dim strSatisfiedCCJ As String
    Dim intUnsatisfiedCCJ As Integer
    Dim intSatisfiedCCJ As Integer
    Dim xmlBureau2List As IXMLDOMNodeList
    Dim xmlBureau2Node As IXMLDOMNode
    Dim strBX18_CCJAmount As String
    Dim xmlSubPurposeList As IXMLDOMNodeList
    Dim xmlSubPurpose As IXMLDOMNode
    Dim intSubPurpose As Integer
    Dim strSubPurpose As String
    
    LogDetails 3, "SetupChargeDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlChargeDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ChargeDetails", "")
    strRuleName = "SetupChargeDetails"
    
    
    
    '** Get Product Information : START **
    Set xmlProductInformationNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProductInformation", "")
    
    
    'LineOfBusiness
    gstrRuleNumber = strRuleName & ".1"
    Set xmlLineOfBusinessNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LineOfBusiness", "")
    
    strLineOfBusiness = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "NATUREOFLOAN")
    If Not strLineOfBusiness = "" Then
        xmlLineOfBusinessNode.Text = GetComboText("NatureOfLoan", CSafeInt(strLineOfBusiness))
        'xmlLineOfBusinessNode.Text = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "NATUREOFLOAN_TEXT")
    End If
    xmlProductInformationNode.appendChild xmlLineOfBusinessNode
    

    'MortgageType - APPLICATIONFACTFIND.TYPEOFAPPLICATION
    gstrRuleNumber = strRuleName & ".2"
    Set xmlApplicationTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageType", "")
    
    'EP694 - 07/06/2006 - Peter Edney
    'Get the Self Certified Indicator
    'EP2_956 - in Epsom2 this is a combo lookup
    Set xmlTempNode = Nothing
    strSelfCertInd = ""
    Set xmlTempNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND")
    If Not xmlTempNode Is Nothing Then
        strSelfCertInd = GetAttributeText(xmlTempNode, "APPLICATIONINCOMESTATUS")
    End If
    
    'Get the type of application
    Set xmlTempNode = Nothing
    strApplicationType = ""
    Set xmlTempNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND")
    If Not xmlTempNode Is Nothing Then
        strApplicationType = GetMandatoryAttributeText(xmlTempNode, "TYPEOFAPPLICATION")
    End If
    
    If IsValidationType("TypeOfMortgage", strApplicationType, "R") Then
        
        If IsValidationType("ApplicationIncomeStatus", strSelfCertInd, "SC") Then
            xmlApplicationTypeNode.Text = "Self Cert Remortgage"
        Else
            xmlApplicationTypeNode.Text = "Full Status Remortgage"
        End If
    Else
        If IsValidationType("ApplicationIncomeStatus", strSelfCertInd, "SC") Then
            xmlApplicationTypeNode.Text = "Self Cert Purchase"
        Else
            xmlApplicationTypeNode.Text = "Full Status Purchase"
        End If
    End If
    
    xmlProductInformationNode.appendChild xmlApplicationTypeNode
    
    'MortgagePurpose
    'This is held on the LoanComponent which for Phase 1 is ok as there will only be 1 LoanComponent, but will need to be amended for the future.
    gstrRuleNumber = strRuleName & ".3"
    Set xmlMortgagePurposeTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgagePurposeType", "")
    xmlMortgagePurposeTypeNode.Text = ""
    
    'EP694 - 07/06/2006 - Peter Edney
    'Get Type of Buyer
    Set xmlTempNode = Nothing
    strTypeOfBuyer = ""
    Set xmlTempNode = vxmlRequestNode.selectSingleNode("APPLICATION")
    If Not xmlTempNode Is Nothing Then
        strTypeOfBuyer = GetAttributeText(xmlTempNode, "TYPEOFBUYER")
    End If
    
    'EP729 - 12/06/2006 - Peter Edney
    If xmlMortgagePurposeTypeNode.Text = "" Then
        If IsValidationType("TypeOfMortgage", strApplicationType, "TOE") Then
            xmlMortgagePurposeTypeNode.Text = "Remortgage - T of E"
        End If
    End If
        
    'EP729 - 12/06/2006 - Peter Edney
    If xmlMortgagePurposeTypeNode.Text = "" Then
        If IsValidationType("TypeOfMortgage", strApplicationType, "R") Then
        
            If xmlMortgagePurposeTypeNode.Text = "" Then
                If IsValidationType("NatureOfLoan", strLineOfBusiness, "BI") Or IsValidationType("NatureOfLoan", strLineOfBusiness, "BR") Then
                    xmlMortgagePurposeTypeNode.Text = "Buy to let"
                End If
            End If

            If xmlMortgagePurposeTypeNode.Text = "" Then
                intSubPurpose = 0
                Set xmlSubPurposeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@SUBPURPOSEOFLOAN")
                For Each xmlSubPurpose In xmlSubPurposeList
                    If IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "HI") Or _
                        IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "DC") Or _
                        IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "CR") Then
                        
                        If Not vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[@SUBPURPOSEOFLOAN!='" & xmlSubPurpose.Text & "']") Is Nothing Then
                            intSubPurpose = intSubPurpose + 1
                        End If
                        
                    End If
                Next
                If intSubPurpose > 1 Then
                    xmlMortgagePurposeTypeNode.Text = "Capital Raise - Mix (Home Improve/ debt consol)"
                End If
            End If
            
            If xmlMortgagePurposeTypeNode.Text = "" Then
                Set xmlSubPurpose = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[1]/@SUBPURPOSEOFLOAN")
                If Not xmlSubPurpose Is Nothing Then
                    Select Case True
                    Case IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "HI")
                        xmlMortgagePurposeTypeNode.Text = "Capital Raise - Home Improvement"
                        
                    Case IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "DC")
                        xmlMortgagePurposeTypeNode.Text = "Capital Raise - Debt Consolidation"
                        
                    Case IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "CR"), _
                        IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "HP"), _
                        IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "PP"), _
                        IsValidationType("SubPurposeOfLoan", xmlSubPurpose.Text, "O")
                        xmlMortgagePurposeTypeNode.Text = "Capital Raise - Other"
                        
                    End Select
                End If
            End If
            
            If xmlMortgagePurposeTypeNode.Text = "" Then
                Set xmlTempNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[1]/@PORTEDLOAN")
                If Not xmlTempNode Is Nothing Then
                    If xmlTempNode.Text = "1" Then
                        xmlMortgagePurposeTypeNode.Text = "Remortgage - Ported"
                    End If
                End If
            End If
             'EP2_956
            If xmlMortgagePurposeTypeNode.Text = "" Then
                If IsValidationType("TypeOfMortgage", strApplicationType, "RTB") Then
                    xmlMortgagePurposeTypeNode.Text = "Right to Buy"
                End If
            End If

            If xmlMortgagePurposeTypeNode.Text = "" Then
                xmlMortgagePurposeTypeNode.Text = "Remortgage - Standard"
            End If

         
            
        End If
    End If
        
    If xmlMortgagePurposeTypeNode.Text = "" Then
        If IsValidationType("TypeOfMortgage", strApplicationType, "HM") Then
            xmlMortgagePurposeTypeNode.Text = "Mover"
        End If
    End If
    
    If xmlMortgagePurposeTypeNode.Text = "" Then
        If IsValidationType("NatureOfLoan", strLineOfBusiness, "BI") Or IsValidationType("NatureOfLoan", strLineOfBusiness, "BR") Then
            If IsValidationType("TypeOfBuyerNewLoan", strTypeOfBuyer, "F") Then
                xmlMortgagePurposeTypeNode.Text = "BTL – 1st Time buyer"
            Else
                xmlMortgagePurposeTypeNode.Text = "Buy to let"
            End If
        End If
    End If
    
    If xmlMortgagePurposeTypeNode.Text = "" Then
        If IsValidationType("TypeOfBuyerNewLoan", strTypeOfBuyer, "F") Then
            xmlMortgagePurposeTypeNode.Text = "First Time buyer"
        End If
    End If
    
    If xmlMortgagePurposeTypeNode.Text = "" Then
        If IsValidationType("TypeOfMortgage", strApplicationType, "RTB") Then
                xmlMortgagePurposeTypeNode.Text = "Right to Buy"
        End If
    End If

    If xmlMortgagePurposeTypeNode.Text = "" Then
        xmlMortgagePurposeTypeNode.Text = "Other"
    End If
    
    xmlProductInformationNode.appendChild xmlMortgagePurposeTypeNode
    


    xmlChargeDetailsNode.appendChild xmlProductInformationNode
    '** Get Product Information : START **





    
    '** Get Payment Terms : START **
    Set xmlPaymentTermsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PaymentTerms", "")


    'First Payment Date
    gstrRuleNumber = strRuleName & ".5"
    Set xmlFirstPaymentDateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FirstPaymentDate", "")
    
    dtCompletionDate = CSafeDate(GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "COMPLETIONDATE"))
    'Add 1 month to the completion date
    strFirstPaymentDate = CStr(DateAdd("m", 1, dtCompletionDate))
    xmlFirstPaymentDateNode.Text = strFirstPaymentDate
    
    xmlPaymentTermsNode.appendChild xmlFirstPaymentDateNode


    xmlChargeDetailsNode.appendChild xmlPaymentTermsNode
    '** Get Payment Terms : END **







    '** Get Miscellaneous Charge Fields - CCJ : START **
    Set xmlMiscellaneousChargeDataNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MiscellaneousChargeData", "")


    Set xmlCCJDataNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CCJDATA", "")


    'get the correct APPLICATIONCREDITCHECK data
    Set xmlApplicationCreditCheckNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK")

    For Each xmlApplicationCreditCheckNode In xmlApplicationCreditCheckNodeList
        'Ordered descending, so the 1st record will be the correct one
        strSequenceNumber = GetAttributeText(xmlApplicationCreditCheckNode, "SEQUENCENUMBER")
        strCreditCheckGUID = GetAttributeText(xmlApplicationCreditCheckNode, "CREDITCHECKGUID")
        Exit For
    Next

    'get the correct BESPOKEBUREAUDATA
    Set xmlBespokeBureauData = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/CREDITCHECK/BESPOKEBUREAUDATA[@CREDITCHECKGUID='" & strCreditCheckGUID & "']")
 
    'Last CCJ before Date of Completion in years - It seems to me this can be taken from the MOSTRECENTREGDATEUNSATCCJ
    gstrRuleNumber = strRuleName & ".6"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LastCCJBeforeDataOfCompletionYears", "")
    If Not xmlBespokeBureauData Is Nothing Then
        strTemp = GetAttributeText(xmlBespokeBureauData, "MOSTRECENTREGDATEUNSATCCJ")
        If Not strTemp = "" Then
            xmlTempNode.Text = strTemp
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    
    
    'Last Satisfied CCJ before Date of Completion in years - It seems to me this can be taken from the MOSTRECENTREGDATESATCCJ
    gstrRuleNumber = strRuleName & ".7"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LastSatisfiedCCJBeforeDateOfCompletionYears", "")
    If Not xmlBespokeBureauData Is Nothing Then
        strTemp = GetAttributeText(xmlBespokeBureauData, "MOSTRECENTREGDATESATCCJ")
        If Not strTemp = "" Then
            xmlTempNode.Text = strTemp
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    
   
    'AdvisedSale
    gstrRuleNumber = strRuleName & ".8"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AdvisedSale", "")
    
    strTemp = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "LEVELOFADVICE")
    If Not strTemp = "" Then
        'check if the Level Of Advice is "Advised"
        If IsValidationType("LevelOfAdvice", CSafeInt(strTemp), "ADV") Then
            xmlTempNode.Text = "Y"
        Else
            xmlTempNode.Text = "N"
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
   
    
    'TotalNumberOfCCJsLessThan3Years
    gstrRuleNumber = strRuleName & ".9"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TotalNumberOfCCJsLessThan3Years", "")
    'EP694 - 07/06/2006 - Peter Edney
    If Not xmlBespokeBureauData Is Nothing Then
        strUnsatisfiedCCJ = GetAttributeText(xmlBespokeBureauData, "NUMBERUNSATISFIEDCCJ")
        strSatisfiedCCJ = GetAttributeText(xmlBespokeBureauData, "NUMBERSATISFIEDCCJ")
        If IsNumeric(strUnsatisfiedCCJ) Then intUnsatisfiedCCJ = CSafeInt(strUnsatisfiedCCJ)
        If IsNumeric(strSatisfiedCCJ) Then intSatisfiedCCJ = CSafeInt(strSatisfiedCCJ)
        xmlTempNode.Text = CStr(intUnsatisfiedCCJ + intSatisfiedCCJ)
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    

    'TotalValueOfCCJsLessThan3Years
    gstrRuleNumber = strRuleName & ".10"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TotalValueOfCCJsLessThan3Years", "")
    'EP694 - 07/06/2006 - Peter Edney
    xmlTempNode.Text = "0"
    Set xmlBureau2List = vxmlRequestNode.selectNodes("//APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/CREDITCHECK/BXCUSTOMERBUREAU2[@PERSONTYPE='1']")
    For Each xmlBureau2Node In xmlBureau2List
        strBX18_CCJAmount = GetAttributeText(xmlBureau2Node, "BX18_CCJAmount")
        If IsNumeric(strBX18_CCJAmount) Then
            xmlTempNode.Text = CStr(CSafeInt(xmlTempNode.Text) + CSafeInt(strBX18_CCJAmount))
        End If
    Next
    xmlCCJDataNode.appendChild xmlTempNode
    
    
    'Value of all unsatisfied CCJs - It seems this can be taken from VALUEUNSATISFIEDCCJ
    gstrRuleNumber = strRuleName & ".11"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValueOfAllUnsatisfiedCCJs", "")
    If Not xmlBespokeBureauData Is Nothing Then
        strTemp = GetAttributeText(xmlBespokeBureauData, "VALUEUNSATISFIEDCCJ")
        If Not strTemp = "" Then
            xmlTempNode.Text = strTemp
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    
    
    'No of all unsatisfied CCJs - It seems this can be taken from NUMBERUNSATISFIEDCCJ
    gstrRuleNumber = strRuleName & ".12"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NoOfAllUnsatisfiedCCJs", "")
    If Not xmlBespokeBureauData Is Nothing Then
        strTemp = GetAttributeText(xmlBespokeBureauData, "NUMBERUNSATISFIEDCCJ")
        If Not strTemp = "" Then
            xmlTempNode.Text = strTemp
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    
    
    'Value of all satisfied CCJs - It seems this can be taken from VALUESATISFIEDCCJ
    gstrRuleNumber = strRuleName & ".13"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValueOfAllSatisfiedCCJs", "")
    If Not xmlBespokeBureauData Is Nothing Then
        strTemp = GetAttributeText(xmlBespokeBureauData, "VALUESATISFIEDCCJ")
        If Not strTemp = "" Then
            xmlTempNode.Text = strTemp
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    
    
    'No of all satisfied CCJs - It seems this can be taken from NUMBERSATISFIEDCCJ
    gstrRuleNumber = strRuleName & ".14"
    Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NoOfAllSatisfiedCCJs", "")
    If Not xmlBespokeBureauData Is Nothing Then
        strTemp = GetAttributeText(xmlBespokeBureauData, "NUMBERSATISFIEDCCJ")
        If Not strTemp = "" Then
            xmlTempNode.Text = strTemp
        End If
    End If
    xmlCCJDataNode.appendChild xmlTempNode
    
    
    xmlMiscellaneousChargeDataNode.appendChild xmlCCJDataNode
    

    xmlChargeDetailsNode.appendChild xmlMiscellaneousChargeDataNode
    '** Get Miscellaneous Charge Fields - CCJ : END **



    vxmlResponseNode.appendChild xmlChargeDetailsNode
    
    LogDetails 3, "SetupChargeDetails : End "
    
    
SetupChargeDetailsExit:
    Set xmlDoc = Nothing
    Set xmlChargeDetailsNode = Nothing
    Set xmlProductInformationNode = Nothing
    Set xmlLineOfBusinessNode = Nothing
    Set xmlApplicationTypeNode = Nothing
    Set xmlMortgagePurposeTypeNode = Nothing
    Set xmlPaymentTermsNode = Nothing
    Set xmlFirstPaymentDateNode = Nothing
    Set xmlMiscellaneousChargeDataNode = Nothing
    Set xmlCCJDataNode = Nothing
    Set xmlApplicationCreditCheckNodeList = Nothing
    Set xmlApplicationCreditCheckNode = Nothing
    Set xmlBespokeBureauData = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetupBrokerDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupBrokerDetails
' Author:       Steve Badman
' Date:         11/05/2006
' Purpose:      Create Broker Details data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupBrokerDetailsExit
    
    Const strFunctionName As String = "SetupBrokerDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlBrokerDetailsNode As IXMLDOMNode
    Dim xmlBrokerFirmNode As IXMLDOMNode
    Dim strBrokerFirmName As String
    Dim strBrokerFirmFSANumber As String
    Dim xmlBrokerNumberNode As IXMLDOMNode
    Dim xmlBrokerCompanyNameNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
    Dim xmlBrokerAddressNode As IXMLDOMNode
    Dim xmlFlatNameOrNumberNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNameNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNumberNode As IXMLDOMNode
    Dim xmlStreetNode As IXMLDOMNode
    Dim xmlDistrictNode As IXMLDOMNode
    Dim xmlTownOrCityNode As IXMLDOMNode
    Dim xmlCountyNode As IXMLDOMNode
    Dim xmlPostCodeNode As IXMLDOMNode
    Dim strAddressGUID As String
    Dim xmlContactDetailsNode As IXMLDOMNode
    Dim xmlContactTelephoneDetailsNodeList As IXMLDOMNodeList
    Dim xmlContactTelephoneDetailsNode As IXMLDOMNode
    Dim strUsage As String
    Dim xmlBrokerWorkPhoneNode As IXMLDOMNode
    Dim xmlBrokerOtherPhoneNode As IXMLDOMNode
    Dim xmlAreaCodeNode As IXMLDOMNode
    Dim xmlLocalNumberNode As IXMLDOMNode
    Dim xmlExtensionNode As IXMLDOMNode
    Dim xmlCoCommissionNum As IXMLDOMNode
    
    LogDetails 3, "SetupBrokerDetails : Start "
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlBrokerDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BrokerDetails", "")
    
    strRuleName = "SetupBrokerDetails"
    
    'EP2_956
    Set xmlBrokerFirmNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM")
    If Not xmlBrokerFirmNode Is Nothing Then
     'We should look for a firmtradingname as well - but this may be a later enhancement ...
        strBrokerFirmName = GetAttributeText(xmlBrokerFirmNode, "ARFIRMNAME")
        strBrokerFirmFSANumber = GetAttributeText(xmlBrokerFirmNode, "FSAARFIRMREF")
    Else
      Set xmlBrokerFirmNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRMBROKER")
      If Not xmlBrokerFirmNode Is Nothing Then
        strBrokerFirmName = GetAttributeText(xmlBrokerFirmNode, "PRINCIPALFIRMNAME")
        strBrokerFirmFSANumber = GetAttributeText(xmlBrokerFirmNode, "FSAREF")
      End If
    End If
    'EP2_956 End
    
    If Not xmlBrokerFirmNode Is Nothing Then
   
        'Broker Number
        gstrRuleNumber = strRuleName & ".1"
        Set xmlBrokerNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BrokerNumber", "")
        'EP2_956
        If Len(strBrokerFirmFSANumber) = 0 Then
          strBrokerFirmFSANumber = "Unknown"
        End If
        xmlBrokerNumberNode.Text = strBrokerFirmFSANumber
        
'        'EP768 - 14/06/2006 - Peter Edney
'        If Not xmlBrokerIntermediaryOrgNode Is Nothing Then
'            Set xmlCoCommissionNum = xmlBrokerIntermediaryOrgNode.selectSingleNode("@COMPANYCOMMISIONNUMBER")
'            If Not xmlCoCommissionNum Is Nothing Then
'                xmlBrokerNumberNode.Text = xmlCoCommissionNum.Text
'            End If
'        End If
       'EP2_956 End
        xmlBrokerDetailsNode.appendChild xmlBrokerNumberNode
   
   
        'Broker Company Name
        gstrRuleNumber = strRuleName & ".2"
        Set xmlBrokerCompanyNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CompanyName", "")
        'EP2_956
        If Len(strBrokerFirmName) = 0 Then
          strBrokerFirmName = "Unknown"
        End If
        xmlBrokerCompanyNameNode.Text = strBrokerFirmName
        '''xmlBrokerCompanyNameNode.Text = GetAttributeText(xmlBrokerIntermediaryOrgNode, "NAME")
        xmlBrokerDetailsNode.appendChild xmlBrokerCompanyNameNode
         
        'EP2_956 - don't need this - its in the xmlBrokerFirmNode
'        'get AddressGUID and Address node
'        strAddressGUID = GetAttributeText(xmlBrokerFirmNode, "ADDRESSGUID")
'
'        If strAddressGUID <> "" Then
'            Set xmlAddressNode = xmlBrokerFirmNode.selectSingleNode("ADDRESS")
'        End If
        
        'set all nodes, dependant on address
        Set xmlBrokerAddressNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BrokerAddress", "")
        Set xmlFlatNameOrNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FlatNameOrNumber", "")
        Set xmlHouseOrBuildingNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingName", "")
        Set xmlHouseOrBuildingNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingNumber", "")
        Set xmlStreetNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Street", "")
        Set xmlDistrictNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "District", "")
        Set xmlTownOrCityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TownOrCity", "")
        Set xmlCountyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "County", "")
        Set xmlPostCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PostCode", "")
        
        
           'FlatNameOrNumber
            gstrRuleNumber = strRuleName & ".3"
            'xmlFlatNameOrNumberNode.Text = GetAttributeText(xmlAddressNode, "FLATNUMBER")
            'xmlFlatNameOrNumberNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE1")
            'HouseOrBuildingName
            gstrRuleNumber = strRuleName & ".4"
            'xmlHouseOrBuildingNameNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENAME")
            xmlHouseOrBuildingNameNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE1")
            'HouseOrBuildingNumber
            gstrRuleNumber = strRuleName & ".5"
            'xmlHouseOrBuildingNumberNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENUMBER")
            xmlHouseOrBuildingNumberNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE2")
            'Street
            gstrRuleNumber = strRuleName & ".6"
            'xmlStreetNode.Text = GetAttributeText(xmlAddressNode, "STREET")
            xmlStreetNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE3")
            'District
            gstrRuleNumber = strRuleName & ".7"
            'xmlDistrictNode.Text = GetAttributeText(xmlAddressNode, "DISTRICT")
            xmlDistrictNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE4")
            'TownOrCity
            gstrRuleNumber = strRuleName & ".8"
            'xmlTownOrCityNode.Text = GetAttributeText(xmlAddressNode, "TOWN")
            xmlTownOrCityNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE5")
            'County
            gstrRuleNumber = strRuleName & ".9"
            'xmlCountyNode.Text = GetAttributeText(xmlAddressNode, "COUNTY")
            xmlCountyNode.Text = GetAttributeText(xmlBrokerFirmNode, "ADDRESSLINE6")
            'PostCode
            gstrRuleNumber = strRuleName & ".10"
            'xmlPostCodeNode.Text = GetAttributeText(xmlAddressNode, "POSTCODE")
            xmlPostCodeNode.Text = GetAttributeText(xmlBrokerFirmNode, "POSTCODE")
        'End If
            
        'xmlBrokerAddressNode.appendChild xmlFlatNameOrNumberNode
        xmlBrokerAddressNode.appendChild xmlHouseOrBuildingNameNode
        xmlBrokerAddressNode.appendChild xmlHouseOrBuildingNumberNode
        xmlBrokerAddressNode.appendChild xmlStreetNode
        xmlBrokerAddressNode.appendChild xmlDistrictNode
        xmlBrokerAddressNode.appendChild xmlTownOrCityNode
        xmlBrokerAddressNode.appendChild xmlCountyNode
        xmlBrokerAddressNode.appendChild xmlPostCodeNode
        
        xmlBrokerDetailsNode.appendChild xmlBrokerAddressNode
    
    
    
    
        'get the Broker contact telephone details
        'Set xmlContactTelephoneDetailsNodeList = xmlBrokerFirmNode.selectNodes("CONTACTTELEPHONEDETAILS")
        
'        For Each xmlContactTelephoneDetailsNode In xmlContactTelephoneDetailsNodeList
'            strUsage = GetAttributeText(xmlContactTelephoneDetailsNode, "USAGE")
'            If Not strUsage = "" Then
'                If IsValidationType("ContactTelephoneUsage", CSafeInt(strUsage), "W") Then  '"W"=Work
                    'work phone number
                    gstrRuleNumber = strRuleName & ".12"
                    Set xmlBrokerWorkPhoneNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BrokerWorkPhoneNumber", "")
                    
                    Set xmlAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
                    Set xmlLocalNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
                    Set xmlExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
                    
                    xmlAreaCodeNode.Text = GetAttributeText(xmlBrokerFirmNode, "TELEPHONEAREACODE")
                    xmlLocalNumberNode.Text = GetAttributeText(xmlBrokerFirmNode, "TELEPHONENUMBER")
                    'xmlExtensionNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "EXTENSIONNUMBER")
                    
                    xmlBrokerWorkPhoneNode.appendChild xmlAreaCodeNode
                    xmlBrokerWorkPhoneNode.appendChild xmlLocalNumberNode
                   'xmlBrokerWorkPhoneNode.appendChild xmlExtensionNode
                
                    xmlBrokerDetailsNode.appendChild xmlBrokerWorkPhoneNode
'                Else
                    'Other Phone number
                    gstrRuleNumber = strRuleName & ".11"
                
                    Set xmlBrokerOtherPhoneNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BrokerOtherPhoneNumber", "")
                    
                    Set xmlAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
                    Set xmlLocalNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
                    Set xmlExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
                    
                    xmlAreaCodeNode.Text = GetAttributeText(xmlBrokerFirmNode, "FAXAREACODE")
                    xmlLocalNumberNode.Text = GetAttributeText(xmlBrokerFirmNode, "FAXNUMBER")
                    'xmlExtensionNode.Text = GetAttributeText(xmlContactTelephoneDetailsNode, "EXTENSIONNUMBER")
                    
                    xmlBrokerOtherPhoneNode.appendChild xmlAreaCodeNode
                    xmlBrokerOtherPhoneNode.appendChild xmlLocalNumberNode
                    xmlBrokerOtherPhoneNode.appendChild xmlExtensionNode
                
                    xmlBrokerDetailsNode.appendChild xmlBrokerOtherPhoneNode
'                End If
'            End If
'        Next
    End If


    vxmlResponseNode.appendChild xmlBrokerDetailsNode
    
    LogDetails 3, "SetupBrokerDetails : End "
    
    
SetupBrokerDetailsExit:
    Set xmlDoc = Nothing
    Set xmlBrokerDetailsNode = Nothing
    Set xmlBrokerFirmNode = Nothing
    Set xmlBrokerNumberNode = Nothing
    Set xmlBrokerCompanyNameNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlBrokerAddressNode = Nothing
    Set xmlFlatNameOrNumberNode = Nothing
    Set xmlHouseOrBuildingNameNode = Nothing
    Set xmlHouseOrBuildingNumberNode = Nothing
    Set xmlStreetNode = Nothing
    Set xmlDistrictNode = Nothing
    Set xmlTownOrCityNode = Nothing
    Set xmlCountyNode = Nothing
    Set xmlPostCodeNode = Nothing
    Set xmlContactDetailsNode = Nothing
    Set xmlContactTelephoneDetailsNodeList = Nothing
    Set xmlContactTelephoneDetailsNode = Nothing
    Set xmlBrokerWorkPhoneNode = Nothing
    Set xmlBrokerOtherPhoneNode = Nothing
    Set xmlAreaCodeNode = Nothing
    Set xmlLocalNumberNode = Nothing
    Set xmlExtensionNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetupAccountCreation(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupAccountCreation
' Author:       Sam Dali
' Date:         18/08/05
' Purpose:      Create account creation data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
On Error GoTo SetupAccountCreationExit

Const strFunctionName As String = "SetupAccountCreation"

Dim xmlDoc As FreeThreadedDOMDocument40
'Dim xmlDoc As New MSXML2.DOMDocument40
Dim xmlSearchNode As IXMLDOMNode 'temp node used for searching

Dim xmlMortgageAccountDetailsNode As IXMLDOMNode 'Outer Node

'all nodes for Account Creation
Dim xmlRelationshipCodeNode As IXMLDOMNode
Dim xmlPrimaryApplicantNode As IXMLDOMNode
Dim xmlSecondaryApplicantNode As IXMLDOMNode
Dim xmlMortgageApplicationNumberNode As IXMLDOMNode
Dim xmlSVRDefaultProductCodeNode As IXMLDOMNode
Dim xmlCurrencyCodeNode As IXMLDOMNode
Dim xmlApplicationTypeNode As IXMLDOMNode
Dim xmlMortgageAmountNode As IXMLDOMNode
Dim xmlInterestOnlyPortionNode As IXMLDOMNode
Dim xmlCompletionDateNode As IXMLDOMNode
Dim xmlPreferredPaymentDayNode As IXMLDOMNode
Dim xmlNumberOfPaymentsNode As IXMLDOMNode
Dim xmlApplicantBankAccountDetailsNode As IXMLDOMNode
Dim xmlBankAccountNameNode As IXMLDOMNode
Dim xmlBankAccountSortCodeNode As IXMLDOMNode
Dim xmlBankAccountNumberNode As IXMLDOMNode
Dim xmlTransposedIndicatorNode As IXMLDOMNode
Dim xmlPanelNumberNode As IXMLDOMNode
Dim xmlSolicitorBankAccountDetailsNode As IXMLDOMNode
Dim xmlSolicitorAccountNameNode As IXMLDOMNode
Dim xmlSolicitorSortCodeNode As IXMLDOMNode
Dim xmlSolicitorAccountNumberNode As IXMLDOMNode
Dim xmlSolicitorFundingMethodNode As IXMLDOMNode
Dim xmlMortgageReferenceNumberNode As IXMLDOMNode
Dim xmlHubReferenceNumberNode As IXMLDOMElement
Dim xmlIncomeBasisNode As IXMLDOMNode
Dim xmlIncomeNode As IXMLDOMNode
Dim xmlMaritalStatusNode As IXMLDOMNode
Dim xmlEmploymentStatusNode As IXMLDOMNode
Dim xmlRiskIndicatorNode As IXMLDOMNode
Dim xmlCreditScoreNode As IXMLDOMNode
Dim xmlSourceChannelNode As IXMLDOMNode

LogDetails 3, "SetupAccountCreation : Start "

Set xmlDoc = New FreeThreadedDOMDocument40
xmlDoc.validateOnParse = False
xmlDoc.setProperty "NewParser", True

Set xmlMortgageAccountDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageAccountDetails", "")

Dim strRuleName As String
strRuleName = "AccountCreation"

Dim strApplicationNumber As String
strApplicationNumber = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONNUMBER")

'get primary customerRole
Dim xmlPrimaryCustomerRoleNode As IXMLDOMNode
Set xmlPrimaryCustomerRoleNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER='1']")

'get secondary CustomerRole
Dim xmlSecondaryCustomerRoleNode As IXMLDOMNode
Set xmlSecondaryCustomerRoleNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER='2']")

Dim strCustomerNumber As String, strCustomerVersionNumber As String, strSecondaryCustomerVersion As String
strCustomerNumber = GetAttributeText(xmlPrimaryCustomerRoleNode, "CUSTOMERNUMBER")
strCustomerVersionNumber = GetAttributeText(xmlPrimaryCustomerRoleNode, "CUSTOMERVERSIONNUMBER")

If Not xmlSecondaryCustomerRoleNode Is Nothing Then
    strSecondaryCustomerVersion = GetAttributeText(xmlSecondaryCustomerRoleNode, "CUSTOMERVERSIONNUMBER")
End If

'RelationshipCode
gstrRuleNumber = strRuleName & ".1"
Set xmlRelationshipCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RelationshipCode", "")
xmlRelationshipCodeNode.Text = GetMandatoryAttributeText(GetMandatoryNode(vxmlRequestNode, "APPLICATION/APPLICATIONFACTFIND"), "NUMBEROFAPPLICANTS")

'PrimaryApplicantId
gstrRuleNumber = strRuleName & ".2"
Set xmlPrimaryApplicantNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PrimaryApplicantId", "")
If Not xmlPrimaryCustomerRoleNode Is Nothing Then
    xmlPrimaryApplicantNode.Text = GetMandatoryAttributeText(xmlPrimaryCustomerRoleNode.selectSingleNode( _
                                    "CUSTOMER"), "CUSTOMERNUMBER")
    'xmlPrimaryApplicantNode.Text = GetMandatoryAttributeText(xmlPrimaryCustomerRoleNode.selectSingleNode( _
    '                                "CUSTOMER[@CUSTOMERNUMBER = '" & _
    '                                GetMandatoryAttributeText(xmlPrimaryCustomerRoleNode, "CUSTOMERNUMBER") & "']"), "OTHERSYSTEMCUSTOMERNUMBER")
End If

'SecondaryApplicantId
gstrRuleNumber = strRuleName & ".3"
Set xmlSecondaryApplicantNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SecondaryApplicantId", "")
If Not xmlSecondaryCustomerRoleNode Is Nothing Then
    xmlSecondaryApplicantNode.Text = GetAttributeText(xmlSecondaryCustomerRoleNode.selectSingleNode( _
                                    "CUSTOMER"), "CUSTOMERNUMBER")
'    xmlSecondaryApplicantNode.Text = GetAttributeText(xmlSecondaryCustomerRoleNode.selectSingleNode( _
'                                    "CUSTOMER[@CUSTOMERNUMBER = '" & _
'                                    GetAttributeText(xmlSecondaryCustomerRoleNode, "CUSTOMERNUMBER") & "']"), "OTHERSYSTEMCUSTOMERNUMBER")
End If

'MortgageApplicationNumber
gstrRuleNumber = strRuleName & ".4"
Set xmlMortgageApplicationNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageApplicationNumber", "")
xmlMortgageApplicationNumberNode.Text = strApplicationNumber

'ProductTypeID
gstrRuleNumber = strRuleName & ".5"
Set xmlSVRDefaultProductCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProductTypeID", "")
xmlSVRDefaultProductCodeNode.Text = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("GLOBALPARAMETER[@NAME='SVRDefaultProductCode']"), "STRING")

'CurrencyCode
gstrRuleNumber = strRuleName & ".6"
Set xmlCurrencyCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CurrencyCode", "")
xmlCurrencyCodeNode.Text = "GBP" 'it's hardcoded

'ApplicationType
gstrRuleNumber = strRuleName & ".7"
Set xmlApplicationTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ApplicationType", "")
Dim strApplicationType As String
Dim strApplicationTypeValue As String
strApplicationType = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "TYPEOFAPPLICATION")

If Not strApplicationType = "" Then
    If IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "N") And IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "FT") Then
        'xmlApplicationTypeNode.Text = "1"
        xmlApplicationTypeNode.Text = "2" 'First Time Buyer
    Else
        If IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "N") And IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "HM") Then
            'xmlApplicationTypeNode.Text = "2"
            xmlApplicationTypeNode.Text = "1" 'Home Mover
        Else
            If IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "R") Then
                xmlApplicationTypeNode.Text = "3"
            End If
        End If
    End If
End If

'MortageAmount
gstrRuleNumber = strRuleName & ".8"
Set xmlMortgageAmountNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageAmount", "")
Dim strMortgageSubQuoteNumber As String

strMortgageSubQuoteNumber = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION[@QUOTATIONNUMBER='" & _
                                   GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "ACCEPTEDQUOTENUMBER") & "']"), "MORTGAGESUBQUOTENUMBER")

xmlMortgageAmountNode.Text = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER = '" & _
                                strMortgageSubQuoteNumber & "']"), "TOTALLOANAMOUNT")
                                

'InterestOnlyPortion
gstrRuleNumber = strRuleName & ".9"
Dim xmlLoanComponentNodeList As IXMLDOMNodeList
Dim xmlLoanComponent As IXMLDOMNode
Dim intPartAndPartInterestAmount, intIOAmount As Long
Dim strRepaymentMethod As String

Set xmlLoanComponentNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGESUBQUOTENUMBER='" & strMortgageSubQuoteNumber & "']")
For Each xmlLoanComponent In xmlLoanComponentNodeList
    strRepaymentMethod = GetMandatoryAttributeText(xmlLoanComponent, "REPAYMENTMETHOD")
    'assumed INTERESTONLYELEMENT & TOTALLOANCOMPONENTAMOUNT r mandatory, if not this may fail
    If (strRepaymentMethod = "3") Then
        intPartAndPartInterestAmount = intPartAndPartInterestAmount + CSafeLng(GetMandatoryAttributeText(xmlLoanComponent, "INTERESTONLYELEMENT"))
    Else
        If (strRepaymentMethod = "1") Then
            intIOAmount = intIOAmount + CSafeLng(GetMandatoryAttributeText(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT"))
        End If
    End If
Next

Set xmlInterestOnlyPortionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InterestOnlyPortion", "")
xmlInterestOnlyPortionNode.Text = CStr(intPartAndPartInterestAmount + intIOAmount) '= 0 + 0, if RepaymentMethod = 2 for all loan components
                            
'CompletionDate
gstrRuleNumber = strRuleName & ".10"
Set xmlCompletionDateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CompletionDate", "")

Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE")
If Not xmlSearchNode Is Nothing Then
    'truncate the time
    'completionDate is mandatory, if not this may fail
    xmlCompletionDateNode.Text = Format(GetMandatoryAttributeText(xmlSearchNode, "COMPLETIONDATE"), "yyyy/mm/dd")
End If
Set xmlSearchNode = Nothing

Dim xmlApplicationBankBuildingSOCNode As IXMLDOMNode
Dim strSortCode As String
Set xmlApplicationBankBuildingSOCNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC")

'create the nodes, which depend on ApplicationBankBuildingSOC
'if xmlApplicantBankAccountDetailsNode = nothing, these nodes will be empty
Set xmlPreferredPaymentDayNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PreferredPaymentDay", "")
Set xmlApplicantBankAccountDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ApplicantBankAccountDetails", "")
Set xmlBankAccountNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AccountName", "")
Set xmlBankAccountSortCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SortCode", "")
Set xmlBankAccountNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AccountNumber", "")
Set xmlTransposedIndicatorNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TransposedIndicator", "")

If Not xmlApplicationBankBuildingSOCNode Is Nothing Then
    'PreferredPaymentDay
    gstrRuleNumber = strRuleName & ".11"
    xmlPreferredPaymentDayNode.Text = GetAttributeText(xmlApplicationBankBuildingSOCNode, "PREFEREDPAYMENTDAY")

    'AccountName
    gstrRuleNumber = strRuleName & ".14"
    xmlBankAccountNameNode.Text = GetAttributeText(xmlApplicationBankBuildingSOCNode, "ACCOUNTNAME")
    xmlApplicantBankAccountDetailsNode.appendChild xmlBankAccountNameNode
    
    'SortCode
    gstrRuleNumber = strRuleName & ".12"
    strSortCode = GetAttributeText(xmlApplicationBankBuildingSOCNode.selectSingleNode("THIRDPARTY"), "THIRDPARTYBANKSORTCODE")
    strSortCode = Replace(strSortCode, "-", "")
    xmlBankAccountSortCodeNode.Text = strSortCode
    xmlApplicantBankAccountDetailsNode.appendChild xmlBankAccountSortCodeNode
    
    'AccountNumber
    gstrRuleNumber = strRuleName & ".13"
    xmlBankAccountNumberNode.Text = GetAttributeText(xmlApplicationBankBuildingSOCNode, "ACCOUNTNUMBER")
    xmlApplicantBankAccountDetailsNode.appendChild xmlBankAccountNumberNode
    
    'TransposedIndicator
    gstrRuleNumber = strRuleName & ".15"
    xmlTransposedIndicatorNode.Text = GetAttributeText(xmlApplicationBankBuildingSOCNode, "TRANSPOSEDINDICATOR")
    If xmlTransposedIndicatorNode.Text <> "" Then
        xmlApplicantBankAccountDetailsNode.appendChild xmlTransposedIndicatorNode
    End If
End If


'get directoryGUID
Dim strDirectoryGUID As String
Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/THIRDPARTYDATA/APPLICATIONLEGALREP")
If Not xmlSearchNode Is Nothing Then
    strDirectoryGUID = GetAttributeText(xmlSearchNode, "DIRECTORYGUID")
End If
Set xmlSearchNode = Nothing

'PanelNumber
Dim strPanelNumber As String
gstrRuleNumber = strRuleName & ".16"
Set xmlPanelNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PanelNumber", "")
If Not strDirectoryGUID = "" Then
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/THIRDPARTYDATA/APPLICATIONLEGALREP/PANEL[@DIRECTORYGUID='" & _
    strDirectoryGUID & "']")
    If Not xmlSearchNode Is Nothing Then
        strPanelNumber = Left(GetAttributeText(xmlSearchNode, "PANELID"), 8)
        xmlPanelNumberNode.Text = strPanelNumber
    End If
End If

Set xmlSolicitorBankAccountDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorBankAccountDetails", "")

'create all elements which are dependant PANELBANKACCOUNT
'if PANELBANKACCOUNT is not there, these will be empty
Set xmlSolicitorAccountNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AccountName", "")
Set xmlSolicitorSortCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SortCode", "")
Set xmlSolicitorAccountNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AccountNumber", "")

Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/THIRDPARTYDATA/APPLICATIONLEGALREP/PANELBANKACCOUNT[@DIRECTORYGUID='" & _
strDirectoryGUID & "']")
If Not xmlSearchNode Is Nothing Then
    'SolicitorAccountName
    gstrRuleNumber = strRuleName & ".19"
    xmlSolicitorAccountNameNode.Text = GetAttributeText(xmlSearchNode, "ACCOUNTNAME")
    
    'SolicitorSortCode
    gstrRuleNumber = strRuleName & ".17"
    strSortCode = GetAttributeText(xmlSearchNode, "BANKSORTCODE")
    strSortCode = Replace(strSortCode, "-", "")
    xmlSolicitorSortCodeNode.Text = strSortCode

    'SolicitorAccountNumber
    gstrRuleNumber = strRuleName & ".18"
    xmlSolicitorAccountNumberNode.Text = GetAttributeText(xmlSearchNode, "ACCOUNTNUMBER")
End If

xmlSolicitorBankAccountDetailsNode.appendChild xmlSolicitorAccountNameNode
xmlSolicitorBankAccountDetailsNode.appendChild xmlSolicitorSortCodeNode
xmlSolicitorBankAccountDetailsNode.appendChild xmlSolicitorAccountNumberNode

Set xmlSearchNode = Nothing

'SolicitorFundingMethod
gstrRuleNumber = strRuleName & ".20"
Dim strPaymentMethodValueID As String, strPaymentSequenceNumber As String
Set xmlSolicitorFundingMethodNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SolicitorFundingMethod", "")

Set xmlSearchNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
If Not xmlSearchNode Is Nothing Then
    strPaymentSequenceNumber = GetMandatoryAttributeText(xmlSearchNode, "PAYMENTSEQUENCENUMBER")
    Set xmlSearchNode = Nothing
End If

Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/PAYMENTDETAILS/PAYMENTRECORD[@APPLICATIONNUMBER='" & strApplicationNumber & "' and @PAYMENTSEQUENCENUMBER='" & strPaymentSequenceNumber & "']")

If Not xmlSearchNode Is Nothing Then
    strPaymentMethodValueID = GetAttributeText(xmlSearchNode, "PAYMENTMETHOD")
    Set xmlSearchNode = Nothing
End If
If Not strPaymentMethodValueID = "" Then
    If IsValidationType("PaymentMethod", CSafeInt(strPaymentMethodValueID), "B") Then
        xmlSolicitorFundingMethodNode.Text = "1"
    Else
        If IsValidationType("PaymentMethod", CSafeInt(strPaymentMethodValueID), "YC") Then
            xmlSolicitorFundingMethodNode.Text = "2"
        End If
    End If
End If
Set xmlSearchNode = Nothing

'get CONVEYANCERREFERENCE & HUBREFERENCE from ApplicationFirstTitle with latest dateTime and messageType 'ACA'
Dim xmlApplicationFirstTitleNodeList As IXMLDOMNodeList
Dim xmlApplicationFirstTitleNode As IXMLDOMNode
Dim dtDateTime, dtLatestDateTime As Date
Dim strConveyencerReference, strHubReference As String

Set xmlApplicationFirstTitleNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/THIRDPARTYDATA/APPLICATIONFIRSTTITLE[@MESSAGETYPE='ACA']")
For Each xmlApplicationFirstTitleNode In xmlApplicationFirstTitleNodeList
    dtDateTime = CSafeDate(GetAttributeText(xmlApplicationFirstTitleNode, "DATETIME"))
    If dtDateTime > dtLatestDateTime Then
        strConveyencerReference = GetAttributeText(xmlApplicationFirstTitleNode, "CONVEYANCERREFERENCE")
        strHubReference = GetAttributeText(xmlApplicationFirstTitleNode, "HUBREFERENCE")
        dtLatestDateTime = dtDateTime
    End If
Next

'MortgageReferenceNumber
gstrRuleNumber = strRuleName & ".21"
Set xmlMortgageReferenceNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageReferenceNumber", "")
If IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "R") Then
    xmlMortgageReferenceNumberNode.Text = strConveyencerReference
Else
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE")
    If Not xmlSearchNode Is Nothing Then
        xmlMortgageReferenceNumberNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "SOLICITORSREFNUMBER")
    End If
End If
Set xmlSearchNode = Nothing

'HubReferenceNumber
gstrRuleNumber = strRuleName & ".22"
Set xmlHubReferenceNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HubReferenceNumber", "")
If IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "R") Then
    xmlHubReferenceNumberNode.Text = strHubReference
End If

'If strHubReference = "" Then
'    Dim xmlNode As MSXML2.IXMLDOMNode
'    Set xmlNode = xmlDoc.createNode(MSXML2.NODE_ATTRIBUTE, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance")
'    xmlNode.Text = "true"
'    xmlHubReferenceNumberNode.setAttributeNode (xmlNode)
    
'    Dim xmlAttrib As IXMLDOMAttribute
'    Set xmlAttrib = xmlNode.ownerDocument.createAttribute(vstrAttribName)
'    xmlAttrib.Value = vstrAttribValue
'
'    vxmlDestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
'End If

'IncomeBasis
gstrRuleNumber = strRuleName & ".23"
Dim strIncomeMultiplierType As String
Set xmlIncomeBasisNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IncomeBasis", "")
Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER='" & strMortgageSubQuoteNumber & "']")
If Not xmlSearchNode Is Nothing Then
    strIncomeMultiplierType = GetAttributeText(xmlSearchNode, "INCOMEMULTIPLIERTYPE")
End If
If strIncomeMultiplierType = "S" Then
    xmlIncomeBasisNode.Text = "1"
Else
    If (strIncomeMultiplierType = "J" Or strIncomeMultiplierType = "H") Then
        xmlIncomeBasisNode.Text = "2"
    End If
End If

'Income
gstrRuleNumber = strRuleName & ".24"
Dim xmlPrimaryCurrentEmployersRefNode, xmlSecondaryCurrentEmployersRefNode As IXMLDOMNode
Dim strGrossAllowableIncome1, strGrossAllowableIncome2 As String
Dim dblIncome As Double
Set xmlIncomeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Income", "")

If Not xmlPrimaryCustomerRoleNode Is Nothing Then
    Set xmlPrimaryCurrentEmployersRefNode = xmlPrimaryCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']/CURRENTEMPLOYERSREF")
End If

If Not xmlSecondaryCustomerRoleNode Is Nothing Then
    Set xmlSecondaryCurrentEmployersRefNode = xmlSecondaryCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strSecondaryCustomerVersion & "']/CURRENTEMPLOYERSREF")
End If

'at least one currentEmployersRef exists
'If Not (xmlPrimaryCurrentEmployersRefNode Is Nothing And xmlSecondaryCurrentEmployersRefNode Is Nothing) Then
'    If Not xmlPrimaryCurrentEmployersRefNode Is Nothing Then
'        strGrossAllowableIncome1 = GetAttributeText(xmlPrimaryCurrentEmployersRefNode, "GROSSALLOWABLEINCOME")
'    End If

'    If Not xmlSecondaryCurrentEmployersRefNode Is Nothing Then
'        strGrossAllowableIncome2 = GetAttributeText(xmlSecondaryCurrentEmployersRefNode, "GROSSALLOWABLEINCOME")
'    End If
'Else
        'get GROSSCONFIRMEDALLOWABLEINCOME/allowableIncome from IncomeSummary
        If Not xmlPrimaryCustomerRoleNode Is Nothing Then
            Set xmlSearchNode = xmlPrimaryCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']/INCOME/INCOMESUMMARY")
            If Not xmlSearchNode Is Nothing Then
                strGrossAllowableIncome1 = GetAttributeText(xmlSearchNode, "GROSSCONFIRMEDALLOWABLEINCOME")
                
                If strGrossAllowableIncome1 = "" Then
                    strGrossAllowableIncome1 = GetAttributeText(xmlSearchNode, "ALLOWABLEANNUALINCOME")
                End If
            End If
            Set xmlSearchNode = Nothing
        End If
        
        If Not xmlSecondaryCustomerRoleNode Is Nothing Then
            Set xmlSearchNode = xmlSecondaryCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strSecondaryCustomerVersion & "']/INCOME/INCOMESUMMARY")
            If Not xmlSearchNode Is Nothing Then
                strGrossAllowableIncome2 = GetAttributeText(xmlSearchNode, "GROSSCONFIRMEDALLOWABLEINCOME")
                
                If strGrossAllowableIncome2 = "" Then
                    strGrossAllowableIncome2 = GetAttributeText(xmlSearchNode, "ALLOWABLEANNUALINCOME")
                End If
            End If
            Set xmlSearchNode = Nothing
        End If
'End If

If Not strGrossAllowableIncome1 = "" Then
    dblIncome = CSafeDbl(strGrossAllowableIncome1)
End If

If Not strGrossAllowableIncome2 = "" Then
    dblIncome = dblIncome + CSafeDbl(strGrossAllowableIncome2)
End If
xmlIncomeNode.Text = CStr(dblIncome)

'MaritalStatus
gstrRuleNumber = strRuleName & ".25"
Dim strMaritalStatusValueID As String
Set xmlMaritalStatusNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MaritalStatus", "")

'assumed customerversion does exist
'strMaritalStatusValueID = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION[@CUSTOMERNUMBER='" & _
'strCustomerNumber & "']"), "MARITALSTATUS")

'Set xmlCustomerVersionNode = xmlCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersion & "']")

strMaritalStatusValueID = GetAttributeText(xmlPrimaryCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']"), "MARITALSTATUS")

If strMaritalStatusValueID <> "" Then
    If IsValidationType("MaritalStatus", CSafeInt(strMaritalStatusValueID), "M") Then
        xmlMaritalStatusNode.Text = "M"
    Else
        If IsValidationType("MaritalStatus", CSafeInt(strMaritalStatusValueID), "S") Then
            xmlMaritalStatusNode.Text = "S"
        Else
            If IsValidationType("MaritalStatus", CSafeInt(strMaritalStatusValueID), "D") Then
                xmlMaritalStatusNode.Text = "D"
            Else
                If IsValidationType("MaritalStatus", CSafeInt(strMaritalStatusValueID), "W") Then
                    xmlMaritalStatusNode.Text = "W"
                Else
                    If IsValidationType("MaritalStatus", CSafeInt(strMaritalStatusValueID), "C") Then
                        xmlMaritalStatusNode.Text = "C"
                    Else
                        If IsValidationType("MaritalStatus", CSafeInt(strMaritalStatusValueID), "O") Then
                            xmlMaritalStatusNode.Text = "O"
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

'EmploymentStatus
gstrRuleNumber = strRuleName & ".26"
Dim strEmploymentStatusValueID As String
Set xmlEmploymentStatusNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "EmploymentStatus", "")
Set xmlSearchNode = xmlPrimaryCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersionNumber & "']/INCOME/EMPLOYMENT[@CUSTOMERNUMBER='" & _
                             strCustomerNumber & "' and @CUSTOMERVERSIONNUMBER ='" & strCustomerVersionNumber & "' and @MAINSTATUS = '1' ]")
If Not xmlSearchNode Is Nothing Then
    strEmploymentStatusValueID = GetAttributeText(xmlSearchNode, "EMPLOYMENTSTATUS")
End If
Set xmlSearchNode = Nothing

If strEmploymentStatusValueID <> "" Then
    If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "EMP") Then
            xmlEmploymentStatusNode.Text = "E"
    Else
        'If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "C") Then
        If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "CON") Then
            xmlEmploymentStatusNode.Text = "T"
        Else
            If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "SELF") Then
                xmlEmploymentStatusNode.Text = "P"
            Else
                If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "N") Then
                    xmlEmploymentStatusNode.Text = "U"
                Else
                    If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "HM") Then
                        xmlEmploymentStatusNode.Text = "H"
                    Else
                        If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "R") Then
                            xmlEmploymentStatusNode.Text = "R"
                        Else
                            If IsValidationType("EmploymentStatus", CSafeInt(strEmploymentStatusValueID), "STU") Then
                                xmlEmploymentStatusNode.Text = "S"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

'get APPLICATIONCREDITCHECK.CREDITCHECKGUID with highest SEQUENCENUMBER
Dim xmlApplicationCreditCheckNodeList As IXMLDOMNodeList
Dim xmlApplicationCreditCheckNode As IXMLDOMNode
Dim intMaxSequenceNumber As Integer, intSequenceNumber As Integer
Dim strCreditCheckGUID As String, strSequenceNumber As String
Set xmlApplicationCreditCheckNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK")
For Each xmlApplicationCreditCheckNode In xmlApplicationCreditCheckNodeList
    strSequenceNumber = GetAttributeText(xmlApplicationCreditCheckNode, "SEQUENCENUMBER")
    If strSequenceNumber <> "" Then
        intSequenceNumber = CSafeInt(strSequenceNumber)
    End If
    If intSequenceNumber > intMaxSequenceNumber Then
        intMaxSequenceNumber = intSequenceNumber
        strCreditCheckGUID = GetAttributeText(xmlApplicationCreditCheckNode, "CREDITCHECKGUID")
    End If
Next

'RiskIndicator
gstrRuleNumber = strRuleName & ".27"
Set xmlRiskIndicatorNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RiskIndicator", "")
Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/BESPOKEDECISION[@CREDITCHECKGUID='" & strCreditCheckGUID & "']")
If Not xmlSearchNode Is Nothing Then
    xmlRiskIndicatorNode.Text = GetAttributeText(xmlSearchNode, "RISKINDICATOR")
End If
Set xmlSearchNode = Nothing

'CreditScore
gstrRuleNumber = strRuleName & ".28"
Set xmlCreditScoreNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CreditScore", "")
'SC 17/03/06 Use the node where the scorecard id = 1
Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/CREDITCHECKSCORE[@CREDITCHECKGUID='" & strCreditCheckGUID & "' and @SCORECARDID='" & "1" & "']")
If Not xmlSearchNode Is Nothing Then
    xmlCreditScoreNode.Text = GetAttributeText(xmlSearchNode, "SCORE")
End If
Set xmlSearchNode = Nothing

'SourceChannel
gstrRuleNumber = strRuleName & ".29"
Set xmlSourceChannelNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SourceChannel", "")
Dim strDirectIndirectBusiness As String
strDirectIndirectBusiness = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "DIRECTINDIRECTBUSINESS")
If strDirectIndirectBusiness <> "" Then
    If IsValidationType("Direct/Indirect", CSafeInt(strDirectIndirectBusiness), "P") Then
        xmlSourceChannelNode.Text = "DAWT"
    Else
        If IsValidationType("Direct/Indirect", CSafeInt(strDirectIndirectBusiness), "W") Then
            xmlSourceChannelNode.Text = "WEB"
        Else
            If IsValidationType("Direct/Indirect", CSafeInt(strDirectIndirectBusiness), "INT") Then
                xmlSourceChannelNode.Text = "INT"
            Else
                If IsValidationType("Direct/Indirect", CSafeInt(strDirectIndirectBusiness), "WF") Then
                   xmlSourceChannelNode.Text = "WF"
                End If
            End If
        End If
    End If
End If

'NumberOfPayments
gstrRuleNumber = strRuleName & ".30"
Set xmlNumberOfPaymentsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfPayments", "")
Dim xmlLoanComponentList As IXMLDOMNodeList
Dim xmlLoanComponentNode As IXMLDOMNode
Dim intTermInYears As Integer, intTermInMonths As Integer, intMaxTermInYears As Integer

'get the list of loan components matching the criteria
Set xmlLoanComponentList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGESUBQUOTENUMBER = '" & strMortgageSubQuoteNumber & "']")
                                     
'get the max of TermInYears and TermInMonths for that loan component
For Each xmlLoanComponentNode In xmlLoanComponentList
    intTermInYears = CSafeInt(GetAttributeText(xmlLoanComponentNode, "TERMINYEARS"))
    If intTermInYears > intMaxTermInYears Then
        intMaxTermInYears = intTermInYears
        intTermInMonths = CSafeInt(GetAttributeText(xmlLoanComponentNode, "TERMINMONTHS"))
    End If
Next
xmlNumberOfPaymentsNode.Text = CStr((12 * intMaxTermInYears) + intTermInMonths)

xmlMortgageAccountDetailsNode.appendChild xmlRelationshipCodeNode
xmlMortgageAccountDetailsNode.appendChild xmlPrimaryApplicantNode
If xmlSecondaryApplicantNode.Text <> "" Then
    xmlMortgageAccountDetailsNode.appendChild xmlSecondaryApplicantNode
End If
xmlMortgageAccountDetailsNode.appendChild xmlMortgageApplicationNumberNode
xmlMortgageAccountDetailsNode.appendChild xmlSVRDefaultProductCodeNode
xmlMortgageAccountDetailsNode.appendChild xmlCurrencyCodeNode
xmlMortgageAccountDetailsNode.appendChild xmlApplicationTypeNode
xmlMortgageAccountDetailsNode.appendChild xmlMortgageAmountNode
xmlMortgageAccountDetailsNode.appendChild xmlInterestOnlyPortionNode
xmlMortgageAccountDetailsNode.appendChild xmlCompletionDateNode
xmlMortgageAccountDetailsNode.appendChild xmlPreferredPaymentDayNode
xmlMortgageAccountDetailsNode.appendChild xmlNumberOfPaymentsNode
xmlMortgageAccountDetailsNode.appendChild xmlApplicantBankAccountDetailsNode
xmlMortgageAccountDetailsNode.appendChild xmlPanelNumberNode
xmlMortgageAccountDetailsNode.appendChild xmlSolicitorBankAccountDetailsNode
xmlMortgageAccountDetailsNode.appendChild xmlSolicitorFundingMethodNode
xmlMortgageAccountDetailsNode.appendChild xmlMortgageReferenceNumberNode
If xmlHubReferenceNumberNode.Text <> "" Then
    xmlMortgageAccountDetailsNode.appendChild xmlHubReferenceNumberNode
End If
If xmlIncomeBasisNode.Text <> "" Then
    xmlMortgageAccountDetailsNode.appendChild xmlIncomeBasisNode
End If
xmlMortgageAccountDetailsNode.appendChild xmlIncomeNode
xmlMortgageAccountDetailsNode.appendChild xmlMaritalStatusNode
xmlMortgageAccountDetailsNode.appendChild xmlEmploymentStatusNode
xmlMortgageAccountDetailsNode.appendChild xmlRiskIndicatorNode
xmlMortgageAccountDetailsNode.appendChild xmlCreditScoreNode
xmlMortgageAccountDetailsNode.appendChild xmlSourceChannelNode

vxmlResponseNode.appendChild xmlMortgageAccountDetailsNode

LogDetails 3, "SetupAccountCreation : End "

SetupAccountCreationExit:
    Set xmlRelationshipCodeNode = Nothing
    Set xmlPrimaryApplicantNode = Nothing
    Set xmlSecondaryApplicantNode = Nothing
    Set xmlMortgageApplicationNumberNode = Nothing
    Set xmlSVRDefaultProductCodeNode = Nothing
    Set xmlCurrencyCodeNode = Nothing
    Set xmlApplicationTypeNode = Nothing
    Set xmlMortgageAmountNode = Nothing
    Set xmlInterestOnlyPortionNode = Nothing
    Set xmlCompletionDateNode = Nothing
    Set xmlPreferredPaymentDayNode = Nothing
    Set xmlApplicantBankAccountDetailsNode = Nothing
    Set xmlPanelNumberNode = Nothing
    Set xmlSolicitorBankAccountDetailsNode = Nothing
    Set xmlSolicitorAccountNameNode = Nothing
    Set xmlSolicitorAccountNumberNode = Nothing
    Set xmlSolicitorSortCodeNode = Nothing
    Set xmlSolicitorFundingMethodNode = Nothing
    Set xmlMortgageReferenceNumberNode = Nothing
    Set xmlHubReferenceNumberNode = Nothing
    Set xmlBankAccountSortCodeNode = Nothing
    Set xmlBankAccountNumberNode = Nothing
    Set xmlBankAccountNameNode = Nothing
    Set xmlIncomeBasisNode = Nothing
    Set xmlIncomeNode = Nothing
    Set xmlMaritalStatusNode = Nothing
    Set xmlEmploymentStatusNode = Nothing
    Set xmlRiskIndicatorNode = Nothing
    Set xmlSourceChannelNode = Nothing
    Set xmlNumberOfPaymentsNode = Nothing
    
    Set xmlMortgageAccountDetailsNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
   
End Sub
Private Sub XXX_SetupPropertyDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupPropertyDetails
' Author:       Sam Dali
' Date:         25/08/05
' Purpose:      Create Property Details in Completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to
'------------------------------------------------------------------------

On Error GoTo SetupPropertyDetailsExit

Const strFunctionName As String = "SetupPropertyDetails"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlSearchNode As IXMLDOMNode 'temp node used for searching

'all nodes used for Property Details
Dim xmlValuationTypeNode As IXMLDOMNode
Dim xmlPropertyAddressNode As IXMLDOMNode
Dim xmlHouseOrBuildingNameNode As IXMLDOMNode
Dim xmlFlatNameOrNumberNode As IXMLDOMNode
Dim xmlHouseOrBuildingNumberNode As IXMLDOMNode
Dim xmlStreetNode As IXMLDOMNode
Dim xmlDistrictNode As IXMLDOMNode
Dim xmlTownOrCityNode As IXMLDOMNode
Dim xmlCountyNode As IXMLDOMNode
Dim xmlPostCodeNode As IXMLDOMNode
Dim xmlIsNewPropertyNode As IXMLDOMNode
Dim xmlTenureOfPropertyNode As IXMLDOMNode
Dim xmlUnexpiredTermNode As IXMLDOMNode
Dim xmlPropertyTypeNode As IXMLDOMNode
Dim xmlExLocalAuthorityNode As IXMLDOMNode
Dim xmlNumberOfLivingRoomNode As IXMLDOMNode
Dim xmlNumberOfBedRoomsNode As IXMLDOMNode
Dim xmlNumberOfKitchensNode As IXMLDOMNode
Dim xmlNumberOfBathroomsNode As IXMLDOMNode
Dim xmlNumberOfGaragesNode As IXMLDOMNode
Dim xmlNumberOfParkingSpacesNode As IXMLDOMNode
Dim xmlApproximateYearPropertyBuildNode As IXMLDOMNode
Dim xmlBuildingStandardIndemnityTypeNode As IXMLDOMNode
Dim xmlGrossExternalFloorAreaNode As IXMLDOMNode
Dim xmlConstructionTypeIndicatorNode As IXMLDOMNode
Dim xmlConstructionTypeNode As IXMLDOMNode
Dim xmlSaleabilityNode As IXMLDOMNode
Dim xmlOverallConditionNode As IXMLDOMNode
Dim xmlLastValuationNode As IXMLDOMNode
Dim xmlBuildingInsuranceReinstallmentAmountNode As IXMLDOMNode
Dim xmlValuerNameNode As IXMLDOMNode
Dim xmlLastValuationDateNode As IXMLDOMNode
Dim xmlPropertyPurchaseNode As IXMLDOMNode
Dim xmlLoanToValueNode As IXMLDOMNode
Dim xmlIsPropertyInScotlandNode As IXMLDOMNode
Dim xmlRegisteredTitleNoOneNode As IXMLDOMNode
Dim xmlRegisteredTitleNoTwoNode As IXMLDOMNode
Dim xmlRegisteredTitleNoThreeNode As IXMLDOMNode
Dim xmlTenureNode As IXMLDOMNode

LogDetails 3, "SetupPropertyDetails : Start"

Set xmlDoc = New FreeThreadedDOMDocument40
xmlDoc.validateOnParse = False
xmlDoc.setProperty "NewParser", True

Dim strRuleName As String
strRuleName = "PropertyDetails"

'get latest ValnRepSummary.InstructionSequenceNo
Dim xmlValnRepSummaryNodeList As IXMLDOMNodeList
Dim intMaxSequenceNo, intSequenceNo, intIndex As Integer
Dim strMaxSequenceNo As String
Set xmlValnRepSummaryNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/VALUATION/VALNREPSUMMARY")

'Find the one with the highest INSTRUCTIONSEQUENCENO
For Each xmlSearchNode In xmlValnRepSummaryNodeList
     intSequenceNo = GetAttributeAsInteger(xmlSearchNode, "INSTRUCTIONSEQUENCENO")
     If intSequenceNo > intMaxSequenceNo Then
         intMaxSequenceNo = intSequenceNo
     End If
Next
Set xmlSearchNode = Nothing
strMaxSequenceNo = CStr(intMaxSequenceNo)

'get valnRepPropertyDetails
Dim xmlvalnRepPropertyDetailsNode As IXMLDOMNode
Set xmlvalnRepPropertyDetailsNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUATION/VALNREPPROPERTYDETAILS[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")

'get NewProperty
Dim xmlNewPropertyNode As IXMLDOMNode
Set xmlNewPropertyNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY")

'get ValnRepValuation
Dim xmlValnRepValuationNode As IXMLDOMNode
'make sure, extracting valnRepValuation this way, works
Set xmlValnRepValuationNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUATION/VALNREPSUMMARY/VALNREPVALUATION[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")

'get ValuerInstruction
Dim xmlValuerInstruction As IXMLDOMNode
Set xmlValuerInstruction = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUATION/VALUERINSTRUCTION[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")

'ValuationType
gstrRuleNumber = strRuleName & ".1"
Dim strValuationTypeValueID As String
Set xmlValuationTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValuationType", "")
If Not xmlValuerInstruction Is Nothing Then
    strValuationTypeValueID = GetAttributeText(xmlValuerInstruction, "VALUATIONTYPE")
End If
If Not strValuationTypeValueID = "" Then
    If IsValidationType("ValuationType", CSafeInt(strValuationTypeValueID), "SC1") Or IsValidationType("ValuationType", CSafeInt(strValuationTypeValueID), "SC2") Or IsValidationType("ValuationType", CSafeInt(strValuationTypeValueID), "SC3") Then
        xmlValuationTypeNode.Text = "1"
    Else
        xmlValuationTypeNode.Text = strValuationTypeValueID
    End If
End If

'START: New Code added by JJ 08-Feb-2006
If xmlValuerInstruction Is Nothing Then
    'If Not xmlNewPropertyNode.Attributes("VALUATIONTYPE") Is Nothing Then
        xmlValuationTypeNode.Text = GetAttributeText(xmlNewPropertyNode, "VALUATIONTYPE")
    'End If
End If
'END: New Code added by JJ 08-Feb-2006

'get AddressGUID and Address node
Dim strAddressGUID As String
Dim xmlAddressNode As IXMLDOMNode
If Not xmlNewPropertyNode Is Nothing Then
    Set xmlSearchNode = xmlNewPropertyNode.selectSingleNode("NEWPROPERTYADDRESS")
End If
If Not xmlSearchNode Is Nothing Then
    strAddressGUID = GetAttributeText(xmlSearchNode, "ADDRESSGUID")
End If
If strAddressGUID <> "" Then
    Set xmlAddressNode = xmlSearchNode.selectSingleNode("ADDRESS[@ADDRESSGUID='" & strAddressGUID & "']")
End If

'set all nodes, dependant on address
Set xmlPropertyAddressNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyAddress", "")
Set xmlHouseOrBuildingNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingName", "")
Set xmlFlatNameOrNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FlatNameOrNumber", "")
Set xmlHouseOrBuildingNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingNumber", "")
Set xmlStreetNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Street", "")
Set xmlDistrictNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "District", "")
Set xmlTownOrCityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TownOrCity", "")
Set xmlCountyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "County", "")
Set xmlPostCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PostCode", "")
Set xmlUnexpiredTermNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "UnexpiredTerm", "")
    
If Not xmlAddressNode Is Nothing Then
    'HouseOrBuildingName
    gstrRuleNumber = strRuleName & ".2"
    xmlHouseOrBuildingNameNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENAME")
    
    'FlatNameOrNumber
    gstrRuleNumber = strRuleName & ".3"
    xmlFlatNameOrNumberNode.Text = GetAttributeText(xmlAddressNode, "FLATNUMBER")
    
    'HouseOrBuildingNumber
    gstrRuleNumber = strRuleName & ".4"
    xmlHouseOrBuildingNumberNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENUMBER")
    
    'Street
    gstrRuleNumber = strRuleName & ".5"
    xmlStreetNode.Text = GetAttributeText(xmlAddressNode, "STREET")
    
    'District
    gstrRuleNumber = strRuleName & ".6"
    xmlDistrictNode.Text = GetAttributeText(xmlAddressNode, "DISTRICT")
    
    'TownOrCity
    gstrRuleNumber = strRuleName & ".7"
    xmlTownOrCityNode.Text = GetAttributeText(xmlAddressNode, "TOWN")
    
    'County
    gstrRuleNumber = strRuleName & ".8"
    xmlCountyNode.Text = GetAttributeText(xmlAddressNode, "COUNTY")
    
    'PostCode
    gstrRuleNumber = strRuleName & ".9"
    xmlPostCodeNode.Text = GetAttributeText(xmlAddressNode, "POSTCODE")
   
End If

'UnexpiredTerm
gstrRuleNumber = strRuleName & ".12"
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    xmlUnexpiredTermNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "UNEXPIREDLEASE")
End If
    
xmlPropertyAddressNode.appendChild xmlFlatNameOrNumberNode
xmlPropertyAddressNode.appendChild xmlHouseOrBuildingNameNode
xmlPropertyAddressNode.appendChild xmlHouseOrBuildingNumberNode
xmlPropertyAddressNode.appendChild xmlStreetNode
xmlPropertyAddressNode.appendChild xmlDistrictNode
xmlPropertyAddressNode.appendChild xmlTownOrCityNode
xmlPropertyAddressNode.appendChild xmlCountyNode
xmlPropertyAddressNode.appendChild xmlPostCodeNode

'IsNewProperty
gstrRuleNumber = strRuleName & ".10"
Set xmlIsNewPropertyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IsNewProperty", "")
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    xmlIsNewPropertyNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "NEWPROPERTYINDICATOR")
Else
    If Not xmlNewPropertyNode Is Nothing Then
        xmlIsNewPropertyNode.Text = GetAttributeText(xmlNewPropertyNode, "NEWPROPERTYINDICATOR")
    End If
End If

'TenureOfProperty
gstrRuleNumber = strRuleName & ".11"
Dim strTenureValueID As String
Set xmlTenureOfPropertyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TenureOfProperty", "")
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    strTenureValueID = GetAttributeText(xmlvalnRepPropertyDetailsNode, "TENURE")
End If

If Not strTenureValueID = "" Then
    If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "F") Then 'Freehold
        xmlTenureOfPropertyNode.Text = "F/H"
    Else
        If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "L") Then 'Leasehold
            xmlTenureOfPropertyNode.Text = "L/H"
        Else
            If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "FE") Then 'Feudal
                xmlTenureOfPropertyNode.Text = "Feu"
            Else
            'SC 17/03/2006 UAT Defect 1754: Look for validation type = 'C'
                If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "C") Then 'Commonhold
                    xmlTenureOfPropertyNode.Text = "C/H"
                End If
            End If
        End If
    End If
End If

'START: New Code added by JJ 08-Feb-2006
If strTenureValueID = "" Then
    'If Not xmlNewPropertyNode.Attributes("TENURETYPE") Is Nothing Then
        strTenureValueID = GetAttributeText(xmlNewPropertyNode, "TENURETYPE")
    'End If
    
    If Not strTenureValueID = "" Then
        If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "F") Then 'Freehold
            xmlTenureOfPropertyNode.Text = "F/H"
        Else
            If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "L") Then 'Leasehold
                xmlTenureOfPropertyNode.Text = "L/H"
            Else
                If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "FE") Then 'Feudal
                    xmlTenureOfPropertyNode.Text = "Feu"
                Else
                    'SC 17/03/2006 UAT Defect 1754: Look for validation type = 'C'
                    If IsValidationType("PropertyTenure", CSafeInt(strTenureValueID), "C") Then 'Commonhold
                        xmlTenureOfPropertyNode.Text = "C/H"
                    End If
                End If
            End If
        End If
    End If
End If
'END: New Code added by JJ 08-Feb-2006

'MAR694
Set xmlTenureNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Tenure", "")
xmlTenureNode.Text = xmlTenureOfPropertyNode.Text


'ExLocalAuthority
gstrRuleNumber = strRuleName & ".14"
Set xmlExLocalAuthorityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ExLocalAuthority", "")
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    xmlExLocalAuthorityNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "EXLOCALAUTHORITY")
End If

'get ValnReppropertyServices
Dim xmlValnRepPropertyServicesNode As IXMLDOMNode
Set xmlValnRepPropertyServicesNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUATION/VALNREPSUMMARY/VALNREPPROPERTYSERVICES[@INSTRUCTIONSEQUENCENO ='" & strMaxSequenceNo & "']")

'NumberOfBedRooms
gstrRuleNumber = strRuleName & ".16"
Set xmlNumberOfBedRoomsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfBedRooms", "")
Dim strRoomType As String
If Not xmlValnRepPropertyServicesNode Is Nothing Then
    xmlNumberOfBedRoomsNode.Text = GetAttributeText(xmlValnRepPropertyServicesNode, "NUMBEROFBEDROOMS")
End If

If xmlNumberOfBedRoomsNode.Text = "" Then 'no data in ValnReppropertyServices
    If Not xmlNewPropertyNode Is Nothing Then
        Set xmlSearchNode = xmlNewPropertyNode.selectSingleNode("NEWPROPERTYROOMTYPE")
    End If
    If Not xmlSearchNode Is Nothing Then
        strRoomType = GetAttributeText(xmlSearchNode, "ROOMTYPE")
        If strRoomType <> "" Then
            If IsValidationType("NewPropertyRoomType", CSafeInt(strRoomType), "BD") Then
                xmlNumberOfBedRoomsNode.Text = GetAttributeText(xmlSearchNode, "NUMBEROFROOMS")
            End If
        End If
    End If
End If

'create all elements which r based on ValnRepPropertyServicesNode
Set xmlNumberOfLivingRoomNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfLivingRoom", "") 'NumberOfLivingRoom - thats how it's spelt in Commomtypes.xsd
Set xmlNumberOfKitchensNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfKitchens", "")
Set xmlNumberOfBathroomsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfBathrooms", "")
Set xmlNumberOfGaragesNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfGarages", "")
Set xmlNumberOfParkingSpacesNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NumberOfParkingSpaces", "")
                            
If Not xmlValnRepPropertyServicesNode Is Nothing Then
    
    gstrRuleNumber = strRuleName & ".15"
    xmlNumberOfLivingRoomNode.Text = GetAttributeText(xmlValnRepPropertyServicesNode, "LIVINGROOMS")
        
    'NumberOfBathrooms
    gstrRuleNumber = strRuleName & ".18"
    xmlNumberOfBathroomsNode.Text = GetAttributeText(xmlValnRepPropertyServicesNode, "BATHROOMS")
    
    'NumberOfGarages
    gstrRuleNumber = strRuleName & ".19"
    xmlNumberOfGaragesNode.Text = GetAttributeText(xmlValnRepPropertyServicesNode, "GARAGES")
    
    'NumberOfParkingSpaces
    gstrRuleNumber = strRuleName & ".20"
    xmlNumberOfParkingSpacesNode.Text = GetAttributeText(xmlValnRepPropertyServicesNode, "PARKINGSPACES")
    
End If

If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    'NumberOfKitchens
    gstrRuleNumber = strRuleName & ".17"
    xmlNumberOfKitchensNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "NUMBEROFKITCHENS")
End If
    
'ApproximateYearPropertyBuild
gstrRuleNumber = strRuleName & ".21"
Set xmlApproximateYearPropertyBuildNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ApproximateYearPropertyBuild", "")
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    xmlApproximateYearPropertyBuildNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "YEARBUILT")
End If

'BuildingStandardIndemnityType
gstrRuleNumber = strRuleName & ".22"
Set xmlBuildingStandardIndemnityTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BuildingStandardIndemnityType", "")
Dim intCertificationType As Integer
Dim strcertificationType As String, strHouseBuldersGuarantee As String

If Not xmlValnRepValuationNode Is Nothing Then
    strcertificationType = GetAttributeText(xmlValnRepValuationNode, "CERTIFICATIONTYPE")
End If

If Not strcertificationType = "" Then 'valuation report exists
    intCertificationType = CSafeInt(strcertificationType)
    If IsValidationType("BuildingsCertificationType", intCertificationType, "N") Then
        xmlBuildingStandardIndemnityTypeNode.Text = "1" 'NHBC
    Else
        If IsValidationType("BuildingsCertificationType", intCertificationType, "Z") Then
            xmlBuildingStandardIndemnityTypeNode.Text = "2" 'Zurich
        Else
            If IsValidationType("BuildingsCertificationType", intCertificationType, "P") Then
                xmlBuildingStandardIndemnityTypeNode.Text = "3" 'premier gurantee
            End If
        End If
    End If
 Else
    Dim intHouseBuildersGuarantee As Integer
    If Not xmlNewPropertyNode Is Nothing Then
        strHouseBuldersGuarantee = GetAttributeText(xmlNewPropertyNode, "HOUSEBUILDERSGUARANTEE")
    End If
    
    If strHouseBuldersGuarantee <> "" Then
        intHouseBuildersGuarantee = CSafeInt(strHouseBuldersGuarantee)
        If IsValidationType("BuildingsCertificationType", intHouseBuildersGuarantee, "N") Then
            xmlBuildingStandardIndemnityTypeNode.Text = "1" 'NHBC
        Else
            If IsValidationType("BuildingsCertificationType", intHouseBuildersGuarantee, "Z") Then
                xmlBuildingStandardIndemnityTypeNode.Text = "2" 'Zurich
            Else
                If IsValidationType("BuildingsCertificationType", intHouseBuildersGuarantee, "P") Then
                    xmlBuildingStandardIndemnityTypeNode.Text = "3" 'premier gurantee
                End If
            End If
        End If
    End If
 End If
  
'PropertyType
gstrRuleNumber = strRuleName & ".13"
Set xmlPropertyTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyType", "")
Dim strPropertyType As String, strPropertyDesc As String
Dim intPropertyType As Integer, intPropertyDesc As Integer
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then 'check ValnRepPropertyDetails
    strPropertyType = GetAttributeText(xmlvalnRepPropertyDetailsNode, "TYPEOFPROPERTY")
    strPropertyDesc = GetAttributeText(xmlvalnRepPropertyDetailsNode, "PROPERTYDESCRIPTION")
Else       'Use NewProperty
    If Not xmlNewPropertyNode Is Nothing Then
        strPropertyType = GetAttributeText(xmlNewPropertyNode, "TYPEOFPROPERTY")
        strPropertyDesc = GetAttributeText(xmlNewPropertyNode, "DESCRIPTIONOFPROPERTY")
    End If
End If

If Not strPropertyType = "" Then
    intPropertyType = CSafeInt(strPropertyType)
End If

If Not strPropertyDesc = "" Then
    intPropertyDesc = CSafeInt(strPropertyDesc)
End If

If (strPropertyType <> "") Then
    
    If IsValidationType("PropertyType", intPropertyType, "H") Then
        If strPropertyDesc <> "" Then
            If IsValidationType("PropertyDescription", intPropertyDesc, "HD") Then
                xmlPropertyTypeNode.Text = "3" 'detatched house
            Else
                If IsValidationType("PropertyDescription", intPropertyDesc, "HS") Then
                    xmlPropertyTypeNode.Text = "2" 'semidetatched house
                Else
                    If IsValidationType("PropertyDescription", intPropertyDesc, "HT") Then
                        xmlPropertyTypeNode.Text = "1" 'house terrace
                    End If
                End If
            End If
        End If
    Else
        If IsValidationType("PropertyType", intPropertyType, "B") Then
            xmlPropertyTypeNode.Text = "14" 'Bungalow
        Else
            If IsValidationType("PropertyType", intPropertyType, "F") Or IsValidationType("PropertyType", CSafeInt(strPropertyType), "M") Then
                xmlPropertyTypeNode.Text = "4" 'Flat or Massionette
            Else
                If IsValidationType("PropertyType", intPropertyType, "O") Then
                    xmlPropertyTypeNode.Text = "99" 'other
                End If
            End If
        End If
    End If
End If

'GrossExternalFloorArea
gstrRuleNumber = strRuleName & ".23"
Set xmlGrossExternalFloorAreaNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "GrossExternalFloorArea", "")
Dim strStructure As String, strArea As String
If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
    strArea = GetAttributeText(xmlvalnRepPropertyDetailsNode, "RESIDENCEAREA")
    'strArea = "100020"
    'This is a temporary work around to restrict the area to 5 digits max
    If strArea <> "" Then
        If CSafeLng(strArea) <= 99999 Then
            xmlGrossExternalFloorAreaNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "RESIDENCEAREA")
        Else
            xmlGrossExternalFloorAreaNode.Text = "99999"
        End If
    End If
    
    strStructure = GetAttributeText(xmlvalnRepPropertyDetailsNode, "STRUCTURE")
End If

'ConstructionTypeIndicator
gstrRuleNumber = strRuleName & ".24"
Set xmlConstructionTypeIndicatorNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ConstructionTypeIndicator", "")
Set xmlConstructionTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ConstructionType", "")
If Not strStructure = "" Then
    If IsValidationType("BuildingConstruction", CSafeInt(strStructure), "TC") Then
        xmlConstructionTypeIndicatorNode.Text = "Y" 'traditional
    Else
        xmlConstructionTypeIndicatorNode.Text = "N" 'NonTraditional
    End If
    
    'ConstructionType
    gstrRuleNumber = strRuleName & ".25"
    If IsValidationType("BuildingConstruction", CSafeInt(strStructure), "NTC") Then
        xmlConstructionTypeNode.Text = GetAttributeText(xmlvalnRepPropertyDetailsNode, "CONSTRUCTIONTYPEDETAILS")
    End If

End If

Dim strSaleability As String, strOverallCondition As String, strPresentValuation As String
If Not xmlValnRepValuationNode Is Nothing Then
    strSaleability = GetAttributeText(xmlValnRepValuationNode, "SALEABILITY")
    strOverallCondition = GetAttributeText(xmlValnRepValuationNode, "OVERALLCONDITION")
    strPresentValuation = GetAttributeText(xmlValnRepValuationNode, "PRESENTVALUATION")
End If

'Saleability
gstrRuleNumber = strRuleName & ".26"
Set xmlSaleabilityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Saleability", "")
If strSaleability <> "" Then
    xmlSaleabilityNode.Text = GetComboText("ValuationSaleability", CSafeInt(strSaleability))
End If

'OverallCondition
gstrRuleNumber = strRuleName & ".27"
Set xmlOverallConditionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "OverallCondition", "")
If strOverallCondition <> "" Then
    xmlOverallConditionNode.Text = GetComboText("ValuationOverAllCondition", CSafeInt(strOverallCondition))
End If

'LastValuation
gstrRuleNumber = strRuleName & ".28"
Set xmlLastValuationNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LastValuation", "")
Dim intTypeOfMortgageValueID As Integer
Dim lngPurchasePriceorEstimatedValue As Long
Dim lngPresentationValuation As Long
Dim lngHomeTrackValuationAmount As Long
Dim strPurchasePriceorEstimatedValue As String, strValuationAmount As String
If strPresentValuation <> "" Then
    lngPresentationValuation = CSafeLng(strPresentValuation)
End If

strPurchasePriceorEstimatedValue = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "PURCHASEPRICEORESTIMATEDVALUE")
If strPurchasePriceorEstimatedValue <> "" Then
    lngPurchasePriceorEstimatedValue = CSafeLng(GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "PURCHASEPRICEORESTIMATEDVALUE"))
End If

intTypeOfMortgageValueID = CSafeInt(GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "TYPEOFAPPLICATION"))
If IsValidationType("TypeOfMortgage", intTypeOfMortgageValueID, "FT") Or IsValidationType("TypeOfMortgage", intTypeOfMortgageValueID, "HM") Then
    If Not xmlValnRepValuationNode Is Nothing Then
        xmlLastValuationNode.Text = GetAttributeText(xmlValnRepValuationNode, "POSTWORKSVALUATION")
    End If
    If xmlLastValuationNode.Text = "" Then 'no postworks
        If lngPresentationValuation > lngPurchasePriceorEstimatedValue Then
            xmlLastValuationNode.Text = CStr(lngPurchasePriceorEstimatedValue)
        Else
            xmlLastValuationNode.Text = CStr(lngPresentationValuation)
        End If
    End If
Else
    If IsValidationType("TypeOfMortgage", intTypeOfMortgageValueID, "R") Then
        Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/HOMETRACKVALUATIONDETAILS")
        If Not xmlSearchNode Is Nothing And xmlValnRepValuationNode Is Nothing Then 'Hometrack valuation but no ESurv valuation
            strValuationAmount = GetAttributeText(xmlSearchNode, "VALUATIONRESULTSVALUATIONAMOUNT")
            If strValuationAmount <> "" Then
                lngHomeTrackValuationAmount = CSafeLng(strValuationAmount)
            End If
            'SC 06/03/2006: UAT Defect 586 Set to the lower of Purchase Price or Estimated Value and Home Track valuation results
            If lngPurchasePriceorEstimatedValue > lngHomeTrackValuationAmount Then
                xmlLastValuationNode.Text = CStr(lngHomeTrackValuationAmount)
            Else
                xmlLastValuationNode.Text = CStr(lngPurchasePriceorEstimatedValue)
            End If
        Else 'have done an ESurv valuation for this remortgage
            xmlLastValuationNode.Text = CStr(lngPresentationValuation)
        End If
    End If
End If

'BuildingInsuranceReinstatementAmount
gstrRuleNumber = strRuleName & ".29"
Set xmlBuildingInsuranceReinstallmentAmountNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "BuildingInsuranceReinstallmentAmount", "")
If Not xmlValnRepValuationNode Is Nothing Then
    xmlBuildingInsuranceReinstallmentAmountNode.Text = GetAttributeText(xmlValnRepValuationNode, "REINSTATEMENTVALUE")
End If

'ValuerName
gstrRuleNumber = strRuleName & ".30"
Set xmlValuerNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValuerName", "")
If Not xmlValuerInstruction Is Nothing Then
    xmlValuerNameNode.Text = GetAttributeText(xmlValuerInstruction, "VALUERNAME")
End If

'LastValuationDate
gstrRuleNumber = strRuleName & ".31"
Set xmlLastValuationDateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LastValuationDate", "")
Dim strAppointmentDate As String
If Not xmlValuerInstruction Is Nothing Then
    strAppointmentDate = GetAttributeText(xmlValuerInstruction, "APPOINTMENTDATE")
Else
    Dim xmlHomeTrackaluationDetailsNode As IXMLDOMNode
    Set xmlHomeTrackaluationDetailsNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/HOMETRACKVALUATIONDETAILS")
    If Not xmlHomeTrackaluationDetailsNode Is Nothing Then
        strAppointmentDate = GetAttributeText(xmlHomeTrackaluationDetailsNode, "VALUATIONRESULTSDATERECEIVED")
    End If
End If

If strAppointmentDate <> "" Then
    xmlLastValuationDateNode.Text = Format(strAppointmentDate, "yyyy/mm/dd")
End If


'PropertyPurchase
gstrRuleNumber = strRuleName & ".32"
Set xmlPropertyPurchaseNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyPurchasePrice", "")
xmlPropertyPurchaseNode.Text = strPurchasePriceorEstimatedValue

'LoanToValue
gstrRuleNumber = strRuleName & ".33"
Dim strLoanToValue As String
Set xmlLoanToValueNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LoanToValue", "")
strLoanToValue = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER = '" & _
                                GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION[@QUOTATIONNUMBER='" & _
                                    GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "ACCEPTEDQUOTENUMBER") & "']"), "MORTGAGESUBQUOTENUMBER") & _
                                "']"), "LTV")
'strLoanToValue = "50.5"
'rounding needed, if it's a fraction
If InStr(strLoanToValue, ".") > 0 Then
    If (CSafeInt(Mid(strLoanToValue, InStr(strLoanToValue, ".") + 1, 1)) >= 5) Then
        xmlLoanToValueNode.Text = CStr(CSafeInt(Mid(strLoanToValue, 1, InStr(strLoanToValue, ".") - 1)) + 1)
    Else
        xmlLoanToValueNode.Text = CStr(CSafeInt(Mid(strLoanToValue, 1, InStr(strLoanToValue, ".") - 1)))
    End If
Else
    xmlLoanToValueNode.Text = strLoanToValue
End If

'IsPropertyInScotland
gstrRuleNumber = strRuleName & ".34"
Dim strPropertyLocationValueID As String
Set xmlIsPropertyInScotlandNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IsPropertyInScotland", "")
If Not xmlNewPropertyNode Is Nothing Then
    strPropertyLocationValueID = GetAttributeText(xmlNewPropertyNode, "PROPERTYLOCATION")
End If
If strPropertyLocationValueID <> "" Then
    If IsValidationType("PropertyLocation", CSafeInt(strPropertyLocationValueID), "S") Then
        xmlIsPropertyInScotlandNode.Text = "1"
    Else
        xmlIsPropertyInScotlandNode.Text = "0"
    End If
End If

Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE")
Set xmlRegisteredTitleNoOneNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RegisteredTitleNoOne", "")
Set xmlRegisteredTitleNoTwoNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RegisteredTitleNoTwo", "")
Set xmlRegisteredTitleNoThreeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "RegisteredTitleNoThree", "")

If Not xmlSearchNode Is Nothing Then
    'RegisteredTitleNoOne
    gstrRuleNumber = strRuleName & ".35"
    If GetAttributeAsBoolean(xmlSearchNode, "UNREGISTEREDRECORDED") Then
        xmlRegisteredTitleNoOneNode.Text = "UNREG"
    Else
        xmlRegisteredTitleNoOneNode.Text = GetAttributeText(xmlSearchNode, "TITLENUMBER")
    End If
    
    'RegisteredTitleNoTwo
    gstrRuleNumber = strRuleName & ".36"
    xmlRegisteredTitleNoTwoNode.Text = GetAttributeText(xmlSearchNode, "TITLENUMBER2")
    
    'RegisteredTitleNoThree
    gstrRuleNumber = strRuleName & ".37"
    xmlRegisteredTitleNoThreeNode.Text = GetAttributeText(xmlSearchNode, "TITLENUMBER3")
End If

'add the created nodes to the response
vxmlResponseNode.appendChild xmlNumberOfBedRoomsNode
vxmlResponseNode.appendChild xmlLastValuationNode
vxmlResponseNode.appendChild xmlLastValuationDateNode
vxmlResponseNode.appendChild xmlPropertyAddressNode
vxmlResponseNode.appendChild xmlTenureOfPropertyNode
vxmlResponseNode.appendChild xmlIsPropertyInScotlandNode
vxmlResponseNode.appendChild xmlPropertyTypeNode
vxmlResponseNode.appendChild xmlGrossExternalFloorAreaNode
vxmlResponseNode.appendChild xmlUnexpiredTermNode
vxmlResponseNode.appendChild xmlExLocalAuthorityNode
vxmlResponseNode.appendChild xmlApproximateYearPropertyBuildNode
vxmlResponseNode.appendChild xmlConstructionTypeIndicatorNode
vxmlResponseNode.appendChild xmlConstructionTypeNode
vxmlResponseNode.appendChild xmlTenureNode
vxmlResponseNode.appendChild xmlRegisteredTitleNoOneNode
vxmlResponseNode.appendChild xmlRegisteredTitleNoTwoNode
vxmlResponseNode.appendChild xmlRegisteredTitleNoThreeNode
vxmlResponseNode.appendChild xmlIsNewPropertyNode
vxmlResponseNode.appendChild xmlBuildingStandardIndemnityTypeNode
vxmlResponseNode.appendChild xmlValuationTypeNode
vxmlResponseNode.appendChild xmlNumberOfLivingRoomNode
vxmlResponseNode.appendChild xmlNumberOfKitchensNode
vxmlResponseNode.appendChild xmlNumberOfBathroomsNode
vxmlResponseNode.appendChild xmlNumberOfGaragesNode
vxmlResponseNode.appendChild xmlNumberOfParkingSpacesNode
vxmlResponseNode.appendChild xmlSaleabilityNode
vxmlResponseNode.appendChild xmlOverallConditionNode
vxmlResponseNode.appendChild xmlValuerNameNode
vxmlResponseNode.appendChild xmlBuildingInsuranceReinstallmentAmountNode
vxmlResponseNode.appendChild xmlPropertyPurchaseNode
vxmlResponseNode.appendChild xmlLoanToValueNode

LogDetails 3, "SetupPropertyDetails : End"

SetupPropertyDetailsExit:
    
    Set xmlDoc = Nothing
    Set xmlSearchNode = Nothing
    Set xmlValuationTypeNode = Nothing
    Set xmlHouseOrBuildingNameNode = Nothing
    Set xmlFlatNameOrNumberNode = Nothing
    Set xmlHouseOrBuildingNumberNode = Nothing
    Set xmlStreetNode = Nothing
    Set xmlDistrictNode = Nothing
    Set xmlTownOrCityNode = Nothing
    Set xmlCountyNode = Nothing
    Set xmlPostCodeNode = Nothing
    Set xmlIsNewPropertyNode = Nothing
    Set xmlTenureOfPropertyNode = Nothing
    Set xmlUnexpiredTermNode = Nothing
    Set xmlPropertyTypeNode = Nothing
    Set xmlExLocalAuthorityNode = Nothing
    Set xmlNumberOfLivingRoomNode = Nothing
    Set xmlNumberOfBedRoomsNode = Nothing
    Set xmlNumberOfKitchensNode = Nothing
    Set xmlNumberOfBathroomsNode = Nothing
    Set xmlNumberOfGaragesNode = Nothing
    Set xmlNumberOfParkingSpacesNode = Nothing
    Set xmlApproximateYearPropertyBuildNode = Nothing
    Set xmlBuildingStandardIndemnityTypeNode = Nothing
    Set xmlGrossExternalFloorAreaNode = Nothing
    Set xmlConstructionTypeIndicatorNode = Nothing
    Set xmlConstructionTypeNode = Nothing
    Set xmlSaleabilityNode = Nothing
    Set xmlOverallConditionNode = Nothing
    Set xmlLastValuationNode = Nothing
    Set xmlBuildingInsuranceReinstallmentAmountNode = Nothing
    Set xmlValuerNameNode = Nothing
    Set xmlLastValuationDateNode = Nothing
    Set xmlPropertyPurchaseNode = Nothing
    Set xmlLoanToValueNode = Nothing
    Set xmlIsPropertyInScotlandNode = Nothing
    Set xmlRegisteredTitleNoOneNode = Nothing
    Set xmlRegisteredTitleNoTwoNode = Nothing
    Set xmlRegisteredTitleNoThreeNode = Nothing
       
    
    CheckError strFunctionName, TypeName(Me)
    
End Sub



Private Sub SetupPropertyDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupPropertyDetails
' Author:       Sam Dali
' Date:         25/08/05
' Purpose:      Create Property Details in Completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to
'------------------------------------------------------------------------

    On Error GoTo SetupPropertyDetailsExit

    Const strFunctionName As String = "SetupPropertyDetails"

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlSearchNode As IXMLDOMNode 'temp node used for searching
    Dim strRuleName As String
    Dim xmlPropertyDetailsNode As IXMLDOMNode
    
    Dim xmlValnInstructionNodeList As IXMLDOMNodeList
    Dim intMaxSequenceNo As Integer
    Dim intSequenceNo As Integer
    Dim intIndex As Integer
    Dim strMaxSequenceNo As String
    Dim xmlValuerInstruction As IXMLDOMNode
    Dim xmlvalnRepPropertyDetailsNode As IXMLDOMNode
    Dim xmlValnRepValuationNode As IXMLDOMNode
    
    Dim xmlPropertyAddressNode As IXMLDOMNode
    Dim xmlNewPropertyNode As IXMLDOMNode
    Dim strAddressGUID As String
    Dim xmlAddressNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNameNode As IXMLDOMNode
    Dim xmlFlatNameOrNumberNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNumberNode As IXMLDOMNode
    Dim xmlStreetNode As IXMLDOMNode
    Dim xmlDistrictNode As IXMLDOMNode
    Dim xmlTownOrCityNode As IXMLDOMNode
    Dim xmlCountyNode As IXMLDOMNode
    Dim xmlPostCodeNode As IXMLDOMNode
    
    Dim xmlPropertyDetailsDescriptionNode As IXMLDOMNode
    Dim xmlTitleNoOneNode As IXMLDOMNode
    Dim xmlTitleNoTwoNode As IXMLDOMNode
    Dim xmlTitleNoThreeNode As IXMLDOMNode
    Dim strPropertyTenureValueID As String
    Dim xmlPropertyTenureNode As IXMLDOMNode
    Dim strPropertyTypeValueID As String
    Dim xmlPropertyTypeNode As IXMLDOMNode
    Dim xmlPropertySubTypeNode As IXMLDOMNode
    Dim strPropertyDescriptionValueID As String
    Dim xmlPropertyDescriptionNode As IXMLDOMNode
    Dim xmlFreeholdTitleNode As IXMLDOMNode
    Dim xmlPropertyLetNode As IXMLDOMNode
    Dim strLineOfBusiness As String
    
    Dim xmlPropertyValuationDetailsNode As IXMLDOMNode
    Dim xmlLastValuationDateNode As IXMLDOMNode
    Dim strAppointmentDate As String
    Dim xmlPropertyValueNode As IXMLDOMNode
    Dim strPropertyValue As String
    Dim xmlTotalValuationNode As IXMLDOMNode
    Dim strTotalValuation As String
    Dim xmlValuationTypeNode As IXMLDOMNode
    Dim strValuationTypeValueID As String
    Dim xmlPropertyPurchaseNode As IXMLDOMNode
    Dim strPurchasePriceorEstimatedValue As String
    Dim xmlServicingBranchNode As IXMLDOMNode
    Dim xmlUnderwriterNode As IXMLDOMNode
    
    Dim xmlMiscSecurityDataNode As IXMLDOMNode
    Dim xmlPropertyPurchaseDateNode As IXMLDOMNode
    Dim xmlLoanToValueNode As IXMLDOMNode
    Dim strLoanToValue As String
    'EP2_956
    Dim strUKPostCode As String
    Dim strSearchPat As String
    Dim xmlLendingBranchComboValueNode As IXMLDOMNode
    
    
    
    LogDetails 3, "SetupPropertyDetails : Start"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strRuleName = "PropertyDetails"
    gstrRuleNumber = strRuleName
    
    Set xmlPropertyDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyDetails", "")
    
    
    
    '** Get latest ValuerInstruction.InstructionSequenceNo : START **
    
    Set xmlValnInstructionNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION")
    
    'Find the one with the highest INSTRUCTIONSEQUENCENO
    For Each xmlSearchNode In xmlValnInstructionNodeList
         intSequenceNo = GetAttributeAsInteger(xmlSearchNode, "INSTRUCTIONSEQUENCENO")
         If intSequenceNo > intMaxSequenceNo Then
             intMaxSequenceNo = intSequenceNo
         End If
    Next
    Set xmlSearchNode = Nothing
    strMaxSequenceNo = CStr(intMaxSequenceNo)
    m_strMaxValInstructionSeqNo = strMaxSequenceNo
    
    'get ValuerInstruction
    Set xmlValuerInstruction = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")
    
    'get valnRepPropertyDetails
    Set xmlvalnRepPropertyDetailsNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION/VALNREPPROPERTYDETAILS[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")
    
    'get ValnRepValuation
    Set xmlValnRepValuationNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/VALUERINSTRUCTION/VALNREPVALUATION[@INSTRUCTIONSEQUENCENO='" & strMaxSequenceNo & "']")
    '** Get latest ValnRepSummary.InstructionSequenceNo : END **
    
    
    
    
    
    
    
    '** Get Security Property Address : START **
    Set xmlNewPropertyNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY")
    
    
    'get AddressGUID and Address node
    If Not xmlNewPropertyNode Is Nothing Then
        Set xmlSearchNode = xmlNewPropertyNode.selectSingleNode("NEWPROPERTYADDRESS")
    End If
    If Not xmlSearchNode Is Nothing Then
        strAddressGUID = GetAttributeText(xmlSearchNode, "ADDRESSGUID")
    End If
    If strAddressGUID <> "" Then
        Set xmlAddressNode = xmlSearchNode.selectSingleNode("ADDRESS[@ADDRESSGUID='" & strAddressGUID & "']")
    End If
    
    
    'set all nodes, dependant on address
    Set xmlPropertyAddressNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyAddress", "")
    Set xmlHouseOrBuildingNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingName", "")
    Set xmlFlatNameOrNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FlatNameOrNumber", "")
    Set xmlHouseOrBuildingNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingNumber", "")
    Set xmlStreetNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Street", "")
    Set xmlDistrictNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "District", "")
    Set xmlTownOrCityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TownOrCity", "")
    Set xmlCountyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "County", "")
    Set xmlPostCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PostCode", "")
        
    If Not xmlAddressNode Is Nothing Then
        'HouseOrBuildingName
        gstrRuleNumber = strRuleName & ".1"
        xmlHouseOrBuildingNameNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENAME")
        
        'FlatNameOrNumber
        gstrRuleNumber = strRuleName & ".2"
        xmlFlatNameOrNumberNode.Text = GetAttributeText(xmlAddressNode, "FLATNUMBER")
        
        'HouseOrBuildingNumber
        gstrRuleNumber = strRuleName & ".3"
        xmlHouseOrBuildingNumberNode.Text = GetAttributeText(xmlAddressNode, "BUILDINGORHOUSENUMBER")
        
        'Street
        gstrRuleNumber = strRuleName & ".4"
        xmlStreetNode.Text = GetAttributeText(xmlAddressNode, "STREET")
        
        'District
        gstrRuleNumber = strRuleName & ".5"
        xmlDistrictNode.Text = GetAttributeText(xmlAddressNode, "DISTRICT")
        
        'TownOrCity
        gstrRuleNumber = strRuleName & ".6"
        xmlTownOrCityNode.Text = GetAttributeText(xmlAddressNode, "TOWN")
        
        'County
        gstrRuleNumber = strRuleName & ".7"
        xmlCountyNode.Text = GetAttributeText(xmlAddressNode, "COUNTY")
        
        'PostCode
        gstrRuleNumber = strRuleName & ".8"
        xmlPostCodeNode.Text = GetAttributeText(xmlAddressNode, "POSTCODE")
       
    End If
    xmlPropertyAddressNode.appendChild xmlFlatNameOrNumberNode
    xmlPropertyAddressNode.appendChild xmlHouseOrBuildingNameNode
    xmlPropertyAddressNode.appendChild xmlHouseOrBuildingNumberNode
    xmlPropertyAddressNode.appendChild xmlStreetNode
    xmlPropertyAddressNode.appendChild xmlDistrictNode
    xmlPropertyAddressNode.appendChild xmlTownOrCityNode
    xmlPropertyAddressNode.appendChild xmlCountyNode
    xmlPropertyAddressNode.appendChild xmlPostCodeNode
    
    xmlPropertyDetailsNode.appendChild xmlPropertyAddressNode
    '** Get Security Property Address : END **
    
    
    
    
    
    '** Get Property Description : START **
    Set xmlPropertyDetailsDescriptionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyDetailsDescription", "")
    
    
    'Get Title Numbers
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE")
    
    Set xmlTitleNoOneNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TitleNoOne", "")
    Set xmlTitleNoTwoNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TitleNoTwo", "")
    Set xmlTitleNoThreeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TitleNoThree", "")
    
    If Not xmlSearchNode Is Nothing Then
        'TitleNoOne
        gstrRuleNumber = strRuleName & ".9"
        xmlTitleNoOneNode.Text = GetAttributeText(xmlSearchNode, "TITLENUMBER")
        
        'TitleNoTwo
        gstrRuleNumber = strRuleName & ".10"
        xmlTitleNoTwoNode.Text = GetAttributeText(xmlSearchNode, "TITLENUMBER2")
        
        'TitleNoThree
        gstrRuleNumber = strRuleName & ".11"
        xmlTitleNoThreeNode.Text = GetAttributeText(xmlSearchNode, "TITLENUMBER3")
    End If
    xmlPropertyDetailsDescriptionNode.appendChild xmlTitleNoOneNode
    xmlPropertyDetailsDescriptionNode.appendChild xmlTitleNoTwoNode
    xmlPropertyDetailsDescriptionNode.appendChild xmlTitleNoThreeNode
    
    
    
    'PropertyTenure
    gstrRuleNumber = strRuleName & ".12"
    Set xmlPropertyTenureNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyTenure", "")
    
    If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
        strPropertyTenureValueID = GetAttributeText(xmlvalnRepPropertyDetailsNode, "TENURE")
    End If
    
    If Not strPropertyTenureValueID = "" Then
        'Use VALNREPVALUATION
        xmlPropertyTenureNode.Text = GetComboText("PropertyTenure", CSafeInt(strPropertyTenureValueID))
    Else
        'Use NEWPROPERTY
        strPropertyTenureValueID = GetAttributeText(xmlNewPropertyNode, "TENURETYPE")
        If Not strPropertyTenureValueID = "" Then
            xmlPropertyTenureNode.Text = GetComboText("PropertyTenure", CSafeInt(strPropertyTenureValueID))
        End If
    End If
    xmlPropertyDetailsDescriptionNode.appendChild xmlPropertyTenureNode
    
    
    
    'PropertyType
    gstrRuleNumber = strRuleName & ".13"
    Set xmlPropertyTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyType", "")
    
    If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
        strPropertyTypeValueID = GetAttributeText(xmlvalnRepPropertyDetailsNode, "TYPEOFPROPERTY")
    End If
    
    If Not strPropertyTypeValueID = "" Then
        'Use VALNREPVALUATION
        xmlPropertyTypeNode.Text = GetComboText("PropertyType", CSafeInt(strPropertyTypeValueID))
    Else
        'Use NEWPROPERTY
        strPropertyTypeValueID = GetAttributeText(xmlNewPropertyNode, "TYPEOFPROPERTY")
        If Not strPropertyTypeValueID = "" Then
            xmlPropertyTypeNode.Text = GetComboText("PropertyType", CSafeInt(strPropertyTypeValueID))
        End If
    End If
    xmlPropertyDetailsDescriptionNode.appendChild xmlPropertyTypeNode
           
    'PropertyDescription
    gstrRuleNumber = strRuleName & ".15"
    Set xmlPropertyDescriptionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyDescription", "")
    
    If Not xmlvalnRepPropertyDetailsNode Is Nothing Then
        strPropertyDescriptionValueID = GetAttributeText(xmlvalnRepPropertyDetailsNode, "PROPERTYDESCRIPTION")
    End If
    
    If Not strPropertyDescriptionValueID = "" Then
        'Use VALNREPVALUATION
        xmlPropertyDescriptionNode.Text = GetComboText("PropertyDescription", CSafeInt(strPropertyDescriptionValueID))
    Else
        'Use NEWPROPERTY
        strPropertyDescriptionValueID = GetAttributeText(xmlNewPropertyNode, "DESCRIPTIONOFPROPERTY")
        If Not strPropertyDescriptionValueID = "" Then
            xmlPropertyDescriptionNode.Text = GetComboText("PropertyDescription", CSafeInt(strPropertyDescriptionValueID))
        End If
    End If
    'xmlPropertyDetailsDescriptionNode.appendChild xmlPropertyDescriptionNode
        
        
    'EP694 - 07/06/2006 - Peter Edney
    'PropertySubType
    gstrRuleNumber = strRuleName & ".14"
    Set xmlPropertySubTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertySubType", "")
    'DRC    07/02/2007  EP2_956 - Added in Maisonette and Change of main description to "Other"
    If IsValidationType("PropertyType", strPropertyTypeValueID, "F") Then
        xmlPropertySubTypeNode.Text = xmlPropertyDescriptionNode.Text
        xmlPropertyDescriptionNode.Text = "Other"
    ElseIf IsValidationType("PropertyType", strPropertyTypeValueID, "M") Then
        xmlPropertySubTypeNode.Text = xmlPropertyDescriptionNode.Text
        xmlPropertyDescriptionNode.Text = "Other"
    Else
        xmlPropertySubTypeNode.Text = "Not applicable"
    End If
    
    xmlPropertyDetailsDescriptionNode.appendChild xmlPropertySubTypeNode
    xmlPropertyDetailsDescriptionNode.appendChild xmlPropertyDescriptionNode

    
    'Freehold Title indicator
    gstrRuleNumber = strRuleName & ".16"
    Set xmlFreeholdTitleNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FreeholdTitle", "")
    'If the PropertyTenure validation type is either "F" = Freehold
    'Use strPropertyTenureValueID from PropertyTenure above
    If IsValidationType("PropertyTenure", CSafeInt(strPropertyTenureValueID), "F") Then
        xmlFreeholdTitleNode.Text = "Y"
    Else
        xmlFreeholdTitleNode.Text = "N"
    End If
    xmlPropertyDetailsDescriptionNode.appendChild xmlFreeholdTitleNode
    
    
    
    'Property Let indicator
    gstrRuleNumber = strRuleName & ".17"
    Set xmlPropertyLetNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyLet", "")
        
    strLineOfBusiness = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "NATUREOFLOAN")
    If Not strLineOfBusiness = "" Then
        'check if the Nature of loan is a Buy To Let
        If IsValidationType("NatureOfLoan", CSafeInt(strLineOfBusiness), "BI") Or _
            IsValidationType("NatureOfLoan", CSafeInt(strLineOfBusiness), "BR") Then
            xmlPropertyLetNode.Text = "Y"
        Else
            xmlPropertyLetNode.Text = "N"
        End If
    End If
    xmlPropertyDetailsDescriptionNode.appendChild xmlPropertyLetNode
    
    
    xmlPropertyDetailsNode.appendChild xmlPropertyDetailsDescriptionNode
    '** Get Security Property Description : END **
    
    
    
    
    
    
    '** Get Property Valuation Details : START **
    Set xmlPropertyValuationDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyValuationDetails", "")
    
    
    'LastValuationDate
    gstrRuleNumber = strRuleName & ".18"
    Set xmlLastValuationDateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LastValuationDate", "")
    If Not xmlValuerInstruction Is Nothing Then
        strAppointmentDate = GetAttributeText(xmlValuerInstruction, "APPOINTMENTDATE")
        If strAppointmentDate <> "" Then
            xmlLastValuationDateNode.Text = Format(strAppointmentDate, "dd/mm/yyyy")
        End If
    End If
    xmlPropertyValuationDetailsNode.appendChild xmlLastValuationDateNode
    
    
    
    'PropertyValue - PRESENTVALUATION
    gstrRuleNumber = strRuleName & ".19"
    Set xmlPropertyValueNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyValue", "")
    
    If Not xmlValnRepValuationNode Is Nothing Then
        strPropertyValue = GetAttributeText(xmlValnRepValuationNode, "PRESENTVALUATION")
        If strPropertyValue <> "" Then
            xmlPropertyValueNode.Text = strPropertyValue
        End If
    End If
    xmlPropertyValuationDetailsNode.appendChild xmlPropertyValueNode
    
    
    
    'TotalValuation - POSTWORKSVALUATION
    gstrRuleNumber = strRuleName & ".20"
    Set xmlTotalValuationNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TotalValuation", "")
    
    If Not xmlValnRepValuationNode Is Nothing Then
        strTotalValuation = GetAttributeText(xmlValnRepValuationNode, "POSTWORKSVALUATION")
        If strTotalValuation <> "" Then
            xmlTotalValuationNode.Text = strTotalValuation
        End If
    End If
    xmlPropertyValuationDetailsNode.appendChild xmlTotalValuationNode
    
    
    
    'ValuationType
    gstrRuleNumber = strRuleName & ".21"
    Set xmlValuationTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValuationType", "")
    If Not xmlValuerInstruction Is Nothing Then
        strValuationTypeValueID = GetAttributeText(xmlValuerInstruction, "VALUATIONTYPE")
    End If
    If Not strValuationTypeValueID = "" Then
        'Use VALNREPVALUATION
        xmlValuationTypeNode.Text = GetComboText("ValuationType", CSafeInt(strValuationTypeValueID))
    Else
        'Use NEWPROPERTY
        xmlValuationTypeNode.Text = GetAttributeText(xmlNewPropertyNode, "VALUATIONTYPE_TEXT")
    End If
    xmlPropertyValuationDetailsNode.appendChild xmlValuationTypeNode
    
  
  
    'PurchasePrice - PURCHASEPRICEORESTIMATEDVALUE
    gstrRuleNumber = strRuleName & ".22"
    Set xmlPropertyPurchaseNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyPurchasePrice", "")
    strPurchasePriceorEstimatedValue = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "PURCHASEPRICEORESTIMATEDVALUE")
    If strPurchasePriceorEstimatedValue <> "" Then
        xmlPropertyPurchaseNode.Text = strPurchasePriceorEstimatedValue
    End If
    xmlPropertyValuationDetailsNode.appendChild xmlPropertyPurchaseNode
  
  
  
    'Servicing Branch
    gstrRuleNumber = strRuleName & ".23"
    Set xmlServicingBranchNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ServicingBranch", "")
     'EP2_956 - for Epsom 2 return the numeric code for the postode of the new property
          ' we only  want the first 2 letters
            strUKPostCode = Left(xmlPostCodeNode.Text, 2)
            ' some postcodes start with 1 letter
            If IsNumeric(Right(strUKPostCode, 1)) Then
                strUKPostCode = Left(strUKPostCode, 1)
            End If
        'EP2_956 do a lookup on the lending branch combo
            strSearchPat = "COMBOVALUE[COMBOVALIDATION[@VALIDATIONTYPE='" & strUKPostCode & "']]"
            Set xmlLendingBranchComboValueNode = m_xmlLendingBranchCombo.selectSingleNode(strSearchPat)
    'xmlServicingBranchNode.Text = GetAttributeText(vxmlRequestNode, "???????")
    'xmlServicingBranchNode.Text = "??? To be identified ???"
    'Set to Postcode for now as agreed by Marcus Bennett
    'xmlServicingBranchNode.Text = m_strAddressPostCode
    xmlServicingBranchNode.Text = GetAttributeText(xmlLendingBranchComboValueNode, "VALUEID")
    'store in module level variable for use in orginating
    m_strAddressPostCode = xmlServicingBranchNode.Text
    xmlPropertyValuationDetailsNode.appendChild xmlServicingBranchNode
  
  
  
    'Underwriter
    gstrRuleNumber = strRuleName & ".24"
    Set xmlUnderwriterNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Underwriter", "")
    'Not currently held - Marcus Bennett 16/05/2006
    xmlUnderwriterNode.Text = "* NOT KNOWN *"
    xmlPropertyValuationDetailsNode.appendChild xmlUnderwriterNode
    
    
    xmlPropertyDetailsNode.appendChild xmlPropertyValuationDetailsNode
    '** Get Property Valuation Details : END **
    
    
    
    
    
    
 
    '** Get Miscellaneous Security Data : START **
    Set xmlMiscSecurityDataNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MiscellaneousSecurityData", "")
    
    
    'PropertyPurchaseDate
    gstrRuleNumber = strRuleName & ".25"
    Set xmlPropertyPurchaseDateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PropertyPurchaseDate", "")
    'use the completion date
    xmlPropertyPurchaseDateNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/REPORTONTITLE"), "COMPLETIONDATE")
    xmlMiscSecurityDataNode.appendChild xmlPropertyPurchaseDateNode
    
    
    'LoanToValue
    gstrRuleNumber = strRuleName & ".26"
    Set xmlLoanToValueNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LoanToValue", "")
    strLoanToValue = GetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE"), "LTV")
    If strLoanToValue <> "" Then
        xmlLoanToValueNode.Text = strLoanToValue
    End If
    xmlMiscSecurityDataNode.appendChild xmlLoanToValueNode
    
    
    
    xmlPropertyDetailsNode.appendChild xmlMiscSecurityDataNode
    '** Get Miscellaneous Security Data : END **
      


    vxmlResponseNode.appendChild xmlPropertyDetailsNode


    LogDetails 3, "SetupPropertyDetails : End"


SetupPropertyDetailsExit:
    
    Set xmlDoc = Nothing
    Set xmlSearchNode = Nothing
    Set xmlPropertyDetailsNode = Nothing
    Set xmlValnInstructionNodeList = Nothing
    Set xmlValuerInstruction = Nothing
    Set xmlvalnRepPropertyDetailsNode = Nothing
    Set xmlValnRepValuationNode = Nothing
    Set xmlPropertyAddressNode = Nothing
    Set xmlNewPropertyNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlHouseOrBuildingNameNode = Nothing
    Set xmlFlatNameOrNumberNode = Nothing
    Set xmlHouseOrBuildingNumberNode = Nothing
    Set xmlStreetNode = Nothing
    Set xmlDistrictNode = Nothing
    Set xmlTownOrCityNode = Nothing
    Set xmlCountyNode = Nothing
    Set xmlPostCodeNode = Nothing
    Set xmlPropertyDetailsDescriptionNode = Nothing
    Set xmlTitleNoOneNode = Nothing
    Set xmlTitleNoTwoNode = Nothing
    Set xmlTitleNoThreeNode = Nothing
    Set xmlPropertyTenureNode = Nothing
    Set xmlPropertyTypeNode = Nothing
    Set xmlPropertySubTypeNode = Nothing
    Set xmlPropertyDescriptionNode = Nothing
    Set xmlFreeholdTitleNode = Nothing
    Set xmlPropertyLetNode = Nothing
    Set xmlPropertyValuationDetailsNode = Nothing
    Set xmlLastValuationDateNode = Nothing
    Set xmlPropertyValueNode = Nothing
    Set xmlTotalValuationNode = Nothing
    Set xmlValuationTypeNode = Nothing
    Set xmlPropertyPurchaseNode = Nothing
    Set xmlServicingBranchNode = Nothing
    Set xmlUnderwriterNode = Nothing
    Set xmlMiscSecurityDataNode = Nothing
    Set xmlPropertyPurchaseDateNode = Nothing
    Set xmlLoanToValueNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub SetupCustomerCreation(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupCustomerCreation
' Author:       Sam Dali
' Date:         23/08/05
' Purpose:      Create Customer Creation data in Completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to
'01/09/05       Started with structural assumptions
'------------------------------------------------------------------------

    On Error GoTo SetupCustomerCreationExit
    
    Const strFunctionName As String = "SetupCustomerCreation"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlSearchNode As IXMLDOMNode 'temp node used for searching
    Dim strRuleName As String
    Dim xmlCustomerRoleNodeList As IXMLDOMNodeList
    Dim xmlCustomerRoleNode As IXMLDOMNode
    Dim xmlCustomerVersionNode As IXMLDOMNode
    Dim strCustomerVersion As String
    
    
    '*** Client Detail - Name and Address Entry ***
    'CustomerNumber
    'CustomerName
    'CustomerAddress
    '
    '*** Client Detail - Identification ***
    'Date Or Birth
    'Gender
    'MaritalStatus
    'National Insurance No.
    '
    '*** Client Detail - Telephones ***
    'HomePhoneNumber
    'WorkPhoneNumber
    'MobilePhoneNumber
    '
    '*** Client Detail - Demographic ***
    'Lending Branch
    'Occupation/Business Type
    'Industry Sector
    '
    '*** Miscellaneous ***
    'Proof of Income
    'Gross Income
    
    
    
    LogDetails 3, "SetupCustomerCreation : Start"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strRuleName = "CustomerCreation"
    
    
    Set xmlCustomerRoleNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER='1' or @CUSTOMERORDER='2']")
    
    For Each xmlCustomerRoleNode In xmlCustomerRoleNodeList
        strCustomerVersion = GetAttributeText(xmlCustomerRoleNode, "CUSTOMERVERSIONNUMBER")
        Set xmlCustomerVersionNode = xmlCustomerRoleNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER='" & strCustomerVersion & "']")
        AddCustomer xmlCustomerVersionNode, vxmlRequestNode, vxmlResponseNode
    Next
    
    LogDetails 3, "SetupCustomerCreation : End"


SetupCustomerCreationExit:
    
    Set xmlDoc = Nothing
    Set xmlSearchNode = Nothing
    Set xmlCustomerRoleNodeList = Nothing
    Set xmlCustomerRoleNode = Nothing
    Set xmlCustomerVersionNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub AddCustomer(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlAppRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    AddCustomer
' Author:       Sam Dali
' Date:         02/09/05
' Purpose:      Adding a Customer Record to completionBlock
' Input parameters:
'               vxmlRequestNode - XML Request, a customer record
'               vxmlResponseNode - XML Response to append to
'------------------------------------------------------------------------

    On Error GoTo AddCustomerExit
    
    Const strFunctionName As String = "AddCustomer"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim xmlCustomerNode As IXMLDOMNode
    
    Dim xmlCustomerNumberNode As IXMLDOMNode
    Dim xmlHomePhoneNumberNode As IXMLDOMNode
    Dim xmlHomePhoneAreaCodeNode As IXMLDOMNode
    Dim xmlHomePhoneLocalNumberNode As IXMLDOMNode
    Dim xmlHomePhoneExtensionNode As IXMLDOMNode
    Dim xmlMobilePhoneNumberNode As IXMLDOMNode
    Dim xmlMobilePhoneAreaCodeNode As IXMLDOMNode
    Dim xmlMobilePhoneLocalNumberNode As IXMLDOMNode
    Dim xmlMobilePhoneExtensionNode As IXMLDOMNode
    Dim xmlWorkPhoneNumberNode As IXMLDOMNode
    Dim xmlWorkPhoneAreaCodeNode As IXMLDOMNode
    Dim xmlWorkPhoneLocalNUmberNode As IXMLDOMNode
    Dim xmlWorkPhoneExtensionNode As IXMLDOMNode
    
    'SAB 10/05/2006 - EPSOM
    Dim xmlCustomerNameNode As IXMLDOMNode
    Dim strTitle As String
    Dim xmlCustomerTitleNode As IXMLDOMNode
    Dim xmlCustomerForeNameNode As IXMLDOMNode
    Dim xmlCustomerSecondForNameNode As IXMLDOMNode
    Dim xmlCustomerOtherForeNamesNode As IXMLDOMNode
    Dim xmlCustomerSurNameNode As IXMLDOMNode
    
    Dim strUsage As String
    Dim xmlDateOfBirthNode As IXMLDOMNode
    Dim strDateofBirth As String
    Dim xmlGenderNode As IXMLDOMNode
    Dim xmlMaritalStatusNode As IXMLDOMNode
    Dim xmlNINumberNode As IXMLDOMNode
    
    Dim xmlPhoneNumberNodeList As IXMLDOMNodeList
    Dim xmlPhoneNumberNode As IXMLDOMNode
    Dim xmlHomeNumberNode As IXMLDOMNode
    Dim xmlMobileNumberNode As IXMLDOMNode
    Dim xmlWorkNumberNode As IXMLDOMNode
    
    Dim xmlLendingBranchNode As IXMLDOMNode
    Dim xmlBusinessTypeNode As IXMLDOMNode
    Dim xmlIndustrySectorNode As IXMLDOMNode
    
    Dim xmlProofOfIncomeNode As IXMLDOMNode
    Dim strProofOfIncomeInd As String
    Dim xmlSearchNode As IXMLDOMNode
    Dim xmlGrossIncomeNode As IXMLDOMNode
    Dim strGrossAllowableIncome1 As String
    'SAB 10/05/2006 - EPSOM
    
    Dim strRuleName As String
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber  As String
    Dim xmlCustomerAddressNodelist As IXMLDOMNodeList
    Dim xmlCustomerAddressNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
    Dim strAddressType As String
    Dim strUKPostCode As String
    'EP2_956
    Dim strSearchPat As String
    Dim xmlLendingBranchComboValueNode  As IXMLDOMNode
    
    'AW 28/03/07    EP2_2125
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNewPropertyAddressNode As IXMLDOMNode
    Dim strApplicationType As String
    Dim strNatureOfLoan As String
    Dim strEvidenceOfAcconts  As String
    'AW 28/03/07    EP2_2125    -   End
    
    LogDetails 3, "AddCustomer : Start"
    
    strRuleName = "CustomerCreation"
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlCustomerNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CustomerDetails", "")
    
    
    
    ' *** Client Detail - Name and Address Entry : START ***
    
    strCustomerNumber = GetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
    strCustomerVersionNumber = GetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
    
    'CustomerNumber
    gstrRuleNumber = strRuleName & "-" & strCustomerNumber & "-" & strCustomerVersionNumber & ".1"
    Set xmlCustomerNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CustomerNumber", "")
    xmlCustomerNumberNode.Text = GetAttributeText(vxmlRequestNode.parentNode, "CUSTOMERNUMBER")
    xmlCustomerNode.appendChild xmlCustomerNumberNode
    
    
    
    
    'CustomerName
    gstrRuleNumber = strRuleName & ".2"
    Set xmlCustomerNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CustomerName", "")
    Set xmlCustomerTitleNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Title", "")
    Set xmlCustomerForeNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ForeName", "")
    Set xmlCustomerSecondForNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SecondForeName", "")
    Set xmlCustomerOtherForeNamesNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "OtherForeNames", "")
    Set xmlCustomerSurNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "SurName", "")
    
    strTitle = GetAttributeText(vxmlRequestNode, "TITLE")
    If IsValidationType("Title", CSafeInt(strTitle), "O") Then    '"0" = Other
        xmlCustomerTitleNode.Text = GetAttributeText(vxmlRequestNode, "TITLEOTHER")
    Else
        xmlCustomerTitleNode.Text = GetAttributeText(vxmlRequestNode, "TITLE_TEXT")
    End If
    
    xmlCustomerForeNameNode.Text = GetAttributeText(vxmlRequestNode, "FIRSTFORENAME")
    xmlCustomerSecondForNameNode.Text = GetAttributeText(vxmlRequestNode, "SECONDFORENAME")
    xmlCustomerOtherForeNamesNode.Text = GetAttributeText(vxmlRequestNode, "OTHERFORENAMES")
    xmlCustomerSurNameNode.Text = GetAttributeText(vxmlRequestNode, "SURNAME")
    
    xmlCustomerNameNode.appendChild xmlCustomerTitleNode
    xmlCustomerNameNode.appendChild xmlCustomerForeNameNode
    xmlCustomerNameNode.appendChild xmlCustomerSecondForNameNode
    xmlCustomerNameNode.appendChild xmlCustomerOtherForeNamesNode
    xmlCustomerNameNode.appendChild xmlCustomerSurNameNode
    xmlCustomerNode.appendChild xmlCustomerNameNode
    
    
    
    'CustomerAddress
    gstrRuleNumber = strRuleName & ".3"
    
    'AW EP2_2125 DBM178 - For a purchase this will the security address, for a remortgage or
    'Buy to let application this will be the current address.
    Set xmlTempNode = vxmlAppRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND")
    If Not xmlTempNode Is Nothing Then
        strApplicationType = GetMandatoryAttributeText(xmlTempNode, "TYPEOFAPPLICATION")
        strNatureOfLoan = GetMandatoryAttributeText(xmlTempNode, "NATUREOFLOAN")
    End If

    If IsValidationType("TypeOfMortgage", strApplicationType, "R") Or _
        (IsValidationType("NatureOfLoan", strNatureOfLoan, "BI") Or IsValidationType("NatureOfLoan", strNatureOfLoan, "BR")) Then

        'get all addresses
        Set xmlCustomerAddressNodelist = vxmlRequestNode.selectNodes("CUSTOMERADDRESS")
        
        'loop thru' all addresses to get current address
        For Each xmlCustomerAddressNode In xmlCustomerAddressNodelist
            strAddressType = GetAttributeText(xmlCustomerAddressNode, "ADDRESSTYPE")
            If Not strAddressType = "" Then
                If IsValidationType("CustomerAddressType", CSafeInt(strAddressType), "H") Then
                    'get address node by matching AddressGUID
                    Set xmlAddressNode = xmlCustomerAddressNode.selectSingleNode("ADDRESS[@ADDRESSGUID='" & _
                        GetAttributeText(xmlCustomerAddressNode, "ADDRESSGUID") & "']")
                    If Not xmlAddressNode Is Nothing Then
                        AddCustomerAddress xmlAddressNode, xmlCustomerNode
                    End If
                End If
            End If
        Next
        
    Else
    
        'Use New Property Address
        Set xmlNewPropertyAddressNode = vxmlAppRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS")
        
        If Not xmlNewPropertyAddressNode Is Nothing Then
            AddCustomerAddress xmlNewPropertyAddressNode, xmlCustomerNode
        End If
    End If
    'AW EP2_2125    -   End
    
    
    ' *** Client Detail - Name and Address Entry : END ***
    
    
    
    
    
    ' *** Client Detail - Identification : START ***
    
    'Date Or Birth
    gstrRuleNumber = strRuleName & ".4"
    Set xmlDateOfBirthNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "DateOfBirth", "")
    
    strDateofBirth = GetAttributeText(vxmlRequestNode, "DATEOFBIRTH")
    If strDateofBirth <> "" Then
        xmlDateOfBirthNode.Text = Format$(strDateofBirth, "dd/mm/yyyy")
    End If
    xmlCustomerNode.appendChild xmlDateOfBirthNode
    
    
    'Gender
    gstrRuleNumber = strRuleName & ".5"
    Set xmlGenderNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Gender", "")
    xmlGenderNode.Text = GetAttributeText(vxmlRequestNode, "GENDER_TEXT")
    xmlCustomerNode.appendChild xmlGenderNode
    
    
    'MaritalStatus
    gstrRuleNumber = strRuleName & ".6"
    Set xmlMaritalStatusNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MaritalStatus", "")
    xmlMaritalStatusNode.Text = GetAttributeText(vxmlRequestNode, "MARITALSTATUS_TEXT")
    xmlCustomerNode.appendChild xmlMaritalStatusNode
    
    
    'National Insurance No.
    gstrRuleNumber = strRuleName & ".7"
    Set xmlNINumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "NationalInsuranceNumber", "")
    xmlNINumberNode.Text = GetAttributeText(vxmlRequestNode, "NATIONALINSURANCENUMBER")
    xmlCustomerNode.appendChild xmlNINumberNode
    
    ' *** Client Detail - Identification : END ***
    
    
    
    
    
    ' *** Client Detail - Telephones : START ***
    
    'get all phonenumbers
    Set xmlPhoneNumberNodeList = vxmlRequestNode.selectNodes("CUSTOMERTELEPHONENUMBER")
    For Each xmlPhoneNumberNode In xmlPhoneNumberNodeList
        strUsage = GetAttributeText(xmlPhoneNumberNode, "USAGE")
        If Not strUsage = "" Then
            If IsValidationType("TelephoneUsage", CSafeInt(strUsage), "H") Then
                Set xmlHomeNumberNode = xmlPhoneNumberNode
            Else
                If IsValidationType("TelephoneUsage", CSafeInt(strUsage), "M") Then
                    Set xmlMobileNumberNode = xmlPhoneNumberNode
                Else
                    If IsValidationType("TelephoneUsage", CSafeInt(strUsage), "W") Then
                        Set xmlWorkNumberNode = xmlPhoneNumberNode
                    End If
                End If
            End If
        End If
    Next
    
    
    'HomePhoneNumber
    gstrRuleNumber = strRuleName & ".8"
    Set xmlHomePhoneNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HomePhoneNumber", "")
    Set xmlHomePhoneAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
    Set xmlHomePhoneLocalNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
    Set xmlHomePhoneExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
     
    If Not xmlHomeNumberNode Is Nothing Then
        xmlHomePhoneAreaCodeNode.Text = GetAttributeText(xmlHomeNumberNode, "AREACODE")
        xmlHomePhoneLocalNumberNode.Text = GetAttributeText(xmlHomeNumberNode, "TELEPHONENUMBER")
        xmlHomePhoneExtensionNode.Text = GetAttributeText(xmlHomeNumberNode, "EXTENSIONNUMBER")
    End If
    
    xmlHomePhoneNumberNode.appendChild xmlHomePhoneAreaCodeNode
    xmlHomePhoneNumberNode.appendChild xmlHomePhoneLocalNumberNode
    xmlHomePhoneNumberNode.appendChild xmlHomePhoneExtensionNode
    xmlCustomerNode.appendChild xmlHomePhoneNumberNode
    
    
    'WorkPhoneNumber
    gstrRuleNumber = strRuleName & ".9"
    Set xmlWorkPhoneNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "WorkPhoneNumber", "")
    Set xmlWorkPhoneAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
    Set xmlWorkPhoneLocalNUmberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
    Set xmlWorkPhoneExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
    
    If Not xmlWorkNumberNode Is Nothing Then
        xmlWorkPhoneAreaCodeNode.Text = GetAttributeText(xmlWorkNumberNode, "AREACODE")
        xmlWorkPhoneLocalNUmberNode.Text = GetAttributeText(xmlWorkNumberNode, "TELEPHONENUMBER")
        xmlWorkPhoneExtensionNode.Text = GetAttributeText(xmlWorkNumberNode, "EXTENSIONNUMBER")
    End If
     
    xmlWorkPhoneNumberNode.appendChild xmlWorkPhoneAreaCodeNode
    xmlWorkPhoneNumberNode.appendChild xmlWorkPhoneLocalNUmberNode
    xmlWorkPhoneNumberNode.appendChild xmlWorkPhoneExtensionNode
    xmlCustomerNode.appendChild xmlWorkPhoneNumberNode
    
    
    'MobilePhoneNumber
    gstrRuleNumber = strRuleName & ".10"
    Set xmlMobilePhoneNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MobilePhoneNumber", "")
    Set xmlMobilePhoneAreaCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AreaCode", "")
    Set xmlMobilePhoneLocalNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LocalNumber", "")
    Set xmlMobilePhoneExtensionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Extension", "")
    
    If Not xmlMobileNumberNode Is Nothing Then
        xmlMobilePhoneAreaCodeNode.Text = GetAttributeText(xmlMobileNumberNode, "AREACODE")
        xmlMobilePhoneLocalNumberNode.Text = GetAttributeText(xmlMobileNumberNode, "TELEPHONENUMBER")
        xmlMobilePhoneExtensionNode.Text = GetAttributeText(xmlMobileNumberNode, "EXTENSIONNUMBER")
    End If
    
    xmlMobilePhoneNumberNode.appendChild xmlMobilePhoneAreaCodeNode
    xmlMobilePhoneNumberNode.appendChild xmlMobilePhoneLocalNumberNode
    xmlMobilePhoneNumberNode.appendChild xmlMobilePhoneExtensionNode
    xmlCustomerNode.appendChild xmlMobilePhoneNumberNode
    
    ' *** Client Detail - Telephones : END ***
    
    
    
    
    
    ' *** Client Detail - Demographic : START ***
    
    'Lending Branch
    gstrRuleNumber = strRuleName & ".11"
    Set xmlLendingBranchNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "LendingBranch", "")
    
    'EP1055 - Use First Customers postcode in case second customer doesn't have one
     If Len(m_strAddressPostCode) > 0 Then
        xmlLendingBranchNode.Text = m_strAddressPostCode
    Else
       'EP2_956 - for Epsom 2 return the numeric code for the postode
          ' we only  want the first 2 letters
            strUKPostCode = Left(GetAttributeText(xmlAddressNode, "POSTCODE"), 2)
            ' some postcodes start with 1 letter
            If IsNumeric(Right(strUKPostCode, 1)) Then
                strUKPostCode = Left(strUKPostCode, 1)
            End If
        'EP2_956 do a lookup on the lending branch combo
            strSearchPat = "COMBOVALUE[COMBOVALIDATION[@VALIDATIONTYPE='" & strUKPostCode & "']]"
            Set xmlLendingBranchComboValueNode = m_xmlLendingBranchCombo.selectSingleNode(strSearchPat)
            xmlLendingBranchNode.Text = GetAttributeText(xmlLendingBranchComboValueNode, "VALUEID")
    End If
    'EP1055 - End
    'Use this module variable to hold the postcode so it can be used again
    m_strAddressPostCode = xmlLendingBranchNode.Text
    
    xmlCustomerNode.appendChild xmlLendingBranchNode
    
    
    'Occupation/Business Type
    gstrRuleNumber = strRuleName & ".12"
    Set xmlBusinessTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "OccupationBusinessType", "")
    
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("EMPLOYMENT[@MAINSTATUS=1]")
    If Not xmlSearchNode Is Nothing Then
        xmlBusinessTypeNode.Text = GetAttributeText(xmlSearchNode, "EMPLOYMENTSTATUS_TEXT")
    End If
   
    
    xmlCustomerNode.appendChild xmlBusinessTypeNode
    
    
    'Industry Sector
    gstrRuleNumber = strRuleName & ".13"
    Set xmlIndustrySectorNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "IndustrySector", "")
    'need to look up in combo Nature of business
    If Not xmlSearchNode Is Nothing Then
        strSearchPat = GetAttributeText(xmlSearchNode, "INDUSTRYTYPE")
        strSearchPat = "COMBOVALUE[@VALUEID='" & strSearchPat & "']"
        Set xmlSearchNode = m_xmlNatureofBusinessCombo.selectSingleNode(strSearchPat)
    End If
    
    'EP2_956
    If Not xmlSearchNode Is Nothing Then
        xmlIndustrySectorNode.Text = GetAttributeText(xmlSearchNode, "VALUENAME")
    Else
        xmlIndustrySectorNode.Text = "Not applicable (99997)"
    End If
    'xmlIndustrySectorNode.Text = GetAttributeText(vxmlRequestNode, "???????")
    'xmlIndustrySectorNode.Text = "??? To be identified ???"
    'Default to "Not applicable (99997)" - Marcus Bennett 16/05/2006
     xmlCustomerNode.appendChild xmlIndustrySectorNode

    ' *** Client Detail - Demographic : END ***
    
    
    
    
    ' *** Miscellaneous : START ***
    
    'Proof of Income
    gstrRuleNumber = strRuleName & ".14"
    Set xmlProofOfIncomeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ProofOfIncome", "")
    'EP2_956 Not applicable if app is self-employed
    'AW EP2_2125    28/03/07
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("EMPLOYMENT[@MAINSTATUS=1]")
    If Not xmlSearchNode Is Nothing Then
        If GetAttributeText(xmlSearchNode, "EMPLOYMENTSTATUS_TYPE_E") = "true" Then
        
            Set xmlTempNode = xmlSearchNode.selectSingleNode("EMPLOYEDDETAILS")
            If Not xmlTempNode Is Nothing Then
                strProofOfIncomeInd = GetAttributeText(xmlTempNode, "WAGESLIPSSEENINDICATOR")
                If strProofOfIncomeInd = "1" Then
                    xmlProofOfIncomeNode.Text = "Y"
                Else
                  If strProofOfIncomeInd = "0" Then
                     xmlProofOfIncomeNode.Text = "N"
                  End If
                End If
             End If
         Else
            
              If GetAttributeText(xmlSearchNode, "EMPLOYMENTSTATUS_TYPE_S") = "true" Then
              
                Set xmlTempNode = xmlSearchNode.selectSingleNode("NETPROFIT")
                If Not xmlTempNode Is Nothing Then
                    strEvidenceOfAcconts = GetAttributeText(xmlTempNode, "EVIDENCEOFACCOUNTS")
                    If strEvidenceOfAcconts = "1" Then
                        xmlProofOfIncomeNode.Text = "Y"
                    Else
                      If strEvidenceOfAcconts = "0" Then
                         xmlProofOfIncomeNode.Text = "N"
                      End If
                    End If
                End If
                
              Else
                    xmlProofOfIncomeNode.Text = "N/A"
              End If
              
              
         End If
    End If
    'AW EP2_2125    28/03/07    -   End
    
    Set xmlSearchNode = Nothing
    xmlCustomerNode.appendChild xmlProofOfIncomeNode
    
    
    
    'Gross Income
    gstrRuleNumber = strRuleName & ".15"
    Set xmlGrossIncomeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "GrossIncome", "")
    
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("INCOMESUMMARY")
    If Not xmlSearchNode Is Nothing Then
        strGrossAllowableIncome1 = GetAttributeText(xmlSearchNode, "ALLOWABLEANNUALINCOME")
    End If
    Set xmlSearchNode = Nothing
    If Not strGrossAllowableIncome1 = "" Then
        m_dblJointGrossAllowableIncome = m_dblJointGrossAllowableIncome + CSafeDbl(strGrossAllowableIncome1)
        xmlGrossIncomeNode.Text = strGrossAllowableIncome1
    End If
    xmlCustomerNode.appendChild xmlGrossIncomeNode
    
    ' *** Miscellaneous : END ***
    
    
    vxmlResponseNode.appendChild xmlCustomerNode
    
    LogDetails 3, "AddCustomer : End"


AddCustomerExit:
    Set xmlDoc = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlCustomerNumberNode = Nothing
    Set xmlHomePhoneNumberNode = Nothing
    Set xmlHomePhoneAreaCodeNode = Nothing
    Set xmlHomePhoneLocalNumberNode = Nothing
    Set xmlHomePhoneExtensionNode = Nothing
    Set xmlMobilePhoneNumberNode = Nothing
    Set xmlMobilePhoneAreaCodeNode = Nothing
    Set xmlMobilePhoneLocalNumberNode = Nothing
    Set xmlMobilePhoneExtensionNode = Nothing
    Set xmlWorkPhoneNumberNode = Nothing
    Set xmlWorkPhoneAreaCodeNode = Nothing
    Set xmlWorkPhoneLocalNUmberNode = Nothing
    Set xmlWorkPhoneExtensionNode = Nothing
    Set xmlCustomerNameNode = Nothing
    Set xmlCustomerTitleNode = Nothing
    Set xmlCustomerForeNameNode = Nothing
    Set xmlCustomerSecondForNameNode = Nothing
    Set xmlCustomerOtherForeNamesNode = Nothing
    Set xmlCustomerSurNameNode = Nothing
    Set xmlDateOfBirthNode = Nothing
    Set xmlGenderNode = Nothing
    Set xmlMaritalStatusNode = Nothing
    Set xmlNINumberNode = Nothing
    Set xmlPhoneNumberNodeList = Nothing
    Set xmlPhoneNumberNode = Nothing
    Set xmlHomeNumberNode = Nothing
    Set xmlMobileNumberNode = Nothing
    Set xmlWorkNumberNode = Nothing
    Set xmlLendingBranchNode = Nothing
    Set xmlBusinessTypeNode = Nothing
    Set xmlIndustrySectorNode = Nothing
    Set xmlProofOfIncomeNode = Nothing
    Set xmlSearchNode = Nothing
    Set xmlGrossIncomeNode = Nothing
    Set xmlCustomerAddressNodelist = Nothing
    Set xmlCustomerAddressNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlNewPropertyAddressNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub AddCustomerAddress(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    AddCustomerAddress
' Author:       Sam Dali
' Date:         07/09/05
' Purpose:      Adding a current address record to CustomerDetails
' Input parameters:
'               vxmlRequestNode - XML Request, an address node
'               vxmlResponseNode - XML Response to append to, CustomerCreation Node
'------------------------------------------------------------------------

    On Error GoTo AddCustomerAddressExit
    
    Const strFunctionName As String = "AddCustomerAddress"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim xmlCustomerAddressNode As IXMLDOMNode
    
    'all nodes for current Address Record
    Dim xmlCurrentAddressDetailsNode As IXMLDOMNode
    Dim xmlAddressDetailsNode As IXMLDOMNode
    Dim xmlFlatNameOrNumberNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNameNode As IXMLDOMNode
    Dim xmlHouseOrBuildingNumberNode As IXMLDOMNode
    Dim xmlStreetNode As IXMLDOMNode
    Dim xmlDistrictNode As IXMLDOMNode
    Dim xmlTownOrCityNode As IXMLDOMNode
    Dim xmlCountyNode As IXMLDOMNode
    Dim xmlPostCodeNode As IXMLDOMNode
    Dim xmlCountryCodeNode As IXMLDOMNode
    Dim strDate As String
    
    LogDetails 3, "AddCustomerAddress : Start"
    
    
    Dim strRuleName As String
    strRuleName = "CustomerCreation"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    
    Set xmlCurrentAddressDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CurrentAddressDetails", "")
    Set xmlAddressDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "AddressDetails", "")
    
    'FlatNameOrNumber
    gstrRuleNumber = strRuleName & ".8"
    Set xmlFlatNameOrNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FlatNameOrNumber", "")
    xmlFlatNameOrNumberNode.Text = GetAttributeText(vxmlRequestNode, "FLATNUMBER")
    xmlAddressDetailsNode.appendChild xmlFlatNameOrNumberNode
    
    'HouseOrBuildingName
    gstrRuleNumber = strRuleName & ".9"
    Set xmlHouseOrBuildingNameNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingName", "")
    xmlHouseOrBuildingNameNode.Text = GetAttributeText(vxmlRequestNode, "BUILDINGORHOUSENAME")
    xmlAddressDetailsNode.appendChild xmlHouseOrBuildingNameNode
    
    'HouseOrBuildingNumber
    gstrRuleNumber = strRuleName & ".10"
    Set xmlHouseOrBuildingNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "HouseOrBuildingNumber", "")
    xmlHouseOrBuildingNumberNode.Text = GetAttributeText(vxmlRequestNode, "BUILDINGORHOUSENUMBER")
    xmlAddressDetailsNode.appendChild xmlHouseOrBuildingNumberNode
    
    'Street
    gstrRuleNumber = strRuleName & ".11"
    Set xmlStreetNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Street", "")
    xmlStreetNode.Text = GetAttributeText(vxmlRequestNode, "STREET")
    xmlAddressDetailsNode.appendChild xmlStreetNode
    
    'District
    gstrRuleNumber = strRuleName & ".12"
    Set xmlDistrictNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "District", "")
    xmlDistrictNode.Text = GetAttributeText(vxmlRequestNode, "DISTRICT")
    xmlAddressDetailsNode.appendChild xmlDistrictNode
    
    'TownOrCity
    gstrRuleNumber = strRuleName & ".13"
    Set xmlTownOrCityNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "TownOrCity", "")
    xmlTownOrCityNode.Text = GetAttributeText(vxmlRequestNode, "TOWN")
    xmlAddressDetailsNode.appendChild xmlTownOrCityNode
    
    'County
    gstrRuleNumber = strRuleName & ".14"
    Set xmlCountyNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "County", "")
    xmlCountyNode.Text = GetAttributeText(vxmlRequestNode, "COUNTY")
    xmlAddressDetailsNode.appendChild xmlCountyNode
    
    'PostCode
    gstrRuleNumber = strRuleName & ".15"
    Set xmlPostCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PostCode", "")
    xmlPostCodeNode.Text = GetAttributeText(vxmlRequestNode, "POSTCODE")
    xmlAddressDetailsNode.appendChild xmlPostCodeNode
    
    'CountryCode
    gstrRuleNumber = strRuleName & ".16"
    Dim strCountryCode As String
    Set xmlCountryCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CountryCode", "")
    strCountryCode = GetAttributeText(vxmlRequestNode, "COUNTRY")
    If strCountryCode <> "" Then
        xmlCountryCodeNode.Text = GetValidationType("Country", strCountryCode)
    End If
    xmlAddressDetailsNode.appendChild xmlCountryCodeNode
    
    xmlCurrentAddressDetailsNode.appendChild xmlAddressDetailsNode
    
    vxmlResponseNode.appendChild xmlCurrentAddressDetailsNode
    
    LogDetails 3, "AddCustomerAddress : End"


AddCustomerAddressExit:
    
    Set xmlDoc = Nothing
    Set xmlCurrentAddressDetailsNode = Nothing
    Set xmlFlatNameOrNumberNode = Nothing
    Set xmlHouseOrBuildingNameNode = Nothing
    Set xmlHouseOrBuildingNumberNode = Nothing
    Set xmlStreetNode = Nothing
    Set xmlDistrictNode = Nothing
    Set xmlTownOrCityNode = Nothing
    Set xmlCountyNode = Nothing
    Set xmlPostCodeNode = Nothing
    Set xmlCountryCodeNode = Nothing
    CheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub XXX_SetupComponentRecord(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupComponentRecord
' Author:       Sam Dali
' Date:         30/08/05
' Purpose:      Create Loan Component data in Completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to
'------------------------------------------------------------------------

On Error GoTo SetupComponentRecordExit

Const strFunctionName As String = "SetupComponentRecord"

Dim xmlDoc As FreeThreadedDOMDocument40

Dim xmlLoanComponentList As IXMLDOMNodeList
Dim xmlLoanComponentNode As IXMLDOMNode
Dim xmlMortgageProductCodeNodeList As IXMLDOMNodeList
Dim xmlMortgageProductCodeNode As IXMLDOMNode
Dim xmlInterestRateTypeNode As IXMLDOMNode

Dim strMortgageSubQuoteNumber As String
Dim dtMortgageProductStartDate As Date, dtInterestRateTypeStartDate As Date, dtLoanComponentStartDate As Date
Dim strMortgageProductStartDate As String, strInterestRateTypeStartDate As String, strLoanComponentStartDate As String
Dim blnToAddComponent As Boolean
    
LogDetails 3, "SetupComponentRecord : Start"

Set xmlDoc = New FreeThreadedDOMDocument40
xmlDoc.validateOnParse = False
xmlDoc.setProperty "NewParser", True

'get mortgagesubquotenumber for accepted quotation
strMortgageSubQuoteNumber = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION[@QUOTATIONNUMBER='" & _
                                    GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "ACCEPTEDQUOTENUMBER") & "']"), "MORTGAGESUBQUOTENUMBER")
                                    
Set xmlLoanComponentList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGESUBQUOTENUMBER = '" & strMortgageSubQuoteNumber & "']")
                                    
For Each xmlLoanComponentNode In xmlLoanComponentList
    Set xmlMortgageProductCodeNodeList = xmlLoanComponentNode.selectNodes("MORTGAGEPRODUCT")
    
    For Each xmlMortgageProductCodeNode In xmlMortgageProductCodeNodeList
        'get interesttype = 'F' and interestratetypesequencenumber = 1
        Set xmlInterestRateTypeNode = xmlMortgageProductCodeNode.selectSingleNode("INTERESTRATETYPE[@RATETYPE='F' and @INTERESTRATETYPESEQUENCENUMBER='1' ]")
        
        If Not xmlInterestRateTypeNode Is Nothing Then
            strInterestRateTypeStartDate = GetAttributeText(xmlInterestRateTypeNode, "STARTDATE")
            If strInterestRateTypeStartDate <> "" Then
                dtInterestRateTypeStartDate = Format(strInterestRateTypeStartDate, "yyyy/mm/dd")
            End If
            
            strMortgageProductStartDate = GetAttributeText(xmlMortgageProductCodeNode, "STARTDATE")
            If strMortgageProductStartDate <> "" Then
                dtMortgageProductStartDate = Format(strMortgageProductStartDate, "yyyy/mm/dd")
            End If
            
            If dtInterestRateTypeStartDate >= dtMortgageProductStartDate Then
                blnToAddComponent = True
                Exit For
            End If
        End If
    Next

'    If blnToAddComponent Then
'        strLoanComponentStartDate = GetAttributeText(xmlLoanComponentNode, "STARTDATE")
'        If strLoanComponentStartDate <> "" Then
'            dtLoanComponentStartDate = Format(strLoanComponentStartDate, "yyyy/mm/dd")
'        End If
'        If dtLoanComponentStartDate >= dtMortgageProductStartDate Then
'            AddLoanComponent xmlLoanComponentNode, vxmlResponseNode
'        End If
'    End If
Next
                                
'ComponentAmount
'InterestOnlyPortion
'FixedRatePeriodEndDate
'InterestRate
'EarlyRepaymentCharges
'PercentageOfBalance
'FixedRateProductCode
'PortingFlag
LogDetails 3, "SetupComponentRecord : End"
    
SetupComponentRecordExit:
    
    Set xmlDoc = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlLoanComponentNode = Nothing
    Set xmlMortgageProductCodeNode = Nothing
    Set xmlInterestRateTypeNode = Nothing
        
    CheckError strFunctionName, TypeName(Me)
    
End Sub





Private Sub XXX_AddLoanComponent(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    AddLoanComponent
' Author:       Sam Dali
' Date:         31/08/05
' Purpose:      Adding a loan component to Component record
' Input parameters:
'               vxmlRequestNode - XML Request, a loan component node
'               vxmlResponseNode - XML Response to append to
'------------------------------------------------------------------------

On Error GoTo AddLoanComponentExit

Const strFunctionName As String = "AddLoanComponent"

Dim xmlDoc As FreeThreadedDOMDocument40

Dim xmlComponentRecordNode As IXMLDOMNode

'all nodes for Component Record
Dim xmlComponentAmountNode As IXMLDOMNode
Dim xmlInterestOnlyPortionNode As IXMLDOMNode
Dim xmlFixedRatePeriodEndDateNode As IXMLDOMNode
Dim xmlInterestRateNode As IXMLDOMNode
Dim xmlEarlyRepaymentChargesNode As IXMLDOMNode
Dim xmlPercentageOfBalanceNode As IXMLDOMNode
Dim xmlFixedRateProductCodeNode As IXMLDOMNode
Dim xmlPortingFlagNode As IXMLDOMNode
Dim xmlSearchNode As IXMLDOMNode 'temp node
'not there in DataCompletion.xls, added to identify a loan component, LOANCOMPONENTSEQUENCENUMBER
Dim strLoanComponentSequenceNumber As String
'Dim xmlLoanComponentSequenceNumberNode As IXMLDOMNode

LogDetails 3, "AddLoanComponent : Start"

Dim strRuleName As String
strRuleName = "ComponentRecord"

Set xmlDoc = New FreeThreadedDOMDocument40
xmlDoc.validateOnParse = False
xmlDoc.setProperty "NewParser", True

strLoanComponentSequenceNumber = GetAttributeText(vxmlRequestNode, "LOANCOMPONENTSEQUENCENUMBER")

Set xmlComponentRecordNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentRecord", "")

'set the sequence number
'Set xmlLoanComponentSequenceNumberNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT,"LoanComponentSequenceNumber")
'xmlLoanComponentSequenceNumberNode.Text = strLoanComponentSequenceNumber
'xmlComponentRecordNode.appendChild xmlLoanComponentSequenceNumberNode

'ComponentAmount
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".1"
Set xmlComponentAmountNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ComponentAmount", "")
xmlComponentAmountNode.Text = GetAttributeText(vxmlRequestNode, "TOTALLOANCOMPONENTAMOUNT")
xmlComponentRecordNode.appendChild xmlComponentAmountNode

'InterestOnlyPortion
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".2"
Dim strRepaymentMethod As String
Set xmlInterestOnlyPortionNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InterestOnlyPortion", "")

strRepaymentMethod = GetMandatoryAttributeText(vxmlRequestNode, "REPAYMENTMETHOD")
If (strRepaymentMethod = "3") Then
    xmlInterestOnlyPortionNode.Text = CSafeLng(GetMandatoryAttributeText(vxmlRequestNode, "INTERESTONLYELEMENT"))
Else
    If (strRepaymentMethod = "1") Then
        xmlInterestOnlyPortionNode.Text = CSafeLng(GetMandatoryAttributeText(vxmlRequestNode, "TOTALLOANCOMPONENTAMOUNT"))
    Else
        xmlInterestOnlyPortionNode.Text = "0"
    End If
End If
xmlComponentRecordNode.appendChild xmlInterestOnlyPortionNode

'FixedRatePeriodEndDate
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".3"
Dim strManualPortedInd As String, strDate As String
Set xmlFixedRatePeriodEndDateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FixedRatePeriodEndDate", "")
 
strManualPortedInd = GetAttributeText(vxmlRequestNode, "MANUALPORTEDLOANIND")
If strManualPortedInd = "1" Then
    'MORTGAGELOANGUID is the primary key, ProductCode and applicationumber need not b matched
    'ported loan, so mortgageLoan is mandatory
    xmlFixedRatePeriodEndDateNode.Text = GetAttributeText(vxmlRequestNode.selectSingleNode("MORTGAGELOAN[@MORTGAGELOANGUID='" & GetAttributeText(vxmlRequestNode, "MORTGAGELOANGUID") & "']"), "CURRENTRATEEXPIRYDATE")
Else
    Set xmlSearchNode = vxmlRequestNode.selectSingleNode("MORTGAGEPRODUCT/INTERESTRATETYPE")
    If Not xmlSearchNode Is Nothing Then
        strDate = GetAttributeText(xmlSearchNode, "INTERESTRATEENDDATE")
        
        If strDate = "" Then 'interestratetype.enddate is not populated
            Dim dtCompletionDate As Date
            ' dtCompletionDate = CSafeDate("01/01/2005")
            'dtCompletionDate = CSafeDate(GetAttributeText(vxmlRequestNode.parentNode.parentNode.parentNode.selectSingleNode("REPORTONTITLE"), "COMPLETIONDATE"))
            dtCompletionDate = CSafeDate(GetAttributeText(vxmlRequestNode.parentNode.parentNode.selectSingleNode("REPORTONTITLE"), "COMPLETIONDATE"))
            
            'set to ReportOnTitle.CompletionDate + InterestRateType.InterestRatePeriod
            strDate = CStr(DateAdd("m", CSafeDbl(GetAttributeText(xmlSearchNode, "INTERESTRATEPERIOD")), dtCompletionDate))

        End If
    End If
    
    If strDate <> "" Then
        xmlFixedRatePeriodEndDateNode.Text = Format(strDate, "yyyy/mm/dd")
    End If

End If
xmlComponentRecordNode.appendChild xmlFixedRatePeriodEndDateNode

'InterestRate
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".4"
Set xmlInterestRateNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "InterestRate", "")
xmlInterestRateNode.Text = GetAttributeAsDouble(vxmlRequestNode, "RESOLVEDRATE")
xmlComponentRecordNode.appendChild xmlInterestRateNode

'EarlyRepaymentCharges
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".5"
Dim xmlRedemptionFeeBandNodeList As IXMLDOMNodeList
Dim xmlMortgageProductNode As IXMLDOMNode, xmlRedemptionFeeSetNode As IXMLDOMNode, xmlRedemptionFeeBandNode As IXMLDOMNode
Dim strEarlyPaymentCharges As String, strMortgageProductCode As String

strMortgageProductCode = GetAttributeText(vxmlRequestNode, "MORTGAGEPRODUCTCODE")
Set xmlMortgageProductNode = vxmlRequestNode.selectSingleNode("MORTGAGEPRODUCT[@MORTGAGEPRODUCTCODE='" & _
                                strMortgageProductCode & "']")

If Not xmlMortgageProductNode Is Nothing Then
    Set xmlRedemptionFeeSetNode = vxmlRequestNode.selectSingleNode("MORTGAGEPRODUCT/REDEMPTIONFEESET[@REDEMPTIONFEESET='" & _
                                GetAttributeText(xmlMortgageProductNode, "REDEMPTIONFEESET") & "']")
End If
                                
If Not xmlRedemptionFeeSetNode Is Nothing Then
    Set xmlRedemptionFeeBandNodeList = vxmlRequestNode.selectNodes("MORTGAGEPRODUCT/REDEMPTIONFEEBAND[@REDEMPTIONFEESET='" & _
                                GetAttributeText(xmlRedemptionFeeSetNode, "REDEMPTIONFEESET") & "']")
    If Not xmlRedemptionFeeBandNodeList Is Nothing Then
        For Each xmlRedemptionFeeBandNode In xmlRedemptionFeeBandNodeList
            strEarlyPaymentCharges = strEarlyPaymentCharges & GetAttributeText(xmlRedemptionFeeBandNode, "REDEMPTIONFEESTEPNUMBER") & "," & _
                                     GetAttributeText(xmlRedemptionFeeBandNode, "PERIOD") & "," & _
                                     GetAttributeText(xmlRedemptionFeeBandNode, "FEEPERCENTAGE") & ","
        Next
        
        strEarlyPaymentCharges = Left(strEarlyPaymentCharges, Len(strEarlyPaymentCharges) - 1) 'cut the last ","
    End If
End If

Set xmlEarlyRepaymentChargesNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "EarlyRepaymentCharges", "")
xmlEarlyRepaymentChargesNode.Text = strEarlyPaymentCharges
xmlComponentRecordNode.appendChild xmlEarlyRepaymentChargesNode

'PercentageOfBalance
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".6"
Set xmlPercentageOfBalanceNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PercentageOfBalance", "")
xmlPercentageOfBalanceNode.Text = GetAttributeText(xmlMortgageProductNode, "FREECHARGEPERCENTAGE")
If xmlPercentageOfBalanceNode.Text = "" Then 'MORTGAGEPRODUCT.FREECHARGEPERCENTAGE not populated
    xmlPercentageOfBalanceNode.Text = "0"
End If
xmlComponentRecordNode.appendChild xmlPercentageOfBalanceNode

'FixedRateProductCode
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".7"
Set xmlFixedRateProductCodeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FixedRateProductCode", "")
xmlFixedRateProductCodeNode.Text = strMortgageProductCode
xmlComponentRecordNode.appendChild xmlFixedRateProductCodeNode

'PortingFlag
gstrRuleNumber = strRuleName & strLoanComponentSequenceNumber & ".8"
Set xmlPortingFlagNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "PortingFlag", "")
xmlPortingFlagNode.Text = GetAttributeText(vxmlRequestNode, "MANUALPORTEDLOANIND")

If Len(xmlPortingFlagNode.Text) < 1 Then
    xmlPortingFlagNode.Text = "0"
End If

xmlComponentRecordNode.appendChild xmlPortingFlagNode

vxmlResponseNode.appendChild xmlComponentRecordNode

LogDetails 3, "AddLoanComponent : End"

AddLoanComponentExit:
    
    Set xmlDoc = Nothing
    Set xmlComponentAmountNode = Nothing
    Set xmlInterestOnlyPortionNode = Nothing
    Set xmlFixedRatePeriodEndDateNode = Nothing
    Set xmlInterestRateNode = Nothing
    Set xmlEarlyRepaymentChargesNode = Nothing
    Set xmlPercentageOfBalanceNode = Nothing
    Set xmlFixedRateProductCodeNode = Nothing
    Set xmlPortingFlagNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub

'AW 28/03/07    EP2_2125    -   Re-write to use Mortgage One Off Costs
Private Sub SetupFees(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupFees
' Author:       Sam Dali
' Date:         12/09/05
' Purpose:      Create Fees data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    On Error GoTo SetupFeesExit
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    
    Dim xmlMortgageFeeDetailsNode As IXMLDOMNode
    Dim xmlFeeNode As IXMLDOMNode
    Dim xmlFeeAmountNode As IXMLDOMNode
    Dim xmlFeeStatusNode As IXMLDOMNode
    Dim xmlFeeTypeNodeList As IXMLDOMNodeList
    Dim xmlFeeTypeNode As IXMLDOMNode
    Dim xmlFeeItemNode As IXMLDOMNode
    
    Dim strFeeType As String
    Dim strFeeAmount As String

    Dim intPayment As Integer
    Dim intReturn As Integer
    
    LogDetails 3, "SetupFees : Start"
    
    Const strFunctionName As String = "SetupFees"
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strRuleName = "Fees"
    
    Set xmlMortgageFeeDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "MortgageFeeDetails", "")
    
    'Only costs added to loan are retrieved from database
    Set xmlFeeTypeNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/MORTGAGEONEOFFCOST")
        
    For Each xmlFeeItemNode In xmlFeeTypeNodeList
    
        strFeeType = GetAttributeText(xmlFeeItemNode, "MORTGAGEONEOFFCOSTTYPE")
            
        'create the next Fee node
        Set xmlFeeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Fee", "")
        Set xmlFeeTypeNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FeeType", "")
        Set xmlFeeAmountNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FeeAmount", "")
        Set xmlFeeStatusNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FeeStatus", "")
            
        'Get FeeType description
        xmlFeeTypeNode.Text = GetComboText("OneOffCost", CSafeInt(strFeeType))
        xmlFeeNode.appendChild xmlFeeTypeNode
        
        'Get Fee Amount
        strFeeAmount = GetAttributeText(xmlFeeItemNode, "AMOUNT")
        If Not strFeeAmount = "" Then
            xmlFeeAmountNode.Text = strFeeAmount
                    
            'Add to the Application Fee amount - Used for calculating the Net Advance Amount
            m_dblApplicationFeeAmount = m_dblApplicationFeeAmount + CSafeDbl(strFeeAmount)
        End If
        
        xmlFeeNode.appendChild xmlFeeAmountNode
        xmlFeeStatusNode.Text = "Added to Loan"
        xmlFeeNode.appendChild xmlFeeStatusNode
        
        xmlMortgageFeeDetailsNode.appendChild xmlFeeNode
    Next
    
    vxmlResponseNode.appendChild xmlMortgageFeeDetailsNode
    
    LogDetails 3, "SetupFees : End"
    
SetupFeesExit:
    Set xmlDoc = Nothing
    Set xmlMortgageFeeDetailsNode = Nothing
    Set xmlFeeNode = Nothing
    Set xmlFeeTypeNode = Nothing
    Set xmlFeeAmountNode = Nothing
    Set xmlFeeStatusNode = Nothing
    Set xmlFeeTypeNodeList = Nothing
    Set xmlFeeTypeNode = Nothing

End Sub
'AW 28/03/07    EP2_2125    -   End

Private Sub XXX_SetupFees(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupFees
' Author:       Sam Dali
' Date:         12/09/05
' Purpose:      Create Fees data in completion block
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
On Error GoTo SetupFeesExit
Dim xmlValuationFeeDetailsNode As IXMLDOMNode
Dim xmlFeeRefundAmountNode As IXMLDOMNode
Dim xmlFeeAmountNotRefundableNode As IXMLDOMNode

Dim xmlApplicationFeeTypeNodeList As IXMLDOMNodeList, xmlApplicationFeeTypeNode As IXMLDOMNode, xmlLatestApplicationFeeTypeNode As IXMLDOMNode
Dim xmlFeePaymentNodeList As IXMLDOMNodeList
Dim xmlFeePaymentNode As IXMLDOMNode
Dim xmlRFVNode As IXMLDOMNode
Dim xmlPaidNode As IXMLDOMNode
Dim xmlLatestFeePaymentNode As IXMLDOMNode
Dim strFeeType As String, strValFeeType As String
Dim strFeeTypeSequenceNumber As String, strMaxFeeTypeSequenceNumber As String
Dim strPaymentEventID As String
Dim strNTRAmount As String
Dim blnP As Boolean, blnRFV As Boolean, blnNTR As Boolean, blnAR As Boolean
Dim intNTRAmount As Integer
Dim intIncome As Integer
Dim strAmountPaid As String
Dim strRefundAmount As String
Dim strPaidEvent As String
Dim strPaidValueID As String, strPOPAValueID As String
Dim xmlNodeList As IXMLDOMNodeList, xmlNodeList1 As IXMLDOMNodeList

LogDetails 3, "SetupFees : Start"

Const strFunctionName As String = "SetupFees"
Dim xmlDoc As FreeThreadedDOMDocument40
Set xmlDoc = New FreeThreadedDOMDocument40
xmlDoc.validateOnParse = False
xmlDoc.setProperty "NewParser", True

Dim strRuleName As String
strRuleName = "Fees"

Set xmlValuationFeeDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "ValuationFeeDetails", "")

'FeeRefundAmount
Set xmlFeeRefundAmountNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FeeRefundAmount", "")

'FeeAmountNotRefundable
Set xmlFeeAmountNotRefundableNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "FeeAmountNotRefundable", "")

'ApplicationType
Dim strApplicationType As String
strApplicationType = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "TYPEOFAPPLICATION")

If Not strApplicationType = "" Then
    If IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "FT") Or IsValidationType("TypeOfMortgage", CSafeInt(strApplicationType), "HM") Then
        Set xmlApplicationFeeTypeNodeList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/FEES/APPLICATIONFEETYPE")
        
        'get the highest FeeTypeSequenceNumber for Valuation Fee
        'The applicationFeeType with highest FeeTypeSequenceNumber is the latest one
        strMaxFeeTypeSequenceNumber = "0"
    
        For Each xmlApplicationFeeTypeNode In xmlApplicationFeeTypeNodeList
            strFeeType = GetAttributeText(xmlApplicationFeeTypeNode, "FEETYPE")
            
            If Not strFeeType = "" Then
                'if it is a valuation fee
                If IsValidationType("OneOffCost", CSafeInt(strFeeType), "VAL") Then
                    strValFeeType = strFeeType
                    strFeeTypeSequenceNumber = GetAttributeText(xmlApplicationFeeTypeNode, "FEETYPESEQUENCENUMBER")
                    If Not strFeeTypeSequenceNumber = "" Then
                        'Pick up the payment events with P or POPA validation types
                        strPaidValueID = ComboHelper.GetFirstComboValueId("PaymentEvent", "P")
                        strPOPAValueID = ComboHelper.GetFirstComboValueId("PaymentEvent", "POPA")
                        'Search for an occurance of a payment event with a validation type of 'P' or 'POPA'
                        Set xmlNodeList = xmlApplicationFeeTypeNode.selectNodes("FEEPAYMENT[@PAYMENTEVENT='" & strPaidValueID & "']")
                        Set xmlNodeList1 = xmlApplicationFeeTypeNode.selectNodes("FEEPAYMENT[@PAYMENTEVENT='" & strPOPAValueID & "']")
                        
                        If xmlNodeList.length <> 0 Or xmlNodeList1.length <> 0 Then
                        
                            'clean the variables
                            blnP = False
                            blnRFV = False
                            blnNTR = False
                            blnAR = False
                            strAmountPaid = ""
                            strRefundAmount = "" 'we only refund the last valuation fee
                            
                            'loop through all the fees
                            Set xmlFeePaymentNodeList = xmlApplicationFeeTypeNode.selectNodes("FEEPAYMENT[@FEETYPE='" & strFeeType & "']")
                            For Each xmlFeePaymentNode In xmlFeePaymentNodeList
    
                                strPaymentEventID = GetAttributeText(xmlFeePaymentNode, "PAYMENTEVENT")
                                If Not strPaymentEventID = "" Then
                                    If IsValidationType("PaymentEvent", CSafeInt(strPaymentEventID), "P") Or _
                                    IsValidationType("PaymentEvent", CSafeInt(strPaymentEventID), "POPA") Then
                                        blnP = True
                                        Set xmlPaidNode = xmlFeePaymentNode.selectSingleNode("PAYMENTRECORD")
                                    End If
    
                                    If IsValidationType("PaymentEvent", CSafeInt(strPaymentEventID), "RFV") Then
                                        blnRFV = True
                                        Set xmlRFVNode = xmlFeePaymentNode
                                    End If
    
                                    If IsValidationType("PaymentEvent", CSafeInt(strPaymentEventID), "NTR") Then
                                        blnNTR = True
                                    End If
                                    
                                    If IsValidationType("PaymentEvent", CSafeInt(strPaymentEventID), "AR") Then
                                        blnAR = True
                                    End If
                                    
                                End If
                            Next
                            
                            'Check to see if any fees need to be taken into income
                            'A fee is taken into income if it is paid and it does not have a 'AR' flag next to it
                            If blnP = True And blnAR <> True Then
                                'Find the amount which was paid
                                strAmountPaid = GetAttributeText(xmlPaidNode, "AMOUNT")
                                If strAmountPaid <> "" Then
                                    intIncome = intIncome + CSafeInt(strAmountPaid)
                                End If
                                If blnNTR = False Then
                                    'Take the latest refund node to calculate the refund amount
                                    strRefundAmount = GetAttributeText(xmlRFVNode, "REFUNDAMOUNT")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        'Now, subtract the refund amount from the amount to be taken into income, if the last fee payment does not have the
        'Already Refunded flag
        
        If blnAR = False Then
            If strRefundAmount <> "" Then
                intIncome = intIncome - CSafeInt(strRefundAmount)
            End If
        End If
        If strRefundAmount = "" Then
            strRefundAmount = "0"
        End If
        xmlFeeRefundAmountNode.Text = strRefundAmount
        xmlFeeAmountNotRefundableNode.Text = CStr(intIncome)
    End If
End If

xmlValuationFeeDetailsNode.appendChild xmlFeeRefundAmountNode
xmlValuationFeeDetailsNode.appendChild xmlFeeAmountNotRefundableNode

vxmlResponseNode.appendChild xmlValuationFeeDetailsNode

LogDetails 3, "SetupFees : End"
    
SetupFeesExit:
    Set xmlValuationFeeDetailsNode = Nothing
    Set xmlFeeRefundAmountNode = Nothing
    Set xmlFeeAmountNotRefundableNode = Nothing

End Sub


Function GetMaximumAttributeValue(ByVal vxmlNodeList As IXMLDOMNodeList, ByVal strAttributeName As String, _
                                    Optional ByVal vntUpperLimit As Variant) As String
'------------------------------------------------------------------------
' Procedure:    GetMaximumAttributeValue
' Author:       Mark Coleman
' Date:         24/09/01
' Purpose:
' Input parameters:
'               vxmlNodeList - Node list within which to find the maximum
'               strAttributeName - Attribute for which maximum value is to be found
' Output parameters:
'               Maximum value found
'------------------------------------------------------------------------

On Error GoTo GetMaximumAttributeValueExit

Const strFunctionName As String = "GetMaximumAttributeValue"

Dim xmlNode As IXMLDOMNode
Dim strValue As String
Dim vntValue As Variant
Dim vntMax As Variant


    If Not vxmlNodeList Is Nothing Then
        For Each xmlNode In vxmlNodeList
            vntValue = GetAttributeText(xmlNode, strAttributeName)
            If vntValue > vntMax Then
                If Not IsMissing(vntUpperLimit) Then
                    If vntValue <= vntUpperLimit Then
                        vntMax = vntValue
                    End If
                Else
                    vntMax = vntValue
                End If
            End If
        Next
    End If
    
    GetMaximumAttributeValue = CStr(vntMax)

GetMaximumAttributeValueExit:
    Set xmlNode = Nothing
    
    
    CheckError strFunctionName, TypeName(Me)


End Function


Function GetMinimumAttributeValue(ByVal vxmlNodeList As IXMLDOMNodeList, ByVal strAttributeName As String, _
                                    Optional ByVal vntLowerLimit As Variant) As String
'------------------------------------------------------------------------
' Procedure:    GetMinimumAttributeValue
' Author:       Mark Coleman
' Date:         24/09/01
' Purpose:
' Input parameters:
'               vxmlNodeList - Node list within which to find the minimum
'               strAttributeName - Attribute for which minimum value is to be found
' Output parameters:
'               Minimum value found
'------------------------------------------------------------------------

On Error GoTo GetMinimumAttributeValueExit

Const strFunctionName As String = "GetMinimumAttributeValue"

Dim xmlNode As IXMLDOMNode
Dim strValue As String
Dim vntValue As Variant
Dim vntMin As Variant


    If Not vxmlNodeList Is Nothing Then
        For Each xmlNode In vxmlNodeList
            vntValue = GetAttributeText(xmlNode, strAttributeName)
            If vntValue < vntMin Then
                If Not IsMissing(vntLowerLimit) Then
                    If vntValue >= vntLowerLimit Then
                        vntMin = vntValue
                    End If
                Else
                    vntMin = vntValue
                End If
            ElseIf IsEmpty(vntMin) Then
                vntMin = vntValue
            End If
        Next
    End If
    
    GetMinimumAttributeValue = CStr(vntMin)

GetMinimumAttributeValueExit:
    Set xmlNode = Nothing
    
    
    CheckError strFunctionName, TypeName(Me)


End Function


'moved this function from XMLHelper
Private Function CreateErrorResponse() As String
    
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlReponseElem As IXMLDOMElement
    Dim xmlErrorElem As IXMLDOMElement
    Dim xmlDescriptionElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlReponseElem = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlReponseElem
    xmlReponseElem.setAttribute "TYPE", "APPERR"

    Set xmlErrorElem = xmlDoc.createElement("ERROR")
    xmlReponseElem.appendChild xmlErrorElem
        
    Set xmlElement = xmlDoc.createElement("NUMBER")
    xmlElement.Text = Err.Number
    xmlErrorElem.appendChild xmlElement
    LogDetails 2, "Error Number : " & Err.Number
    
    Set xmlElement = xmlDoc.createElement("SOURCE")
    xmlElement.Text = "omAdminRulesBO"
        
    If Len(gstrRuleNumber) > 0 Then
        xmlElement.Text = xmlElement.Text & ".Rule" & gstrRuleNumber
    End If
    xmlErrorElem.appendChild xmlElement
    LogDetails 2, "Error Source : " & xmlElement.Text
    
    Set xmlDescriptionElem = xmlDoc.createElement("DESCRIPTION")
    xmlDescriptionElem.Text = Err.Description
    xmlErrorElem.appendChild xmlDescriptionElem
    LogDetails 2, "Error Description : " & Err.Description
    
    Set xmlElement = xmlDoc.createElement("VERSION")
    xmlElement.Text = App.Comments
    xmlErrorElem.appendChild xmlElement
            
    CreateErrorResponse = xmlReponseElem.xml
        
    Set xmlDoc = Nothing
    Set xmlReponseElem = Nothing
    Set xmlErrorElem = Nothing
    Set xmlDescriptionElem = Nothing
    Set xmlElement = Nothing

End Function
'AW 28/03/07    EP2_2125
Private Sub SetupCashbackDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                        ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    SetupCashbackDetails
' Author:       A Woodhouse
' Date:         28/03/2007
' Purpose:      Create Cashback Details data in completion block
'               Only to be displayed if there is a Cashback amount
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to i.e. CompletionBlock
'------------------------------------------------------------------------
    
    On Error GoTo SetupCashbackDetailsExit
    
    Const strFunctionName As String = "SetupCashbackDetails"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRuleName As String
    Dim xmlLoanComponentList As IXMLDOMNodeList
    Dim xmlLoanIncentiveList As IXMLDOMNodeList
    Dim xmlLoanComponentNode As IXMLDOMNode
    Dim xmlMortgageIncentiveNode As IXMLDOMNode
    Dim xmlCashbackDetailsNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim strMortgageSubQuoteNumber As String
    Dim lngTotalIncentive As Long
    
    LogDetails 3, "SetupCashbackDetails : Start "
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    lngTotalIncentive = 0

    strRuleName = "SetupCashbackDetails"
    'Cashback amount
    gstrRuleNumber = strRuleName & ".2"
    
    'The Cashback amount should be the sum of Incentive Amounts across loan components
    strMortgageSubQuoteNumber = GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/QUOTATION[@QUOTATIONNUMBER='" & _
                                   GetMandatoryAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND"), "ACCEPTEDQUOTENUMBER") & "']"), "MORTGAGESUBQUOTENUMBER")
                                   
    Set xmlLoanComponentList = vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGESUBQUOTENUMBER = '" & strMortgageSubQuoteNumber & "']")

    For Each xmlLoanComponentNode In xmlLoanComponentList
        
        Set xmlLoanIncentiveList = xmlLoanComponentNode.selectNodes("MORTGAGEINCENTIVE")
        
        For Each xmlMortgageIncentiveNode In xmlLoanIncentiveList
            lngTotalIncentive = lngTotalIncentive + CSafeLng(GetAttributeText(xmlMortgageIncentiveNode, "INCENTIVEAMOUNT"))
        Next
    Next
    
    If lngTotalIncentive > 0 Then
    
        Set xmlCashbackDetailsNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "Cashback", "")
        Set xmlTempNode = xmlDoc.createNode(MSXML2.NODE_ELEMENT, "CashbackAmount", "")
        xmlTempNode.Text = CStr(lngTotalIncentive)
        xmlCashbackDetailsNode.appendChild xmlTempNode
        vxmlResponseNode.appendChild xmlCashbackDetailsNode
        
    End If
    
    LogDetails 3, "SetupCashbackDetails : End "
    
    
SetupCashbackDetailsExit:

    Set xmlDoc = Nothing
    Set xmlCashbackDetailsNode = Nothing
    Set xmlMortgageIncentiveNode = Nothing
    Set xmlTempNode = Nothing
    
    CheckError strFunctionName, TypeName(Me)
End Sub
'AW 28/03/07    EP2_2125    -   End
