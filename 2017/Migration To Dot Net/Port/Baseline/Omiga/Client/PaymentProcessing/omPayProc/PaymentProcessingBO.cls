VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PaymentProcessingBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProcessingBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   PaymentProcessing Business Object.

'Dependencies:
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MC                 Created
'MV     21/03/01    SYS2072 Changed GetPaymentSummary Proc
'SR     23/05/01    SYS2298 Modified methods 'CancelBalance' and 'GetPaymentSummary'
'CL     06/06/01    SYS2340 Added mehtod 'UpdateFirstPaymentDate'
'SR     01/08/01    SYS2545 Modified method FindFeePaymentTotals - Do not consider 'RebateOrAddition'
'                   While calculating the Outstanding Amount
'SR     01/08/01    SYS2552 Modified methods 'CreateDisbursement' and 'UpdateDisbursement'
'SR     06/09/01    SYS2412 New methods 'DeletePayeeHistory' and 'CreateReturnOfFunds'
'                   Modified methods 'GetPaymentSummary' and 'FindDisbursementPaymentList'
'MC     19/11/01    SYS3027 Improve error handling in batch process
'DS     26/11/01    SYS2857 Update to DoRequest for CompleteInterfacing
'MC     12/12/01    SYS3027 Further improve error handling in batch process
'SR     19/12/2001  SYS2547 New methods UpdateFirstPaymentDateAndBankBuildingSoc and UpdateFirstPaymentDateAndCreateBankBuildingSoc
'MC     20/12/2001  SYS3505 Add audit reporting to batch
'MC     20/12/2001  SYS3541 Write error details to Event Log
'DM     03/01/02    SYS3185 Implement CompleteInterfacing
'RF     03/01/02    Complete SYS3185 (Implement CompleteInterfacing)
'RF     13/02/02    SYS4079 Handle record not found correctly in RunBatchProgramme
'JLD    19/02/02    SYS4115 Reset local variable.
'RF     25/02/02    SYS4155 Error in RunBatchProgramme
'RF     28/02/02    SYS4197 Faulty error handling in OnMessage. Also an error in call to
'                           CompleteBatch in PaySanctionedPayment.
'RF     27/02/02    SYS4193 Add new workaround method TemporaryCompleteInterfacing
'PSC    01/03/02    SYS4214 Add enhance error info in RunBatchProgramme
'JLD    13/03/02    SYS4114 if there are no fees, still carry on in GetPaymentSummary
'JLD    19/03/02    SYS4177 Only subtract fees already deducted from the balance, not those yet to be deducted.
'SA     19/03/02    SYS4287 Call to BatchscheduleBo.PrintDocuments in CompleteBatch
'SA     05/04/02    SYS4340 Call to PrintDocs was missing full request string
'GD     09/04/2002  SYS4290 Change OnMessage to SetComplete(), and return success.
'DM     12/04/2002  SYS4389 Added CompleteBatch to the DoRequest method.
'SA     18/04/02    SYS4271 Whilst investigating this AQR, found some debug code which needed removing
'SA     23/05/02    SYS4516 In InterfacePaidPayment_Exit need to check for final transaction
'------------------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements MessageQueueComponentVC2

Private gobjContext As ObjectContext
Private gobjTxBO As IPaymentProcessingTxBO
Private gobjDO As IPAymentProcessingDO


'------------------------------------------------------------------------------------
'   BMIDS
'
'   Prog    Date        AQR         Description
'   AW      23/08/02    BM029       Added CreateChequeAuditRecord, SanctionPrintCheques,
'                                   SetPaidChequePayments
'   PSC     20/09/2002  BMIDS00480  Amend RunBatchProgramme to process Intitial Advances first
'   PSC     01/10/2002  BMIDS00501  Improve error reporting in complete batch
'   DPF     03/10/2002  BM037       Have amended GetPaymentSummary method to check for manual incentive
'                                   amount in MortgageSubQuote table and changed way we build Oustanding
'                                   Incentives amount to pass back to GUI
'   PSC     08/10/2002  BMIDS00595  Amend to process Incentive Releases
'   PSC     14/10/2002  BMIDS00337  Amend UpdateDisbursement to get correct Building
'                                   Society details
'   PSC     22/10/2002  BMIDS00700  Amend GetPaymentSummary to handle Manual Incentives  and
'                                   drawdowns  correctly
'   GD      15/10/2002  BMIDS00614  Allow reprinting of cheques
'   GD      07/11/2002  BMIDS00870  Ensure ChequeNumbers are calc'd correctly for RePrints
'   PSC     11/11/2002  BMIDS00599  Take into account retention returns when calculating how much
'                                   retention has been paid in GetPaymentSummary
'   GD      15/11/2002  BMIDS0922   Locking issues in cheque printing, plus change datatype used to hold cheque numbers
'   PSC     28/11/2002  BMIDS01099  Amend to process incentive returns
'   MV      11/02/2003  BM0336      Amended GetPaymentSummary
'   GD      22/04/2003  BM0198      Added New Method : SaveCostsFeesDisbursements()
'   INR     14/10/2003  BMIDS628    Changes to logging in PaySanctionedPayment & InterfacePaidPayment
'   MC      15/06/2004  BMIDS763    GetProductSwitchFee() & GetTTFeeAmount() Added
'   GHun    15/07/2004  BMIDS793    Amended GetPaymentSummary
'   SR      06/08/2004  BMIDS813    Modified methods GetProductSwitchFee, GetTTFeeAmount
'   KRW     25/10/2004  BMIDS930    Changed GetPaymentsummary to ignore zero manual incentives
'   JD      22/11/2004  BMIDS604    Amended CompleteBatch to look for Cancelled Batch error.
'   JD      29/03/2005  BMIDS977    Added ApplyApplicationFeeRebatesAndAdditions
'------------------------------------------------------------------------------------
'MARS Specific History:
'   HMA     25/08/2005  MAR28       Add GetValuationFee and PayFees requests.
'   HM      26/09/2005  MAR29       Add ValidateCompletionDate, ValidateSolicitor, SetAdvanceDate, SetPaymentMethod, CreateDelayedCompletionTask requests.
'   HMA     07/10/2005  MAR75       Amend PayFees for Direct Gateway call.
'   DRC     02/03/2006 MAR1334      Add CancelDisbursements method
'   JD      04/05/2006  MAR1616     PayFees - pass fee types and amounts
'   JD      10/05/2006  MAR1754     PayFees - send customernumber to interface
'   JD      27/06/2006  MAR1844     PayFees - treat return code 'F' as an 'A'
'   JD      12/07/2006  MAR1844     PayFees - treat return code 'F' as a 'D' (not an A)
'-------------------------------------------------------------------------------------------------------
'Epsom Specific History:
'   PSC     26/02/2007  EP2_1347    Refund of valuation processing
'   PSC     13/03/2007  EP2_1347    Treat valuation refunds the same as incentives
'-------------------------------------------------------------------------------------------------------


Private Function MessageQueueComponentVC2_OnMessage( _
    ByVal vstrXmlConfig As String, _
    ByVal vstrXMLData As String) As Long
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'RF     03/01/02    Changes for SYS3185 (Implement CompleteInterfacing) - never abort the
'                   queue txn if the target method is CompleteInterfacing.
'RF     28/02/02    SYS4197 Faulty error handling in OnMessage.
'GD     09/04/2002  SYS4290 Change OnMessage to SetComplete(), and return success - above 2 AQRs backed out
'------------------------------------------------------------------------------------------
On Error GoTo MessageQueueComponentVC2_OnMessageVbErr

    Const cstrFunctionName As String = "MessageQueueComponentVC2_OnMessage"
    
    Dim strErrMsg As String
    Dim strResponse As String


    strResponse = omPayProcRequest(vstrXMLData)
    
    'Check the response
    errCheckXMLResponse strResponse, True
    
    'Complete so remove message from queue
    MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS
    
MessageQueueComponentVC2_OnMessage_Exit:

    Exit Function
    
MessageQueueComponentVC2_OnMessageVbErr:


    strErrMsg = "Error processing message." & vbCrLf & _
        "Number: " & CStr(Err.Number) & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        "Description: " & Err.Description
    App.LogEvent strErrMsg, vbLogEventTypeError
    
    
    gobjContext.SetComplete
    MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS

    GoTo MessageQueueComponentVC2_OnMessage_Exit

End Function


Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function omPayProcRequest(ByVal vstrXmlIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo omPayProcRequestVbErr

    Const strFunctionName As String = "omPayProcRequest"
    Dim xmlindoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute

    If gobjContext Is Nothing Then
        errThrowError strFunctionName, oeMTSNotFound
    End If
    gobjContext.SetComplete
    
    Set gobjTxBO = Nothing
    
    Set xmlindoc = New FreeThreadedDOMDocument40
    xmlindoc.validateOnParse = False
    xmlindoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    xmlindoc.async = False
    xmlOutDoc.async = False
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    Debug.Print vstrXmlIn
    xmlindoc.loadXML vstrXmlIn
    
    Set xmlRequestNode = xmlGetMandatoryNode(xmlindoc, "REQUEST")
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        Set xmlOperationNodeList = xmlindoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    omPayProcRequest = xmlOutDoc.xml
    Debug.Print omPayProcRequest
    
omPayProcRequestExit:
  
    Set gobjTxBO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlindoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

omPayProcRequestVbErr:
    
    omPayProcRequest = errCreateErrorResponse()
    Debug.Print omPayProcRequest
    
    GoTo omPayProcRequestExit
    
End Function

Private Sub DoRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DoRequestExit

    Const strFunctionName As String = "DoRequest"
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    strOperation = UCase$(strOperation)
    
    Select Case strOperation
        Case "CREATEAPPLICATIONFEETYPES"
            CreateApplicationFeeTypes vxmlRequestNode, vxmlResponseNode
            
        Case "CREATEFEETYPEPAYMENT"
            CreateFeeTypePayment vxmlRequestNode, vxmlResponseNode
            
        Case "FINDFEETYPELIST"
            FindFeeTypeList vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEFEETYPEPAYMENT"
            UpdateFeeTypePayment vxmlRequestNode, vxmlResponseNode
        
        Case "DELETEAPPLICATIONFEETYPES"
            DeleteApplicationFeeTypes vxmlRequestNode, vxmlResponseNode
        
        Case "DELETEFEETYPEPAYMENT"
            DeleteFeeTypePayment vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEAPPLICATIONFEETYPE"
            UpdateApplicationFeeType vxmlRequestNode, vxmlResponseNode
        
        Case "CREATEAPPLICATIONFEETYPE"
            CreateApplicationFeeType vxmlRequestNode, vxmlResponseNode
        
        Case "FINDPAYMENTHISTORYLIST"
            FindPaymentHistoryList vxmlRequestNode, vxmlResponseNode
        
        Case "FINDPAYEEHISTORYLIST"
            FindPayeeHistoryList vxmlRequestNode, vxmlResponseNode
        
        Case "FINDPAYMENTSFORPAYEELIST"
            FindPaymentsForPayeeList vxmlRequestNode, vxmlResponseNode
        
        Case "CREATEPAYEEHISTORYDETAILS"
            CreatePayeeHistoryDetails vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEPAYEEHISTORYDETAILS"
            UpdatePayeeHistoryDetails vxmlRequestNode, vxmlResponseNode
        
        Case "CANCELBALANCE"
            CancelBalance vxmlRequestNode, vxmlResponseNode
        
        Case "CREATEDISBURSEMENT"
            CreateDisbursement vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEDISBURSEMENT"
            UpdateDisbursement vxmlRequestNode, vxmlResponseNode
        
        Case "FINDFULLPAYMENTLIST"
            FindFullPaymentList vxmlRequestNode, vxmlResponseNode
        
        Case "GETPAYMENTSUMMARY"
            GetPaymentSummary vxmlRequestNode, vxmlResponseNode
        
        '*=[MC]BMIDS763/CC075
        Case "GETPRODUCTSWITCHFEE"
            Call GetProductSwitchFee(vxmlRequestNode, vxmlResponseNode)
        
        Case "GETTTFEEAMOUNT"
            Call GetTTFeeAmount(vxmlRequestNode, vxmlResponseNode)
        '*=SECTION END
        
        Case "FINDPAYMENTSFORBATCH"
            FindPaymentsForBatch vxmlRequestNode, vxmlResponseNode
        
        Case "GETFULLPAYMENTDETAILS"
            GetFullPaymentDetails vxmlRequestNode, vxmlResponseNode
            
        Case "FINDAPPLICATIONTHIRDPARTYLIST"
            FindApplicationThirdPartyList vxmlRequestNode, vxmlResponseNode
        
        Case "FINDFEEPAYMENTTOTALS"
            FindFeePaymentTotals vxmlRequestNode, vxmlResponseNode
        
        Case "FINDSANCTIONINGLIST"
            FindSanctioningList vxmlRequestNode, vxmlResponseNode
        
        Case "RUNBATCHPROGRAMME"
            RunBatchProgramme vxmlRequestNode, vxmlResponseNode
        
        Case "INTERFACEPAIDPAYMENT"
            InterfacePaidPayment vxmlRequestNode, vxmlResponseNode
        
        Case "PAYSANCTIONEDPAYMENT"
            PaySanctionedPayment vxmlRequestNode, vxmlResponseNode
            
        Case "GETINTERFACINGDETAILS"
            GetInterfacingDetails vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEFIRSTPAYMENTDATE"
            UpdateFirstPaymentDate vxmlRequestNode, vxmlResponseNode
                    
        Case "FINDLOANCOMPONENTPAYMENTLIST"
            FindLoanComponentPaymentList vxmlRequestNode, vxmlResponseNode
            
        Case "DELETEPAYEEHISTORY"
            DeletePayeeHistory vxmlRequestNode, vxmlResponseNode
            
        Case "CREATERETURNOFFUNDS"
            CreateReturnOfFunds vxmlRequestNode, vxmlResponseNode
            
        Case "COMPLETEINTERFACING"
            CompleteInterfacing vxmlRequestNode, vxmlResponseNode
            
        ' SR 18/12/01 : SYS2547
        Case "UPDATEFIRSTPAYMENTDATEANDCREATEBANKBUILDINGSOC"
            UpdateFirstPaymentDateAndCreateBankBuildingSoc vxmlRequestNode, vxmlResponseNode
            
        Case "UPDATEFIRSTPAYMENTDATEANDBANKBUILDINGSOC"
            UpdateFirstPaymentDateAndBankBuildingSoc vxmlRequestNode, vxmlResponseNode
        ' SR 18/12/01 : SYS2547 END
        
'BMIDS628 INR Removed duplicate entry
'        Case "COMPLETEINTERFACING"
'            CompleteInterfacing vxmlRequestNode, vxmlResponseNode
        
        'DM     12/04/2002  SYS4389 Added CompleteBatch to the DoRequest method.
        Case "COMPLETEBATCH"
            CompleteBatch vxmlRequestNode, vxmlResponseNode
        ' END DM     12/04/2002  SYS4389 Added CompleteBatch to the DoRequest method.
            
        '   AW  23/08/02    BM029
        Case "CREATECHEQUEAUDITRECORD"
            CreateChequeAuditRecord vxmlRequestNode, vxmlResponseNode
            
        Case "SANCTIONPRINTCHEQUES"
            SanctionPrintCheques vxmlRequestNode, vxmlResponseNode
            
        Case "SETPAIDCHEQUEPAYMENTS"
            SetPaidChequePayments vxmlRequestNode, vxmlResponseNode
        '   AW  23/08/02    BM029 END
        
        'GD BM0198 START
        Case "SAVECOSTSFEESDISBURSEMENTS"
            SaveCostsFeesDisbursements vxmlRequestNode, vxmlResponseNode
        'GD BM0198 END
        
        'JD BMIDS977
        Case "APPLYAPPLICATIONFEEREBATESANDADDITIONS"
            ApplyApplicationFeeRebatesAndAdditions vxmlRequestNode, vxmlResponseNode
        'JD BMIDS977 end
        
        'MAR28
        Case "GETVALUATIONFEE"
            GetValuationFee vxmlRequestNode, vxmlResponseNode
            
        Case "PAYFEES"
            PayFees vxmlRequestNode, vxmlResponseNode
        
        'MAR29
        Case "VALIDATECOMPLETIONDATE"
            ValidateCompletionDate vxmlRequestNode, vxmlResponseNode

        Case "VALIDATESOLICITOR"
            ValidateSolicitor vxmlRequestNode, vxmlResponseNode

        Case "SETADVANCEDATE"
            SetAdvanceDate vxmlRequestNode, vxmlResponseNode

        Case "SETPAYMENTMETHOD"
            SetPaymentMethod vxmlRequestNode, vxmlResponseNode
        
        Case "RESETPOSTCOMPLETIONTASKS"
            ResetPostCompletionTasks vxmlRequestNode, vxmlResponseNode
        '/MAR29
        ' PSC 15/11/2005 MAR140 - Start
        Case "GETPOSTCOMPLETIONCASETASKSLIST"
            GetPostCompletionCaseTasksList vxmlRequestNode, vxmlResponseNode
        ' PSC 15/11/2005 MAR140 - End
        ' DRC  02/03/2006 MAR1334 - Start
        Case "CANCELDISBURSEMENTS"
            CancelDisbursements vxmlRequestNode, vxmlResponseNode
        ' DRC  02/03/2006 MAR1334 - End
        Case Else
            errThrowError "DoRequest", oeNotImplemented, vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text & _
                " not implemented"
    
    End Select

DoRequestExit:

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetProductSwitchFee(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo GetProductSwitchFeeExit

    Const strFunctionName As String = "GetProductSwitchFee"
    
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    Dim strMortgageSQNumber As String
    Dim xmlNode As IXMLDOMNode
        
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "GETPRODUCTSWITCHFEE")
    
    'Get Application Number, Application Fact Find Number and Mortgage Sub Quote Number
    strAppNumber = xmlGetNodeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetNodeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strAppNumber) = 0 Or strAppFFNumber = 0 Then
        errThrowError strFunctionName, oeMissingElementValue, _
                "APPLICATIONNUMBER=" & strAppNumber & "  APPLICATIONFACTFINDNUMBER=" & strAppFFNumber
    End If
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    Call gobjDO.GetProductSwitchFee(xmlNode, vxmlResponseNode)
    
GetProductSwitchFeeExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Sub

'MAR28 Add function to get Valuation Fee
Private Sub GetValuationFee(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo GetValuationFeeExit

    Const strFunctionName As String = "GetValuationFee"
    
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    Dim strMortgageSQNumber As String
    Dim xmlNode As IXMLDOMNode
        
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "GETVALUATIONFEE")
    
    'Get Application Number, Application Fact Find Number and Mortgage Sub Quote Number
    strAppNumber = xmlGetNodeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetNodeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    strMortgageSQNumber = xmlGetNodeText(xmlNode, "MORTGAGESUBQUOTENUMBER")
    
    If (Len(strAppNumber) = 0) Or (Len(strAppFFNumber) = 0) Or (Len(strMortgageSQNumber) = 0) Then
        errThrowError strFunctionName, oeMissingElementValue, _
          "APPLICATIONNUMBER=" & strAppNumber & "  APPLICATIONFACTFINDNUMBER=" & strAppFFNumber & " MORTGAGESUBQUOTENUMBER=" & strMortgageSQNumber
    End If
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    Call gobjDO.GetValuationFee(xmlNode, vxmlResponseNode)
    
GetValuationFeeExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Sub

'Add function to Pay Fees
Private Sub PayFees(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo PayFeesExit

    Const strFunctionName As String = "PayFees"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strValue As String

    Dim xmlNode As IXMLDOMNode
    Dim xmlDGRequestNode As IXMLDOMNode
    Dim xmlDGRequestDetailsNode As IXMLDOMNode
    Dim xmlCreditCardNode As IXMLDOMNode
    Dim xmlCardPaymentResponseNode As IXMLDOMNode
    Dim xmlPaymentResponseNode As IXMLDOMNode
    Dim xmlNewRequestNode As IXMLDOMNode
    Dim xmlFeePaymentNode As IXMLDOMNode
    Dim xmlPaymentRecordNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlFeeListResponse As IXMLDOMNode
    Dim xmlFeeListResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlInPaymentList As IXMLDOMNodeList
    Dim intPaymentEvent As Integer
    Dim strPaymentSequenceNumber As String
    Dim strPaymentResult As String
    Dim strAppNo As String
    Dim strTransactionMode As String
    Dim strTransactionReference As String
    Dim strAuthorisationReference As String
    Dim strCardFailResultCode As String
    Dim strCardFailReason As String
    Dim strPayment As String
    Dim strCardPaymentResponse As String
    Dim xmlPaymentDetailsNode As IXMLDOMNode 'MAR1616
    Dim xmlInFeeList As IXMLDOMNodeList 'MAR1616
    Dim xmlInFee As IXMLDOMNode 'MAR1616
    Dim xmlCustomerNode As IXMLDOMNode 'MAR1754
    Dim strResponse As String 'MAR1754
    Dim strCIFNumber As String 'MAR1754
             
    'MAR397 Use Late Binding
    Dim objCardPayment As Object
    Dim objCustomer As Object 'MAR1754
    'Dim objCardPayment As omCardPayment.CardPaymentBO
    'Set objCardPayment = gobjContext.CreateInstance("omCardPayment.CardPaymentBO")
     
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    Set xmlFeeListResponseElem = xmlDoc.createElement("RESPONSE")
    Set xmlFeeListResponse = xmlDoc.appendChild(xmlFeeListResponseElem)
    
    
    'Check for mandatory data
    Set xmlCreditCardNode = xmlGetMandatoryNode(vxmlRequestNode, "CreditCardDetails")
    strAppNo = xmlGetMandatoryNode(vxmlRequestNode, "ApplicationNumber").Text
    strTransactionMode = xmlGetMandatoryNode(vxmlRequestNode, "TransactionMode").Text
    
    'MAR1754 find the OtherSystemCustomerNumber to pass as CIF
    Set xmlCustomerNode = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    Set objCustomer = gobjContext.CreateInstance("omCust.CustomerBO")

    'Check that the object has been created correctly
    If (objCustomer Is Nothing) Then
        errThrowError strFunctionName, oeObjectNotCreatable, "omCust could not be created"
    End If

    strResponse = objCustomer.GetCustomerNumbers(xmlCustomerNode.xml)
    xmlDoc.loadXML strResponse
    Set xmlCustomerNode = xmlDoc.selectSingleNode("RESPONSE/CUSTOMER")
    strCIFNumber = xmlGetNodeText(xmlCustomerNode, "OTHERSYSTEMCUSTOMERNUMBER")
    If (strCIFNumber = "") Then
        errThrowError strFunctionName, 170, "CIF number neccessary to pay fees."
    End If
    
    'Build up the XML to send to the Direct Gateway
    Set xmlNewRequestNode = xmlGetRequestNode(vxmlRequestNode)
    
    'Send the customer number MAR1754
    Set xmlCustomerNode = xmlDoc.createElement("CUSTOMER")
    Set xmlTempNode = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlTempNode.Text = strCIFNumber
    xmlCustomerNode.appendChild xmlTempNode
    xmlNewRequestNode.appendChild xmlCustomerNode
    
    Set xmlDGRequestNode = xmlDoc.createElement("TakeCardPaymentRequestType")
    
    Set xmlDGRequestDetailsNode = xmlDoc.createElement("TakeCardPaymentRequestDetails")
    xmlDGRequestDetailsNode.appendChild xmlCreditCardNode.cloneNode(True)
  
    Set xmlTempNode = xmlDoc.createElement("MasterAccountNumber")
    xmlDGRequestDetailsNode.appendChild xmlTempNode
    xmlTempNode.Text = "0"
    
    Set xmlTempNode = xmlDoc.createElement("MortgageApplicationNumber")
    xmlDGRequestDetailsNode.appendChild xmlTempNode
    xmlTempNode.Text = strAppNo
    
    Set xmlTempNode = xmlDoc.createElement("PaymentMechanism")
    xmlDGRequestDetailsNode.appendChild xmlTempNode
    xmlTempNode.Text = "CreditCard"
    
    Set xmlTempNode = xmlDoc.createElement("TransactionMode")
    xmlDGRequestDetailsNode.appendChild xmlTempNode
    xmlTempNode.Text = strTransactionMode
    
    Set xmlTempNode = xmlDoc.createElement("TransactionType")
    xmlDGRequestDetailsNode.appendChild xmlTempNode
    xmlTempNode.Text = "Auth"
    
    Set xmlTempNode = xmlDoc.createElement("Pipeline")
    xmlDGRequestDetailsNode.appendChild xmlTempNode
    xmlTempNode.Text = "Payment"
      
    xmlDGRequestNode.appendChild xmlDGRequestDetailsNode.cloneNode(True)
    
    'MAR1616 Create PaymentDetails node
    Set xmlInFeeList = vxmlRequestNode.selectNodes("./FEE")
    For Each xmlInFee In xmlInFeeList
        Set xmlPaymentDetailsNode = xmlDoc.createElement("PaymentDetails")
        Set xmlTempNode = xmlDoc.createElement("FeeType")
        xmlTempNode.Text = xmlGetNodeText(xmlInFee, "FEETYPE")
        xmlPaymentDetailsNode.appendChild xmlTempNode
        
        Set xmlTempNode = xmlDoc.createElement("Amount")
        xmlTempNode.Text = xmlGetNodeText(xmlInFee, "FEEAMOUNT")
        xmlPaymentDetailsNode.appendChild xmlTempNode
        
        xmlDGRequestNode.appendChild xmlPaymentDetailsNode
    Next


    xmlNewRequestNode.appendChild xmlDGRequestNode.cloneNode(True)

    'Send a Card Payment request across the link.
    Set objCardPayment = gobjContext.CreateInstance("omCardPayment.CardPaymentBO")

    'Check that the object has been created correctly
    If (objCardPayment Is Nothing) Then
        errThrowError strFunctionName, oeObjectNotCreatable, "omCardPayment could not be created"
    End If

    strCardPaymentResponse = objCardPayment.RunTakeCardPayment(xmlNewRequestNode.xml)
    xmlDoc.loadXML strCardPaymentResponse
    
    Set xmlCardPaymentResponseNode = xmlDoc.selectSingleNode("RESPONSE")
    
    xmlChangeNodeName xmlCardPaymentResponseNode, "RESPONSE", "PAYMENTRESPONSE"
 
    vxmlResponseNode.appendChild xmlCardPaymentResponseNode.cloneNode(True)
    
    Set xmlPaymentResponseNode = xmlDoc.selectSingleNode(".//TakeCardPaymentResponseDetails")

    'If a response has been returned, save the details
    If Not xmlPaymentResponseNode Is Nothing Then
        
        ' PaymentResult     A = Success
        '                   D = Not successfully authorised
        '                   E = Card details invalid or incomplete
        '                   F = D MAR1844
    
        ' Create a payment if necessary

        strPaymentResult = xmlPaymentResponseNode.selectSingleNode("PaymentResult").Text
        strTransactionReference = xmlPaymentResponseNode.selectSingleNode("TransactionReference").Text
        strAuthorisationReference = xmlPaymentResponseNode.selectSingleNode("AuthorisationReference").Text
        strCardFailResultCode = strPaymentResult
        strCardFailReason = xmlPaymentResponseNode.selectSingleNode("UpFrontPaymentReason").Text
        strPayment = GetFirstComboValueId("PaymentEvent", "P")

        'Save the details to the database

        Set xmlNewRequestNode = xmlGetRequestNode(vxmlRequestNode)

        Set xmlTempNode = xmlDoc.createElement("APPLICATIONFEETYPE")
        xmlSetAttributeValue xmlTempNode, "APPLICATIONNUMBER", strAppNo
        xmlSetAttributeValue xmlTempNode, "_COMBOLOOKUP_", "1"

        xmlNewRequestNode.appendChild xmlTempNode.cloneNode(True)

        'Get Fee Type data
        Call FindFeeTypeList(xmlNewRequestNode, xmlFeeListResponse)

        Set xmlNodeList = xmlFeeListResponse.selectNodes("APPLICATIONFEETYPE")
        If xmlNodeList.length < 1 Then
            errThrowError strFunctionName, oeRecordNotFound
        End If

        Set xmlNewRequestNode = xmlGetRequestNode(vxmlRequestNode)

        For Each xmlNode In xmlNodeList

            Set xmlInPaymentList = xmlNode.selectNodes("FEEPAYMENT")
            For Each xmlFeePaymentNode In xmlInPaymentList

                Set xmlNewRequestNode = xmlGetRequestNode(vxmlRequestNode)

                intPaymentEvent = xmlGetAttributeAsInteger(xmlFeePaymentNode, "PAYMENTEVENT")
                If IsValidationType("PaymentEvent", intPaymentEvent, "AW") Then
                    strPaymentSequenceNumber = xmlGetAttributeText(xmlFeePaymentNode, "PAYMENTSEQUENCENUMBER")

                    xmlSetAttributeValue xmlFeePaymentNode, "CARDFAILRESULTCODE", strCardFailResultCode
                    xmlSetAttributeValue xmlFeePaymentNode, "CARDFAILREASON", strCardFailReason

                    'If the payment was successful, set the Payment Event to "Payment"
                    If (strPaymentResult = "A") Then 'MAR1844
                        xmlSetAttributeValue xmlFeePaymentNode, "PAYMENTEVENT", strPayment
                    End If

                    xmlNewRequestNode.appendChild xmlFeePaymentNode.cloneNode(True)

                    'Get the corresponding Payment Record
                    Set xmlPaymentRecordNode = xmlNode.selectSingleNode("PAYMENTRECORD[@PAYMENTSEQUENCENUMBER='" & strPaymentSequenceNumber & "']")

                    xmlSetAttributeValue xmlPaymentRecordNode, "TRANSACTIONREFERENCE", strTransactionReference
                    xmlSetAttributeValue xmlPaymentRecordNode, "AUTHORISATIONREFERENCE", strAuthorisationReference

                    xmlNewRequestNode.appendChild xmlPaymentRecordNode.cloneNode(True)

                    Call UpdateFeeTypePayment(xmlNewRequestNode, xmlTempResponse)

                End If

            Next

        Next
    End If
    
PayFeesExit:
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlDGRequestNode = Nothing
    Set xmlDGRequestDetailsNode = Nothing
    Set xmlCreditCardNode = Nothing
    Set xmlCardPaymentResponseNode = Nothing
    Set xmlPaymentResponseNode = Nothing
    Set xmlNewRequestNode = Nothing
    Set xmlFeePaymentNode = Nothing
    Set xmlPaymentRecordNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlFeeListResponse = Nothing
    Set xmlFeeListResponseElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlInPaymentList = Nothing
    Set objCardPayment = Nothing
    Set xmlPaymentDetailsNode = Nothing 'MAR1616
    Set xmlInFeeList = Nothing 'MAR1616
    Set xmlInFee = Nothing
    Set xmlCustomerNode = Nothing 'MAR1754
    Set objCustomer = Nothing 'MAR1754
 
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub GetTTFeeAmount(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo GetTTFeeAmountExit
    
    Const strFunctionName As String = "GetTTFeeAmount"
    
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    Dim strMortgageSQNumber As String
    Dim xmlNode As IXMLDOMNode
        
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "GETTTFEEAMOUNT")
    
    'Get Application Number, Application Fact Find Number and Mortgage Sub Quote Number
    strAppNumber = xmlGetNodeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetNodeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strAppNumber) = 0 Or strAppFFNumber = 0 Then
        errThrowError strFunctionName, oeMissingElementValue, _
                "APPLICATIONNUMBER=" & strAppNumber & "  APPLICATIONFACTFINDNUMBER=" & strAppFFNumber
    End If
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    Call gobjDO.GetTTFeeAmount(xmlNode, vxmlResponseNode)
    
GetTTFeeAmountExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Sub


Private Sub CreateApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationFeeTypesExit

    Const strFunctionName As String = "CreateApplicationFeeTypes"
    
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    Dim strMortgageSQNumber As String
    Dim xmlNode As IXMLDOMNode
    
    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFEETYPE")
    
    'Get Application Number, Application Fact Find Number and Mortgage Sub Quote Number
    strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    strMortgageSQNumber = xmlGetMandatoryAttributeText(xmlNode, "MORTGAGESUBQUOTENUMBER")
    
'    If Len(strAppNumber) = 0 Or Len(strAppFFNumber) = 0 Or _
'        Len(strMortgageSQNumber) = 0 Then
'
'         errThrowError strFunctionName, oeMissingPrimaryTag, "Missing attribute: APPLICATION = " & strAppNumber & _
'                "APPLICATIONFACTFINDNUMBER = " & strAppFFNumber & _
'                "MORTAGESUBQUOTENUMBER = " & strMortgageSQNumber
'
'    End If
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreateApplicationFeeTypes vxmlRequestNode, vxmlResponseNode
    
CreateApplicationFeeTypesExit:
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub GetInterfacingDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo GetInterfacingDetails_Exit

Const strFunctionName As String = "GetInterfacingDetails"

Dim objPayProcDO As IPAymentProcessingDO

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlDisbNode As IXMLDOMNode
Dim xmlLCPNode As IXMLDOMNode
Dim xmlDPNode As IXMLDOMNode
Dim xmlPayeeNode As IXMLDOMNode
Dim strPaymentType As String        ' PSC 20/09/2002 BMIDS00480

' PSC 13/03/2007 EP2_1347 - Start
Dim blnIsIncentive As Boolean
Dim blnIsReturnOfFunds As Boolean
Dim blnIsValuationRefund As Boolean
' PSC 13/03/2007 EP2_1347 - End

    'Validation
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Find Payment Record(s)
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    objPayProcDO.FindPaymentRecordList xmlNode, vxmlResponseNode
    
    Set xmlNodeList = vxmlResponseNode.selectNodes("PAYMENTRECORD")
    If xmlNodeList.length > 0 Then
        Set xmlDisbNode = xmlDoc.createElement("DISBURSEMENTPAYMENT")
        Set xmlLCPNode = xmlDoc.createElement("LOANCOMPONENTPAYMENT")
        Set xmlPayeeNode = xmlDoc.createElement("PAYEEHISTORY")
        
        For Each xmlNode In xmlNodeList
            'Get Disbursement Payment
            xmlCopyMandatoryAttribute xmlNode, xmlDisbNode, "APPLICATIONNUMBER"
            xmlCopyMandatoryAttribute xmlNode, xmlDisbNode, "PAYMENTSEQUENCENUMBER"
            objPayProcDO.FindDisbursementPaymentList xmlDisbNode, xmlNode
            
            'Disbursement Payment has to exist for this payment
            Set xmlDPNode = xmlGetMandatoryNode(xmlNode, "DISBURSEMENTPAYMENT")
            
            ' PSC 20/09/2002 BMIDS00480 - Start
            strPaymentType = xmlGetMandatoryAttributeText(xmlDPNode, "PAYMENTTYPE")
            
            ' Only need disbursement record for Return Of Funds
            ' PSC 13/03/2007 EP2_1347 - Start
            blnIsReturnOfFunds = IsValidationType("PaymentType", CSafeInt(strPaymentType), "N")
            blnIsIncentive = IsValidationType("PaymentType", CSafeInt(strPaymentType), "C")
            blnIsValuationRefund = IsValidationType("PaymentType", CSafeInt(strPaymentType), "VALREFUND")
            ' PSC 13/03/2007 EP2_1347 - End
            
            ' PSC 13/03/2007 EP2_1347
            If blnIsReturnOfFunds = False Then
                
                ' PSC 08/10/2002 BMIDS00595 - Start
                ' Incentive Releases and valuation refunds do not have loan component payments
                ' PSC 13/03/2007 EP2_1347
                If blnIsIncentive = False And blnIsValuationRefund = False Then
                    'Get Loan Component Payment(s)
                    xmlCopyMandatoryAttribute xmlNode, xmlLCPNode, "APPLICATIONNUMBER"
                    xmlCopyMandatoryAttribute xmlNode, xmlLCPNode, "PAYMENTSEQUENCENUMBER"
                    objPayProcDO.FindLoanComponentPaymentList xmlDisbNode, xmlNode
                End If
                ' PSC 08/10/2002 BMIDS00595 - End
                
                'Get Payee Details
                
                xmlCopyMandatoryAttribute xmlNode, xmlPayeeNode, "APPLICATIONNUMBER"
                xmlCopyMandatoryAttribute xmlDPNode, xmlPayeeNode, "PAYEEHISTORYSEQNO"
                objPayProcDO.FindPayeeHistoryList xmlPayeeNode, xmlNode
            End If
            ' PSC 20/09/2002 BMIDS00480 - End
        Next
    Else
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetInterfacingDetails_Exit:
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlDisbNode = Nothing
    Set xmlLCPNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub InterfacePaidPayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo InterfacePaidPayment_Exit

Const strFunctionName As String = "InterfacePaidPayment"

Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
Dim objBatchSched As BatchScheduleBO
Dim objPayProcDO As IPAymentProcessingDO

Dim xmlBatchSched As IXMLDOMNode
Dim xmlPayment As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlFinishTrans As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strAppNo As String
Dim strValue As String
Dim strResponse As String
Dim strBatchNumber As String
Dim strBatchRunNumber As String
Dim strBatchAuditGuid As String
Dim strPaySeqNumber As String
Dim blnInitialising As Boolean
Dim blnRetry As Boolean
Dim strErrMsg As String 'SYS3541 MDC 20/12/2001

    blnInitialising = True
    
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    
    'Validation
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    
    Set xmlPayment = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    strAppNo = xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER")
    strPaySeqNumber = xmlGetMandatoryAttributeText(xmlPayment, "PAYMENTSEQUENCENUMBER")

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Find the BatchAuditGuid for this payment
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlTempRequest = xmlDoc.createElement("PAYPROCBATCHAUDIT")
    xmlSetAttributeValue xmlTempRequest, "APPLICATIONNUMBER", strAppNo
    xmlSetAttributeValue xmlTempRequest, "PAYMENTSEQUENCENUMBER", strPaySeqNumber
    xmlSetAttributeValue xmlTempRequest, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlTempRequest, "BATCHRUNNUMBER", strBatchRunNumber
    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    objPayProcDO.FindPaymentAuditList xmlTempRequest, xmlTempResponse
    Set xmlNode = xmlGetMandatoryNode(xmlTempResponse, "PAYPROCBATCHAUDIT")
    strBatchAuditGuid = xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID")
    'Append the Batch Audit Guid to the request for use in the NTTxBO
    Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
    xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGuid
    vxmlRequestNode.appendChild xmlNode
    blnInitialising = False
    
    'Call the InterfacePaidPayment method on the NTTxBO therefore executing in separate transaction
    Set objPayProcNTTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
    objPayProcNTTxBO.InterfacePaidPayment vxmlRequestNode, vxmlResponseNode
    'BMIDS628
    strValue = xmlGetAttributeText(vxmlResponseNode, "TYPE")

InterfacePaidPayment_Exit:

    'If error is DB failure or occurs during initialisation we cannot update the batch status
    'If Err.Number <> 0 And (Err = DB_FAIL Or blnInitialising) Then
    If Err.Number <> 0 And blnInitialising Then
    '   Raise the error back so that the message queue is informed and message moved to holding queue
       Err.Raise Err.Number, Err.Source, Err.Description
    'BMIDS628 Check if objPayProcNTTxBO.InterfacePaidPayment failed
    ElseIf ((Err.Number <> 0) Or (StrComp(Trim$(strValue), "FAIL") = 0)) Then
    
        'BMIDS628 Do this all in the one place, moved from start of InterfacePaidPayment_Exit:
        Dim strBatchMsg As String
        Dim strErrNumber As String
        Dim strErrSource As String
        Dim strErrDesc As String
        
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Audit Guid: " & strBatchAuditGuid
        strErrMsg = strErrMsg & vbCrLf & "Application Number: " & strAppNo
        strErrMsg = strErrMsg & vbCrLf & "Payment Seq No: " & strPaySeqNumber
        strBatchMsg = strErrMsg
        If Err.Number <> 0 Then
            'Need to save these so they can be used in the call to incrementbatchschedfailure
            strErrNumber = Err.Number
            strErrSource = Err.Source
            strErrDesc = Err.Description
            strErrMsg = strErrMsg & vbCrLf & "Error Number: " & strErrNumber
            strErrMsg = strErrMsg & vbCrLf & "Error Source: " & strErrSource
            strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & strErrDesc
        Else
            strErrMsg = strErrMsg & " Call to objPayProcNTTxBO.InterfacePaidPayment failed.  See event log for details."
            strErrDesc = " Call to objPayProcNTTxBO.InterfacePaidPayment failed.  See event log for details."
        End If
        Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strErrMsg)
    
        'Increment Batch Schedule Failure
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "IncrementBatchScheduleFailure"
        xmlSetAttributeValue xmlBatchSched, "BATCHAUDITGUID", strBatchAuditGuid
        xmlSetAttributeValue xmlBatchSched, "ERRORNUMBER", strErrNumber
        xmlSetAttributeValue xmlBatchSched, "ERRORSOURCE", strErrSource
        xmlSetAttributeValue xmlBatchSched, "ERRORDESCRIPTION", strErrDesc
        xmlTempRequest.appendChild xmlBatchSched
        strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
        If errCheckXMLResponse(strResponse, False) <> 0 Then
            '#TASK - Write to event log?
            'Unable to update batchaudit status - Write Reject Report or event log ???
'            Call WriteRejectReport
            'BMIDS628 Log failure to IncrementBatchScheduleFailure
            strBatchMsg = "Failed to Increment BatchSchedule Failure: " & strBatchMsg
            Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strBatchMsg)

        Else
            xmlDoc.loadXML strResponse
            Set xmlFinishTrans = xmlDoc.selectSingleNode(".//BATCHSCHEDULE")
        End If
        'SYS4516 If final transaction, update the batch status
        If Not xmlFinishTrans Is Nothing Then
            If xmlGetAttributeAsBoolean(xmlFinishTrans, "LASTTRANSACTION", "0") Then
                '------------------------------------------------------------------------------------------
                ' call completebatch
                '------------------------------------------------------------------------------------------
                Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlTempRequest.appendChild xmlBatchSched
                CompleteBatch xmlTempRequest, xmlTempResponse
            End If
        End If
    End If
    
    Set objPayProcNTTxBO = Nothing
    Set objBatchSched = Nothing
    Set objPayProcDO = Nothing
    Set xmlBatchSched = Nothing
    Set xmlPayment = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlFinishTrans = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub PaySanctionedPayment( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo PaySanctionedPayment_Exit

    Const strFunctionName As String = "PaySanctionedPayment"
    
    Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
    Dim objBatchSched As BatchScheduleBO
    Dim objPayProcDO As IPAymentProcessingDO
    
    Dim xmlBatchSched As IXMLDOMNode
    Dim xmlPayment As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlFinishTrans As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strAppNo As String
    Dim strValue As String
    Dim strResponse As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strBatchAuditGuid As String
    Dim strPaySeqNumber As String
    Dim blnInitialising As Boolean
    Dim blnStatusUpdateNotReqd As Boolean
    Dim strErrMsg As String 'SYS3541 MDC 20/12/2001

    blnInitialising = True
    
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    
    'Validation
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    
    Set xmlPayment = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    strAppNo = xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER")
    strPaySeqNumber = xmlGetMandatoryAttributeText(xmlPayment, "PAYMENTSEQUENCENUMBER")

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Find the BatchAuditGuid for this payment
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlTempRequest = xmlDoc.createElement("PAYPROCBATCHAUDIT")
    xmlSetAttributeValue xmlTempRequest, "APPLICATIONNUMBER", strAppNo
    xmlSetAttributeValue xmlTempRequest, "PAYMENTSEQUENCENUMBER", strPaySeqNumber
    xmlSetAttributeValue xmlTempRequest, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlTempRequest, "BATCHRUNNUMBER", strBatchRunNumber
    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    objPayProcDO.FindPaymentAuditList xmlTempRequest, xmlTempResponse
    Set xmlNode = xmlGetMandatoryNode(xmlTempResponse, "PAYPROCBATCHAUDIT")
    strBatchAuditGuid = xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID")
    'Append the Batch Audit Guid to the request for use in the NTTxBO
    Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
    xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGuid
    vxmlRequestNode.appendChild xmlNode
    
    blnInitialising = False
    
    'Call the PaySanctionedPayment method on the NTTxBO therefore executing in separate transaction
    Set objPayProcNTTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
    objPayProcNTTxBO.PaySanctionedPayment vxmlRequestNode, vxmlResponseNode
    strValue = xmlGetAttributeText(vxmlResponseNode, "TYPE")
    If ((StrComp(Trim$(strValue), "RETRY") = 0) Or (StrComp(Trim$(strValue), "CANCEL") = 0) _
                 Or (StrComp(Trim$(strValue), "FAIL") = 0)) Then
        'Retrying this transaction as record If it is currently locked OR it failed
        'and we want to increment BatchSchedule Failure.
        blnStatusUpdateNotReqd = True
    End If
    
    If Not blnStatusUpdateNotReqd Then
        'Increment Batch Schedule Success
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "IncrementBatchScheduleSuccess"
        xmlSetAttributeValue xmlBatchSched, "BATCHAUDITGUID", strBatchAuditGuid
        xmlTempRequest.appendChild xmlBatchSched
        strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
        If errCheckXMLResponse(strResponse, False) <> 0 Then
            'Write Reject Report ???
'            Call WriteRejectReport
            'BMIDS628 WriteRejectReport doesn't do anything write failure to event & Batch logs
            strErrMsg = vbCrLf & vbCrLf & "Failed to Increment BatchSchedule Success : " & TypeName(Me) & "_" & strFunctionName
            strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
            strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
            strErrMsg = strErrMsg & vbCrLf & "Batch Audit Guid: " & strBatchAuditGuid
            strErrMsg = strErrMsg & vbCrLf & "Application Number: " & strAppNo
            strErrMsg = strErrMsg & vbCrLf & "Payment Seq No: " & strPaySeqNumber & strResponse
            Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strErrMsg)

        Else
            xmlDoc.loadXML strResponse
            Set xmlFinishTrans = xmlDoc.selectSingleNode(".//BATCHSCHEDULE")
        End If
    End If
    
PaySanctionedPayment_Exit:

    'If error is DB failure or occurs during initialisation we cannot update the batch status
    'If Err.Number <> 0 And (Err = DB_FAIL Or blnInitialising) Then
    If Err.Number <> 0 And blnInitialising Then
    '   Raise the error back so that the message queue is informed and message moved to holding queue
       Err.Raise Err.Number, Err.Source, Err.Description
    ElseIf ((Err.Number <> 0) Or (StrComp(Trim$(strValue), "FAIL") = 0)) Then
    
        'BMIDS628 Do this all in the one place, moved from start of PaySanctionedPayment_Exit:
        Dim strBatchMsg As String
        Dim strErrNumber As String
        Dim strErrSource As String
        Dim strErrDesc As String
        
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Audit Guid: " & strBatchAuditGuid
        strErrMsg = strErrMsg & vbCrLf & "Application Number: " & strAppNo
        strErrMsg = strErrMsg & vbCrLf & "Payment Seq No: " & strPaySeqNumber
        strBatchMsg = strErrMsg
        If Err.Number <> 0 Then
            'Need to save these so they can be used in the call to incrementbatchschedfailure
            strErrNumber = Err.Number
            strErrSource = Err.Source
            strErrDesc = Err.Description
            strErrMsg = strErrMsg & vbCrLf & "Error Number: " & strErrNumber
            strErrMsg = strErrMsg & vbCrLf & "Error Source: " & strErrSource
            strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & strErrDesc
        Else
            strErrMsg = strErrMsg & " Call to objPayProcNTTxBO.PaySanctionedPayment failed. See event log for details."
            strErrDesc = " Call to objPayProcNTTxBO.PaySanctionedPayment failed.  See event log for details."
        End If
        Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strErrMsg)
    
        'Increment Batch Schedule Failure
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "IncrementBatchScheduleFailure"
        xmlSetAttributeValue xmlBatchSched, "BATCHAUDITGUID", strBatchAuditGuid
        xmlSetAttributeValue xmlBatchSched, "ERRORNUMBER", strErrNumber
        xmlSetAttributeValue xmlBatchSched, "ERRORSOURCE", strErrSource
        xmlSetAttributeValue xmlBatchSched, "ERRORDESCRIPTION", strErrDesc
        xmlTempRequest.appendChild xmlBatchSched
        strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
        If errCheckXMLResponse(strResponse, False) <> 0 Then
            'Write Reject Report ???
'            Call WriteRejectReport
            'BMIDS628 Log failure to IncrementBatchScheduleFailure
            strBatchMsg = "Failed to Increment BatchSchedule Failure: " & strBatchMsg & strResponse
            Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strBatchMsg)

        Else
            xmlDoc.loadXML strResponse
            Set xmlFinishTrans = xmlDoc.selectSingleNode(".//BATCHSCHEDULE")
        End If
    End If
    
    'Check if final transaction
    If Not xmlFinishTrans Is Nothing Then
        If xmlGetAttributeAsBoolean(xmlFinishTrans, "LASTTRANSACTION", "0") Then
            
            '------------------------------------------------------------------------------------------
            ' call completebatch
            '------------------------------------------------------------------------------------------
            
            ' RF 19/02/02 Send REQUEST header into CompleteBatch
            Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlTempRequest.appendChild xmlBatchSched
            ' RF 28/02/02 SYS4197
            CompleteBatch xmlTempRequest, xmlTempResponse
        End If
    End If

    Set objPayProcNTTxBO = Nothing
    Set objBatchSched = Nothing
    Set objPayProcDO = Nothing
    Set xmlBatchSched = Nothing
    Set xmlPayment = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlFinishTrans = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CompleteBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CompleteBatch_Exit

    Const strFunctionName As String = "CompleteBatch"
    
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlBatchSched As IXMLDOMNode
    Dim xmlBatchPayment As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlDoc As FreeThreadedDOMDocument40
        
    Dim objBatchSched As BatchScheduleBO
    Dim objPayProcDO As IPAymentProcessingDO
    
    Dim strCompletedValueId As String
    Dim strCWEValueId As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strValue As String
    Dim strResponse As String
    
    'SYS3505 MDC 20/12/2001
    Dim xmlNode As IXMLDOMNode
    Dim xmlPPAuditNodeList As IXMLDOMNodeList
    Dim xmlPPAuditNode As IXMLDOMNode
    Dim intFileNo As Integer
    Dim intValue As Integer
    Dim strTextLine As String
    Dim strBatchLogFile As String
    Dim strBatchLogFilePath As String
    'SYS3505 End
    
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40   'SYS4287
    
    'Validation
    
    'Set xmlBatchSched = vxmlRequestNode 'xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Find List of Payments processed
    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    Set xmlBatchPayment = xmlDoc.createElement("BATCHPAYMENT")
    xmlSetAttributeValue xmlBatchPayment, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlBatchPayment, "BATCHRUNNUMBER", strBatchRunNumber

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    objPayProcDO.FindBatchPaymentList xmlBatchPayment, xmlTempResponse

    
    '#TASK - Printing
    'Evaluate For Printing
    '**** DO PRINTING STUFF HERE!!! ****
    
    'Write Batch Process Report
    Call ProducePaymentReport(xmlBatchSched, xmlTempResponse)
    
    'Validate overall success by checking if any transactions errored
    'Get the combovalue for F - 'Failure'
    strValue = GetFirstComboValueId("BatchAuditStatus", "F")
    Set xmlNodeList = xmlTempResponse.selectNodes("PAYMENTRECORD[@STATUS='" & strValue & "' or @STATUS='']")
    If xmlNodeList.length > 0 Then
        strValue = "CWE"
        'JD BMIDS604 check the error number on the failed applications. We are looking for a 'Cancelled' app
        Dim nCancelledErr As Long
        Dim xmlCancelledNodeList As IXMLDOMNodeList
        nCancelledErr = 7025 + vbObjectError + 512
        Set xmlCancelledNodeList = xmlTempResponse.selectNodes("PAYMENTRECORD[@ERRORNUMBER='" & CStr(nCancelledErr) & "']")
        If xmlCancelledNodeList.length > 0 Then
            strValue = "CA"
        End If
    Else
        strValue = "C"
    End If
    
    'Get the combovalue for status
    strValue = GetFirstComboValueId("BatchScheduleStatus", strValue)
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    
    'Update the Batch Status
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "SetBatchStatus"
    xmlTempRequest.appendChild xmlBatchSched
    xmlSetAttributeValue xmlBatchSched, "STATUS", strValue
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    
    errCheckXMLResponse strResponse, False, vxmlResponseNode

    'SYS3505 MDC 20/12/2001. Complete audit report for rate change
    'Get RateChangeBatchAudit records
    Set xmlTempRequest = xmlDoc.createElement("PAYPROCBATCHAUDIT")
    xmlSetAttributeValue xmlTempRequest, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlTempRequest, "BATCHRUNNUMBER", strBatchRunNumber
    objPayProcDO.FindPaymentAuditList xmlTempRequest, xmlTempResponse
    
    'Get the BatchLogFilePath GlobalParameter Value from GlobalParameter DB Table
    strBatchLogFilePath = GetGlobalParamString("BatchLogFilePath")
    
    'Checking for the Directory Existence to write the outputFile
    'If does not Exist then Create One
    If Dir(strBatchLogFilePath, vbDirectory) = "" Then
        MkDir strBatchLogFilePath
    End If
    
    strBatchLogFile = strBatchLogFilePath & "\Batch_" & strBatchNumber & "_" & strBatchRunNumber
    
    intFileNo = FreeFile
    Open strBatchLogFile & ".log" For Append As #intFileNo
    Print #intFileNo, ""
    Print #intFileNo, "Batch Audit:"
    Print #intFileNo, ""
    Print #intFileNo, "BatchAuditGuid" & vbTab & vbTab & vbTab & vbTab & "Application" & vbTab _
                    & "Payment Sequence Number" & vbTab & "Status"
    Print #intFileNo, "------------------------------------------------------" _
                        & "--------------------------------"
    
    Set xmlNodeList = xmlTempResponse.selectNodes(".//PAYMENTRECORD")
    For Each xmlNode In xmlNodeList
        strValue = xmlGetAttributeText(xmlNode, "BATCHAUDITGUID")
        strTextLine = strValue & vbTab
        
        Set xmlPPAuditNodeList = xmlTempResponse.selectNodes( _
            ".//PAYPROCBATCHAUDIT[@BATCHAUDITGUID='" & strValue & "']")
        If xmlPPAuditNodeList.length > 0 Then
            Set xmlPPAuditNode = xmlPPAuditNodeList.Item(0)

            strValue = xmlGetAttributeText(xmlPPAuditNode, "APPLICATIONNUMBER")
            strTextLine = strTextLine & strValue & vbTab
            strValue = xmlGetAttributeText(xmlPPAuditNode, "PAYMENTSEQUENCENUMBER")
            strTextLine = strTextLine & strValue & vbTab & vbTab & vbTab
        End If
        
        intValue = xmlGetAttributeAsInteger(xmlNode, "STATUS")
        strValue = GetComboText("BatchAuditStatus", intValue)
        strTextLine = strTextLine & strValue
        
        Print #intFileNo, strTextLine
    Next
    
    Print #intFileNo, ""
    Print #intFileNo, "--- End of Report ---"
    Close #intFileNo
    'SYS3505 End
    
    'SYS4287 Now Print the Report
    'SYS4340 Need the full Request string
   ' Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "PrintDocuments"
    
    Set xmlBatchSched = xmlDoc.createElement("BATCH")
    xmlSetAttributeValue xmlBatchSched, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlBatchSched, "BATCHRUNNUMBER", strBatchRunNumber
    xmlTempRequest.appendChild xmlBatchSched
    
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    'Check for errors
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponse xmlTempResponseDoc.xml
    'SYS4287 end
    
CompleteBatch_Exit:

    ' PSC 01/10/2002 BMIDS00501 - Start
    ' If an error has occurred then set the betch status to complete with error and
    ' write out the error info
    If Err.Number <> 0 Then
    
        strValue = "CWE"
    
        strValue = GetFirstComboValueId("BatchScheduleStatus", strValue)
        Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
        
        'Update the Batch Status
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        Set xmlTempRequest = xmlDoc.createElement("REQUEST")
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "SetBatchStatus"
        xmlTempRequest.appendChild xmlBatchSched
        xmlSetAttributeValue xmlBatchSched, "STATUS", strValue
        xmlSetAttributeValue xmlBatchSched, "ERRORNUMBER", CStr(Err.Number)
        xmlSetAttributeValue xmlBatchSched, "ERRORSOURCE", CStr(Err.Source)
        xmlSetAttributeValue xmlBatchSched, "ERRORDESCRIPTION", CStr(Err.Description)
        
        objBatchSched.omBatchRequest xmlTempRequest.xml
    
    End If
    ' PSC 01/10/2002 BMIDS00501 - End

    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlBatchSched = Nothing
    Set xmlBatchPayment = Nothing
    Set xmlNodeList = Nothing
 
    Set xmlDoc = Nothing
    Set objBatchSched = Nothing
    Set objPayProcDO = Nothing
    Set xmlTempResponseDoc = Nothing    'SYS4287
    
    errCheckError strFunctionName, TypeName(Me)

End Sub



Private Sub FindFailedPayments( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'RF     13/02/02    SYS4079 Handle record not found correctly
'------------------------------------------------------------------------------------------
On Error GoTo FindFailedPayments_Exit

    Const cstrFunctionName As String = "FindFailedPayments"
    
    Dim objBatchSched As BatchScheduleBO
    Dim objPayProcDO As IPAymentProcessingDO
    
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    Dim strValue As String
    Dim strValueId As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Validate
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")
    'strValue = xmlGetMandatoryAttributeText(xmlNode, "STATUS")
    
    'Ensure we are looking for failed payments
    strValueId = GetFirstComboValueId("BatchAuditStatus", "F")
    xmlSetAttributeValue xmlNode, "STATUS", strValueId
    
    'Create request
    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "FindBatchAuditList"
    
    'Call FindBatchAuditList to find failed items
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    
    'Validate Response
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    'Get Batch Audit records found and retrieve corresponding PaymentAudit records
    Set xmlNodeList = xmlDoc.selectNodes(".//BATCHAUDIT")
    Set xmlTempRequest = xmlDoc.createElement("PAYPROCBATCHAUDIT")
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    For Each xmlNode In xmlNodeList
        xmlCopyMandatoryAttribute xmlNode, xmlTempRequest, "BATCHAUDITGUID"
        objPayProcDO.FindPaymentAuditList xmlTempRequest, vxmlResponseNode
    Next
    
    
    ' PSC 18/09/2002 SYS5181 - Start
    'Need to manipulate XML returned so it's in the same format as returned from FindPaymentsForBatch
    '<PAYMENTRECORD ApplicationNumber PaymentSequenceNumber>
    'Currently returning <PAYPROCBATCHAUDIT> with attribs of ApplciationNumber, PaymentSequenceNumber, BatchNumber, BatchRunNumber, BatchAuditGuid
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNodeList = vxmlResponseNode.selectNodes(".//PAYPROCBATCHAUDIT")
    For Each xmlNode In xmlNodeList
        xmlChangeNodeName xmlNode, "PAYPROCBATCHAUDIT", "PAYMENTRECORD"
        'Remove the unnecessary attribs
        xmlNode.Attributes.removeNamedItem "BATCHNUMBER"
        xmlNode.Attributes.removeNamedItem "BATCHRUNNUMBER"
        xmlNode.Attributes.removeNamedItem "BATCHAUDITGUID"
    Next
    ' PSC 18/09/2002 SYS5181 - End

    
FindFailedPayments_Exit:
    
    Set objBatchSched = Nothing
    Set objPayProcDO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub FindPaymentAuditList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindPaymentAuditList_Exit

Const strFunctionName As String = "FindPaymentAuditList"

Dim objPayProcDO As IPAymentProcessingDO
Dim xmlNode As IXMLDOMNode

    ' Raise error if no PAYPROCBATCHAUDIT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYPROCBATCHAUDIT")
    
    'Validate BATCHAUDITGUID attribute is specified
    xmlCheckMandatoryAttribute xmlNode, "BATCHAUDITGUID"

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'Find PAYPROCBATCHAUDIT records
    objPayProcDO.FindPaymentAuditList vxmlRequestNode, vxmlResponseNode

FindPaymentAuditList_Exit:
    Set objPayProcDO = Nothing
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePaymentAudit_Exit

Const strFunctionName As String = "UpdatePaymentAudit"

Dim xmlNode As IXMLDOMNode
Dim strValue As String

    ' Raise error if no PAYPROCBATCHAUDIT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYPROCBATCHAUDIT")
    
    'Validate BATCHAUDITGUID attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID")

    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdatePaymentAudit vxmlRequestNode, vxmlResponseNode

UpdatePaymentAudit_Exit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ProducePaymentReport(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo ProducePaymentReport_Exit

Const strFunctionName As String = "ProducePaymentReport"

    'Not implemented yet

ProducePaymentReport_Exit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub RunBatchProgramme( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'RF     13/02/02    SYS4079 Handle record not found correctly
'------------------------------------------------------------------------------------------
On Error GoTo RunBatchProgrammeExit

    Dim objBatchSchedBO As omBatch.BatchScheduleBO
    Dim objMessageQueue As omPayProc.IomMessageQueue
    Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim nodeTempRequest As IXMLDOMNode
    Dim nodeTempResponse As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    Dim nodelistPaymentRecords As IXMLDOMNodeList
    Dim nodeBatch As IXMLDOMNode
    Dim nodeBatchAudit As IXMLDOMNode
    Dim nodeBatchSchedule As IXMLDOMElement
    Dim nodeErrRequest As IXMLDOMNode    'SYS3027
    
    ' PSC 20/09/2002 BMIDS00480 - Start
    Dim xmlPayments As IXMLDOMNode
    Dim strIAValue As String
    Dim intCount As Integer
    ' PSC 20/09/2002 BMIDS00480 - End

    
    Dim strValue As String
    Dim strResponse As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strErrMsg As String 'SYS3541 MDC 20/12/2001
    Dim strApplicationNo As String
    Dim strPaymentSeqNo As String
    Dim strDescription As String
    
    'AM BMIDS722 - Start
    Dim lngBatchInterval As Long
    Dim dtStartTime As Date
    Dim lngCurrentItem As Long
    Dim strExecuteAfterDate As String
    'AM BMIDS722 - End
    
    Const cstrFunctionName As String = "RunBatchProgramme"

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Validation
    Set nodeBatch = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    strBatchNumber = xmlGetMandatoryAttributeText(nodeBatch, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(nodeBatch, "BATCHRUNNUMBER")
    
    Dim blnContinue As Boolean
    blnContinue = True
    
    '------------------------------------------------------------------------------------------
    ' Get batch schedule info
    '------------------------------------------------------------------------------------------

    'Create request to get Batch schedule data
    
    Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
    
    ' RF 13/02/02
    'xmlSetAttributeValue nodeTempRequest, "OPERATION", "GetBatch"
    xmlSetAttributeValue nodeTempRequest, "OPERATION", "GetBatchSchedule"
    
    ' RF 13/02/02
    'Set nodeBatch = xmlDoc.createElement("BATCH")
    Set nodeBatchSchedule = xmlDoc.createElement("BATCHSCHEDULE")
    
    xmlCopyMandatoryAttribute nodeBatch, nodeBatchSchedule, "BATCHNUMBER"
    xmlCopyMandatoryAttribute nodeBatch, nodeBatchSchedule, "BATCHRUNNUMBER"
    nodeTempRequest.appendChild nodeBatchSchedule

    'Create and call the Batch Scheduler component
    Set objBatchSchedBO = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    strResponse = objBatchSchedBO.omBatchRequest(nodeTempRequest.xml)
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    Set nodeTempResponse = xmlDoc.documentElement
    errCheckXMLResponseNode nodeTempResponse, , True
    
    Set nodeBatchSchedule = xmlGetMandatoryNode(nodeTempResponse, "BATCHSCHEDULE")
    
    '------------------------------------------------------------------------------------------
    'Check that Batch has not been cancelled
    '------------------------------------------------------------------------------------------
    
    strValue = xmlGetMandatoryAttributeText(nodeBatchSchedule, "STATUS")
    If strValue = GetFirstComboValueId("BatchScheduleStatus", "CA") Then
        'Batch has been cancelled
        blnContinue = False
    End If
    
    If blnContinue Then
    
        '------------------------------------------------------------------------------------------
        'Find Payment records to be processed
        '------------------------------------------------------------------------------------------
        
        Dim strFailedBatchNo As String
        strFailedBatchNo = xmlGetAttributeText(nodeBatchSchedule, "FAILEDBATCHRUNNUMBER")
        
        Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
        nodeTempRequest.appendChild nodeBatch
        
        Set nodeTempResponse = vxmlResponseNode.cloneNode(False)
        
        Dim lngErrNo As Long
        Dim strErrDesc As String
        Dim strErrSource As String
        
        If Len(Trim$(strFailedBatchNo)) = 0 Then
            
            'We are running a new Batch Job
            
            '------------------------------------------------------------------------------------------
            ' Get batch information
            '------------------------------------------------------------------------------------------
            
            Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlSetAttributeValue nodeTempRequest, "OPERATION", "GetBatch"
            nodeTempRequest.appendChild nodeBatch.cloneNode(True)
            strResponse = objBatchSchedBO.omBatchRequest(nodeTempRequest.xml)
            Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
            Set nodeTempResponse = xmlDoc.documentElement
            errCheckXMLResponseNode nodeTempResponse, , True
            
            '------------------------------------------------------------------------------------------
            ' FindPaymentsForBatch
            '------------------------------------------------------------------------------------------
            
            Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlSetAttributeValue nodeTempRequest, "OPERATION", "GetBatch"
            nodeTempRequest.appendChild xmlGetMandatoryNode(nodeTempResponse, "BATCH")
            
            Set nodeTempResponse = vxmlResponseNode.cloneNode(False)
            
            Err.Clear
            On Error Resume Next
            FindPaymentsForBatch nodeTempRequest, nodeTempResponse
            lngErrNo = Err.Number
            strErrDesc = Err.Description
            strErrSource = Err.Source
            On Error GoTo RunBatchProgrammeExit
            
            If lngErrNo <> 0 Then
                If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
                
                    ' no records found
                    
                    blnContinue = False
                    
                    '------------------------------------------------------------------------------------------
                    ' call completebatch
                    '------------------------------------------------------------------------------------------
                    
                    Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
                    xmlSetAttributeValue nodeTempRequest, "OPERATION", "CompleteBatch"
    
                    Set nodeBatchSchedule = xmlDoc.createElement("BATCHSCHEDULE")
                    xmlSetAttributeValue nodeBatchSchedule, "BATCHNUMBER", strBatchNumber
                    xmlSetAttributeValue nodeBatchSchedule, "BATCHRUNNUMBER", strBatchRunNumber
                    
                    nodeTempRequest.appendChild nodeBatchSchedule
                    
                    CompleteBatch nodeTempRequest, nodeTempResponse
                    errCheckXMLResponseNode nodeTempResponse, , True
                Else
                    ' re-raise unhandled error
                    Err.Raise lngErrNo, strErrSource, strErrDesc
                End If
            End If
            
        Else
            'We are re-running a failed batch
            
            xmlChangeNodeName nodeTempRequest, "BATCH", "BATCHAUDIT"
            
            Err.Clear
            On Error Resume Next
                        
            ' PSC 18/09/2002 SYS5181 - Start
            Set nodeTemp = xmlGetMandatoryNode(nodeTempRequest, ".//BATCHAUDIT")
            xmlSetAttributeValue nodeTemp, "BATCHRUNNUMBER", strFailedBatchNo
            ' PSC 18/09/2002 SYS5181 - End

            FindFailedPayments nodeTempRequest, nodeTempResponse
            lngErrNo = Err.Number
            strErrDesc = Err.Description
            strErrSource = Err.Source
            On Error GoTo RunBatchProgrammeExit
        
            If lngErrNo <> 0 Then
                If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
                    blnContinue = False
                    errThrowError cstrFunctionName, 575
                Else
                    ' re-raise unhandled error
                    Err.Raise lngErrNo, strErrSource, strErrDesc
                End If
            End If
        End If
    
        If blnContinue Then
        
            '------------------------------------------------------------------------------------------
            ' Update the Batch Schedule
            '------------------------------------------------------------------------------------------
            ' PSC 20/09/2002 BMIS00480 - Start
            Set xmlPayments = nodeTempResponse
            Set nodelistPaymentRecords = xmlPayments.selectNodes(".//PAYMENTRECORD")
            ' PSC 20/09/2002 BMIS00480 - End
            
            
            Set nodeTempRequest = vxmlRequestNode.cloneNode(False)
            xmlSetAttributeValue nodeTempRequest, "OPERATION", "UpdateBatchSchedule"
            Set nodeBatchSchedule = xmlDoc.createElement("BATCHSCHEDULE")
            xmlSetAttributeValue nodeBatchSchedule, "BATCHNUMBER", strBatchNumber
            xmlSetAttributeValue nodeBatchSchedule, "BATCHRUNNUMBER", strBatchRunNumber
            xmlSetAttributeValue nodeBatchSchedule, "RETRY", "1"
            xmlSetAttributeValue nodeBatchSchedule, "TOTALNUMBEROFRECORDS", _
                CStr(nodelistPaymentRecords.length)
            nodeTempRequest.appendChild nodeBatchSchedule
            
            strResponse = objBatchSchedBO.omBatchRequest(nodeTempRequest.xml)
            errCheckXMLResponse strResponse, True
            
            ' RF 25/02/02 SYS4155 Start
            Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
            Set nodeTempResponse = vxmlResponseNode.cloneNode(False)
            'Set nodeBatch = xmlGetMandatoryNode(nodeTempRequest, "BATCH")
            nodeTempRequest.appendChild nodeBatch
            ' RF 25/02/02 End
                        
            'AM BMIDS722 - Start
            'get the batch interval and set the start time
            lngBatchInterval = GetBatchInterval(nodelistPaymentRecords.length)
            'set the start time, but be aware that if the application server and
            'database server are not synchronised then the batch may not be
            'processed when expected. It may be preferable to use the date/time
            'from the database server rather than the machine this component is
            'running on
            dtStartTime = Now
            lngCurrentItem = 0
            'AM BMIDS722 - End
            
            '------------------------------------------------------------------------------------------
            ' Process each payment
            '------------------------------------------------------------------------------------------
            
            Set objPayProcNTTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
            
            ' PSC 20/09/2002 BMIDS00480 - Start
            strIAValue = GetFirstComboValueId("PaymentType", "I")
                        
            ' Process initial advances first and then the remainder
            For intCount = 0 To 1
            
                If intCount = 0 Then ' Initial advance payments
                    Set nodelistPaymentRecords = xmlPayments.selectNodes(".//PAYMENTRECORD[@PAYMENTTYPE='" & strIAValue & "']")
                Else ' Non initial advance payments
                    Set nodelistPaymentRecords = xmlPayments.selectNodes(".//PAYMENTRECORD[@PAYMENTTYPE!='" & strIAValue & "']")
                End If
                
                For Each nodeTemp In nodelistPaymentRecords
                    'AM BMIDS722 - Start
                    'set the execute after time to slow down the batch processing
                    lngCurrentItem = lngCurrentItem + 1
                    strExecuteAfterDate = GetExecuteAfterDate(dtStartTime, lngBatchInterval, lngCurrentItem)
                    xmlSetAttributeValue nodeTemp, "EXECUTEAFTER", strExecuteAfterDate
                    'AM BMIDS722 - End
                    
                    'Call NTTxBO to create BatchAudit, PayProcBatchAudit records and
                    'send a corresponding message to the queue for each payment
                    On Error Resume Next
                    nodeBatch.appendChild nodeTemp
                    objPayProcNTTxBO.ProcessPayment nodeTempRequest, nodeTempResponse
                    If Err.Number <> 0 Then
                    
                        '------------------------------------------------------------------------------------------
                        ' CreateBatchAudit
                        '------------------------------------------------------------------------------------------
                        
                        'SYS3027 MDC Fix error handling
                        strValue = Err.Source
                        If Len(Trim$(strValue)) = 0 Or strValue = App.Title Then
                            strValue = App.Title & ".ProcessPayment"
                        End If
                        
                        ' PSC 01/03/2002 SYS4214 - Start
                        strApplicationNo = xmlGetAttributeText(nodeTemp, "APPLICATIONNUMBER")
                        strPaymentSeqNo = xmlGetAttributeText(nodeTemp, "PAYMENTSEQUENCENUMBER")
                        strDescription = "Application No: " & strApplicationNo & _
                                         " Payment Sequence No: " & strPaymentSeqNo & " " & _
                                         Err.Description
                                         
                        ' PSC 01/03/2002 SYS4214 - End
                        
                        
                        Set nodeErrRequest = vxmlRequestNode.cloneNode(False)
                        xmlSetAttributeValue nodeErrRequest, "OPERATION", "CreateBatchAudit"
                        Set nodeBatchAudit = xmlDoc.createElement("BATCHAUDIT")
                        xmlSetAttributeValue nodeBatchAudit, "ERRORNUMBER", CStr(Err.Number)
                        xmlSetAttributeValue nodeBatchAudit, "ERRORSOURCE", strValue
                        ' PSC 01/03/2002 SYS4214
                        xmlSetAttributeValue nodeBatchAudit, "ERRORDESCRIPTION", strDescription
                        Err.Clear
                        On Error GoTo RunBatchProgrammeExit
                        xmlSetAttributeValue nodeBatchAudit, "BATCHNUMBER", strBatchNumber
                        xmlSetAttributeValue nodeBatchAudit, "BATCHRUNNUMBER", strBatchRunNumber
                        nodeErrRequest.appendChild nodeBatchAudit
                        strResponse = objBatchSchedBO.omBatchRequest(nodeErrRequest.xml)
                        errCheckXMLResponse strResponse, True
                        'SYS3027 End
                    End If
                    nodeBatch.removeChild nodeTemp
                Next
            Next
            ' PSC 20/09/2002 BMIDS00480 - End
            On Error GoTo RunBatchProgrammeExit
        End If
    End If
    
RunBatchProgrammeExit:

    Set objBatchSchedBO = Nothing
    Set xmlDoc = Nothing
    Set nodeTempRequest = Nothing
    Set nodeTempResponse = Nothing
    Set nodeTemp = Nothing
    Set nodelistPaymentRecords = Nothing
    Set nodeBatch = Nothing
    Set nodeBatchAudit = Nothing
    Set nodeBatchSchedule = Nothing
    Set nodeErrRequest = Nothing 'SYS3027
    Set xmlPayments = Nothing       ' PSC 20/09/2002 BMIDS00480
    
    'SYS3541 MDC 20/12/2001
    If Err.Number <> 0 Then
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & cstrFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
    End If
    'SYS3541 End
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub GetFullPaymentDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo GetFullPaymentDetailsExit

Const strFunctionName As String = "GetFullPaymentDetails"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlAttrib As IXMLDOMAttribute
Dim xmlElement As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlPayee As IXMLDOMNode
Dim objPayProcDO As IPAymentProcessingDO

Dim strPattern As String

    ' Raise error if no PAYMENTRECORD node exists
    Set xmlElement = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    'If PAYMENTSEQUENCENUMBER attribute specified, remove it, saving for later use
    Set xmlAttrib = xmlElement.getAttributeNode("PAYMENTSEQUENCENUMBER")
    If Not xmlAttrib Is Nothing Then
        xmlElement.removeAttribute "PAYMENTSEQUENCENUMBER"
    End If

    'Get Payment Records
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempResponse = vxmlResponseNode.cloneNode(True)
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlTempRequest.appendChild xmlElement
    Call FindPaymentHistoryList(xmlTempRequest, xmlTempResponse)
'    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
'    objPayProcDO.FindPaymentHistoryList xmlElement, xmlTempResponse

    'Filter out non disbursement payments
    If xmlAttrib Is Nothing Then
        strPattern = "PAYMENTRECORD[DISBURSEMENTPAYMENT]"
    Else
        'Also remove any PaymentRecord's not matching the specified Payment Sequence Number
        strPattern = "PAYMENTRECORD[DISBURSEMENTPAYMENT and @PAYMENTSEQUENCENUMBER=""" & xmlAttrib.Text & """]"
    End If

    Set xmlNodeList = xmlTempResponse.selectNodes(strPattern)
    If xmlNodeList.length = 0 Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    For Each xmlNode In xmlNodeList
        'Find Payee History Details
        Set xmlTempRequest = xmlDoc.createElement("REQUEST")
        Set xmlElement = xmlDoc.createElement("PAYEEHISTORY")
        xmlElement.setAttribute "APPLICATIONNUMBER", xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        xmlElement.setAttribute "PAYEEHISTORYSEQNO", xmlGetMandatoryNodeText(xmlNode, "DISBURSEMENTPAYMENT/@PAYEEHISTORYSEQNO")
        xmlTempRequest.appendChild xmlElement
        Set xmlTempResponse = vxmlResponseNode.cloneNode(True)
        Call FindPayeeHistoryList(xmlTempRequest, xmlTempResponse, False)
        
        'Append to the PaymentRecord
        Set xmlPayee = xmlGetNode(xmlTempResponse, "PAYEEHISTORY")
        If Not xmlPayee Is Nothing Then
            xmlNode.appendChild xmlPayee
        End If
        
        'Append PaymentRecord to the main Response
        vxmlResponseNode.appendChild xmlNode
    Next


GetFullPaymentDetailsExit:
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlNodeList = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlPayee = Nothing
    Set objPayProcDO = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindSanctioningList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindSanctioningList_Exit

Const strFunctionName As String = "FindSanctioningList"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim objPayProcDO As IPAymentProcessingDO

Dim strValue As String

    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "FINDSANCTIONINGLIST")

    'Check if Application Number has been specified
    strValue = xmlGetAttributeText(xmlNode, "APPLICATIONNUMBER")
    If Len(Trim$(strValue)) > 0 Then
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        Set xmlNode = xmlDoc.createElement("FINDSANCTIONINGLIST")
        xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strValue
        'SR 10-12-2001 : SYS3400
        xmlSetAttributeValue xmlNode, "_COMBOLOOKUP_", "1"
    End If
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    objPayProcDO.FindSanctioningList xmlNode, vxmlResponseNode
    
FindSanctioningList_Exit:
    Set objPayProcDO = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CreateFeeTypePaymentExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "CreateFeeTypePayment"

    ' Raise error if no FEEPAYMENT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "FEEPAYMENT")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate FEETYPE attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPE")
    
    'Validate FEETYPESEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPESEQUENCENUMBER")
    
    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    'strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreateFeeTypePayment vxmlRequestNode, vxmlResponseNode
    
CreateFeeTypePaymentExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindPaymentsForBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindPaymentsForBatchExit

    Dim nodeBatch As IXMLDOMNode
    Dim objPayProcDO As IPAymentProcessingDO
    
    Const strFunctionName As String = "FindPaymentsForBatch"

    ' Raise error if no BATCH node exists
    Set nodeBatch = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'Find all Fee Payment types for specified application
    objPayProcDO.FindPaymentsForBatch vxmlRequestNode, vxmlResponseNode
    
FindPaymentsForBatchExit:
    Set nodeBatch = Nothing
    Set objPayProcDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreatePaymentAudit( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CreatePaymentAuditExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "CreatePaymentAudit"

    ' Raise error if no PAYMENTPROCESSINGBATCHAUDIT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYPROCBATCHAUDIT")
    
    'Validate BATCHAUDIT attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID")
    
    'Validate BATCHNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
    
    'Validate BATCHRUNNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreatePaymentAudit vxmlRequestNode, vxmlResponseNode
    
CreatePaymentAuditExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub CreateChequeAuditRecord( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BM029   AW  15/09/02
On Error GoTo CreateChequeAuditRecordExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "CreateChequeAuditRecord"

    ' Raise error if no PAYMENTCHEQUEAUDIT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTCHEQUEAUDIT")
    
    'Validate CACHEQUENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "CACHEQUENUMBER")
    
    'Validate CAUSERID attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "CAUSERID")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreateChequeAuditRecord vxmlRequestNode, vxmlResponseNode
    
CreateChequeAuditRecordExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub SanctionPrintCheques( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BM029   AW  15/09/02
On Error GoTo SanctionPrintChequesExit

Const strFunctionName As String = "SanctionPrintCheques"

Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
Dim nodeResponseFromGetLockData As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlResponse As IXMLDOMElement
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
'GD BMIDS0922
Dim nNextCheqNum As Double
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strAppNo As String
    'GD BMIDS00614 START
    Dim blnRePrintStatus As Boolean
    Dim blnRePrintMode As Boolean
    blnRePrintMode = False
    If xmlGetAttributeText(vxmlRequestNode, "REPRINTMODE") = "1" Then
        blnRePrintMode = True
    End If
    'GD BMIDS00614 END
    Set objPayProcNTTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
    
    Dim lngErrNo As Long
    Dim blnLockExists As Boolean
    Dim colAppsLocked As Collection
    Set colAppsLocked = New Collection
    Dim lNum As Long
    Dim bLockedByProcess As Boolean
    Dim bExternalLock As Boolean
    Dim strThisUser As String
    Dim strUser As String
    Dim xmlAppLockElem As IXMLDOMElement
    'GD BMIDS0922
    nNextCheqNum = CDbl(xmlGetMandatoryAttributeText(vxmlRequestNode, "CHEQUENUMBER")) - 1
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    xmlResponse.setAttribute "TYPE", "SUCCESS"
      
    Set xmlNodeList = vxmlRequestNode.selectNodes("PAYMENTRECORD")
    strThisUser = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
    
    For Each xmlNode In xmlNodeList
    
        bLockedByProcess = False
        blnLockExists = True
        
        'Validate APPLICATIONNUMBER attribute is specified
        strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        
       'If there are multiple payments for a single app, have to make sure subsequent payments
       'are not excluded due to lock applied for first payment
        For lNum = 1 To colAppsLocked.Count
            If (colAppsLocked(lNum) = strAppNo) Then
                bLockedByProcess = True
                Exit For
            End If
        Next
        
        If Not bLockedByProcess Then
            ' Need to pass an existing response node.
            Set nodeResponseFromGetLockData = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
            
            objPayProcNTTxBO.CheckLocks xmlNode, nodeResponseFromGetLockData
        
            lngErrNo = errCheckXMLResponseNode(nodeResponseFromGetLockData, , False)
               
            If lngErrNo = 0 Then
                blnLockExists = True
            ElseIf errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
                    blnLockExists = False
                
            End If
        End If
        
        
        
        'GD BMIDS00614 START Check if the App is locked to another user
        Set nodeResponseFromGetLockData = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
        objPayProcNTTxBO.CheckLocks xmlNode, nodeResponseFromGetLockData
        lngErrNo = errCheckXMLResponseNode(nodeResponseFromGetLockData, , False)
        bExternalLock = False
        If lngErrNo = 0 Then
            Set xmlAppLockElem = nodeResponseFromGetLockData.selectSingleNode(".//APPLICATIONLOCK/USERID")
            strUser = xmlAppLockElem.Text
                If strUser <> strThisUser Then
                    bExternalLock = True
                End If
        ElseIf errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
            bExternalLock = False
        End If
        'GD BMIDS00614 END
        
        
        
        'Check the ReprintStatus flag of the individual Node in the list
        blnRePrintStatus = CSafeBool(xmlGetAttributeText(xmlNode, "REPRINTSTATUS"))
        'GD BMIDS00614 START
        If ((Not blnLockExists) Or (bLockedByProcess = True) Or blnRePrintStatus) Then
        'If ((Not blnLockExists) Or (bLockedByProcess = True)) Then
        'GD BMIDS00614 END
        'If Not blnLockExists Then
            Set xmlDoc = New FreeThreadedDOMDocument40
            xmlDoc.validateOnParse = False
            xmlDoc.setProperty "NewParser", True
            Set xmlTempRequest = xmlDoc.createElement("REQUEST")
            xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "USERID", "USERID"
            xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "UNITID", "UNITID"
            xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "USERAUTHORITYLEVEL", "USERAUTHORITYLEVEL"
            xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "MACHINEID", "MACHINEID"
            xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "CHANNELID", "CHANNELID"
            'GD BMIDS00614 START
            xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "REPRINTMODE", "REPRINTMODE"
            
            'nNextCheqNum = nNextCheqNum + 1
            'GD BMIDS00614 END
            If bExternalLock Then
                'GD BMIDS00870 START
                'xmlSetAttributeValue xmlNode, "CHEQUENUMBER", Str$(nNextCheqNum)
                'GD BMIDS00870 END
            
                xmlSetAttributeValue xmlNode, "CHEQUENUMBER", "Locked"
            End If
            xmlTempRequest.appendChild xmlNode
            
            'Build a list of apps locked by this process
            colAppsLocked.Add strAppNo

            'GD BMIDS00614 START
            If (blnRePrintMode And blnRePrintStatus) Or (blnRePrintMode = False) Then
                'Call the SanctionPrintCheque method on the NTTxBO therefore executing in separate transaction
                If Not bExternalLock Then
                    'GD BMIDS00870 START
                    nNextCheqNum = nNextCheqNum + 1
                    xmlSetAttributeValue xmlNode, "CHEQUENUMBER", Str$(nNextCheqNum)
                    'GD BMIDS00870 END
                    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
                    objPayProcNTTxBO.SanctionPrintCheque xmlTempRequest, xmlTempResponse, bLockedByProcess
                    'Check for errors
                    errCheckXMLResponse xmlTempResponse.xml, True
                End If
            End If
            'GD BMIDS00614 END
        Else
            If (blnLockExists And (Not blnRePrintMode)) Or bExternalLock Then
                xmlSetAttributeValue xmlNode, "CHEQUENUMBER", "Locked"
            End If
        End If
        
        xmlResponse.appendChild xmlNode
        Set nodeResponseFromGetLockData = Nothing
    Next
    
'    'TODO  Save high cheque number
'    Dim objOrgBO As omOrg.OrganisationBO
'    Set objOrgBO = gobjContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
'
'    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
'    Set xmlElement = xmlDoc.createElement("UNIT")
'    xmlTempRequest.appendChild xmlElement
'    Set xmlNode = xmlDoc.createElement("UNITID")
'    xmlElement.appendChild xmlNode
'    xmlNode.Text = strApplicationNumber
'    Set xmlNode = xmlDoc.createElement("UNHIGHCHEQUENUMBER")
'    xmlElement.appendChild xmlNode
'
'    strResponse = objOrgBO.UpdateHighChequeNum(vxmlRequestNode.xml)
'    errCheckXMLResponse strResponse, True
    
    'Return list of payments with allocated cheque numbers
    Dim xmlChildNode As IXMLDOMNode
    Dim xmlPaymentList As IXMLDOMNodeList

    Set xmlPaymentList = xmlResponse.selectNodes("PAYMENTRECORD")
    For Each xmlChildNode In xmlPaymentList
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
    
SanctionPrintChequesExit:
    Set nodeResponseFromGetLockData = Nothing
    Set xmlNodeList = Nothing
    Set xmlChildNode = Nothing
    Set xmlPaymentList = Nothing
    Set xmlResponse = Nothing
    Set colAppsLocked = Nothing

    Set objPayProcNTTxBO = Nothing

    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub SetPaidChequePayments( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMElement)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BM029   AW  15/09/02
On Error GoTo SetPaidChequePaymentsExit

Const strFunctionName As String = "SetPaidChequePayments"

Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMElement
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strAppNo As String
    
    Set objPayProcNTTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
    
    Set xmlNodeList = vxmlRequestNode.selectNodes("PAYMENTRECORD")
    
    For Each xmlNode In xmlNodeList
        
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        Set xmlTempRequest = xmlDoc.createElement("REQUEST")
        xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "USERID", "USERID"
        xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "UNITID", "UNITID"
        xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "USERAUTHORITYLEVEL", "USERAUTHORITYLEVEL"
        xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "MACHINEID", "MACHINEID"
        xmlCopyMandatoryAttributeValue vxmlRequestNode, xmlTempRequest, "CHANNELID", "CHANNELID"
            
        xmlTempRequest.appendChild xmlNode
        'Validate APPLICATIONNUMBER attribute is specified
        strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

        'Call the SanctionPrintCheque method on the NTTxBO therefore executing in separate transaction
        Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
        objPayProcNTTxBO.SetPaidChequePayment xmlTempRequest, xmlTempResponse
        'Check for errors
        errCheckXMLResponse xmlTempResponse.xml, True
        
    Next
    
    vxmlResponseNode.setAttribute "TYPE", "SUCCESS"
    
SetPaidChequePaymentsExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    Set objPayProcNTTxBO = Nothing

    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindFeeTypeList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindFeeTypeListExit

Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlFeePayNodeList As IXMLDOMNodeList
Dim xmlElement As IXMLDOMElement
Dim xmlSchemaNode As IXMLDOMNode
Dim xmlFeePayNode As IXMLDOMNode
Dim objPayProcDO As IPAymentProcessingDO

Dim strValue As String
Dim blnDoComboLookUp As Boolean

Const strFunctionName As String = "FindFeeTypeList"

    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFEETYPE")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'Find all Fee Payment types for specified application
    objPayProcDO.FindFeeTypeList vxmlRequestNode, vxmlResponseNode
    
    Set xmlSchemaNode = adoGetSchema("APPLICATIONFEETYPE")
    blnDoComboLookUp = xmlGetAttributeAsBoolean(xmlNode, "_COMBOLOOKUP_")
    
    Set xmlNodeList = vxmlResponseNode.selectNodes("APPLICATIONFEETYPE")
    For Each xmlNode In xmlNodeList
        'Find the Fee Payments made for each of this application's fee types
        Set xmlFeePayNode = xmlCreateChildRequest(xmlNode, xmlSchemaNode, "FEEPAYMENT", blnDoComboLookUp)
        objPayProcDO.FindFeePaymentList xmlFeePayNode, xmlNode
        Set xmlFeePayNodeList = xmlNode.selectNodes("FEEPAYMENT")
        For Each xmlElement In xmlFeePayNodeList
            'Find the associated PaymentRecord's
            Set xmlFeePayNode = xmlCreateChildRequest(xmlNode, xmlSchemaNode, "PAYMENTRECORD", blnDoComboLookUp)
            xmlCopyAttribute xmlElement, xmlFeePayNode, "PAYMENTSEQUENCENUMBER"
            objPayProcDO.FindPaymentRecordList xmlFeePayNode, xmlNode
        Next
    Next
    
FindFeeTypeListExit:
    Set objPayProcDO = Nothing
    Set xmlSchemaNode = Nothing
    Set xmlFeePayNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindFeePaymentTotals(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindFeePaymentTotalsExit

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlFeePayment As IXMLDOMNode
Dim xmlFeePaymentList As IXMLDOMNodeList
Dim xmlFeeTypeNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode

Dim strValue As String
Dim lngAmountPaid As Long
Dim lngAmountOS As Long

Const strFunctionName As String = "FindFeePaymentTotals"

    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFEETYPE")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    
    'Get Fee Type data
    Call FindFeeTypeList(xmlTempRequest, xmlTempResponse)
    
    'Reformat to create Response in correct form
    Set xmlNodeList = xmlTempResponse.selectNodes("APPLICATIONFEETYPE")
    If xmlNodeList.length < 1 Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    For Each xmlNode In xmlNodeList
        Set xmlFeeTypeNode = xmlDoc.createElement("APPLICATIONFEETYPE")
        'Get Fee Type
        xmlCopyMandatoryAttribute xmlNode, xmlFeeTypeNode, "FEETYPE"
        xmlCopyMandatoryAttribute xmlNode, xmlFeeTypeNode, "FEETYPE_TEXT"
        xmlCopyMandatoryAttribute xmlNode, xmlFeeTypeNode, "FEETYPESEQUENCENUMBER"
        
        ' SR 01/08/01 : SYS2545 - Do not conside the column RebateOrAddition from table 'ApplicationFeeType'
        'Get Fee Amount = Amount
        lngAmountOS = xmlGetMandatoryAttributeAsLong(xmlNode, "AMOUNT") '+ xmlGetMandatoryAttributeAsLong(xmlNode, "REBATEORADDITION")
        xmlSetAttributeValue xmlFeeTypeNode, "AMOUNT", CStr(lngAmountOS)
        
        'Total up Fee Payments made for this FeeType
        lngAmountPaid = 0
        Set xmlFeePaymentList = xmlNode.selectNodes("FEEPAYMENT")
        For Each xmlFeePayment In xmlFeePaymentList
            lngAmountPaid = lngAmountPaid + xmlGetMandatoryAttributeAsLong(xmlFeePayment, "AMOUNTPAID")
        Next
        xmlSetAttributeValue xmlFeeTypeNode, "AMOUNTPAID", CStr(lngAmountPaid)
        
        'Calculate Amount Outstanding
        lngAmountOS = lngAmountOS - lngAmountPaid
        xmlSetAttributeValue xmlFeeTypeNode, "AMOUNTOUTSTANDING", CStr(lngAmountOS)
        
        If lngAmountOS <> 0 Then
            'Append to the Response
            vxmlResponseNode.appendChild xmlFeeTypeNode
        End If
    Next
    
FindFeePaymentTotalsExit:
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlNodeList = Nothing
    Set xmlFeePayment = Nothing
    Set xmlFeePaymentList = Nothing
    Set xmlFeeTypeNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindApplicationThirdPartyList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationThirdPartyListExit

Dim xmlNode As IXMLDOMNode
Dim objPayProcDO As IPAymentProcessingDO

Dim strValue As String

Const strFunctionName As String = "FindApplicationThirdPartyList"

    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONTHIRDPARTY")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'Find all Fee Payment types for specified application
    objPayProcDO.FindApplicationThirdPartyList vxmlRequestNode, vxmlResponseNode
    
FindApplicationThirdPartyListExit:
    Set objPayProcDO = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindPaymentHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode, _
                                    Optional ByVal blnRaiseErrorIfNotFound As Boolean = True)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindPaymentHistoryListExit

Const strFunctionName As String = "FindPaymentHistoryList"

Dim objPayProcDO As IPAymentProcessingDO
Dim xmlNode As IXMLDOMNode
Dim xmlRequest As IXMLDOMElement
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlResponse As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlAppFeeTypeNode As IXMLDOMElement
Dim objMortSQBO As MortgageSubQuoteBO
Dim colValueId As Collection
Dim xmlSchemaNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim xmlTempResponse As IXMLDOMNode
Dim xmlAttrib As IXMLDOMAttribute
Dim xmlTPNode As IXMLDOMNode
Dim xmlPayeeResponse As IXMLDOMNode
Dim xmlDisbNode As IXMLDOMNode
Dim xmlLoanCompList As IXMLDOMNodeList
Dim xmlLoanComp As IXMLDOMNode

Dim strValue As String
Dim intTTCount As Integer
Dim intTTFeeCount As Integer
Dim intIndex As Integer
Dim intValue As Integer
Dim blnDoComboLookUp As Boolean
Dim strAppNo As String
Dim blnUseAppNo As Boolean
Dim strPattern As String
Dim strSubPattern As String
Dim lngValue As Long

    ' Raise error if no PAYMENTRECORD node exists
    Set xmlTempNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    strAppNo = xmlGetAttributeText(xmlTempNode, "APPLICATIONNUMBER")
    If Len(Trim$(strAppNo)) > 0 Then
        blnUseAppNo = True
    End If
    
    If Not blnUseAppNo Then
        'If ISSUEDATE attribute specified, remove it, saving for later use
        Set xmlElement = xmlGetNode(vxmlRequestNode, "DISBURSEMENTPAYMENT")
        If Not xmlElement Is Nothing Then
            Set xmlAttrib = xmlElement.getAttributeNode("ISSUEDATE")
            If Not xmlAttrib Is Nothing Then
                xmlElement.removeAttribute "ISSUEDATE"
            End If
        End If
    End If
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    'Find PAYMENTRECORD items
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    objPayProcDO.FindPaymentRecordList xmlTempNode, xmlTempResponse
        
    Set xmlSchemaNode = adoGetSchema("PAYMENTRECORD")
    blnDoComboLookUp = xmlGetAttributeAsBoolean(xmlTempNode, "_COMBOLOOKUP_")
    
    Set objMortSQBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    
    Set xmlNodeList = xmlTempResponse.selectNodes("PAYMENTRECORD")
    For Each xmlNode In xmlNodeList
        'Find the FEEPAYMENT's for each PAYMENTRECORD
        Set xmlRequest = xmlCreateChildRequest(xmlNode, xmlSchemaNode, "FEEPAYMENT", blnDoComboLookUp)
        objPayProcDO.FindFeePaymentList xmlRequest, xmlNode

        'Find DISBURSEMENTPAYMENT items
        If blnUseAppNo Then
            Set xmlRequest = xmlCreateChildRequest(xmlNode, xmlSchemaNode, "DISBURSEMENTPAYMENT", blnDoComboLookUp)
        Else
            ' Raise error if no DISBURSEMENTPAYMENT node exists
            Set xmlRequest = xmlGetMandatoryNode(vxmlRequestNode, "DISBURSEMENTPAYMENT")
            strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
            xmlRequest.setAttribute "APPLICATIONNUMBER", strAppNo
            strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
            xmlRequest.setAttribute "PAYMENTSEQUENCENUMBER", strValue
        End If
        objPayProcDO.FindDisbursementPaymentList xmlRequest, xmlNode
        
        'If DisbursementPayment found, get payee, loan component and loan component payment data
        Set xmlDisbNode = xmlGetNode(xmlNode, ".//DISBURSEMENTPAYMENT")
        If Not xmlDisbNode Is Nothing Then
            strAppNo = xmlGetMandatoryAttributeText(xmlDisbNode, "APPLICATIONNUMBER")
            strValue = xmlGetMandatoryAttributeText(xmlDisbNode, "PAYEEHISTORYSEQNO")
            Set xmlRequest = xmlDoc.createElement("REQUEST")
            Set xmlElement = xmlDoc.createElement("PAYEEHISTORY")
            xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
            xmlElement.setAttribute "PAYEEHISTORYSEQNO", strValue
            xmlRequest.appendChild xmlElement
            Set xmlPayeeResponse = xmlDoc.createElement("RESPONSE")
            Call FindPayeeHistoryList(xmlRequest, xmlPayeeResponse, False)
            'Get the Third Party details and append them to the main response
            Set xmlTPNode = xmlGetNode(xmlPayeeResponse, "PAYEEHISTORY/THIRDPARTY")
            If Not xmlTPNode Is Nothing Then
                xmlDisbNode.appendChild xmlTPNode
            End If
        
            Set xmlRequest = xmlCreateChildRequest(xmlNode, xmlSchemaNode, "LOANCOMPONENTPAYMENT", blnDoComboLookUp)
            If xmlGetAttributeAsInteger(xmlNode, "PAYMENTSEQUENCENUMBER") = 15 Then
                Debug.Print "stop"
            End If
            
            ' SR 10/07/01 : Handle record not found error
            On Error Resume Next
            Dim lngErrNo As Long
            objPayProcDO.FindLoanComponentPaymentList xmlRequest, xmlNode
            
            lngErrNo = Err.Number
            
            On Error GoTo FindPaymentHistoryListExit
            
            If lngErrNo <> 0 Then
                lngErrNo = errGetOmigaErrorNumber(lngErrNo)
                If lngErrNo <> oeRecordNotFound Then
                    errThrowError strFunctionName, lngErrNo
                End If
            End If
            
            If Not blnUseAppNo Then
                'Get Loan Components for this application
                Call GetLoanComponents(strAppNo, xmlTempResponse)
            End If
        End If
    Next

    If blnUseAppNo Then
        'Get Loan Components for the application
        Call GetLoanComponents(strAppNo, xmlTempResponse)
    End If

    'Do not include any PAYMENTRECORD's without DISBURSEMENTPAYMENT or
    'FEEPAYMENT records that match the parameters passed in the REQUEST
    If Not xmlAttrib Is Nothing Then
        'If an Issue Date has been specified, remove any PaymentRecord's  where the associated
        'DISBURSEMENTPAYMENT has an ISSUEDATE after the date requested
        strValue = xmlAttrib.Text
        If IsDate(strValue) Then
            strSubPattern = "(date(DISBURSEMENTPAYMENT/@ISSUEDATE) <= date(""" & Format$(strValue, "yyyy-mm-dd") & """)"
        End If
    End If
    
    If Len(strSubPattern) = 0 Then
        strSubPattern = "DISBURSEMENTPAYMENT"
    Else
        strSubPattern = strSubPattern & ")"
    End If
    
    strPattern = "PAYMENTRECORD[" & strSubPattern & " or FEEPAYMENT]"
    Set xmlNodeList = xmlTempResponse.selectNodes(strPattern)
    If xmlNodeList.length = 0 And blnRaiseErrorIfNotFound Then
        'Raise a Record Not Found error
        errThrowError strFunctionName, oeRecordNotFound, "No PAYMENTRECORD's found"
    End If
    For Each xmlNode In xmlNodeList
        vxmlResponseNode.appendChild xmlNode
    Next
    
    Set xmlNodeList = xmlTempResponse.selectNodes("LOANCOMPONENT")
    For Each xmlNode In xmlNodeList
        vxmlResponseNode.appendChild xmlNode
    Next
    
    'Find any Payments by CHAPS and if required create a CHAPS Charge Fee Type record
    Set xmlNodeList = vxmlResponseNode.selectNodes("PAYMENTRECORD")
    For Each xmlNode In xmlNodeList
        'Count the number of CHAPS/TT payments
        intValue = xmlGetAttributeAsInteger(xmlNode, "PAYMENTMETHOD")
        If IsValidationType("PaymentMethod", intValue, "Y") Then
            intTTCount = intTTCount + 1
        End If
    Next
    
    If intTTCount > 0 Then
        'See if the correct number of CHAPS Charge Fee Types already exist
        Set xmlRequest = xmlDoc.createElement("REQUEST")
        Set xmlResponse = xmlDoc.createElement("RESPONSE")
        Set xmlAppFeeTypeNode = xmlDoc.createElement("APPLICATIONFEETYPE")
        xmlRequest.appendChild xmlAppFeeTypeNode
        objPayProcDO.FindFeeTypeList xmlRequest, xmlResponse
        Set xmlNodeList = xmlResponse.selectNodes("APPLICATIONFEETYPE")
        For Each xmlNode In xmlNodeList
            intValue = xmlGetAttributeAsInteger(xmlNode, "FEETYPE")
            If IsValidationType("OneOffCost", intValue, "TTF") Then
                intTTFeeCount = intTTFeeCount + 1
            End If
        Next
        If intTTFeeCount < intTTCount Then
            'Get the standard CHAPS Charge Amount
            lngValue = GetGlobalParamAmount("StandardCHAPSCharge")
            xmlAppFeeTypeNode.setAttribute "AMOUNT", CStr(lngValue)
            'Get the combovalue for CHAPS/TT
            Set colValueId = New Collection
            GetValueIdsForValidationType "OneOffCost", "TTF", colValueId
            If colValueId.Count > 0 Then
                strValue = CStr(colValueId.Item(1))
            Else
                errThrowError strFunctionName, oePPComboValueIdNotFound
            End If
            xmlAppFeeTypeNode.setAttribute "FEETYPE", strValue
            'Create the required number of CHAPS Charge Fee Type records
            Set xmlResponse = xmlDoc.createElement("RESPONSE")
            xmlResponse.setAttribute "TYPE", "SUCCESS"
            For intIndex = 1 To intTTCount - intTTFeeCount
                Call CreateApplicationFeeType(xmlRequest, xmlResponse)
                errCheckXMLResponseNode xmlResponse, xmlResponse, True
            Next
        End If
    End If
    
FindPaymentHistoryListExit:
    Set objPayProcDO = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlAppFeeTypeNode = Nothing
    Set colValueId = Nothing
    Set xmlElement = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlAttrib = Nothing
    Set xmlTPNode = Nothing
    Set xmlPayeeResponse = Nothing
    Set xmlDisbNode = Nothing
    Set xmlLoanCompList = Nothing
    Set xmlLoanComp = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
        
End Sub

Private Sub GetLoanComponents(ByVal strApplicationNumber As String, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanComponentsExit

Const strFunctionName As String = "GetLoanComponents"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim xmlLoanList As IXMLDOMNodeList
Dim objMortSQBO As MortgageSubQuoteBO

Dim strAppFFNo As String
Dim strResponse As String
Dim lngRet As Long

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objMortSQBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    
    'Get the ApplicationFactFindNumber for this application
    strAppFFNo = GetFactFindNumberForApplication(strApplicationNumber)

    'Create REQUEST
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("FINDACTIVELOANCOMPONENTS")
    xmlTempRequest.appendChild xmlElement
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strApplicationNumber
    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strAppFFNo
    
    'Find LOANCOMPONENT's
    strResponse = objMortSQBO.FindActiveLoanComponentsList(xmlTempRequest.xml)
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    If Not (lngRet = 0 Or lngRet = oeRecordNotFound) Then
        errThrowError strFunctionName, lngRet
    End If
    xmlDoc.loadXML strResponse
    
    'Add LOANCOMPONENT's to the RESPONSE
    Set xmlLoanList = xmlDoc.selectNodes(".//LOANCOMPONENT")
    For Each xmlNode In xmlLoanList
        vxmlResponseNode.appendChild xmlNode
    Next
    
GetLoanComponentsExit:
    Set xmlNode = Nothing
    Set xmlLoanList = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlElement = Nothing
    Set objMortSQBO = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub
Private Sub ApplyApplicationFeeRebatesAndAdditions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo ApplyRebatesAndAdditionsExit

Const strFunctionName As String = "ApplyApplicationFeeRebatesAndAdditions"

    Dim xmlInAppFeeNode As IXMLDOMNode
    Dim xmlOutAppFeeNode As IXMLDOMNode
    Dim xmlInAppFeeNodeList As IXMLDOMNodeList
    Dim xmlInPaymentList As IXMLDOMNodeList
    Dim xmlFeePaymentNode As IXMLDOMNode
    Dim dblFeeAmount As Double
    Dim dblOutFeeAmount As Double
    Dim dblRebateOrAddition As Double
    Dim intPaymentEvent As Integer
    
    'Process each fee in ApplicationFeeTypeXML (passed in) and adjust the value by adding additions
    ' or subtracting rebates
    Set xmlInAppFeeNodeList = vxmlRequestNode.selectNodes("APPLICATIONFEETYPE")
    For Each xmlInAppFeeNode In xmlInAppFeeNodeList
        Set xmlOutAppFeeNode = xmlInAppFeeNode.cloneNode(True)
        dblFeeAmount = xmlGetAttributeAsDouble(xmlInAppFeeNode, "AMOUNT")
        Set xmlInPaymentList = xmlInAppFeeNode.selectNodes("FEEPAYMENT")
        For Each xmlFeePaymentNode In xmlInPaymentList
            intPaymentEvent = xmlGetAttributeAsInteger(xmlFeePaymentNode, "PAYMENTEVENT")
            If IsValidationType("PaymentEvent", intPaymentEvent, "O") And _
               IsValidationType("PaymentEvent", intPaymentEvent, "RA") Then
               dblRebateOrAddition = xmlGetAttributeAsDouble(xmlFeePaymentNode, "AMOUNTPAID")
               dblFeeAmount = dblFeeAmount - dblRebateOrAddition
            End If
        Next
        xmlSetAttributeValue xmlOutAppFeeNode, "AMOUNT", CStr(dblFeeAmount)
        If dblFeeAmount > 0# Then
            'add this applicationFeeType to the outputxml
            vxmlResponseNode.appendChild xmlOutAppFeeNode
        End If
    Next

ApplyRebatesAndAdditionsExit:

    Set xmlInAppFeeNode = Nothing
    Set xmlOutAppFeeNode = Nothing
    Set xmlInAppFeeNodeList = Nothing
    Set xmlInPaymentList = Nothing
    Set xmlFeePaymentNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateFeeTypePaymentExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "UpdateFeeTypePayment"

    ' Raise error if no FEEPAYMENT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "FEEPAYMENT")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate FEETYPE attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPE")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    'Validate FEETYPESEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPESEQUENCENUMBER")
    
    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdateFeeTypePayment vxmlRequestNode, vxmlResponseNode
    
UpdateFeeTypePaymentExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationFeeTypeExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "UpdateApplicationFeeType"

    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFEETYPE")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    'Validate FEETYPE attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPE")

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdateApplicationFeeType vxmlRequestNode, vxmlResponseNode
    
UpdateApplicationFeeTypeExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationFeeTypeExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "CreateApplicationFeeType"

    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFEETYPE")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate FEETYPE attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPE")

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreateApplicationFeeType vxmlRequestNode, vxmlResponseNode

CreateApplicationFeeTypeExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindPayeeHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode, _
                                    Optional ByVal blnErrorIfNotFound As Boolean = True)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindPayeeHistoryListExit

Const strFunctionName As String = "FindPayeeHistoryList"

Dim objPayProcDO As IPAymentProcessingDO
Dim objThirdPartyBO As ThirdPartyBO
Dim xmlTempResponse As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMElement 'IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlTPNode As IXMLDOMElement
Dim xmlChildNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMElement
Dim xmlTPRequest As FreeThreadedDOMDocument40
Dim xmlGuidNode As IXMLDOMElement
Dim xmlThirdParty As IXMLDOMElement
    
Dim strGuid As String
Dim strResponse As String
Dim strValue As String

    ' Raise error if no PAYEEHISTORY node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set objThirdPartyBO = gobjContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")

    'Create a temporary Request for the Third Party method call
    Set xmlTPRequest = New FreeThreadedDOMDocument40
    xmlTPRequest.validateOnParse = False
    xmlTPRequest.setProperty "NewParser", True
    Set xmlTempRequest = xmlTPRequest.createElement("REQUEST")
    xmlTPRequest.appendChild xmlTempRequest
    Set xmlTPNode = xmlTPRequest.createElement("THIRDPARTY")
    xmlTempRequest.appendChild xmlTPNode
    
    'Find payee history records
    objPayProcDO.FindPayeeHistoryList vxmlRequestNode, vxmlResponseNode, blnErrorIfNotFound

    If Not vxmlResponseNode Is Nothing Then
        'For each Payee History record, get the associated thirdparty details
        Set xmlNodeList = vxmlResponseNode.selectNodes(".//PAYEEHISTORY")
        For Each xmlNode In xmlNodeList
            'Extract the Third Party Guid...
            strGuid = xmlNode.getAttribute("THIRDPARTYGUID")
            If Len(Trim$(strGuid)) > 0 Then
                '... and add it to the Third Party Request
                Set xmlGuidNode = xmlTPRequest.selectSingleNode(".//THIRDPARTYGUID")
                If xmlGuidNode Is Nothing Then
                    Set xmlGuidNode = xmlTPRequest.createElement("THIRDPARTYGUID")
                End If
                xmlGuidNode.Text = strGuid
                xmlTPNode.appendChild xmlGuidNode
                
                'Find the third party
                strResponse = objThirdPartyBO.GetThirdParty(xmlTPRequest.xml)
                errCheckXMLResponse strResponse, True
                If xmlDoc.loadXML(strResponse) Then
                    Set xmlTempResponse = xmlDoc.documentElement
                    'Remove ContactDetails node if it exists
                    Set xmlThirdParty = xmlTempResponse.selectSingleNode(".//THIRDPARTY")
                    If Not xmlThirdParty Is Nothing Then
                        For Each xmlChildNode In xmlThirdParty.childNodes
                            If xmlChildNode.nodeName = "CONTACTDETAILS" Then
                                xmlThirdParty.removeChild xmlChildNode
                            End If
                        Next
                    End If
                    
                    'Convert back to Phase 2 style response
                    Set xmlTempResponse = xmlCreateAttributeBasedResponse(xmlTempResponse, True)
                    Set xmlTempResponse = xmlTempResponse.selectSingleNode("THIRDPARTY")
                    
                    'Append it to the main response
                    xmlNode.appendChild xmlTempResponse
                Else
                    errThrowError strFunctionName, oeXMLParserError, "Invalid Response from ThirdPartyBO.GetThirdParty"
                End If
            End If
        Next
    End If

FindPayeeHistoryListExit:
    Set objPayProcDO = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlTPNode = Nothing
    Set xmlThirdParty = Nothing
    Set xmlChildNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTPRequest = Nothing
    Set xmlGuidNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindPaymentsForPayeeList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo FindPaymentsForPayeeListExit

Const strFunctionName As String = "FindPaymentsForPayeeList"

Dim xmlNode As IXMLDOMNode
Dim objPayProcDO As IPAymentProcessingDO
Dim strValue As String

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    ' Raise error if no PAYMENT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYEEHISTORYSEQNO attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYEEHISTORYSEQNO")
    
    objPayProcDO.FindPaymentsForPayeeList vxmlRequestNode, vxmlResponseNode
    
FindPaymentsForPayeeListExit:
    Set objPayProcDO = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CreatePayeeHistoryDetailsExit

Const strFunctionName As String = "CreatePayeeHistoryDetails"

Dim xmlNode As IXMLDOMNode
Dim strValue As String

    ' Raise error if no PAYEEHISTORY node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    ' Raise error if no THIRDPARTY node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY/THIRDPARTY")
    
    ' Raise error if no ADDRESS node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY/THIRDPARTY/ADDRESS")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.CreatePayeeHistoryDetails vxmlRequestNode, vxmlResponseNode
    
CreatePayeeHistoryDetailsExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePayeeHistoryDetailsExit

Const strFunctionName As String = "UpdatePayeeHistoryDetails"

Dim xmlNode As IXMLDOMNode
Dim strValue As String

    ' Raise error if no PAYEEHISTORY node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    'Validate PAYEEHISTORYSEQNO attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYEEHISTORYSEQNO")

    ' Raise error if no THIRDPARTY node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY/THIRDPARTY")
    
    ' Raise error if no ADDRESS node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYEEHISTORY/THIRDPARTY/ADDRESS")

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdatePayeeHistoryDetails vxmlRequestNode, vxmlResponseNode
    
UpdatePayeeHistoryDetailsExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CancelBalance(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo CancelBalanceExit

Const strFunctionName As String = "CancelBalance"

Dim strValue As String
Dim curTotalAmountPaid As Currency
Dim curTotalAmountDue As Currency

Dim xmlNode As IXMLDOMElement
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlPaidList As IXMLDOMNodeList
Dim xmlPaidNode As IXMLDOMElement
Dim xmlAttrib As IXMLDOMAttribute
Dim xmlRequest As IXMLDOMElement
Dim xmlElement As IXMLDOMElement

    ' Raise error if no APPLICATIONFEETYPE node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFEETYPE")
    
    'Validate AMOUNT attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")

    'Validate AMOUNT attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "AMOUNT")

    'Validate USERID attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "USERID")

    ' Raise error if no DISBURSEMENTPAYMENT node exists
    Set xmlElement = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD/DISBURSEMENTPAYMENT")

    'Check authority to create Cancellation Payment record
    strValue = vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text
    Set xmlRequest = vxmlRequestNode.cloneNode(False)
    If Not IsUserAuthorised(xmlRequest, strValue, "PProcCancelBalanceRole") Then
        errThrowError strFunctionName, oePPNotAuthCancelBalance
    End If

    Call FindFeeTypeList(vxmlRequestNode, vxmlResponseNode)
    
    'Loop through each APPLICATIONFEETYPE record found
    Set xmlNodeList = vxmlResponseNode.selectNodes("APPLICATIONFEETYPE")
    For Each xmlNode In xmlNodeList
        'Find total amount paid of all FEEPAYMENT's for this APPLICATIONFEETYPE
        curTotalAmountPaid = 0
        Set xmlPaidList = xmlNode.selectNodes("FEEPAYMENT[@AMOUNTPAID != """"]")
        For Each xmlPaidNode In xmlPaidList
            strValue = xmlGetAttributeText(xmlPaidNode, "AMOUNTPAID", "0")
            curTotalAmountPaid = curTotalAmountPaid + CCur(strValue)
        Next
        
        'Check if Total Amount Paid covers the Amount Due
        'SR 22/05/01 : SYS2298 - The column 'RebateOrAddition' in table ApplicationFeeType is not going to be used
        curTotalAmountDue = GetCurrencyAmount(xmlNode.getAttribute("AMOUNT")) '+ GetCurrencyAmount(xmlNode.getAttribute("REBATEORADDITION"))
        If curTotalAmountDue > curTotalAmountPaid Then
            errThrowError strFunctionName, oePPFeesOutstanding
        End If
    Next
    
    'Create Cancellation Payment Record
    xmlElement.setAttribute "ISSUEDATE", Format$(Now, "dd/mm/yyyy hh:mm:ss")
    xmlElement.setAttribute "CANCELLATIONDATE", Format$(Now, "dd/mm/yyyy hh:mm:ss")
    xmlElement.setAttribute "PAYEEHISTORYSEQNO", "0"
    Call CreateDisbursement(vxmlRequestNode, vxmlResponseNode)
    
CancelBalanceExit:
    
    Set xmlPaidList = Nothing
    Set xmlAttrib = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlRequest = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Function IsUserAuthorised( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal intUserRole As Integer, _
    ByVal strAuthorityRequirement As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Returns True/false depending whether the user
'   is authorised to perform the task requested
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo IsUserAuthorisedExit

Const strFunctionName As String = "IsUserAuthorised"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlChild As IXMLDOMElement
Dim xmlValidation As IXMLDOMElement

Dim objOrgBO As omOrg.OrganisationBO
Dim strResponse As String
Dim strValue As String

    If gobjContext Is Nothing Then
        errThrowError strFunctionName, oeMTSNotFound
    End If
    Set objOrgBO = gobjContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
'    Set xmlRequest = vxmlRequestNode.cloneNode(False)
    Set xmlValidation = xmlDoc.createElement("VALIDATION")
    vxmlRequestNode.appendChild xmlValidation
    Set xmlChild = xmlDoc.createElement("USERROLE")
    xmlChild.Text = CStr(intUserRole)
    xmlValidation.appendChild xmlChild
    Set xmlChild = xmlDoc.createElement("AUTHORITYREQUIREMENT")
    xmlChild.Text = strAuthorityRequirement
    xmlValidation.appendChild xmlChild
    
    strResponse = objOrgBO.ValidateUserAuthority(vxmlRequestNode.xml)
    If xmlDoc.loadXML(strResponse) Then
        strValue = xmlDoc.selectSingleNode(".//VALIDITY").Text
        If Not CBool(strValue) Then
            IsUserAuthorised = False
        Else
            IsUserAuthorised = True
        End If
    Else
        errThrowError strFunctionName, oeXMLParserError, "Invalid response from omOrg ValidateUserAuthority"
    End If

IsUserAuthorisedExit:
    
    Set xmlValidation = Nothing
    Set xmlChild = Nothing
    Set xmlDoc = Nothing
    Set objOrgBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Sub CreateReturnOfFunds(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateReturnOfFundsExit

    Const cstrFunctionName As String = "CreateReturnOfFunds"
    
    Dim xmlROFPaymentRecord As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim strValue As String

    Set xmlROFPaymentRecord = xmlGetMandatoryNode(vxmlRequestNode, ".//ROFPAYMENTRECORD")
    strValue = xmlGetMandatoryAttributeText(xmlROFPaymentRecord, "APPLICATIONNUMBER")
    
    strValue = vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text
    Set xmlNode = vxmlRequestNode.cloneNode(False)
    
    If Not IsUserAuthorised(vxmlRequestNode, strValue, "PProcDisbursementRole") Then
        errThrowError cstrFunctionName, oePPNotAuthCreateDisb
    End If
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreateReturnOfFunds vxmlRequestNode, vxmlResponseNode
    
CreateReturnOfFundsExit:

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub CreateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateDisbursementExit

Dim xmlNode As IXMLDOMElement
Dim strValue As String
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlElement As IXMLDOMElement
Dim xmlTempRequest As IXMLDOMNode
Dim xmlBankBuildingSocNode As IXMLDOMNode
Dim objAppTP As ApplicationThirdPartyBO
Dim objGlobalParameterDO As GlobalParameterDO
Dim xmlChildNode As IXMLDOMNode

Dim strResponse As String
Dim strPreferedPaymentDay As String
Dim strRepaymentBankAccountIndicator As String
Dim strAppNo As String
Dim strAppFFNo As String


Const strFunctionName As String = "CreateDisbursement"

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    'Get the APPLICATIONFACTFINDNUMBER
    strAppFFNo = xmlGetAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    
    'SYS2558 MDC 03/08/2001.
    If Len(Trim$(strAppFFNo)) = 0 Then
        'Get the ApplicationFactFind number
        strAppFFNo = GetFactFindNumberForApplication(strAppNo)
    End If
    'SYS2558 End

    'Set the CREATIONDATETIME to system date/time
    xmlNode.setAttribute "CREATIONDATETIME", Format$(Now, "dd/mm/yy hh:mm:ss")
    
    ' Raise error if no DISBURSEMENTPAYMENT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD/DISBURSEMENTPAYMENT")
    
    'Validate ISSUEDATE attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "ISSUEDATE")

    'Validate PAYEEHISTORYSEQNO attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYEEHISTORYSEQNO")
    
    Dim tmpNode As IXMLDOMNode
    Set tmpNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    
    'Validate APPLICATIONFACTFINDNUMBER attribute is specified
    'SYS2558 MDC 03/08/2001.
    'Not required as already found above
    'strAppFFNo = xmlGetAttributeText(tmpNode, "APPLICATIONFACTFINDNUMBER")
    'SYS2558 End
 
    'Check authority to create Disbursement records
    strValue = vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text
    Set xmlNode = vxmlRequestNode.cloneNode(False)
    If Not IsUserAuthorised(xmlNode, strValue, "PProcDisbursementRole") Then
        errThrowError strFunctionName, oePPNotAuthUpdateDisb
    End If

    Set objAppTP = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
    Set objGlobalParameterDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
   
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlElement = xmlDoc.createElement("APPLICATIONBANKBUILDINGSOC")
    xmlTempRequest.appendChild xmlElement
    Set xmlChildNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlChildNode.Text = strAppNo
    xmlElement.appendChild xmlChildNode
    Set xmlChildNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlChildNode.Text = strAppFFNo
    xmlElement.appendChild xmlChildNode
     
    strResponse = objAppTP.FindBankBuildingSocietyList(xmlTempRequest.xml)
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    ' SR 01/08/01 : SYS2552 - handle error noRecordFound
    Dim lngErrNo As Long
    
    'SYS2558 MDC 03/08/2001.
    'Disabling of error handler not required
    'On Error Resume Next
    strPreferedPaymentDay = ""  'SYS2558 MDC 14/08/2001. Initialise Payment Day
    'SYS2558 End
    lngErrNo = errCheckXMLResponseNode(xmlDoc.documentElement)
    If lngErrNo <> 0 Then
        lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        If lngErrNo <> oeRecordNotFound Then
            errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
        End If
    Else
        ' Take 'PreferredPaymentDay' from the record with 'RepaymentBankIndicator' as 1
        Dim strCondition As String
        strCondition = "//APPLICATIONBANKBUILDINGSOC[REPAYMENTBANKACCOUNTINDICATOR=1]"
        Set xmlBankBuildingSocNode = xmlDoc.selectSingleNode(strCondition)
        If Not xmlBankBuildingSocNode Is Nothing Then
            strPreferedPaymentDay = xmlGetNodeText(xmlBankBuildingSocNode, "PREFEREDPAYMENTDAY")
        End If
    End If
    
    'SYS2558 MDC 14/08/2001. Catch all for Preferred Payment Day
    If Len(Trim$(strPreferedPaymentDay)) = 0 Then
        'Default to the global parameter value
        strPreferedPaymentDay = objGlobalParameterDO.GetCurrentParameterByType("PreferredPaymentDay", "AMOUNT")
    End If
    'SYS2558 End
    
    Set xmlElement = xmlDoc.createElement("PREFERREDPAYMENTDAY")
    xmlElement.Text = strPreferedPaymentDay
    
    vxmlRequestNode.appendChild xmlElement
       
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CreateDisbursement vxmlRequestNode, vxmlResponseNode
    
CreateDisbursementExit:
    Set xmlNode = Nothing
    Set xmlElement = Nothing
    Set objAppTP = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlBankBuildingSocNode = Nothing
    Set xmlChildNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub UpdateDisbursement( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'RF     27/02/02    SYS4193 Add new workaround method TemporaryCompleteInterfacing
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDisbursementExit

    Const cstrFunctionName As String = "UpdateDisbursement"
  
    'Dim nodeRequest As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim elemTemp As IXMLDOMElement
    Dim nodeTempRequest As IXMLDOMNode
    Dim nodeBankBuildingSoc As IXMLDOMNode
    Dim objAppTP As ApplicationThirdPartyBO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlRoot As IXMLDOMNode                      ' PSC 14/10/2002 BMIDS00337
    
    Dim strUserAuthLevel As String
    Dim strResponse As String
    Dim strPreferedPaymentDay As String
    Dim strAppNo As String
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    '------------------------------------------------------------------------------------------
    'Check authority to create Cancellation Payment record
    '------------------------------------------------------------------------------------------
    
    strUserAuthLevel = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    'Set nodeRequest = xmlGetRequestNode(vxmlRequestNode)
    Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
    'If Not IsUserAuthorised(nodeRequest, strUserAuthLevel, "PProcDisbursementRole") Then
    If Not IsUserAuthorised(nodeTempRequest, strUserAuthLevel, "PProcDisbursementRole") Then
        errThrowError cstrFunctionName, oePPNotAuthUpdateDisb
    End If
    
    '------------------------------------------------------------------------------------------
    ' get PreferredPaymentDay
    '------------------------------------------------------------------------------------------
    
    ' Raise error if no PAYMENTRECORD node exists
    xmlCheckMandatoryNode vxmlRequestNode, "PAYMENTRECORD"

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
        
    Set objAppTP = gobjContext.CreateInstance( _
        gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
    Set objGlobalParameterDO = gobjContext.CreateInstance( _
        gstrBASE_COMPONENT & ".GlobalParameterDO")

    Set nodeTempRequest = xmlGetRequestNode(vxmlRequestNode)
    
    ' PSC 14/10/2002 BMIDS00337 - Start
    Set xmlRoot = xmlDoc.createElement("APPLICATIONBANKBUILDINGSOC")
    nodeTempRequest.appendChild xmlRoot
    strAppNo = xmlGetMandatoryNodeText(vxmlRequestNode, "PAYMENTRECORD/@APPLICATIONNUMBER")
    Set elemTemp = xmlDoc.createElement("APPLICATIONNUMBER")
    elemTemp.Text = strAppNo
    xmlRoot.appendChild elemTemp
    ' PSC 14/10/2002 BMIDS00337 - End
    
    strResponse = objAppTP.FindBankBuildingSocietyList(nodeTempRequest.xml)
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    
    ' SR 01/08/01 : SYS2552 - handle error noRecordFound
    Dim lngErrNo As Long
    
    lngErrNo = errCheckXMLResponseNode(xmlDoc.documentElement, , False)
    If lngErrNo <> 0 Then
        lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        
        ' PSC 14/10/2002 BMIDS00337 - Start
        If lngErrNo <> oeRecordNotFound Then
            errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
        End If
        ' PSC 14/10/2002 BMIDS00337 - End
    Else
        ' Take 'PreferredPaymentDay' from the record with 'RepaymentBankIndicator' as 1
        Dim strCondition As String
        strCondition = "//APPLICATIONBANKBUILDINGSOC[REPAYMENTBANKACCOUNTINDICATOR=1]"
        Set nodeBankBuildingSoc = xmlDoc.selectSingleNode(strCondition)
            
        ' PSC 14/10/2002 BMIDS00337 - Start
        If Not nodeBankBuildingSoc Is Nothing Then
            strPreferedPaymentDay = xmlGetNodeText(nodeBankBuildingSoc, "PREFEREDPAYMENTDAY")
        End If
        ' PSC 14/10/2002 BMIDS00337 - End
    End If
    
    ' PSC 14/10/2002 BMIDS00337 - Start
    If Len(Trim$(strPreferedPaymentDay)) = 0 Then
        strPreferedPaymentDay = objGlobalParameterDO.GetCurrentParameterByType("PreferredPaymentDay", "AMOUNT")
    End If
    ' PSC 14/10/2002 BMIDS00337 - End
    
    Set elemTemp = xmlDoc.createElement("PREFERREDPAYMENTDAY")
    elemTemp.Text = strPreferedPaymentDay
    
    vxmlRequestNode.appendChild elemTemp
    
    gobjTxBO.UpdateDisbursement vxmlRequestNode, vxmlResponseNode

UpdateDisbursementExit:

    'Set nodeRequest = Nothing
    Set xmlDoc = Nothing
    Set elemTemp = Nothing
    Set nodeTempRequest = Nothing
    Set nodeBankBuildingSoc = Nothing
    Set objAppTP = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub DeleteApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteApplicationFeeTypesExit

Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim strValue As String

Const strFunctionName As String = "DeleteApplicationFeeTypes"

    Set xmlNodeList = vxmlRequestNode.selectNodes("APPLICATIONFEETYPE")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONFEETYPE element has an APPLICATIONNUMBER attribute value
        For Each xmlNode In xmlNodeList
            strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        Next
    Else
        errThrowError strFunctionName, oeXMLMissingElement, "Missing APPLICATIONFEETYPE element"
    End If
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.DeleteApplicationFeeTypes vxmlRequestNode, vxmlResponseNode
    
DeleteApplicationFeeTypesExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub DeletePayeeHistory(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeletePayeeHistoryExit

Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim strValue As String

Const strFunctionName As String = "DeletePayeeHistory"

    Set xmlNodeList = vxmlRequestNode.selectNodes("PAYEEHISTORY")
    If xmlNodeList.length > 0 Then
        'Check each APPLICATIONFEETYPE element has an APPLICATIONNUMBER attribute value
        For Each xmlNode In xmlNodeList
            strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
            strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYEEHISTORYSEQNO")
        Next
    Else
        errThrowError strFunctionName, oeXMLMissingElement, "Missing PAYEEHISTORY element"
    End If
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.DeletePayeeHistory vxmlRequestNode, vxmlResponseNode
    
DeletePayeeHistoryExit:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub DeleteFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteFeeTypePaymentExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "DeleteFeeTypePayment"

    ' Raise error if no FEEPAYMENT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "FEEPAYMENT")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    'Validate FEETYPE attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPE")
    
    'Validate FEETYPESEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "FEETYPESEQUENCENUMBER")
    
    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.DeleteFeeTypePayment vxmlRequestNode, vxmlResponseNode
    
DeleteFeeTypePaymentExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindFullPaymentList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindFullPaymentListExit

Const strFunctionName As String = "FindFullPaymentList"

Dim objPayProcDO As IPAymentProcessingDO
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlRequest As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempNode As IXMLDOMElement
Dim xmlAttrib As IXMLDOMAttribute
Dim xmlElement As IXMLDOMElement

Dim strValue As String
Dim strPattern As String

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlTempResponse
    
    ' Raise error if no PAYMENTRECORD node exists
    Set xmlElement = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    
    'If ISSUEDATE attribute specified, remove it, saving for later use
    Set xmlAttrib = xmlElement.Attributes.getNamedItem("ISSUEDATE")
    If Not xmlAttrib Is Nothing Then
        xmlElement.removeAttribute "ISSUEDATE"
    End If
    
    'Find all matching PAYMENTRECORD records
    objPayProcDO.FindPaymentHistoryList vxmlRequestNode, vxmlResponseNode
    
    'If required find subset of PAYMENT nodes (where ISSUEDATE <= date entered if specified)
    strPattern = "PAYMENTRECORD"
    If Not xmlAttrib Is Nothing Then
        strValue = xmlAttrib.Text
        If IsDate(strValue) Then
            strPattern = "PAYMENTRECORD[date(@ISSUEDATE) <= date(""" & Format$(strValue, "yyyy-mm-dd") & """)]"
        End If
    End If
    Set xmlNodeList = vxmlResponseNode.selectNodes(strPattern)
    If xmlNodeList.length = 0 Then
        strValue = "No Payment records found"
        If Not xmlAttrib Is Nothing Then
            strValue = strValue & " prior to Issue Date"
        End If
        errThrowError strFunctionName, oeRecordNotFound, strValue
    End If
    
    'For each Payment record, find associated Payee History records
    For Each xmlNode In xmlNodeList
        'Create a PayeeHistory Request
        Set xmlRequest = xmlDoc.createElement("REQUEST")
        Set xmlTempNode = xmlDoc.createElement("PAYEEHISTORY")
        xmlRequest.appendChild xmlTempNode
        strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        xmlTempNode.setAttribute "APPLICATIONNUMBER", strValue
        strValue = xmlGetMandatoryAttributeText(xmlNode, "DISBURSEMENTPAYMENT/PAYEEHISTORYSEQNO")
        strValue = xmlNode.Attributes.getNamedItem("PAYEEHISTORYSEQNO").Text
        xmlTempNode.setAttribute "PAYEEHISTORYSEQNO", strValue
        
        FindPayeeHistoryList xmlRequest, xmlTempResponse
        
        'Append results to current PAYMENT node
        Set xmlTempNode = xmlTempResponse.selectSingleNode("PAYEEHISTORY")
        If Not xmlTempNode Is Nothing Then
            xmlNode.appendChild xmlTempNode
        End If
    Next
    
FindFullPaymentListExit:
    Set objPayProcDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlRequest = Nothing
    Set xmlDoc = Nothing
    Set xmlTempNode = Nothing
    Set xmlAttrib = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub GetPaymentSummary(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo GetPaymentSummaryExit

Const strFunctionName As String = "GetPaymentSummary"

Dim objMSQBO As MortgageSubQuoteBO
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlChildNode As IXMLDOMNode
Dim xmlElement As IXMLDOMElement
Dim xmlCalcElement As IXMLDOMElement
Dim xmlTempResponse As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlMortgageSubQuote As IXMLDOMNode 'CPWP1/BM020 - DPF 20/09/02
Dim xmlDrawDownAmount As IXMLDOMNode 'CPWP1/BM037 - DPF 30/09/02
Dim xmlOneOffCosts As IXMLDOMNode
Dim xmlLoanComponents As IXMLDOMNode
Dim xmlMortgageIncentives As IXMLDOMNode
Dim xmlFeeTypes As IXMLDOMNode
Dim xmlAttrib As IXMLDOMAttribute
Dim xmlPayment As IXMLDOMNode
Dim xmlPaymentList As IXMLDOMNodeList
Dim xmlBankBuildingSocNode As IXMLDOMNode
Dim xmlFeeTypeNode As IXMLDOMNode
Dim xmlFeePaymentNode As IXMLDOMNode
Dim xmlFeeNode As IXMLDOMNode
Dim objAppTP As ApplicationThirdPartyBO
Dim objComboDO As ComboDO
Dim objAppProcBO As Object

Dim strAppNo As String
Dim strAppFFNo As String
Dim strValue As String
Dim strResponse As String
Dim strMSQNumber As String
Dim lngValue As Long, lngAmount As Long, lngErrNo As Long, lngDrawDown As Long
Dim lngTotal As Long, lngAmountRequested As Long
Dim lngBalance As Long, lngTotalRefundAmount As Long, lngRetentionPaymentAmount As Long
Dim lngIncentivePaymentAmount As Long
Dim strPattern As String
Dim strPreferedPaymentDay As String
Dim strRepaymentBankAccountIndicator As String

Dim blnHasManualIncentive As Boolean            ' PSC 22/10/2002 BMIDS00700

' PSC 11/11/2002 BMIDS00599 - Start
Dim strPaymentSeqNo As String
Dim xmlReturns As IXMLDOMNodeList
Dim xmlReturn As IXMLDOMNode
' PSC 11/11/2002 BMIDS00599 - End

' PSC 28/11/2002 BMIDS01099 - Start
Dim strAssocPaySeqNo As String
Dim xmlIncentiveNode As IXMLDOMNode
' PSC 28/11/2002 BMIDS01099 - End

' PSC 26/02/2007 EP2_1347 - Start
Dim blnValRefund As Boolean
Dim lngValRefundAmount As Long
Dim lngValRefundPaymentAmount As Long
Dim intFeeType As String
Dim strValRefundPaymentId As String
Dim xmlValRefundNode As IXMLDOMNode
' PSC 26/02/2007 EP2_1347 - End

    ' Raise error if no PAYMENTRECORD node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")

    'Validate APPLICATIONNUMBER attribute is specified
    strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")

    'Get the ApplicationFactFindNumber for this Application
    strAppFFNo = GetFactFindNumberForApplication(strAppNo)
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    
    'Get the Mortgage Sub Quote Number for this ApplicationFactFind
    Call GetLoanComponents(strAppNo, xmlTempResponse)
    strMSQNumber = xmlGetMandatoryNodeText(xmlTempResponse, ".//MORTGAGESUBQUOTENUMBER")
    
    'Get MortgageSubQuote & OneOffCosts details for the Application/MortgageSubQuote
    Set objMSQBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlNode = xmlDoc.createElement("MORTGAGESUBQUOTE")
    Set xmlChildNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlChildNode.Text = strAppNo
    xmlNode.appendChild xmlChildNode
    Set xmlChildNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlChildNode.Text = strAppFFNo
    xmlNode.appendChild xmlChildNode
    Set xmlChildNode = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlChildNode.Text = strMSQNumber
    xmlNode.appendChild xmlChildNode
    xmlTempRequest.appendChild xmlNode
    
    '--- CPWP1/BM020 - DPF 20/09/2002 - Start ---
    strResponse = objMSQBO.GetMortgageSubQuoteData(xmlTempRequest.xml)
    lngValue = errCheckXMLResponse(strResponse, True)
    xmlDoc.loadXML strResponse
    Set xmlMortgageSubQuote = xmlGetNode(xmlDoc.documentElement, "//MANUALINCENTIVEAMOUNT")
    Set xmlDrawDownAmount = xmlGetNode(xmlDoc.documentElement, "//DRAWDOWN")
                                                                  
    '---- CPWP1/BM020 - DPF 20/09/2002 - End ----
    xmlChangeNodeName xmlTempRequest, "MORTGAGESUBQUOTE", "ONEOFFCOST"
    strResponse = objMSQBO.GetOneOffCostsDetails(xmlTempRequest.xml)
    lngValue = errCheckXMLResponse(strResponse, True)
    xmlDoc.loadXML strResponse
    Set xmlOneOffCosts = xmlDoc.documentElement
    
    '------------------------------------------------------------
    'SR 11/07/01 : SYS2412 - Get details of Mortgage Incentives
    '------------------------------------------------------------
    ' First, call FindLoanComponentList
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlNode = xmlDoc.createElement("MORTGAGESUBQUOTE")
    Set xmlChildNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlChildNode.Text = strAppNo
    xmlNode.appendChild xmlChildNode
    Set xmlChildNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlChildNode.Text = strAppFFNo
    xmlNode.appendChild xmlChildNode
    Set xmlChildNode = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlChildNode.Text = strMSQNumber
    xmlNode.appendChild xmlChildNode
    xmlTempRequest.appendChild xmlNode
    
    strResponse = objMSQBO.FindLoanComponentList(xmlTempRequest.xml)
    lngValue = errCheckXMLResponse(strResponse, True)
    xmlDoc.loadXML strResponse
    Set xmlLoanComponents = xmlDoc.selectSingleNode("//LOANCOMPONENTLIST")
    
    'Call FindMortgageIncentivesList
    xmlChangeNodeName xmlTempRequest, "MORTGAGESUBQUOTE", "MORTGAGEINCENTIVE"
    xmlTempRequest.appendChild xmlLoanComponents
    
    strResponse = objMSQBO.FindMortgageIncentiveList(xmlTempRequest.xml)
    xmlDoc.loadXML strResponse
    lngErrNo = errCheckXMLResponseNode(xmlDoc.documentElement, , False)

    If lngErrNo <> 0 Then
        lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        If lngErrNo <> oeRecordNotFound Then
            errThrowError strFunctionName, lngErrNo
        End If
    Else
        Set xmlMortgageIncentives = xmlDoc.selectSingleNode("//MORTGAGEINCENTIVELIST")
    End If
    
    '----------------------------
    'SR 11/07/01 : SYS2412 - End
    '-----------------------------
    
    'Get Fee details
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("APPLICATIONFEETYPE")
    xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlTempRequest.appendChild xmlElement
    Set xmlTempResponse = xmlDoc.createElement("APPLICATIONFEETYPELIST")
    On Error GoTo FindFeeError
    Call FindFeeTypeList(xmlTempRequest, xmlTempResponse)
    On Error GoTo GetPaymentSummaryExit
    
    'Get Disbursement details
    Call FindPaymentHistoryList(vxmlRequestNode, vxmlResponseNode, False)
    Set xmlNode = vxmlResponseNode.selectSingleNode(".//PAYMENTRECORD")
    If Not xmlNode Is Nothing Then
        'Find each PAYMENTRECORD/FEEPAYMENT
        'For each Fee Payment, Get the application fee types for each Fee Payment record
        Set xmlNodeList = vxmlResponseNode.selectNodes("PAYMENTRECORD/FEEPAYMENT")
        For Each xmlFeePaymentNode In xmlNodeList
            'Find associated APPLICATIONFEETYPE
            strPattern = "APPLICATIONFEETYPE[@APPLICATIONNUMBER=""" & strAppNo & """ and @FEETYPE=""" _
                        & xmlGetAttributeText(xmlFeePaymentNode, "FEETYPE") & """ and @FEETYPESEQUENCENUMBER=""" _
                        & xmlGetAttributeText(xmlFeePaymentNode, "FEETYPESEQUENCENUMBER") & """]"
            Set xmlFeeTypeNode = xmlGetNode(xmlTempResponse, strPattern)
            If Not xmlFeeTypeNode Is Nothing Then
                Set xmlFeeNode = xmlFeeTypeNode.cloneNode(False)
                xmlFeePaymentNode.parentNode.appendChild xmlFeeNode
            End If
        Next
    End If
       
    'Make Calculations
    '~~~~~~~~~~~~~~~~~
    Set xmlCalcElement = xmlDoc.createElement("CALCULATIONS")
    
    'Amount Requested
    strValue = xmlGetNodeText(xmlOneOffCosts, ".//MORTGAGESUBQUOTE/AMOUNTREQUESTED")
    If Len(Trim$(strValue)) = 0 Then
        strValue = "0"
    End If
    lngAmountRequested = CSafeLng(strValue)
    xmlCalcElement.setAttribute "AMOUNTREQUESTED", strValue
    
    'Fees To Be Added To Loan
    lngValue = 0
    Set xmlNodeList = xmlOneOffCosts.selectNodes("ONEOFFCOST/MORTGAGEONEOFFCOSTLIST/MORTGAGEONEOFFCOST[ADDTOLOAN=""1""]/AMOUNT")
    For Each xmlNode In xmlNodeList
        lngValue = lngValue + CSafeLng(xmlNode.Text)
    Next
    xmlCalcElement.setAttribute "FEESTOBEADDEDTOLOAN", CStr(lngValue)
    
    'Total To Be Disbursed
    strValue = CStr(lngValue + CSafeLng(strValue))
    xmlCalcElement.setAttribute "TOTALTOBEDISBURSED", strValue
    
    'Balance
    lngValue = 0: lngRetentionPaymentAmount = 0: lngIncentivePaymentAmount = 0
    Dim xmlTempNode As IXMLDOMNode
    Dim lngReturnFunds  As Long, lngTemp As Long
    Dim strReturnOfFundsId As String, strCancelledStatusID As String, strRetentionReleaseId As String
    Dim strIncentivePaymentId As String, strPaymentType As String
   
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    strReturnOfFundsId = objComboDO.GetFirstComboValueId("PaymentType", "N")
    strCancelledStatusID = objComboDO.GetFirstComboValueId("PaymentStatus", "C")
    strRetentionReleaseId = objComboDO.GetFirstComboValueId("PaymentType", "R")
    strIncentivePaymentId = objComboDO.GetFirstComboValueId("PaymentType", "C")
    strValRefundPaymentId = objComboDO.GetFirstComboValueId("PaymentType", "VALREFUND")
   
    Set xmlNodeList = vxmlResponseNode.selectNodes("PAYMENTRECORD[DISBURSEMENTPAYMENT]")
    For Each xmlElement In xmlNodeList
        Set xmlTempNode = xmlElement.selectSingleNode("DISBURSEMENTPAYMENT")
       'SR 11-07-01 : SYS2412 - The returned funds are stored as -ve amounts, so simply add all the
       '              amounts (for the amount paid) and, deduct it from amount due to find amount outstanding
       ' If xmlGetAttributeText(xmlTempNode, "PAYMENTTYPE") = strRetunOfFundsId Then
       '     lngReturnFunds = lngReturnFunds + CSafeLng(xmlElement.getAttribute("AMOUNT"))
       ' Else
       '     lngValue = lngValue + CSafeLng(xmlElement.getAttribute("AMOUNT"))
       ' End If
        If xmlGetAttributeText(xmlTempNode, "PAYMENTSTATUS") <> strCancelledStatusID Then
            lngTemp = CSafeLng(xmlElement.getAttribute("AMOUNT"))
            'SR 10-12-2001 : SYS3394
            'lngValue = lngValue + lngTemp
            strPaymentType = xmlGetAttributeText(xmlTempNode, "PAYMENTTYPE")
            
            ' PSC 28/11/2002 BMIDS01099 - Start
            ' If return of funds check whether it is for an incentive release or a valuation fee refund
            If strPaymentType = strReturnOfFundsId Then
                strAssocPaySeqNo = xmlGetAttributeText(xmlElement, "ASSOCPAYSEQNUMBER")
                
                strPattern = "PAYMENTRECORD/DISBURSEMENTPAYMENT[@PAYMENTSEQUENCENUMBER=" & strAssocPaySeqNo & _
                             " and @PAYMENTTYPE=" & strIncentivePaymentId & "]"
                Set xmlIncentiveNode = vxmlResponseNode.selectSingleNode(strPattern)
            
                ' PSC 26/02/2007 EP2_1347 - Start
                strPattern = "PAYMENTRECORD/DISBURSEMENTPAYMENT[@PAYMENTSEQUENCENUMBER=" & strAssocPaySeqNo & _
                " and @PAYMENTTYPE=" & strValRefundPaymentId & "]"

                Set xmlValRefundNode = vxmlResponseNode.selectSingleNode(strPattern)
                ' PSC 26/02/2007 EP2_1347 - End
            End If
            
            ' PSC 26/02/2007 EP2_1347
            ' Only add on if this is not an incentive release and not a return of funds for an incentive release
            If strPaymentType <> strIncentivePaymentId And xmlIncentiveNode Is Nothing And _
               strPaymentType <> strValRefundPaymentId And xmlValRefundNode Is Nothing Then
                lngValue = lngValue + lngTemp
            End If
            'SR 10-12-2001 : SYS3394
            ' PSC 28/11/2002 BMIDS01099 - End
            
            'Sum up all payments that are of type 'RetentionRelease' and IncentivePayment
            If strPaymentType = strRetentionReleaseId Then
                lngRetentionPaymentAmount = lngRetentionPaymentAmount + lngTemp
                
                ' PSC 11/11/2002 BMIDS00599 - Start
                ' Find returned retentions and take them into account
                strPaymentSeqNo = xmlGetMandatoryAttributeText(xmlElement, "PAYMENTSEQUENCENUMBER")

                strPattern = "//PAYMENTRECORD[@ASSOCPAYSEQNUMBER=" + strPaymentSeqNo + "]" & _
                             "[DISBURSEMENTPAYMENT/@PAYMENTTYPE=" + strReturnOfFundsId + "]"

                Set xmlReturns = vxmlResponseNode.selectNodes(strPattern)

                For Each xmlReturn In xmlReturns
                    lngRetentionPaymentAmount = lngRetentionPaymentAmount + CSafeLng(xmlGetAttributeText(xmlReturn, "AMOUNT"))
                Next
                ' PSC 11/11/2002 BMIDS00599 - End
                ' PSC 28/11/2002 BMIDS01099
            ElseIf strPaymentType = strIncentivePaymentId Or Not xmlIncentiveNode Is Nothing Then
                lngIncentivePaymentAmount = lngIncentivePaymentAmount + lngTemp
            ' PSC 26/02/2007 EP2_1347
            ElseIf strPaymentType = strValRefundPaymentId Or Not xmlValRefundNode Is Nothing Then
                lngValRefundPaymentAmount = lngValRefundPaymentAmount + lngTemp
            End If
            
        End If
        Set xmlIncentiveNode = Nothing
        ' PSC 26/02/2007 EP2_1347
        Set xmlValRefundNode = Nothing
    Next
    
    lngBalance = lngAmountRequested - lngValue
    
    'Fees To Be Deducted From Advance and Total Refunds
    lngTotal = 0: lngTotalRefundAmount = 0
    
    ' PSC 26/02/2007 EP2_1347
    blnValRefund = GetGlobalParamBoolean("CalculateValuationRefund")
    
    Set xmlNodeList = xmlTempResponse.selectNodes("APPLICATIONFEETYPE")
    For Each xmlElement In xmlNodeList
        lngValue = 0
        'Get the Fee Type Amount
        'SR 22/05/01 : SYS2298 - The column 'RebateOrAddition' in the table ApplicationFeeType is not going to be used
        lngAmount = xmlGetAttributeAsLong(xmlElement, "AMOUNT", "0") '+ xmlGetAttributeAsLong(xmlElement, "REBATEORADDITION", "0")
        'Find the total of all fee payments made for this fee type
        Set xmlPaymentList = xmlElement.selectNodes("FEEPAYMENT")
        For Each xmlPayment In xmlPaymentList
            lngValue = lngValue + xmlGetAttributeAsLong(xmlPayment, "AMOUNTPAID")
            'SR 11/07/01 : SYS2412
            ' PSC 26/02/2007 EP2_1347 - Start
            intFeeType = xmlGetAttributeAsInteger(xmlPayment, "FEETYPE")
            
            If blnValRefund = False Or (blnValRefund = True And Not IsValidationType("OneOffCost", intFeeType, "VAL")) Then
                lngTotalRefundAmount = lngTotalRefundAmount + xmlGetAttributeAsLong(xmlPayment, "REFUNDAMOUNT")
            End If
            ' PSC 26/02/2007 EP2_1347 - End
        Next
        lngTotal = lngTotal + lngAmount - lngValue
    Next
    
    ' PSC 26/02/2007 EP2_1347 - Start
    Set xmlNodeList = xmlOneOffCosts.selectNodes("ONEOFFCOST/MORTGAGEONEOFFCOSTLIST/MORTGAGEONEOFFCOST[REFUNDAMOUNT!='']")
    For Each xmlNode In xmlNodeList
        If IsValidationType("OneOffCost", xmlGetMandatoryNodeText(xmlNode, "MORTGAGEONEOFFCOSTTYPE"), "VAL") Or _
           IsValidationType("OneOffCost", xmlGetMandatoryNodeText(xmlNode, "MORTGAGEONEOFFCOSTTYPE"), "TPV") Or _
           IsValidationType("OneOffCost", xmlGetMandatoryNodeText(xmlNode, "MORTGAGEONEOFFCOSTTYPE"), "TPVA") Then
    
            lngValRefundAmount = lngValRefundAmount + xmlGetNodeAsLong(xmlNode, "REFUNDAMOUNT")
        End If
    Next
    ' PSC 26/02/2007 EP2_1347 - End
    
    xmlCalcElement.setAttribute "FEESTOBEDEDUCTEDFROMADVANCE", CStr(lngTotal)
    
    ' PSC 26/02/2007 EP2_1347 - Start
    xmlCalcElement.setAttribute "TOTALREFUND", CStr(lngTotalRefundAmount + lngValRefundAmount)
    xmlCalcElement.setAttribute "OUTSTANDINGVALUATIONREFUND", CStr(lngValRefundAmount - lngValRefundPaymentAmount)
    ' PSC 26/02/2007 EP2_1347 - End
    
    'lngBalance = lngBalance - lngTotal  JLD SYS4177
    
    '--- CPWP1/BM020 - DPF 20/09/02 - START ---
    'Calculate Total Incentives oustanding, a manual incentive amount has been entered this
    'over-rides the incentive values for the selected sub quote
    
    Dim strFinancialBenefitId As String
    strFinancialBenefitId = objComboDO.GetFirstComboValueId("IncentiveBenefitType", "F")
    lngValue = 0 ' JLD SYS4115
    
    ' PSC 22/10/2002 BMIDS00700 - Start
    blnHasManualIncentive = False
    
    If Not xmlMortgageSubQuote Is Nothing Then
        If Len(xmlMortgageSubQuote.Text) > 0 Then
            If (CLng(xmlMortgageSubQuote.Text)) > 0 Then 'BMIDS930 KRW 25/10/04
                blnHasManualIncentive = True
            End If
        End If
    End If
    
    If Not xmlMortgageIncentives Is Nothing And (Not blnHasManualIncentive) Then
        Set xmlNodeList = xmlMortgageIncentives.selectNodes(".//MORTGAGEINCENTIVE")
        lngValue = 0
        For Each xmlElement In xmlNodeList
            If xmlGetNodeText(xmlElement, ".//INCENTIVEBENEFITTYPE") = strFinancialBenefitId Then
                lngValue = lngValue + xmlGetNodeText(xmlElement, ".//INCENTIVEAMOUNT")
            End If
        Next xmlElement
    ElseIf blnHasManualIncentive Then 'If a Manual Incentive Amount has been entered
        lngValue = 0
        lngValue = xmlMortgageSubQuote.Text
    End If
    ' PSC 22/10/2002 BMIDS00700 - End
    
    'assign attribute to send back to the GUI
    xmlCalcElement.setAttribute "MORTGAGEINCENTIVE", CStr(lngValue) - lngIncentivePaymentAmount
      
    '---- CPWP1/BM020 - DPF 20/09/02 - END ----
      
    '------------------------------
    'Calculate Retention Amount
    '------------------------------
    ' SR 10-12-2001 : SYS3378 - Get data from Valuation Report
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
     
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETVALUATIONREPORT"
    
    Set xmlNode = xmlDoc.createElement("VALUATION")
    xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strAppNo
    xmlSetAttributeValue xmlNode, "APPLICATIONFACTFINDNUMBER", strAppFFNo
    xmlTempRequest.appendChild xmlNode
     
    strResponse = objAppProcBO.OmAppProcRequest(xmlTempRequest.xml)
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    
    lngErrNo = errCheckXMLResponseNode(xmlDoc.documentElement, , False)
    If lngErrNo <> 0 Then
        lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        
        If lngErrNo <> oeRecordNotFound Then errCheckXMLResponseNode xmlDoc.documentElement, , True
    End If
    
    Set xmlTempNode = xmlDoc.selectSingleNode("//GETVALUATIONREPORT")
    If Not xmlTempNode Is Nothing Then
        lngValue = xmlGetAttributeAsLong(xmlTempNode, "RETENTIONSROADS") + xmlGetAttributeAsLong(xmlTempNode, "RETENTIONWORKS")
        lngValue = lngValue - lngRetentionPaymentAmount
    Else
        lngValue = (-1) * lngRetentionPaymentAmount
    End If
    
    xmlCalcElement.setAttribute "RETENTIONAMOUNT", CStr(lngValue)
    
    '---- CPWP1/BM037 - DPF 30/09/02 - START ----
    ' PSC 22/10/2002 BMIDS00700 - Start
    If Not xmlDrawDownAmount Is Nothing Then
        If CSafeDbl(xmlDrawDownAmount.Text) > 0 Then    'BMID793 GHun
            lngDrawDown = xmlDrawDownAmount.Text
            lngBalance = lngBalance - lngDrawDown
            xmlCalcElement.setAttribute "DRAWDOWN", CStr(lngDrawDown)
        End If
    End If
    ' PSC 22/10/2002 BMIDS00700 - End
    '---- CPWP1/BM037 - DPF 30/09/02 - END ----
    
    lngBalance = lngBalance - lngValue
    xmlCalcElement.setAttribute "BALANCE", CStr(lngBalance)
        
    vxmlResponseNode.appendChild xmlCalcElement

GetPaymentSummaryExit:
    Set objMSQBO = Nothing
    Set objAppTP = Nothing
    Set objComboDO = Nothing
    Set objAppProcBO = Nothing
    
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlChildNode = Nothing
    Set xmlElement = Nothing
    Set xmlCalcElement = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempRequest = Nothing
    Set xmlOneOffCosts = Nothing
    Set xmlLoanComponents = Nothing
    Set xmlMortgageIncentives = Nothing
    Set xmlMortgageSubQuote = Nothing
    Set xmlDrawDownAmount = Nothing
    Set xmlFeeTypes = Nothing
    Set xmlPayment = Nothing
    Set xmlPaymentList = Nothing
    Set xmlFeeTypeNode = Nothing
    Set xmlFeePaymentNode = Nothing
    Set xmlFeeNode = Nothing
    Set xmlIncentiveNode = Nothing          ' PSC 28/11/2002 BMIDS01099
    
    ' PSC 11/11/2002 BMIDS00599 - Start
    Set xmlReturns = Nothing
    Set xmlReturn = Nothing
    ' PSC 11/11/2002 BMIDS00599 - End
    
    Set xmlValRefundNode = Nothing ' PSC 26/02/2007 EP2_1347
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
FindFeeError:
    If errGetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    Else
        GoTo GetPaymentSummaryExit
    End If
End Sub

Private Function GetCurrencyAmount(ByVal strAmount As String) As Currency

On Error GoTo GetCurrencyAmountExit

Dim curAmount As Currency

    curAmount = CCur(strAmount)
    GetCurrencyAmount = curAmount
    
GetCurrencyAmountExit:
    If Err.Number <> 0 Then
        GetCurrencyAmount = 0
    End If
    
End Function

Private Sub FindLoanComponentPaymentList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified CASETASK entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CASETASK node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo FindLoanComponentPaymentListExit
    
    Const strFunctionName As String = "FindLoanComponentPaymentList"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
        
    ' validate REQUEST & select CASETASK node
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "LOANCOMPONENTPAYMENT")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    gobjDO.FindLoanComponentPaymentList xmlCaseTaskNode, vxmlResponseNode
    
    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindLoanComponentPaymentListExit:

    Set xmlCaseTaskNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateFirstPaymentDateAndCreateBankBuildingSoc( _
            ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateExit

Const strFunctionName As String = "UpdateFirstPaymentDateAndCreateBankBuildingSoc"

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdateFirstPaymentDateAndCreateBankBuildingSoc vxmlRequestNode, vxmlResponseNode
     

UpdateExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateFirstPaymentDateAndBankBuildingSoc( _
            ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateExit

Const strFunctionName As String = "UpdateFirstPaymentDateAndBankBuildingSoc"

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdateFirstPaymentDateAndBankBuildingSoc vxmlRequestNode, vxmlResponseNode
     

UpdateExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateFirstPaymentDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateFirstPaymentDateExit

Dim xmlNode As IXMLDOMNode
Dim xmlRequest As IXMLDOMNode
Dim strValue As String
Dim strCheckValue As String
Dim xmlLCNode As IXMLDOMNode

Const strFunctionName As String = "UpdateFirstPaymentDate"

    ' Check to see if attribute ApplicationNumber has a value
    Set xmlLCNode = vxmlRequestNode.selectSingleNode("LOANCOMPONENTPAYMENT")
    strCheckValue = xmlGetAttributeText(xmlLCNode, "APPLICATIONNUMBER")
    If Len(strCheckValue) = 0 Then
        errThrowError strFunctionName, 559
    End If
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.UpdateFirstPaymentDate vxmlRequestNode, vxmlResponseNode
    
UpdateFirstPaymentDateExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CompleteInterfacing( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'DM     03/01/02    SYS3185 Implement CompleteInterfacing
'RF     03/01/02    Complete SYS3185 (Implement CompleteInterfacing)
'------------------------------------------------------------------------------------------
On Error GoTo CompleteInterfacingExit

    Const cstrFunctionName As String = "CompleteInterfacing"

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.CompleteInterfacing vxmlRequestNode, vxmlResponseNode
    
CompleteInterfacingExit:

    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub

End Sub

'GD ADDED BM0198 START
Private Sub SaveCostsFeesDisbursements(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description: Created for BM0198
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
On Error GoTo SaveCostsFeesDisbursementsExit

Const strFunctionName As String = "SaveCostsFeesDisbursements"

    Dim strValue As String
    Dim xmlNode As IXMLDOMNode
    

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    gobjTxBO.SaveCostsFeesDisbursements vxmlRequestNode, vxmlResponseNode
    
SaveCostsFeesDisbursementsExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub
'GD ADDED BM0198 END

'AM BMIDS722 - START
'********************************************************************************
'** Function:       GetBatchInterval
'** Created by:     Andy Maggs
'** Date:           05/03/2004
'** Description:    Gets the interval at which to space payment processing
'**                 transactions in a batch.
'** Parameters:     v_intNumTransactions - the number of payment processing
'**                 transactions in the batch.
'** Returns:        The interval (in seconds) at which to space payment
'**                 processing transactions.
'** Errors:         None Expected
'********************************************************************************
Private Function GetBatchInterval(ByVal v_lngNumTransactions As Long) As Long
   
    Dim lngBatchInterval As Long
    Dim lngMaxDuration As Long
    
    'get the value of PProcBatchInterval (interval is enetered by the user in
    'minutes, but we are converting to seconds to get a finer resolution)
    lngBatchInterval = GetGlobalParamAmount("PProcBatchInterval") * 60
    If lngBatchInterval > 0 Then
        'get the value of PProcBatchMaxDuration (the maximum duration in minutes)
        lngMaxDuration = GetGlobalParamAmount("PProcBatchMaxDuration") * 60
        
        'ensure that the interval applied to the transactions will not exceed
        'the maximum duration, if it does, reduce the interval to a value that
        'will not exceed it
        If (v_lngNumTransactions * lngBatchInterval) > lngMaxDuration Then
            'NB if the number of transactions is greater than the maximum
            'duration then the interval is reduced to zero and the code operates
            'as it did before this change was implemented.
            lngBatchInterval = lngMaxDuration \ v_lngNumTransactions
        End If
    End If
    
    GetBatchInterval = lngBatchInterval
    
End Function

'********************************************************************************
'** Function:       GetExecuteAfterDate
'** Created by:     Andy Maggs
'** Date:           05/03/2004
'** Description:    Calculates the date from which an item on the queue should be
'**                 processed given the start date, the transaction interval and
'**                 the current item being processed.
'** Parameters:     v_dtStartDate - the date/time the batch was started.
'**                 v_lngInterval - the number of seconds to allow between each
'**                 transaction.
'**                 v_lngCurrentItem - the current transaction being processed in
'**                 the batch.
'** Returns:        The date/time from which to start processing the transaction
'**                 from the queue formatted yyyy-mm-dd hh:nn:ss.
'** Errors:         None Expected
'********************************************************************************
Private Function GetExecuteAfterDate(ByVal v_dtStartDate As Date, _
        ByVal v_lngInterval As Long, ByVal v_lngCurrentItem As Long) As String
    Dim lngAddTime As Long
    Dim dtExecuteDate As Date
    '*-calculate the number of seconds to add to the start date
    lngAddTime = v_lngInterval * v_lngCurrentItem
    dtExecuteDate = DateAdd("s", lngAddTime, v_dtStartDate)
    GetExecuteAfterDate = Format$(dtExecuteDate, "yyyy-mm-dd hh:nn:ss")
End Function
'AM BMIDS722 - END


Private Sub ValidateCompletionDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo ValidateCompletionDateExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "ValidateCompletionDate"

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "VALIDATECOMPLETIONDATE")
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    gobjDO.ValidateCompletionDate vxmlRequestNode, vxmlResponseNode

ValidateCompletionDateExit:

    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub ValidateSolicitor(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo ValidateSolicitorExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "ValidateSolicitor"

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "SOLICITOR")
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    gobjDO.ValidateSolicitor vxmlRequestNode, vxmlResponseNode

ValidateSolicitorExit:
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetAdvanceDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo SetAdvancedDateExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "SetAdvanceDate"

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "SETADVANCEDATE")
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTMETHOD")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    gobjDO.SetAdvanceDate vxmlRequestNode, vxmlResponseNode
    
SetAdvancedDateExit:
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub SetPaymentMethod(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo SetPaymentMethodExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "SetPaymentMethod"

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "SETPAYMENTMETHOD")
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    gobjDO.SetPaymentMethod vxmlRequestNode, vxmlResponseNode

SetPaymentMethodExit:
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub ResetPostCompletionTasks(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo ResetPostCompletionTasksExit

Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "ResetPostCompletionTasks"

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "FindCaseTaskList")
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "FindCaseTaskList/CASETASK")
    
    'validate presents of attribute depending on UPDATEPROPERTY attribute value
    strValue = xmlGetMandatoryAttributeText(vxmlRequestNode, "UPDATEPROPERTY")
    Select Case UCase(strValue)
    Case "DUEDATE"
        strValue = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")
    Case "TASKSTATUS"
        strValue = xmlGetMandatoryAttributeText(xmlNode, "TASKSTATUSID")
    End Select
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "SOURCEAPPLICATION")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "CASEID")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "ACTIVITYID")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "ACTIVITYINSTANCE")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "STAGEID")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "TASKSTATUS")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "TASKTYPE")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "STAGEID")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    
    gobjDO.ResetPostCompletionTasks vxmlRequestNode, vxmlResponseNode

ResetPostCompletionTasksExit:
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub GetPostCompletionCaseTasksList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo GetPostCompletionCaseTasksListExit
Const strFunctionName As String = "GetPostCompletionCaseTasksList"

Dim xmlNode As IXMLDOMNode
Dim strValue As String

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "STAGEID")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "COMPLETIONDATE")
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    gobjTxBO.GetPostCompletionCaseTasksList vxmlRequestNode, vxmlResponseNode
  
    
GetPostCompletionCaseTasksListExit:
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub CancelDisbursements(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'MAR29
On Error GoTo CancelDisbursementsExit

Dim xmlNode As IXMLDOMNode
Dim xmlDisbursementNode As IXMLDOMNode
Dim xmlTempResponseNode As IXMLDOMNode
Dim xmlNewRequestNode As IXMLDOMNode
Dim xmlDisbursementPaymentList As IXMLDOMNodeList
Dim strValue As String
Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
Dim colValueId As Collection
Dim lngRet As Long

Const strFunctionName As String = "CancelDisbursements"

    ' Raise error if no node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "HEADER")
    
    'Validate mandatory attributes are specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    End If
    Set xmlTempResponseNode = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
    On Error Resume Next
    gobjDO.FindPaymentsToCancel vxmlRequestNode, xmlTempResponseNode
    lngRet = Err.Number
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    On Error GoTo CancelDisbursementsExit
    If Not (lngRet = 0 Or lngRet = oeRecordNotFound) Then
        errThrowError strFunctionName, lngRet
    End If
        
    'Update the status of the Disbursement payments to Cancelled
    Set xmlDisbursementPaymentList = xmlTempResponseNode.selectNodes(".//PAYMENTRECORD")
    If xmlDisbursementPaymentList.length > 0 Then
        Set xmlNewRequestNode = vxmlRequestNode.cloneNode(False)
        Set colValueId = New Collection
        GetValueIdsForValidationType "PaymentStatus", "C", colValueId
         
        If colValueId.Count > 0 Then
            strValue = CStr(colValueId.Item(1))
        Else
            errThrowError strFunctionName, oePPComboValueIdNotFound
        End If
        For Each xmlNode In xmlDisbursementPaymentList
            Set xmlDisbursementNode = xmlNode.selectSingleNode("./DISBURSEMENTPAYMENT")
            xmlSetAttributeValue xmlDisbursementNode, "PAYMENTSTATUS", strValue
            xmlNewRequestNode.appendChild xmlNode
        Next
        If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
        gobjTxBO.UpdateDisbursement xmlNewRequestNode, vxmlResponseNode
    End If
CancelDisbursementsExit:
    Set xmlNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Sub
