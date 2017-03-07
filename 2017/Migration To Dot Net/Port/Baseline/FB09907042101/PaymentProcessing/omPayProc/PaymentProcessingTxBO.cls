VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PaymentProcessingTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProcessingTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   PaymentProcessing Business Object.
'
'Dependencies:
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     23/05/01    SYS2298 - Modified methods 'CreateApplicationFeeType', 'CreateFeeTypePayment' and
'                   'UpdateFeeTypePayment'
'CL     05/06/01    SYS2340 - Modified methods 'AllocateDisbursementsToLoanComponents'. Added methods
'                           'CheckFirstPaymentDate', 'UpdatePaymentDate' and
'                           'IPaymentProcessingTxBO_UpdateFirstPaymentDate'
'SR     06/07/01    SYS2412 - Modified private methods 'AllocateDisbursementsToLoanComponents' and
'                   'GetLCAllocatedAmount'
'SR      01/08/01   SYS2552 - Modified private methods 'AllocateDisbursementsToLoanComponents
'SR     06/09/01    SYS2412 - New method 'CreateReturnOfFunds' and 'DeletePayeeHistory'
'                   Modified method 'UpdateFirstPaymentDate'
'PSC    26/11/01    SYS3164 Amend CreateFeeTypePayment to default the CompletionIndicator to 1
'SR     12/12/01    SYS3391 Modified method 'UpdateDisbursement' - delete all feePayments if the amount=0
'SR     19/12/2001  SYS2547 Modified method AllocateDisbursementToLoanComponents, added new methods
'                   UpdateFirstPaymentDateAndBankBuildingSoc and UpdateFirstPaymentDateAndCreateBankBuildingSoc
'                   modifed UpdateFirstPaymentDate - added new optional parameter
'DM     03/01/02    SYS3185 Implement CompleteInterfacing
'RF     03/01/02    Complete SYS3185 (Implement CompleteInterfacing)
'JLD    26/02/02    SYS4168 don't create all fees as application costs. Use validation type to decide.
'DM     14/03/02    SYS4274 CompleteInterfacing changes.
'DM     27/03/02    SYS4313 CompleteInterfacing changes
'SG     28/05/02    SYS4767 MSMS to Core integration.
'STB    10/06/02    SYS4835 More meaningful error recorded for CBA business errors.
'--------------------------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'PSC    14/11/2005  MAR520 Amend CompleteInterfacing to get error description from correct place
'PSC    15/11/2005  MAR140 Add GetPostCompletionCaseTasksList
'PSC    29/11/2005  MAR733 Amend CompleteInterfacing to get error message correctly
'HMA    05/04/2006  MAR1408  Correction to GetPostCompletionCaseTasksList
'GHun   20/04/2006  MAR1651 Changed CreateReturnOfFunds to comment out stage move
'DRC    03/05/2006  MAR1696(EP441) Changed CreateApplicationFeeType to check if a refund of the valuation fee is needed
'DRC    08/05/2006  MAR1696(EP441) Applied above change to  CreateApplicationFeeTypes also
'--------------------------------------------------------------------------------------------------------------
'Epsom History:
'
'Prog   Date        Description
'PSC    13/03/2007  EP2_1347 Treat valuation refunds the same as incentives
'--------------------------------------------------------------------------------------------------------------

Option Explicit

Implements IPaymentProcessingTxBO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub IPaymentProcessingTxBO_CreateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationFeeTypesVbErr

    Dim xmlFeeTypeNode As IXMLDOMNode
    Dim xmlPaymentRecord As IXMLDOMElement   ' MAR28
    Dim xmlNewNode As IXMLDOMNode            ' MAR28
    Dim strSeqNo As String                   ' MAR28
    Dim objPayProcDO As IPAymentProcessingDO
    Dim iFeeType As Integer
    Dim strAppNo As String
    Dim strAmount As String
    Dim strRefundAmount As String
    Dim xmlDomDoc As FreeThreadedDOMDocument40
    Dim blnValRefund As Boolean              ' MAR1696

    Set xmlDomDoc = New FreeThreadedDOMDocument40
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlFeeTypeNode = vxmlRequestNode.selectSingleNode("APPLICATIONFEETYPE")
    
    If Not xmlFeeTypeNode Is Nothing Then
        objPayProcDO.CreateApplicationFeeType xmlFeeTypeNode, vxmlResponseNode
        
        'MAR28
        'If this is a Valuation Fee, set up PAYMENTRECORD and FEEPAYMENT records for the Refund
        iFeeType = CSafeInt(xmlFeeTypeNode.Attributes.getNamedItem("FEETYPE").Text)

        
        If IsValidationType("OneOffCost", iFeeType, "VAL") = True Then
        'MAR1696 check global parameter to see if we need a refund
            blnValRefund = GetGlobalParamBoolean("CalculateValuationRefund")
            If blnValRefund Then
           
                'MAR28 Get Application Number and Amount
                strAppNo = xmlFeeTypeNode.Attributes.getNamedItem("APPLICATIONNUMBER").Text
                strAmount = xmlFeeTypeNode.Attributes.getNamedItem("AMOUNT").Text
                strRefundAmount = xmlFeeTypeNode.Attributes.getNamedItem("REFUNDAMOUNT").Text
    
                Set xmlPaymentRecord = xmlDomDoc.createElement("PAYMENTRECORD")
                xmlPaymentRecord.setAttribute "CREATIONDATETIME", Format$(Now, "dd/mm/yy hh:mm:ss")
                xmlPaymentRecord.setAttribute "APPLICATIONNUMBER", strAppNo
                xmlPaymentRecord.setAttribute "AMOUNT", strRefundAmount
    
                objPayProcDO.CreatePaymentRecord xmlPaymentRecord, vxmlResponseNode
    
                'Extract the PAYMENTSEQUENCENUMBER generated
                strSeqNo = xmlGetAttributeText(xmlPaymentRecord, "PAYMENTSEQUENCENUMBER")
    
                'Create a FEEPAYMENT record
                Set xmlPaymentRecord = xmlDomDoc.createElement("FEEPAYMENT")
                
                xmlPaymentRecord.setAttribute "APPLICATIONNUMBER", strAppNo
                xmlPaymentRecord.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
                xmlPaymentRecord.setAttribute "FEETYPE", xmlFeeTypeNode.Attributes.getNamedItem("FEETYPE").Text
                xmlPaymentRecord.setAttribute "FEETYPESEQUENCENUMBER", xmlFeeTypeNode.Attributes.getNamedItem("FEETYPESEQUENCENUMBER").Text
                xmlPaymentRecord.setAttribute "REFUNDAMOUNT", strRefundAmount
                xmlPaymentRecord.setAttribute "AMOUNTPAID", "0"
                xmlPaymentRecord.setAttribute "PAYMENTEVENT", GetFirstComboValueId("PaymentEvent", "RFV")
    
                objPayProcDO.CreateFeeTypePayment xmlPaymentRecord, vxmlResponseNode
            End If
        'MAR1696 - end
        End If
      
    End If
    
    gobjContext.SetComplete
    
CreateApplicationFeeTypesExit:
    Set objPayProcDO = Nothing
    Set xmlFeeTypeNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreateApplicationFeeTypesVbErr:
    gobjContext.SetAbort
    GoTo CreateApplicationFeeTypesExit

End Sub

Private Sub IPaymentProcessingTxBO_CreateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateDisbursementVbErr

Dim xmlPaymentNode As IXMLDOMElement
Dim objPayProcDO As IPAymentProcessingDO
Dim xmlNodeList As IXMLDOMNodeList

Dim strSeqNo As String
Dim strAppNo As String

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlPaymentNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    If Not xmlPaymentNode Is Nothing Then
        objPayProcDO.CreatePaymentRecord xmlPaymentNode, vxmlResponseNode
        
        'Extract the PAYMENTSEQUENCENUMBER generated
        strSeqNo = xmlGetAttributeText(xmlPaymentNode, "PAYMENTSEQUENCENUMBER")
        'Get Application Number from PAYMENTRECORD element
        strAppNo = xmlGetAttributeText(xmlPaymentNode, "APPLICATIONNUMBER")
        
        'If a node exists, create a FEEPAYMENT record
        Set xmlNodeList = vxmlRequestNode.selectNodes("PAYMENTRECORD/FEEPAYMENT")
        For Each xmlPaymentNode In xmlNodeList
            xmlPaymentNode.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
            xmlPaymentNode.setAttribute "APPLICATIONNUMBER", strAppNo
            objPayProcDO.CreateFeeTypePayment xmlPaymentNode, vxmlResponseNode
        Next
        
        'Create a DISBURSEMENTPAYMENT record
        Set xmlPaymentNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD/DISBURSEMENTPAYMENT")
        xmlPaymentNode.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
        xmlPaymentNode.setAttribute "APPLICATIONNUMBER", strAppNo
        If Not xmlPaymentNode Is Nothing Then
            objPayProcDO.CreateDisbursement xmlPaymentNode, vxmlResponseNode
        End If
        
        'Allocate the Payment to Loan Components
        AllocateDisbursementsToLoanComponents vxmlRequestNode
    End If
    
    gobjContext.SetComplete
    
CreateDisbursementExit:
    Set objPayProcDO = Nothing
    Set xmlPaymentNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreateDisbursementVbErr:
    gobjContext.SetAbort
    GoTo CreateDisbursementExit

End Sub

Private Sub IPaymentProcessingTxBO_CreateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationFeeTypesVbErr

Dim xmlPaymentNode As IXMLDOMElement
Dim objPayProcDO As IPAymentProcessingDO
Dim strSeqNo As String, strRefundDate As String

    Const csrtFunctionName = "CreateFeeTypePayment"
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'SR 22/05/01 : SYS2298 - If RefundDate is mentioned in FEEPAYMENT, it should not be a past date
    Set xmlPaymentNode = vxmlRequestNode.selectSingleNode("FEEPAYMENT")
    strRefundDate = xmlGetAttributeText(xmlPaymentNode, "REFUNDDATE")
    If strRefundDate <> "" Then
        If Not IsDate(strRefundDate) Then
            errThrowError csrtFunctionName, 533
        End If
        
        If (CDate(strRefundDate) < Date) Then
            errThrowError csrtFunctionName, 565
        End If
    End If
    'SR 22/05/01 - End
    
    Dim colValueIds As Collection
    Set colValueIds = New Collection
    
    'Create a PAYMENTRECORD record
    Set xmlPaymentNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    If Not xmlPaymentNode Is Nothing Then
        xmlPaymentNode.setAttribute "CREATIONDATETIME", Format$(Now, "dd/mm/yy hh:mm:ss")
        objPayProcDO.CreatePaymentRecord xmlPaymentNode, vxmlResponseNode
        
        'Extract the PAYMENTSEQUENCENUMBER generated
        strSeqNo = xmlGetAttributeText(xmlPaymentNode, "PAYMENTSEQUENCENUMBER")
        
        'Create a FEEPAYMENT record
        Set xmlPaymentNode = vxmlRequestNode.selectSingleNode("FEEPAYMENT")
        If Not xmlPaymentNode Is Nothing Then
            ' PSC 26/11/01 SYS3164 - Start
            xmlPaymentNode.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
            GetValueIdsForValidationType "CompletionIndicator", "N", colValueIds
            xmlPaymentNode.setAttribute "COMPLETIONINDICATOR", CStr(colValueIds(1))
            objPayProcDO.CreateFeeTypePayment xmlPaymentNode, vxmlResponseNode
            ' PSC 26/11/01 SYS3164 - End
        End If
    End If
    
    gobjContext.SetComplete
    
CreateApplicationFeeTypesExit:
    Set objPayProcDO = Nothing
    Set xmlPaymentNode = Nothing
    Set colValueIds = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreateApplicationFeeTypesVbErr:
    gobjContext.SetAbort
    GoTo CreateApplicationFeeTypesExit

End Sub

Private Sub IPaymentProcessingTxBO_CreatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreatePayeeHistoryDetailsVbErr

Dim xmlPayeeHistoryNode As IXMLDOMElement
Dim objPayProcDO As IPAymentProcessingDO
Dim objThirdPartyBO As ThirdPartyBO
    
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlDomDoc As FreeThreadedDOMDocument40
Dim xmlKeysNode As IXMLDOMNode

Dim strRet As String

    Set xmlDomDoc = New FreeThreadedDOMDocument40
    xmlDomDoc.validateOnParse = False
    xmlDomDoc.setProperty "NewParser", True

    'Convert to old style element based request
    Set xmlDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "PAYEEHISTORY/THIRDPARTY", True)
    If Not xmlDoc Is Nothing Then
        
        'Create a Third Party record (which will create an Address record also)
        Set objThirdPartyBO = gobjContext.CreateInstance("omTP.ThirdPartyBO")
        strRet = objThirdPartyBO.CreateThirdParty(xmlDoc.documentElement.xml)
        errCheckXMLResponse strRet, True
        
        'Get the ThirdParty Guid of the newly created record
        xmlDomDoc.loadXML strRet
        strRet = xmlGetMandatoryNodeText(xmlDomDoc.documentElement, "GENERATEDKEYS/THIRDPARTYGUID")
        
        'Create a Payee History record
        Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
        Set xmlPayeeHistoryNode = vxmlRequestNode.selectSingleNode("PAYEEHISTORY")
        If Not xmlPayeeHistoryNode Is Nothing Then
            xmlPayeeHistoryNode.setAttribute "THIRDPARTYGUID", strRet
            objPayProcDO.CreatePayeeHistoryDetails xmlPayeeHistoryNode, vxmlResponseNode
        End If
        
    End If
    
    gobjContext.SetComplete
    
CreatePayeeHistoryDetailsExit:
    Set objPayProcDO = Nothing
    Set xmlPayeeHistoryNode = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlDoc = Nothing
    Set xmlDomDoc = Nothing
    Set xmlKeysNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreatePayeeHistoryDetailsVbErr:
    gobjContext.SetAbort
    GoTo CreatePayeeHistoryDetailsExit

End Sub

Private Sub IPaymentProcessingTxBO_CreatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreatePaymentAuditVbErr

    Dim xmlPayAuditNode As IXMLDOMNode
    Dim objPayProcDO As IPAymentProcessingDO
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlPayAuditNode = vxmlRequestNode.selectSingleNode("PAYPROCBATCHAUDIT")
    If Not xmlPayAuditNode Is Nothing Then
        objPayProcDO.CreatePaymentAudit xmlPayAuditNode, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
CreatePaymentAuditExit:
    Set objPayProcDO = Nothing
    Set xmlPayAuditNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreatePaymentAuditVbErr:
    gobjContext.SetAbort
    GoTo CreatePaymentAuditExit

End Sub

Private Sub IPaymentProcessingTxBO_CreateReturnOfFunds( _
        ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateReturnOfFundsVbErr

    Const cstrFunctionName As String = "IPaymentProcessingTxBO_CreateReturnOfFunds"

    Dim xmlOrigPayment As IXMLDOMNode, xmlROFPayment As IXMLDOMNode, xmlNewRequest As IXMLDOMNode
    Dim xmlROFs As IXMLDOMNode, xmlROFList As IXMLDOMNodeList, xmlNode As IXMLDOMNode
    Dim xmlFeePayments As IXMLDOMNode, xmlFeePaymentsList As IXMLDOMNodeList, xmlCaseStage As IXMLDOMNode
    Dim xmlPaymentRecors As IXMLDOMNode, xmlPaymentRecordList As IXMLDOMNodeList
    Dim xmlDoc As FreeThreadedDOMDocument40, xmlPaymentRecord As IXMLDOMNode
    
    Dim strAssocPaySeqNo As String, strApplicationNumber As String
    Dim lngPreviousROFAmount As Long, lngFeePaymentAmount As Long, lngROFCurrentAmount As Long, lngErrNo As Long
    Dim lngOrigPaymentAmount As Long, strPaymentSeqNo As String ', strApplCompleteStageId As String
    'Dim strResponse As String
    
    'Dim objComboDO As ComboDO
    Dim objPaymentProcessingDO As IPAymentProcessingDO
    'BMIDS697 GHun 23/01/2004
    'Dim objTmBO As OmTmBO
    'Dim objTmNoTxBO As omTmNoTxBO
    'BMIDS697 End
    
    ' PSC 28/11/2002 BMIDS01099 - Start
    Dim xmlDisb As IXMLDOMNode
    Dim blnIsIncentive As Boolean
    Dim strPaymentType As String
    ' PSC 28/11/2002 BMIDS01099 - End
    
    ' PSC 13/03/2007 EP2_1347
    Dim blnIsValuationRefund As Boolean

    '-------------------------------------------------------------------------------------------------------
    ' Validate for existing ROF payments against selected PaymentRecord and find the amount already returned
    '--------------------------------------------------------------------------------------------------------
    Set xmlOrigPayment = xmlGetMandatoryNode(vxmlRequestNode, ".//ORIGPAYMENTRECORD")
    Set xmlPaymentRecord = xmlGetMandatoryNode(vxmlRequestNode, ".//ORIGPAYMENTRECORD/PAYMENTRECORD")
    Set xmlROFPayment = xmlGetMandatoryNode(vxmlRequestNode, ".//ROFPAYMENTRECORD")
    
    Dim xmlROFDisbursementPayment As IXMLDOMNode
    Set xmlROFDisbursementPayment = xmlGetMandatoryNode(vxmlRequestNode, ".//ROFPAYMENTRECORD/ROFDISBURSEMENTPAYMENT")
    
    strAssocPaySeqNo = xmlGetMandatoryAttributeText(xmlOrigPayment, "PAYMENTSEQUENCENUMBER")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlOrigPayment, "APPLICATIONNUMBER")
    lngOrigPaymentAmount = xmlGetMandatoryAttributeAsLong(xmlPaymentRecord, "AMOUNT")
    
    'Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    Set objPaymentProcessingDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'strApplCompleteStageId = objComboDO.GetFirstComboValueId("ApplicationStage", "MC")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlROFs = xmlDoc.createElement("ROFRECORDS")
    Set xmlNewRequest = xmlDoc.createElement("PAYMENTRECORD")
    xmlSetAttributeValue xmlNewRequest, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlNewRequest, "ASSOCPAYSEQNUMBER", strAssocPaySeqNo
    
    On Error Resume Next
    
    objPaymentProcessingDO.FindPaymentRecordList xmlNewRequest, xmlROFs
    
    lngErrNo = Err.Number
    On Error GoTo CreateReturnOfFundsVbErr
    If lngErrNo <> 0 Then
        lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        If lngErrNo <> oeRecordNotFound Then
            errThrowError cstrFunctionName, lngErrNo
        End If
    End If
    
    ' Find the amount Returned already, against this PaymentRecord
    lngPreviousROFAmount = 0
    Set xmlROFList = xmlROFs.selectNodes(".//PAYMENTRECORD")
    For Each xmlNode In xmlROFList
        If Len(xmlGetAttributeText(xmlNode, "ASSOCPAYSEQNUMBER")) > 0 Then
            lngPreviousROFAmount = lngPreviousROFAmount + xmlGetMandatoryAttributeAsLong(xmlNode, "AMOUNT")
        End If
    Next xmlNode
    
    '-----------------------------------------------------------------
    'Find the FeePayments made against this payment record
    '(Record with PaySeqNumber as the AssocPaySeqNumber passed in)
    '-----------------------------------------------------------------
    Set xmlNewRequest = xmlDoc.createElement("FEEPAYMENT")
    xmlSetAttributeValue xmlNewRequest, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlNewRequest, "PAYMENTSEQUENCENUMBER", strAssocPaySeqNo
    
    Set xmlFeePayments = xmlDoc.createElement("FEEPAYMENTS")
    
    On Error Resume Next
    
    objPaymentProcessingDO.FindFeePaymentList xmlNewRequest, xmlFeePayments
    
    lngErrNo = Err.Number
    On Error GoTo CreateReturnOfFundsVbErr
    If lngErrNo <> 0 Then
        lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        If lngErrNo <> oeRecordNotFound Then
            errThrowError cstrFunctionName, lngErrNo
        End If
    End If
    
    ' PSC 22/08/2003 BM0198 - Start
    Dim intPaymentEvent As Integer
    
    lngFeePaymentAmount = 0
    
    Set xmlFeePaymentsList = xmlFeePayments.selectNodes(".//FEEPAYMENT")
    For Each xmlNode In xmlFeePaymentsList
        intPaymentEvent = xmlGetAttributeAsInteger(xmlNode, "PAYMENTEVENT")
        
        If IsValidationType("PaymentEvent", intPaymentEvent, "D") = True Then
            lngFeePaymentAmount = lngFeePaymentAmount + xmlGetMandatoryAttributeAsLong(xmlNode, "AMOUNTPAID")
        End If

    Next xmlNode
    ' PSC 22/08/2003 BM0198 - End
    
    '------------------------------------------------------------------
    ' Validate payment amount and create Payment Record (if necessary)
    '------------------------------------------------------------------
    
    lngROFCurrentAmount = xmlGetMandatoryAttributeAsLong(xmlROFDisbursementPayment, "NETPAYMENTAMOUNT")
    
    If lngROFCurrentAmount > lngOrigPaymentAmount - lngFeePaymentAmount Then
        errThrowError cstrFunctionName, oeInvalidParameter, "Amount of ROF cannot exceed original payment amount"
    End If
    
    'Create a row in Payment record table
    CreateROF_CreatePaymentRecord vxmlRequestNode, vxmlResponseNode, strPaymentSeqNo
    
    If lngROFCurrentAmount = lngOrigPaymentAmount - lngFeePaymentAmount Then
    
        'Create rows in table FeePayment
        CreateROF_CreateFeePayments vxmlRequestNode, vxmlResponseNode, strPaymentSeqNo
        
        ' PSC 22/08/2003 BM0198 - Start
        ' Delete the TT Application fee type if we have returned the fee
        Dim xmlTTFees As IXMLDOMNodeList
        Dim xmlTTFee As IXMLDOMNode
        Dim xmlAppFeeType As IXMLDOMNode
        Dim strTTFeeValue As String
        Dim strFeeTypeSeqNo As String
        strTTFeeValue = GetFirstComboValueId("OneOffCost", "TTF")
        
        Set xmlTTFees = xmlFeePayments.selectNodes(".//FEEPAYMENT[@FEETYPE=""" & strTTFeeValue & """]")
        
        ' Delete the relevant fee payment and application fee types
        For Each xmlTTFee In xmlTTFees
        
            ' Delete fee payment
            objPaymentProcessingDO.DeleteFeePayments xmlTTFee, vxmlResponseNode
            
            ' Delete application fee type
            strFeeTypeSeqNo = xmlGetAttributeText(xmlTTFee, "FEETYPESEQUENCENUMBER")
            Set xmlAppFeeType = xmlDoc.createElement("APPLICATIONFEETYPE")
            xmlSetAttributeValue xmlAppFeeType, "APPLICATIONNUMBER", strApplicationNumber
            xmlSetAttributeValue xmlAppFeeType, "FEETYPE", strTTFeeValue
            xmlSetAttributeValue xmlAppFeeType, "FEETYPESEQUENCENUMBER", strFeeTypeSeqNo
            objPaymentProcessingDO.DeleteApplicationFeeTypes xmlAppFeeType, vxmlResponseNode
        Next
        ' PSC 22/08/2003 BM0198 - End
    
    End If
    
    'Validate Loan Component Payment records and create entries in table LoanComponentPayment (if required)
    
    ' PSC 28/11/2002 BMIDS01099 - Start
    Set xmlDisb = xmlGetMandatoryNode(xmlPaymentRecord, "DISBURSEMENTPAYMENT")
    
    strPaymentType = xmlGetMandatoryAttributeText(xmlDisb, "PAYMENTTYPE")
    
    ' PSC 13/03/2007 EP2_1347 - Start
    blnIsIncentive = IsValidationType("PaymentType", CSafeInt(strPaymentType), "C")
    blnIsValuationRefund = IsValidationType("PaymentType", CSafeInt(strPaymentType), "VALREFUND")
    ' PSC 13/03/2007 EP2_1347 - End
    
    ' PSC 13/03/2007 EP2_1347
    If blnIsIncentive = False And blnIsValuationRefund = False Then
        CreateROF_CreateLoanComponentPayments vxmlRequestNode, vxmlResponseNode, xmlROFs, strPaymentSeqNo
    End If
    ' PSC 28/11/2002 BMIDS01099 - End
    
    'validate and move the application back to pre-completion stage
    Set xmlNewRequest = xmlDoc.createElement("PAYMENTRECORD")
    xmlSetAttributeValue xmlNewRequest, "APPLICATIONNUMBER", strApplicationNumber
    
    Set xmlPaymentRecors = xmlDoc.createElement("PAYMENTRECORDS")
    objPaymentProcessingDO.FindPaymentRecordList xmlNewRequest, xmlPaymentRecors
    
    Set xmlPaymentRecordList = xmlPaymentRecors.selectNodes(".//PAYMENTRECORD")
    Dim lngTotalDisbursement As Long
    lngTotalDisbursement = 0
    For Each xmlNode In xmlPaymentRecordList
        lngTotalDisbursement = lngTotalDisbursement + xmlGetMandatoryAttributeAsLong(xmlNode, "AMOUNT")
    Next xmlNode
    
    'MAR1651 GHun No longer required
    'If lngTotalDisbursement = 0 Then
    '    ' Move application to pre-completion stage
    '
    '    'BMIDS697 GHun 23/01/2004
    '    'Set objTmBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".OmTmBO")
    '    Set objTmNoTxBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".OmTmNoTxBO")
    '    'BMIDS697 End
    '
    '    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    '    xmlSetAttributeValue xmlNewRequest, "OPERATION", "MoveToStage"
    '
    '    Set xmlNode = xmlDoc.createElement("APPLICATION")
    '    xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strApplicationNumber
    '    xmlSetAttributeValue xmlNode, "APPLICATIONFACTFINDNUMBER", GetFactFindNumberForApplication(strApplicationNumber)
    '    xmlNewRequest.appendChild xmlNode.cloneNode(True)
    '
    '    'If total Disbursement is zero, move the application back to pre-completion stage
    '    Set xmlCaseStage = vxmlRequestNode.selectSingleNode(".//CASESTAGE")
    '    xmlSetAttributeValue xmlCaseStage, "STAGEID", strApplCompleteStageId
    '    xmlNewRequest.appendChild xmlCaseStage
   
    '    'BMIDS697 GHun 23/01/2004
    '    'strResponse = objTmBO.OmTmRequest(xmlNewRequest.xml)
    '    strResponse = objTmNoTxBO.OmTmNoTxRequest(xmlNewRequest.xml)
    '    'BMIDS697 End
    '    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    '    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode
    '
    'End If
    'MAR1651 End
    
    gobjContext.SetComplete

CreateReturnOfFundsExit:
    
    Set xmlOrigPayment = Nothing
    Set xmlROFPayment = Nothing
    Set xmlNewRequest = Nothing
    Set xmlROFs = Nothing
    Set xmlROFList = Nothing
    Set xmlFeePayments = Nothing
    Set xmlFeePaymentsList = Nothing
    Set xmlPaymentRecordList = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlCaseStage = Nothing
    Set xmlPaymentRecord = Nothing
    
    'Set objComboDO = Nothing
    Set objPaymentProcessingDO = Nothing
    'Set objTmNoTxBO = Nothing
    
    ' PSC 28/11/2002 BMIDS01099 - Start
    Set xmlDisb = Nothing
    ' PSC 28/11/2002 BMIDS01099 - End
    
    ' PSC 22/08/2003 BM0198 - Start
    Set xmlTTFees = Nothing
    Set xmlTTFee = Nothing
    Set xmlAppFeeType = Nothing
    ' PSC 22/08/2003 BM0198 - End
        
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreateReturnOfFundsVbErr:
    gobjContext.SetAbort
    GoTo CreateReturnOfFundsExit

End Sub

Private Sub CreateROF_CreateLoanComponentPayments(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode, _
                                            ByVal xmlROFs As IXMLDOMNode, _
                                            ByVal strPaymentSeqNo As String)

On Error GoTo CreateROF_CreateLoanComponentPaymentsVbErr
    
    Const cstrFunctionName As String = "CreateROF_CreateLoanComponentPayments"
    
    'LCP recods for the original payment (passed in the request)
    Dim xmlLCPList As IXMLDOMNodeList, xmlLCPSorted As IXMLDOMNode, xmlLCP As IXMLDOMNode
    
    Dim xmlFullLCPs As IXMLDOMNode, xmlFullLCPList As IXMLDOMNodeList 'LCP records for this application
    Dim xmlPaymentRecord As IXMLDOMNode
    Dim xmlROFPaymentRecord As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode, xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim strApplicationNumber As String, lngROFAmount As Long, strLCPSeqNo As String
    Dim lngOrigLCPAmount As Long  ' Actual LoanComponentPaymentAmount
    Dim lngLCPReturned As Long  ' Amount returned so far against a LoanComponentPayment
    ' PSC 26/11/2002 BMIDS01052 - Start
    Dim lngOrigIntOnlyAmt As Long          ' Amount of interest only on original payment
    Dim lngIntOnlyReturned As Long         ' Total of interest only part returned so far
    Dim lngOrigCapAndIntAmt As Long        ' Amount of capital and interest on original payment
    Dim lngCapAndIntReturned As Long       ' Total of capital and interest part returned so far
    Dim lngIntOnlyToReturn As Long         ' Amount of interest only available to return
    Dim lngCapAndIntToReturn As Long       ' Amount of capital and interest available to return
    Dim lngIntOnlyAmt As Long              ' Amount of interest only to be returned by this ROF
    Dim lngCapAndIntAmt As Long            ' Amount of capital and interest to be returned by this ROF
    ' PSC 26/11/2002 BMIDS01052 - End
    Dim objDO As IPAymentProcessingDO
    
    ' PSC 20/11/2002 - BMIDS01013 - Start
    Dim xmlROFTotals As IXMLDOMNode
    Dim xmlROFTotal As IXMLDOMNode
    Dim strAssocPaymentSeqNo As String
    ' PSC 20/11/2002 - BMIDS01013 - End
          
    ' Get the Sorted List of LoanComponentPayments
    Set xmlLCPSorted = xmlDoc.createElement("LCPSORTED")
    Set xmlLCPList = vxmlRequestNode.selectNodes(".//PAYMENTRECORD/LOANCOMPONENTPAYMENT")
    For Each xmlLCP In xmlLCPList
        xmlLCPSorted.appendChild xmlLCP.cloneNode(True)
    Next xmlLCP
    
    Set xmlLCPSorted = CreateROF_SortLCPs(xmlLCPSorted)  ' Sort the Loan Component Payments in descending order
    
    Set xmlPaymentRecord = vxmlRequestNode.selectSingleNode(".//PAYMENTRECORD")
    strApplicationNumber = xmlGetAttributeText(xmlPaymentRecord, "APPLICATIONNUMBER")
    
    Set xmlROFPaymentRecord = vxmlRequestNode.selectSingleNode(".//ROFPAYMENTRECORD")
    lngROFAmount = xmlGetMandatoryAttributeAsLong(xmlROFPaymentRecord, "AMOUNT")
    
    ' PSC 20/11/2002 BMIDS01013 - Start

'    'Build request and fetch all the LCP records corresponding to this application
'    Set xmlTableNode = xmlDoc.createElement("LOANCOMPONENTPAYMENT")
'    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
'
'    Set xmlFullLCPs = xmlDoc.createElement("LOANCOMPONENTPAYMENTLIST")
    
    strAssocPaymentSeqNo = xmlGetMandatoryAttributeAsLong(xmlROFPaymentRecord, "ASSOCPAYSEQNUMBER")
    Set xmlROFTotals = xmlDoc.createElement("ROFTOTALS")
    Set xmlTableNode = xmlDoc.createElement("ROFTOTAL")
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    On Error Resume Next
    objDO.FindROFTotals xmlTableNode, xmlROFTotals
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
    
    On Error GoTo CreateROF_CreateLoanComponentPaymentsVbErr
    
    If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
        Err.Raise lngErrNo, strSource, strDescription
    End If
    ' PSC 20/11/2002 BMIDS01013 - End
    
    '-----------------------------------------------------------------------------------------------
    ' For each loan component payment, create one more entry with -ve amount (where ever required).
    ' See that sum of the amount in the created records is not more than current ROF Amount
    '-----------------------------------------------------------------------------------------------
    Set xmlTableNode = xmlDoc.createElement("LOANCOMPONENTPAYMENT")
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlTableNode, "APPLICATIONFACTFINDNUMBER", GetFactFindNumberForApplication(strApplicationNumber)
    xmlSetAttributeValue xmlTableNode, "PAYMENTSEQUENCENUMBER", strPaymentSeqNo
    
    Set xmlLCPList = xmlLCPSorted.selectNodes(".//LOANCOMPONENTPAYMENT")
    Dim lngTemp As Long
    For Each xmlLCP In xmlLCPList
        If lngROFAmount <= 0 Then Exit For
        
        lngOrigLCPAmount = xmlGetMandatoryAttributeAsLong(xmlLCP, "AMOUNT")
        
        ' PSC 26/11/2002 BMIDS01052 - Start
        lngOrigIntOnlyAmt = xmlGetAttributeAsLong(xmlLCP, "INTERESTONLYELEMENT", "0")
        lngOrigCapAndIntAmt = xmlGetAttributeAsLong(xmlLCP, "CAPITALANDINTERESTELEMENT", "0")
        ' PSC 26/11/2002 BMIDS01052 - End
        
        strLCPSeqNo = xmlGetMandatoryAttributeText(xmlLCP, "LOANCOMPONENTSEQUENCENUMBER")
        
        ' Find the Total Loan Component Payment amount that was returned
        
        ' PSC 20/11/2002 BMIDS01013 - Start
'        Set xmlFullLCPList = _
'                xmlFullLCPs.selectNodes(".//LOANCOMPONENTPAYMENT[@LOANCOMPONENTSEQUENCENUMBER=" & strLCPSeqNo & "]")
        lngLCPReturned = 0
        ' PSC 26/11/2002 BMIDS01052 - Start
        lngIntOnlyReturned = 0
        lngIntOnlyToReturn = 0
        lngIntOnlyAmt = 0
        lngCapAndIntReturned = 0
        lngCapAndIntToReturn = 0
        lngCapAndIntAmt = 0
        ' PSC 26/11/2002 BMIDS01052 - End
'        For Each xmlNode In xmlFullLCPList
'            lngTemp = xmlGetMandatoryAttributeAsLong(xmlNode, "AMOUNT")
'            If lngTemp < 0 Then lngLCPReturned = lngLCPReturned - lngTemp
'        Next xmlNode
         
         Set xmlROFTotal = xmlROFTotals.selectSingleNode(".//ROFTOTAL[@LOANCOMPONENTSEQUENCENUMBER=" & strLCPSeqNo & " and @ASSOCPAYSEQNUMBER = " & strAssocPaymentSeqNo & "]")
         
         If Not xmlROFTotal Is Nothing Then
            lngLCPReturned = Abs(xmlGetAttributeAsLong(xmlROFTotal, "TOTALRETURNED", "0"))
            ' PSC 26/11/2002 BMIDS01052 - Start
            lngIntOnlyReturned = Abs(xmlGetAttributeAsLong(xmlROFTotal, "INTERESTONLYRETURNED", "0"))
            lngCapAndIntReturned = Abs(xmlGetAttributeAsLong(xmlROFTotal, "CAPITALANDINTERESTRETURNED", "0"))
            ' PSC 26/11/2002 BMIDS01052 - Start
         End If

        ' PSC 20/11/2002 BMIDS01013 - End
        
        lngIntOnlyToReturn = lngOrigIntOnlyAmt - lngIntOnlyReturned
        lngCapAndIntToReturn = lngOrigCapAndIntAmt - lngCapAndIntReturned

        If lngOrigLCPAmount > lngLCPReturned Then
            ' find the amount for which the LCP record is to be created
            If lngOrigLCPAmount - lngLCPReturned >= lngROFAmount Then
                lngTemp = lngROFAmount
            Else
                lngTemp = lngOrigLCPAmount - lngLCPReturned
            End If
            
            ' PSC 26/11/2002 BMIDS01052 - Start
            If lngOrigIntOnlyAmt > 0 Or lngOrigCapAndIntAmt > 0 Then
                If lngIntOnlyToReturn >= lngTemp Then
                    lngIntOnlyAmt = lngTemp
                Else
                    lngIntOnlyAmt = lngIntOnlyToReturn
                    lngCapAndIntAmt = lngTemp - lngIntOnlyAmt
                End If
            End If
            ' PSC 26/11/2002 BMIDS01052 - End
   
            'Create LCP record with -ve amount
            xmlCopyAttribute xmlLCP, xmlTableNode, "LOANCOMPONENTSEQUENCENUMBER"
            xmlCopyAttribute xmlLCP, xmlTableNode, "MORTGAGESUBQUOTENUMBER"
            xmlSetAttributeValue xmlTableNode, "AMOUNT", (-1) * lngTemp
            
            ' PSC 26/11/2002 BMIDS01052 - Start
            If lngOrigIntOnlyAmt > 0 Or lngOrigCapAndIntAmt > 0 Then
                xmlSetAttributeValue xmlTableNode, "INTERESTONLYELEMENT", (-1) * lngIntOnlyAmt
                xmlSetAttributeValue xmlTableNode, "CAPITALANDINTERESTELEMENT", (-1) * lngCapAndIntAmt
            End If
            ' PSC 26/11/2002 BMIDS01052 - End
            
            objDO.CreateLoanComponentPaymentDetails xmlTableNode, vxmlResponseNode
            
            lngROFAmount = lngROFAmount - lngTemp
        End If

    Next xmlLCP

CreateROF_CreateLoanComponentPaymentsExit:

    Set xmlLCPList = Nothing
    Set xmlLCPSorted = Nothing
    Set xmlLCP = Nothing
    Set xmlFullLCPs = Nothing
    Set xmlFullLCPList = Nothing
    Set xmlPaymentRecord = Nothing
    Set xmlROFPaymentRecord = Nothing
    Set xmlNode = Nothing
    Set xmlTableNode = Nothing
    Set objDO = Nothing
    ' PSC 20/11/2002 - BMIDS01013 - Start
    Set xmlROFTotals = Nothing
    Set xmlROFTotal = Nothing
    ' PSC 20/11/2002 - BMIDS01013 - End


    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateROF_CreateLoanComponentPaymentsVbErr:
    gobjContext.SetAbort
    GoTo CreateROF_CreateLoanComponentPaymentsExit

End Sub


Private Function CreateROF_SortLCPs(ByVal xmlLCPSorted As IXMLDOMNode) As IXMLDOMNode
    
    Const cstrFunctionName As String = "CreateROF_SortLCPs"
    
    ' PSC 26/11/2002 BMIDS01052

    Dim xslDoc As FreeThreadedDOMDocument40, xmlTransformedDoc As FreeThreadedDOMDocument40
    Dim strXslPath As String
    Dim blnLoadedOk As Boolean
    
    Set xslDoc = New FreeThreadedDOMDocument40
    xslDoc.validateOnParse = False
    xslDoc.setProperty "NewParser", True
    xslDoc.async = False
    
    strXslPath = App.Path & "\CreateROF_SortLCPs.xslt"
    strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
    blnLoadedOk = xslDoc.Load(strXslPath)
    
    If Not blnLoadedOk Then
        errThrowError cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
    End If
    
    Set xmlTransformedDoc = New FreeThreadedDOMDocument40
    xmlTransformedDoc.validateOnParse = False
    xmlTransformedDoc.setProperty "NewParser", True
    xmlTransformedDoc.async = False
    
    xmlLCPSorted.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
    
    Set CreateROF_SortLCPs = xmlTransformedDoc.documentElement
    
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    
End Function


Private Sub CreateROF_CreatePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode, _
                                    ByRef rstrPaymentSeqNo As String)

    On Error GoTo CreateROF_CreatePaymentRecordVbErr

    Const cstrFunctionName As String = "CreateROF_CreatePaymentRecord"
    
    Dim xmlOrigPaymentRecord As IXMLDOMNode, xmlPaymentRecord As IXMLDOMNode, xmlROFPayment As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode, xmlDisbPayment As IXMLDOMNode, xmlROFDisbPayment As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim strApplicationNumber As String, strPaymentSeqNo As String, strPaymentMethod As String
    Dim strChequeNumber As String, strUserId As String, strPayeeType As String, strPaymentStatus As String
    Dim strIssueDate As String, strCompletionDate As String, strNotes As String, strPayeeHistorySeqNo As String
    Dim strROFID As String
    
    Dim lngPaymentAmount As Long, lngNetPaymentAmount As Long
    
    Dim objDO As IPAymentProcessingDO
    Dim objComboDO As ComboDO
    
    '--------------------------------------------
    ' Build Request and create Payment Record
    '--------------------------------------------
    Set xmlTableNode = xmlDoc.createElement("PAYMENTRECORD")
    
    Set xmlOrigPaymentRecord = xmlGetMandatoryNode(vxmlRequestNode, ".//ORIGPAYMENTRECORD")
    Set xmlROFPayment = xmlGetMandatoryNode(vxmlRequestNode, ".//ROFPAYMENTRECORD")
    strApplicationNumber = xmlGetAttributeText(xmlOrigPaymentRecord, "APPLICATIONNUMBER")
    strPaymentSeqNo = xmlGetAttributeText(xmlOrigPaymentRecord, "PAYMENTSEQUENCENUMBER")
    
    Set xmlPaymentRecord = xmlGetMandatoryNode(vxmlRequestNode, ".//PAYMENTRECORD")
    strPaymentMethod = xmlGetAttributeText(xmlROFPayment, "PAYMENTMETHOD")
    lngPaymentAmount = xmlGetAttributeAsLong(xmlROFPayment, "AMOUNT")
    strChequeNumber = xmlGetAttributeText(xmlROFPayment, "CHEQUENUMBER")
    strUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
    
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlTableNode, "ASSOCPAYSEQNUMBER", strPaymentSeqNo
    xmlSetAttributeValue xmlTableNode, "PAYMENTMETHOD", strPaymentMethod
    xmlSetAttributeValue xmlTableNode, "AMOUNT", (-1) * lngPaymentAmount
    xmlSetAttributeValue xmlTableNode, "CHEQUENUMBER", strChequeNumber
    xmlSetAttributeValue xmlTableNode, "USERID", strUserId
    xmlSetAttributeValue xmlTableNode, "CREATIONDATETIME", Format$(Now, "dd/mm/yy hh:mm:ss")
    
    Set objDO = gobjContext.CreateInstance(gstrPAYMENTPROCESSING & ".PaymentProcessingDO")
    objDO.CreatePaymentRecord xmlTableNode, vxmlResponseNode
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    
    strPaymentSeqNo = xmlGetAttributeText(xmlTableNode, "PAYMENTSEQUENCENUMBER")
    rstrPaymentSeqNo = strPaymentSeqNo
    '----------------------------------------------------------------------
    ' Build Request and create an entry in the table 'DisbursementPayment'
    '----------------------------------------------------------------------
    Set xmlTableNode = xmlDoc.createElement("DISBURSEMENTPAYMENT")
   
    Set xmlDisbPayment = vxmlRequestNode.selectSingleNode(".//PAYMENTRECORD/DISBURSEMENTPAYMENT")
    Set xmlROFDisbPayment = vxmlRequestNode.selectSingleNode(".//ROFPAYMENTRECORD/ROFDISBURSEMENTPAYMENT")
    strPayeeHistorySeqNo = xmlGetAttributeText(xmlDisbPayment, "PAYEEHISTORYSEQNO")
    strPayeeType = xmlGetAttributeText(xmlDisbPayment, "PAYEETYPE")
    strROFID = objComboDO.GetFirstComboValueId("PaymentType", "N")
    strPaymentStatus = xmlGetAttributeText(xmlROFDisbPayment, "PAYMENTSTATUS")
    strIssueDate = Format$(Now, "dd/mm/yy hh:mm:ss")
    strNotes = xmlGetAttributeText(xmlROFDisbPayment, "PAYMENTNOTES")
    strCompletionDate = xmlGetAttributeText(xmlROFDisbPayment, "COMPLETIONDATE")
    lngNetPaymentAmount = xmlGetAttributeAsLong(xmlROFDisbPayment, "NETPAYMENTAMOUNT")
    
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlTableNode, "PAYMENTSEQUENCENUMBER", strPaymentSeqNo
    xmlSetAttributeValue xmlTableNode, "PAYEEHISTORYSEQNO", strPayeeHistorySeqNo
    xmlSetAttributeValue xmlTableNode, "PAYEETYPE", strPayeeType
    xmlSetAttributeValue xmlTableNode, "PAYMENTTYPE", strROFID
    xmlSetAttributeValue xmlTableNode, "PAYMENTSTATUS", strPaymentStatus
    xmlSetAttributeValue xmlTableNode, "ISSUEDATE", strIssueDate
    xmlSetAttributeValue xmlTableNode, "COMPLETIONDATE", strCompletionDate
    xmlSetAttributeValue xmlTableNode, "PAYMENTNOTES", strNotes
    xmlSetAttributeValue xmlTableNode, "NETPAYMENTAMOUNT", (-1) * lngNetPaymentAmount
    objDO.CreateDisbursement xmlTableNode, vxmlResponseNode
    
    ' Append PaymentSequenceNumber (generated) to response
    Set xmlPaymentRecord = xmlDoc.createElement("PAYMENTRECORD")
    xmlSetAttributeValue xmlPaymentRecord, "PAYMENTSEQUENCENUMBER", strPaymentSeqNo
    
    vxmlResponseNode.appendChild xmlPaymentRecord
    
    gobjContext.SetComplete
    
CreateROF_CreatePaymentRecordExit:
    Set objDO = Nothing
    Set objComboDO = Nothing
    Set xmlOrigPaymentRecord = Nothing
    Set xmlPaymentRecord = Nothing
    Set xmlDisbPayment = Nothing
    Set xmlTableNode = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateROF_CreatePaymentRecordVbErr:
    gobjContext.SetAbort
    GoTo CreateROF_CreatePaymentRecordExit
    
End Sub

Private Sub CreateROF_CreateFeePayments(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode, ByVal strPaymentSeqNo As String)

On Error GoTo CreateROF_CreateFeePaymentsVbErr

    Const cstrFunctionName As String = "CreateROF_CreateFeePayments"
    
    Dim objDO As IPAymentProcessingDO, objComboDO As ComboDO

    Dim xmlPaymentRecord As IXMLDOMNode, xmlNewRequest As IXMLDOMNode, xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlFeePaymentList As IXMLDOMNodeList, xmlFeePayment As IXMLDOMNode
    
    Dim strApplicationNumber As String, strROFID As String, strOrigPaymentSeqNo As String
    Dim lngAmount As Long
        
    Set xmlPaymentRecord = vxmlRequestNode.selectSingleNode(".//ORIGPAYMENTRECORD")
    strOrigPaymentSeqNo = xmlGetMandatoryAttributeText(xmlPaymentRecord, "PAYMENTSEQUENCENUMBER")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlPaymentRecord, "APPLICATIONNUMBER")
        
    Set xmlFeePaymentList = vxmlRequestNode.selectNodes(".//FEEPAYMENT[@PAYMENTSEQUENCENUMBER=" & strOrigPaymentSeqNo & "]")
    
    If xmlFeePaymentList.length > 0 Then
        Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
        
        'BM0198/Bmids656 Set Event Type = Return of Fund
        strROFID = objComboDO.GetFirstComboValueId("PaymentEvent", "F")
    End If
    
    ' PSC 22/08/2003 BM0198
    Dim intFeeType As Integer
    
    For Each xmlFeePayment In xmlFeePaymentList
    
        ' PSC 22/08/2003 BM0198 - Start
        intFeeType = xmlGetAttributeAsInteger(xmlFeePayment, "FEETYPE")
        
        ' Only create a new fee payment if not TT as the TT fee will be deleted
        If IsValidationType("OneOffCost", intFeeType, "TTF") = False Then
            Set xmlNewRequest = xmlDoc.createElement("FEEPAYMENT")
            xmlSetAttributeValue xmlNewRequest, "APPLICATIONNUMBER", strApplicationNumber
            xmlSetAttributeValue xmlNewRequest, "PAYMENTSEQUENCENUMBER", strPaymentSeqNo
            xmlCopyAttribute xmlFeePayment, xmlNewRequest, "FEETYPE"
            xmlCopyAttribute xmlFeePayment, xmlNewRequest, "FEETYPESEQUENCENUMBER"
            xmlSetAttributeValue xmlNewRequest, "PAYMENTEVENT", strROFID
            
            lngAmount = xmlGetMandatoryAttributeAsLong(xmlFeePayment, "AMOUNTPAID")
            xmlSetAttributeValue xmlNewRequest, "AMOUNTPAID", (-1) * lngAmount
            
            objDO.CreateFeeTypePayment xmlNewRequest, vxmlResponseNode
        End If
        ' PSC 22/08/2003 BM0198 - End
        
    Next xmlFeePayment
        
CreateROF_CreateFeePaymentsExit:
    
    Set objDO = Nothing
    Set objComboDO = Nothing
    
    Set xmlPaymentRecord = Nothing
    Set xmlFeePaymentList = Nothing
    Set xmlFeePayment = Nothing
    Set xmlNewRequest = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateROF_CreateFeePaymentsVbErr:
    gobjContext.SetAbort
    GoTo CreateROF_CreateFeePaymentsExit

End Sub


Private Sub IPaymentProcessingTxBO_DeleteApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteApplicationFeeTypesVbErr

Const strFunctionName As String = "IPaymentProcessingTxBO_DeleteApplicationFeeTypes"

Dim xmlFeeTypeNode As IXMLDOMNode
Dim xmlFeeTypeNodeList As IXMLDOMNodeList
Dim objDO As IPAymentProcessingDO
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    Set xmlFeeTypeNodeList = vxmlRequestNode.selectNodes("APPLICATIONFEETYPE")
    For Each xmlFeeTypeNode In xmlFeeTypeNodeList
        objDO.DeleteApplicationFeeTypes xmlFeeTypeNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
DeleteApplicationFeeTypesExit:
    Set objDO = Nothing
    Set xmlFeeTypeNode = Nothing
    Set xmlFeeTypeNodeList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
DeleteApplicationFeeTypesVbErr:
    gobjContext.SetAbort
    GoTo DeleteApplicationFeeTypesExit

End Sub

Private Sub IPaymentProcessingTxBO_DeleteFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DeleteFeeTypePaymentVbErr
    
    Const strFunctionName As String = "IPaymentProcessingTxBO_DeleteFeeTypePayment"
    
    Dim xmlFeePayNode As IXMLDOMNode
    Dim objDO As IPAymentProcessingDO
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    'Delete the FEEPAYMENT record
    Set xmlFeePayNode = vxmlRequestNode.selectSingleNode("FEEPAYMENT")
    objDO.DeleteFeeTypePayment xmlFeePayNode, vxmlResponseNode
    
    'Delete the PAYMENTRECORD record
    Set xmlFeePayNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    objDO.DeletePaymentRecord xmlFeePayNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
DeleteFeeTypePaymentExit:
    Set objDO = Nothing
    Set xmlFeePayNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
DeleteFeeTypePaymentVbErr:
    gobjContext.SetAbort
    GoTo DeleteFeeTypePaymentExit

End Sub

Private Sub IPaymentProcessingTxBO_DeletePayeeHistory( _
        ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeletePayeeHistoryVbErr

Const strFunctionName As String = "IPaymentProcessingTxBO_DeletePayeeHistory"

Dim xmlFeeTypeNode As IXMLDOMNode
Dim xmlFeeTypeNodeList As IXMLDOMNodeList
Dim objDO As IPAymentProcessingDO
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    Set xmlFeeTypeNodeList = vxmlRequestNode.selectNodes("PAYEEHISTORY")
    For Each xmlFeeTypeNode In xmlFeeTypeNodeList
        objDO.DeletePayeeHistory xmlFeeTypeNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
DeletePayeeHistoryExit:
    Set objDO = Nothing
    Set xmlFeeTypeNode = Nothing
    Set xmlFeeTypeNodeList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
DeletePayeeHistoryVbErr:
    gobjContext.SetAbort
    GoTo DeletePayeeHistoryExit

End Sub



' GD BM0198 START
Private Sub IPaymentProcessingTxBO_SaveCostsFeesDisbursements(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo SaveCostsFeesDisbursementsVbErr

Const strFunctionName As String = "IPaymentProcessingTxBO_SaveCostsFeesDisbursements"
    
    'Constants
    Const cstrFunctionName As String = "IPaymentProcessingTxBO_SaveCostsFeesDisbursements"
    
    'Simple Types
    Dim strTTFValueId As String
    Dim strApplicationNumber As String
    Dim blnTTFeeAdded As Boolean
    Dim blnWaiveCheckBox As Boolean
    Dim blnMethodCreate As Boolean
    Dim blnStandingData As Boolean
    Dim intLength As Integer
    Dim strResponse As String
    Dim strRAValueId  As String
    Dim strFeeTypeSequenceNumber As String
    
    'XML Objects
    Dim xmlSettingsElement As IXMLDOMElement
    Dim xmlRequestClone As IXMLDOMElement
    Dim xmlFeePaymentList As IXMLDOMNodeList
    Dim xmlTTFeeElement As IXMLDOMElement
    Dim xmlPaymentElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCreateApplicationFeeTypeRequest As IXMLDOMElement
    Dim xmlApplicationFeeTypeElement As IXMLDOMElement
    Dim xmlCreateFeeTypePaymentRequest As IXMLDOMElement
    Dim xmlFeeTypePaymentElement As IXMLDOMElement
    Dim xmlResponse As IXMLDOMElement
    Dim xmlTempRequestElem As IXMLDOMElement
    
    'Other Objects
    Dim objPayProcBO As PaymentProcessingBO
    
    'Instantiate objects
    Set xmlRequestClone = vxmlRequestNode.cloneNode(True)
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objPayProcBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingBO")

    Set xmlSettingsElement = xmlRequestClone.selectSingleNode("SETTINGS")
    If Not (xmlSettingsElement Is Nothing) Then
    
        blnTTFeeAdded = CSafeBool(xmlGetAttributeText(xmlSettingsElement, "TTFEEADDED"))
        blnWaiveCheckBox = CSafeBool(xmlGetAttributeText(xmlSettingsElement, "WAIVECHECKBOX"))
        blnMethodCreate = CSafeBool(xmlGetAttributeText(xmlSettingsElement, "METHODCREATE"))
        strApplicationNumber = xmlGetMandatoryAttributeText(xmlRequestClone.selectSingleNode("PAYMENTRECORD"), "APPLICATIONNUMBER")
        blnStandingData = CSafeBool(xmlGetAttributeText(xmlSettingsElement, "STANDINGDATA"))
    Else
        errThrowError cstrFunctionName, oeMissingElement, "'SETTINGS' Element missing from Request"
    End If
    
    Set xmlFeePaymentList = vxmlRequestNode.selectNodes("//FEEPAYMENT")
    strTTFValueId = GetFirstComboValueId("OneOffCost", "TTF")
    If blnTTFeeAdded Then
        
        
        'Call PaymentProcessingBO.CreateApplicationFeeType with
        'ApplicationNumber,
        'ApplicationFeeType = Combo.ValueId (where validationtype = TTF),
        'Amount = FEEPAYMENT/@AMOUNTPAID of the LAST FEEPAYMENT IN THE LIST, since this will be the TTFee that was added in the GUI
        'The length of the FEEPAYMENT list
        'ONLY Create an APPLICATIONFEETYPE record if we have NO standing data
        If Not blnStandingData Then
            intLength = xmlFeePaymentList.length
            If intLength > 0 Then
                'Get the last element in the FEEPAYMENT list
                '***** THIS MAY FAIL Don't make this assumption - need to do a selectSingleNode from the list...
                Set xmlTTFeeElement = xmlFeePaymentList.Item(intLength - 1)
            Else
                'We could have zero in the list?
                'errThrowError cstrFunctionName, oeMissingElement, "No 'FEEPAYMENT' Elements exist"
            End If
    
    
            Set xmlCreateApplicationFeeTypeRequest = vxmlRequestNode.cloneNode(False)
            '
            xmlCreateApplicationFeeTypeRequest.setAttribute "OPERATION", "CREATEAPPLICATIONFEETYPE"
            Set xmlApplicationFeeTypeElement = xmlDoc.createElement("APPLICATIONFEETYPE")
            xmlApplicationFeeTypeElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber
            xmlApplicationFeeTypeElement.setAttribute "FEETYPE", strTTFValueId
            xmlApplicationFeeTypeElement.setAttribute "AMOUNT", xmlTTFeeElement.Attributes.getNamedItem("AMOUNTPAID").Text
            xmlCreateApplicationFeeTypeRequest.appendChild xmlApplicationFeeTypeElement
            Set xmlResponse = xmlDoc.createElement("RESPONSE")
            
            'Call PaymentProcessingBO.CreateApplicationFeeType
            strResponse = objPayProcBO.omPayProcRequest(xmlCreateApplicationFeeTypeRequest.xml)
            'CHECK RESPONSE
            errCheckXMLResponse strResponse, True
        
            'Get FEETYPESEQUENCENUMBER from the latest created, or the one that already exists
            strFeeTypeSequenceNumber = GetNextFeeTypeSequenceNumber(strApplicationNumber, strTTFValueId)
            'Add this sequencenumber to the TTFee element
            xmlTTFeeElement.Attributes.getNamedItem("FEETYPESEQUENCENUMBER").Text = strFeeTypeSequenceNumber
        Else 'Is Standing data, so TT info lives in request
            Set xmlTTFeeElement = vxmlRequestNode.selectSingleNode("//FEEPAYMENT[@FEETYPE = '" & strTTFValueId & "']")
            strFeeTypeSequenceNumber = xmlTTFeeElement.Attributes.getNamedItem("FEETYPESEQUENCENUMBER").Text
        End If
        
        
        If blnWaiveCheckBox Then

            'CreateFeeTypePayment with
                'ApplicationNumber,
                'PaymentEvent,
                'PaymentDate.
                
            
            'BM0198
            'Mark as rebated
            ' Look for 'O' as this will uniquely identify an Omiga fee rebate record
            strRAValueId = GetFirstComboValueId("PaymentEvent", "O")
            'Use existing TT Fee and mark it as rebated.
            'xmlTTFeeElement
            xmlTTFeeElement.setAttribute "PAYMENTEVENT", strRAValueId
            
            
            
            
            
''''''            Set xmlCreateFeeTypePaymentRequest = vxmlRequestNode.cloneNode(False)
''''''            ''''
''''''            xmlCreateFeeTypePaymentRequest.setAttribute "OPERATION", "CREATEFEETYPEPAYMENT"
''''''
''''''            Set xmlFeeTypePaymentElement = xmlDoc.createElement("FEEPAYMENT")
''''''
''''''            xmlFeeTypePaymentElement.setAttribute "APPLICATIONNUMBER", strApplicationNumber 'Primary Key
''''''            xmlFeeTypePaymentElement.setAttribute "PAYMENTEVENT", strRAValueId
''''''            xmlFeeTypePaymentElement.setAttribute "PAYMENTDATE", Format$(Now, "dd/mm/yy hh:mm:ss")
''''''            xmlFeeTypePaymentElement.setAttribute "FEETYPE", strTTFValueId 'Primary Key
''''''            xmlFeeTypePaymentElement.setAttribute "FEETYPESEQUENCENUMBER", strFeeTypeSequenceNumber 'Primary Key
''''''            xmlFeeTypePaymentElement.setAttribute "AMOUNTPAID", xmlTTFeeElement.Attributes.getNamedItem("AMOUNTPAID").Text
''''''            xmlCreateFeeTypePaymentRequest.appendChild xmlFeeTypePaymentElement
''''''
''''''
''''''            'Set xmlPaymentElement = xmlRequestClone.cloneNode(False).selectSingleNode(".//PAYMENTRECORD")
''''''            Set xmlPaymentElement = xmlRequestClone.selectSingleNode(".//PAYMENTRECORD").cloneNode(False)
''''''            xmlCreateFeeTypePaymentRequest.appendChild xmlPaymentElement
''''''
''''''            'Call PaymentProcessingBO with xmlCreateFeeTypePaymentRequest
''''''            strResponse = objPayProcBO.omPayProcRequest(xmlCreateFeeTypePaymentRequest.xml)
''''''
''''''            'CHECK STRING RESPONSE, and raise error if applicable.
''''''            errCheckXMLResponse strResponse, True
''''''
''''''            'Remove TT APPLICATIONFEETYPE from REQUEST - use vxmlRequestNode
''''''            Set xmlPaymentElement = vxmlRequestNode.selectSingleNode(".//PAYMENTRECORD")
''''''            If Not (xmlPaymentElement Is Nothing) Then
''''''                xmlPaymentElement.removeChild xmlTTFeeElement
''''''            Else
''''''                errThrowError cstrFunctionName, oeMissingElement, "'PAYMENTRECORD' Element missing from Request"
''''''            End If
        End If

    End If
    
    'Continue with save processing - Create or Update Disbursement and Fees
    'Use vxmlRequestNode

    If blnMethodCreate Then
        vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "CREATEDISBURSEMENT"
    Else
        vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "UPDATEDISBURSEMENT"
    End If
    strResponse = objPayProcBO.omPayProcRequest(vxmlRequestNode.xml)
      
    'CHECK RESPONSE
    errCheckXMLResponse strResponse, True
    
    
    gobjContext.SetComplete
    
SaveCostsFeesDisbursementsExit:
    Set xmlSettingsElement = Nothing
    Set xmlRequestClone = Nothing
    Set xmlFeePaymentList = Nothing
    Set xmlTTFeeElement = Nothing
    Set xmlPaymentElement = Nothing
    Set xmlDoc = Nothing
    Set xmlCreateApplicationFeeTypeRequest = Nothing
    Set xmlApplicationFeeTypeElement = Nothing
    Set xmlCreateFeeTypePaymentRequest = Nothing
    Set xmlFeeTypePaymentElement = Nothing
    Set xmlResponse = Nothing
    Set xmlTempRequestElem = Nothing
    Set objPayProcBO = Nothing

    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
SaveCostsFeesDisbursementsVbErr:
    gobjContext.SetAbort
    GoTo SaveCostsFeesDisbursementsExit
End Sub
' GD BM0198 END

Private Sub IPaymentProcessingTxBO_UpdateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateApplicationFeeTypeVbErr
    
    Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateApplicationFeeType"
    
    Dim xmlFeePayNode As IXMLDOMNode
    Dim objDO As IPAymentProcessingDO
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    Set xmlFeePayNode = vxmlRequestNode.selectSingleNode("APPLICATIONFEETYPE")
    objDO.UpdateApplicationFeeType xmlFeePayNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
UpdateApplicationFeeTypeExit:
    Set objDO = Nothing
    Set xmlFeePayNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateApplicationFeeTypeVbErr:
    gobjContext.SetAbort
    GoTo UpdateApplicationFeeTypeExit

End Sub

Private Sub IPaymentProcessingTxBO_UpdateChequeDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AW     15/09/2002  BM029
'
On Error GoTo UpdateDisbursementVbErr
    
Dim xmlRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlElement As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMElement
Dim xmlDisbNode As IXMLDOMElement
Dim objDO As IPAymentProcessingDO

Dim strSeqNo As String
Dim strAppNo As String
Dim strAppFFNo As String
Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateChequeDisbursement"
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlNodeList = vxmlRequestNode.selectNodes("PAYMENTRECORD")
    For Each xmlNode In xmlNodeList
        'Validate APPLICATIONNUMBER attribute is specified
        strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        'Validate PAYMENTSEQUENCENUMBER attribute is specified
        strSeqNo = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
        ' Raise error if no DISBURSEMENTPAYMENT node exists
        Set xmlDisbNode = xmlGetMandatoryNode(xmlNode, "DISBURSEMENTPAYMENT")
        
        'Update PAYMENTRECORD record
        'Set xmlNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
        objDO.UpdatePaymentRecord xmlNode, vxmlResponseNode
        
        'Update DISBURSEMENTPAYMENT record
        xmlDisbNode.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
        xmlDisbNode.setAttribute "APPLICATIONNUMBER", strAppNo
        
        objDO.UpdateDisbursement xmlDisbNode, vxmlResponseNode
        
        xmlDisbNode.removeAttribute "PAYMENTSEQUENCENUMBER"
        xmlDisbNode.removeAttribute "APPLICATIONNUMBER"
    Next
    
    gobjContext.SetComplete
    
UpdateDisbursementExit:
    Set objDO = Nothing
    Set xmlNode = Nothing
    Set xmlRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateDisbursementVbErr:
    gobjContext.SetAbort
    GoTo UpdateDisbursementExit

End Sub


Private Sub IPaymentProcessingTxBO_UpdateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateDisbursementVbErr
    
Dim xmlRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlElement As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMElement
Dim xmlDisbNode As IXMLDOMElement
Dim objDO As IPAymentProcessingDO

Dim strSeqNo As String
Dim strAppNo As String
Dim strAppFFNo As String
Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateDisbursement"
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlNodeList = vxmlRequestNode.selectNodes("PAYMENTRECORD")
    For Each xmlNode In xmlNodeList
        'Validate APPLICATIONNUMBER attribute is specified
        strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
        'Validate PAYMENTSEQUENCENUMBER attribute is specified
        strSeqNo = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
        ' Raise error if no DISBURSEMENTPAYMENT node exists
        Set xmlDisbNode = xmlGetMandatoryNode(xmlNode, "DISBURSEMENTPAYMENT")
        
        'Update PAYMENTRECORD record
        'Set xmlNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
        objDO.UpdatePaymentRecord xmlNode, vxmlResponseNode
        
        'Update DISBURSEMENTPAYMENT record
        xmlDisbNode.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
        xmlDisbNode.setAttribute "APPLICATIONNUMBER", strAppNo
        objDO.UpdateDisbursement xmlDisbNode, vxmlResponseNode
        
        'If the Amount has changed, reallocated the payments to the Loan Components
        If xmlAttributeValueExists(xmlNode, "AMOUNT") Then
            'Delete all LOANCOMPONENTPAYMENT records for this application and
            ' payment sequence number
            'and reallocate the payments to the Loan Components
            strAppFFNo = GetFactFindNumberForApplication(strAppNo)
            Set xmlRequest = xmlDoc.createElement("REQUEST")
            Set xmlElement = xmlDoc.createElement("LOANCOMPONENTPAYMENT")
            xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
            xmlElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
            'PSC 20/11/2002 BMIDS01035
            xmlElement.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
            xmlRequest.appendChild xmlElement
            objDO.DeleteLoanComponentPaymentDetails xmlElement, xmlTempResponse
            
            ' SR 12-10-01 : SYS3391 - if the amount is 0, delete all the corresponding FeePayment records
            If xmlGetAttributeAsLong(xmlNode, "AMOUNT") = 0 Then
                Set xmlElement = xmlDoc.createElement("FEEPAYMENT")
                xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
                xmlElement.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
                                    
                Set xmlTempResponse = xmlDoc.createElement("FEEPAYMENTLIST")
                objDO.FindFeePaymentList xmlElement, xmlTempResponse
                
                ' PSC 22/08/2003 BM0198 - Start
                Dim xmlFeePayments As IXMLDOMNodeList
                Dim xmlFeePayment As IXMLDOMNode
                Dim intFeeType As Integer
                Dim strFeeTypeSeqNo As String
                Dim xmlAppFeeType As IXMLDOMNode
                Dim intCount As Integer
         
                Set xmlFeePayments = xmlTempResponse.selectNodes(".//FEEPAYMENT")
                
                If xmlFeePayments.length > 0 Then
                    ' Delete FeePayment records
                    objDO.DeleteFeePayments xmlElement, xmlTempResponse
                                         
                    ' Go through each fee payment and delete the Application Fee Type
                    ' if it is a TT fee
                    For Each xmlFeePayment In xmlFeePayments
                        intFeeType = xmlGetAttributeAsInteger(xmlFeePayment, "FEETYPE")
                    
                        ' If a TT fee then set up request to delete the appropriate application fee
                        If IsValidationType("OneOffCost", intFeeType, "TTF") = True Then
                            strFeeTypeSeqNo = xmlGetAttributeText(xmlFeePayment, "FEETYPESEQUENCENUMBER")
                            Set xmlAppFeeType = xmlDoc.createElement("APPLICATIONFEETYPE")
                            xmlSetAttributeValue xmlAppFeeType, "APPLICATIONNUMBER", strAppNo
                            xmlSetAttributeValue xmlAppFeeType, "FEETYPE", CStr(intFeeType)
                            xmlSetAttributeValue xmlAppFeeType, "FEETYPESEQUENCENUMBER", strFeeTypeSeqNo
                            objDO.DeleteApplicationFeeTypes xmlAppFeeType, xmlTempResponse
                        End If
                    Next
                End If
                ' PSC 22/08/2003 BM0198 - End
                
            End If
            ' SR 12-10-01 : SYS3391 - END
            
            'Allocate the Payment to Loan Components
            AllocateDisbursementsToLoanComponents vxmlRequestNode
        End If
    Next
    
    gobjContext.SetComplete
    
UpdateDisbursementExit:
    Set objDO = Nothing
    Set xmlNode = Nothing
    Set xmlRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    
    ' PSC 22/08/2003 BM0198 - Start
    Set xmlFeePayments = Nothing
    Set xmlFeePayment = Nothing
    Set xmlAppFeeType = Nothing
    ' PSC 22/08/2003 BM0198 - End

    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateDisbursementVbErr:
    gobjContext.SetAbort
    GoTo UpdateDisbursementExit

End Sub


Private Sub IPaymentProcessingTxBO_UpdateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateFeeTypePaymentVbErr
    
    Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateFeeTypePayment"
    
    Dim xmlFeePayNode As IXMLDOMNode
    Dim objDO As IPAymentProcessingDO
    Dim strRefundDate As String
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    Set xmlFeePayNode = vxmlRequestNode.selectSingleNode("FEEPAYMENT")
    'SR 22/05/01 : SYS2298 - If RefundDate is mentioned in FEEPAYMENT, it should not be a past date
    strRefundDate = xmlGetAttributeText(xmlFeePayNode, "REFUNDDATE")
    If strRefundDate <> "" Then
        If Not IsDate(strRefundDate) Then
            errThrowError strFunctionName, 533
        End If
        
        If (CDate(strRefundDate) < Date) Then
            errThrowError strFunctionName, 565
        End If
    End If
    'SR 22/05/01 - End
    
    objDO.UpdateFeeTypePayment xmlFeePayNode, vxmlResponseNode
    
    Set xmlFeePayNode = vxmlRequestNode.selectSingleNode("PAYMENTRECORD")
    objDO.UpdatePaymentRecord xmlFeePayNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
UpdateFeeTypePaymentExit:
    Set objDO = Nothing
    Set xmlFeePayNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateFeeTypePaymentVbErr:
    gobjContext.SetAbort
    GoTo UpdateFeeTypePaymentExit

End Sub

Private Sub IPaymentProcessingTxBO_UpdatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdatePayeeHistoryDetailsVbErr

Dim objPayProcDO As IPAymentProcessingDO
Dim objThirdPartyBO As ThirdPartyBO
Dim xmlPayeeHistNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40
Dim strRet As String

    'Convert to old style element based request
    Set xmlDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "PAYEEHISTORY/THIRDPARTY", True)
    If Not xmlDoc Is Nothing Then
        'Update the Third Party and its associated Address record
        Set objThirdPartyBO = gobjContext.CreateInstance("omTP.ThirdPartyBO")
        strRet = objThirdPartyBO.UpdateThirdParty(xmlDoc.xml)
        
        'Update the Payee History record
        Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
        Set xmlPayeeHistNode = vxmlRequestNode.selectSingleNode("PAYEEHISTORY")
        objPayProcDO.UpdatePayeeHistoryDetails xmlPayeeHistNode, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
UpdatePayeeHistoryDetailsExit:
    Set objPayProcDO = Nothing
    Set objThirdPartyBO = Nothing
    Set xmlPayeeHistNode = Nothing
    Set xmlDoc = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
UpdatePayeeHistoryDetailsVbErr:
    gobjContext.SetAbort
    GoTo UpdatePayeeHistoryDetailsExit

End Sub

Private Sub IPaymentProcessingTxBO_UpdatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdatePaymentAuditVbErr

Const strFunctionName As String = "IPaymentProcessingTxBO_UpdatePaymentAudit"

Dim objPayProcDO As IPAymentProcessingDO
Dim xmlNode As IXMLDOMNode

    'Update the Payee History record
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlNode = vxmlRequestNode.selectSingleNode("PAYPROCBATCHAUDIT")
    objPayProcDO.UpdatePaymentAudit xmlNode, vxmlResponseNode

UpdatePaymentAudit_Exit:

    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
UpdatePaymentAuditVbErr:
    gobjContext.SetAbort
    GoTo UpdatePaymentAudit_Exit

End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub IPaymentProcessingTxBO_CreateApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationFeeTypesVbErr

Dim objPayProcDO As IPAymentProcessingDO
Dim objMortgageSQBO As MortgageSubQuoteBO
Dim xmlTempNode As IXMLDOMNode
Dim xmlNewNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
Dim xmlDomDoc As FreeThreadedDOMDocument40

Dim strAppNo As String                   ' MAR28
Dim strAmount As String                  ' MAR28
Dim strSeqNo As String                   ' MAR28
Dim strRefundAmount As String            ' MAR28
Dim xmlPaymentRecord As IXMLDOMElement   ' MAR28
Dim iFeeType As Integer                  ' MAR28
Dim strFeeType As String                 ' MAR28
Dim blnValRefund As Boolean              ' MAR1696
Dim strResponseXML As String
    
    'Create element based request renaming APPLICATIONFEETYPE to ONEOFFCOST
    Set xmlDomDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATIONFEETYPE", False, "ONEOFFCOST")
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set objMortgageSQBO = gobjContext.CreateInstance("omCM.MortgageSubQuoteBO")
    
    strResponseXML = objMortgageSQBO.GetOneOffCostsDetails(xmlDomDoc.xml)
    xmlDomDoc.loadXML (strResponseXML)
    
    'Select each OneOffCost where the amount > 0 and is not being added to the loan
    Set xmlNodeList = xmlDomDoc.selectNodes("//ONEOFFCOST/MORTGAGEONEOFFCOSTLIST/MORTGAGEONEOFFCOST[AMOUNT != ""0"" and ADDTOLOAN != ""1""]")
    Dim intValue As Integer
    For Each xmlNode In xmlNodeList
        
        'JLD SYS4168 Check the oneOffCost does not have a validationType of XAC
        intValue = xmlGetMandatoryNodeAsInteger(xmlNode, "MORTGAGEONEOFFCOSTTYPE")
        If IsValidationType("OneOffCost", intValue, "XAC") = False Then
            
            'Create FEETYPE node from MORTGAGEONEOFFCOSTTYPE
            Set xmlNewNode = xmlDomDoc.createElement("FEETYPE")
            strFeeType = xmlGetMandatoryNodeText(xmlNode, "MORTGAGEONEOFFCOSTTYPE")
            xmlNewNode.Text = strFeeType
            xmlNode.appendChild xmlNewNode
            
            'SR 22/05/01 : SYS2298 - The column 'RebateOrAddition' has been removed from table - 'ApplicationFeeType'
            'Add REBATEORADDITION with default value
            'Set xmlNewNode = xmlDomDoc.createElement("REBATEORADDITION")
            'xmlNewNode.Text = "0"
            'xmlNode.appendChild xmlNewNode
        
            'Convert to attribute based
            Set xmlTempNode = xmlCreateAttributeBasedResponse(xmlNode, False)
            
            'Create an APPLICATIONFEETYPE record
            objPayProcDO.CreateApplicationFeeTypes xmlTempNode, vxmlResponseNode
        
            'MAR28
            'If this is a Valuation Fee, set up PAYMENTRECORD and FEEPAYMENT records for the Refund
            'iFeeType = CSafeInt(xmlNode.Attributes.getNamedItem("FEETYPE").Text)
        
            If IsValidationType("OneOffCost", intValue, "VAL") = True Then
                'MAR28 Get Application Number and Amount
        'MAR1696 check global parameter to see if we need a refund
                blnValRefund = GetGlobalParamBoolean("CalculateValuationRefund")
                If blnValRefund Then
                    strAppNo = xmlGetMandatoryNodeText(xmlNode, "APPLICATIONNUMBER")
                    strAmount = xmlGetMandatoryNodeText(xmlNode, "AMOUNT")
                    strRefundAmount = "0"
                    strRefundAmount = xmlGetNodeText(xmlNode, "REFUNDAMOUNT")
    
                    Set xmlPaymentRecord = xmlDomDoc.createElement("PAYMENTRECORD")
                    xmlPaymentRecord.setAttribute "CREATIONDATETIME", Format$(Now, "dd/mm/yy hh:mm:ss")
                    xmlPaymentRecord.setAttribute "APPLICATIONNUMBER", strAppNo
                    xmlPaymentRecord.setAttribute "AMOUNT", strRefundAmount
    
                    objPayProcDO.CreatePaymentRecord xmlPaymentRecord, vxmlResponseNode
    
                    'Extract the PAYMENTSEQUENCENUMBER generated
                    strSeqNo = xmlGetAttributeText(xmlPaymentRecord, "PAYMENTSEQUENCENUMBER")
    
                    'Create a FEEPAYMENT record.
                    'Set the Amount Paid to 0
                    Set xmlPaymentRecord = xmlDomDoc.createElement("FEEPAYMENT")
                
                    xmlPaymentRecord.setAttribute "APPLICATIONNUMBER", strAppNo
                    xmlPaymentRecord.setAttribute "PAYMENTSEQUENCENUMBER", strSeqNo
                    xmlPaymentRecord.setAttribute "FEETYPE", strFeeType
                    xmlPaymentRecord.setAttribute "FEETYPESEQUENCENUMBER", xmlGetAttributeText(xmlTempNode, "FEETYPESEQUENCENUMBER")
                    xmlPaymentRecord.setAttribute "AMOUNTPAID", "0"
                    xmlPaymentRecord.setAttribute "REFUNDAMOUNT", strRefundAmount
                    xmlPaymentRecord.setAttribute "PAYMENTEVENT", GetFirstComboValueId("PaymentEvent", "RFV")
    
                    objPayProcDO.CreateFeeTypePayment xmlPaymentRecord, vxmlResponseNode
                End If  'MAR1696
            End If
       
        End If
    
     Next
    
    gobjContext.SetComplete
    
CreateApplicationFeeTypesExit:
    Set objPayProcDO = Nothing
    Set objMortgageSQBO = Nothing
    Set xmlTempNode = Nothing
    Set xmlNewNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
    Exit Sub
    
CreateApplicationFeeTypesVbErr:
    gobjContext.SetAbort
    GoTo CreateApplicationFeeTypesExit

End Sub


Private Sub AllocateDisbursementsToLoanComponents(ByVal vxmlRequest As IXMLDOMNode)
On Error GoTo AllocateDisbursementsToLoanComponentsVbErr

Const strFunctionName As String = "AllocateDisbursementsToLoanComponents"

Dim objPayProcDO As IPAymentProcessingDO
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlLoanCompList As IXMLDOMNodeList
Dim xmlLoanComp As IXMLDOMNode
Dim xmlPaymentList As IXMLDOMNodeList
Dim xmlPayment As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMElement
Dim xmlTempResponse As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim objMortSQBO As MortgageSubQuoteBO 'omCM.MortgageSubQuoteBO
Dim objComboDO As ComboDO

Dim strResponse As String
Dim strRequest As String
Dim strValue As String
Dim strAppNo As String
Dim strAppFFNo As String
Dim strMSQNo As String
Dim strLCSeqNo As String
Dim strPaySeqNo As String
Dim dteFirstPaymentDate As Date
Dim strPreferedPaymentDay As String
Dim dteIssueDate As Date
Dim lngLCUnallocated As Long
Dim lngPRUnallocated As Long
Dim lngCapitalAndInterestAmount As Long, lngInterestOnlyAmount As Long, lngTotalLoanCompAmount As Long
Dim lngAllocCapitalAndInterestAmount As Long, lngAllocInterestOnlyAmount As Long
Dim strRepaymentMethod As String, blnPartAndPartPayment As Boolean
Dim lngRet As Long
Dim lngAllocAmt As Long
Dim xmlTempNode As IXMLDOMNode
Dim blnRaiseError As Boolean
Dim XMLLoanCompPay As IXMLDOMNode
Dim iPaySeqNo As Integer
Dim iLowestPaySeqNo As Integer
Dim lngAllocedAmt As Long

' PSC 04/11/2002 BMIDS00721 - Start
Dim strPaymentType As String
' PSC 04/11/2002 BMIDS00721 - End

' PSC 13/03/2007 EP2_1347 - Start
Dim blnIsIncentive As Boolean
Dim blnIsValuationRefund As Boolean
' PSC 13/03/2007 EP2_1347 - End

blnRaiseError = True

    Set xmlPayment = xmlGetMandatoryNode(vxmlRequest, "PAYMENTRECORD")
    
    'Create Request to retrieve a list of Loan Components for the Active Quotation
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("FINDACTIVELOANCOMPONENTS")
    xmlTempRequest.appendChild xmlElement
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.appendChild xmlNode
    strAppNo = xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER")
    xmlNode.Text = strAppNo
    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.appendChild xmlNode
    strAppFFNo = GetFactFindNumberForApplication(strAppNo)
    xmlNode.Text = strAppFFNo
    
    Set objMortSQBO = gobjContext.CreateInstance("omCM.MortgageSubQuoteBO")
    strResponse = objMortSQBO.FindActiveLoanComponentsList(xmlTempRequest.xml)
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        lngRet = errGetOmigaErrorNumber(lngRet)
    End If
    If Not (lngRet = 0 Or lngRet = oeRecordNotFound) Then
        errThrowError strFunctionName, lngRet
    End If

    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    
    'Loop through each Loan Component
    xmlDoc.loadXML strResponse
    Set xmlLoanCompList = xmlDoc.selectNodes("RESPONSE/FINDACTIVELOANCOMPONENTSLIST/LOANCOMPONENT")
    'Dim lngFees As Long
    'Dim blnAddFees As Boolean
    'lngFees = 0
    For Each xmlLoanComp In xmlLoanCompList
        'Get the Unallocated amount for this Loan Component
        strMSQNo = xmlGetMandatoryNodeText(xmlLoanComp, "MORTGAGESUBQUOTENUMBER")
        strLCSeqNo = xmlGetMandatoryNodeText(xmlLoanComp, "LOANCOMPONENTSEQUENCENUMBER")
        'JD BMIDS735 if this is LC 1 then work out the fees added to loan
        ' Remove fees from LoanComponentPayment amount completely
        'If strLCSeqNo = "1" Then
        '    lngFees = xmlGetMandatoryNodeAsLong(xmlLoanComp, "TOTALLOANCOMPONENTAMOUNT") - _
        '              xmlGetMandatoryNodeAsLong(xmlLoanComp, " LOANAMOUNT")
        'End If

        lngAllocAmt = GetAllocatedLCAmount(strAppNo, strAppFFNo, strMSQNo, strLCSeqNo)
        'JD BMIDS735 ignore fee amount if LC 1
        'If lngAllocAmt > 0 And strLCSeqNo = "1" Then
        '    lngAllocAmt = lngAllocAmt - lngFees
        'End If
        'JD BMIDS735 TotalLoanComponentAmount includes any fees. Use LoanAmount instead.
        'lngLCUnallocated = xmlGetMandatoryNodeAsLong(xmlLoanComp, "TOTALLOANCOMPONENTAMOUNT") - lngAllocAmt
        lngLCUnallocated = xmlGetMandatoryNodeAsLong(xmlLoanComp, "LOANAMOUNT") - lngAllocAmt
        
        If lngLCUnallocated > 0 Then
                    
            'Get all Payment Records for this application that are related to DisbursementPayments i.e. not Fee Payments
            Set xmlElement = xmlDoc.createElement("GETDISBURSEMENTPAYMENTS")
            xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
            Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
            objPayProcDO.GetDisbursementPayments xmlElement, xmlTempResponse
            
            'For each PaymentRecord item, allocate payment to one or more Loan Components
            Set xmlPaymentList = xmlTempResponse.selectNodes("PAYMENTRECORD")
            For Each xmlPayment In xmlPaymentList
                
                ' PSC 04/11/2002 BMIDS00721 - Start
                ' Don't process incentive releases
                strPaymentType = xmlGetMandatoryAttributeText(xmlPayment, "PAYMENTTYPE")
                
                ' PSC 13/03/2007 EP2_1347 - Start
                blnIsIncentive = IsValidationType("PaymentType", CSafeInt(strPaymentType), "C")
                blnIsValuationRefund = IsValidationType("PaymentType", CSafeInt(strPaymentType), "VALREFUND")
                ' PSC 13/03/2007 EP2_1347 - End

                ' PSC 13/03/2007 EP2_1347
                If blnIsIncentive = False And blnIsValuationRefund = False Then
                    'Get Unallocated Payment Amount
                    strPaySeqNo = xmlGetMandatoryAttributeText(xmlPayment, "PAYMENTSEQUENCENUMBER")
                    'SR 01/08/01 : SYS2552
                    Set xmlTempNode = xmlGetMandatoryNode(vxmlRequest, "PREFERREDPAYMENTDAY")
                    strPreferedPaymentDay = xmlTempNode.Text
                    Set xmlTempNode = vxmlRequest.selectSingleNode("PAYMENTRECORD/DISBURSEMENTPAYMENT")
                    dteIssueDate = xmlGetMandatoryAttributeText(xmlTempNode, "ISSUEDATE")
                    
                    lngAllocCapitalAndInterestAmount = 0: lngAllocInterestOnlyAmount = 0
                    lngAllocAmt = GetAllocatedLCAmount(strAppNo, strAppFFNo, strMSQNo, strLCSeqNo, _
                                               lngAllocCapitalAndInterestAmount, lngAllocInterestOnlyAmount)
                    'JD BMIDS735 remove fees totally from LoanComponentAmount
                    'If lngAllocAmt > 0 And strLCSeqNo = "1" Then  'JD  BMIDS735 ignore fees if LC1
                    '    lngAllocAmt = lngAllocAmt - lngFees
                    'End If
                    'JD BMIDS735 TotalLoanComponentAmount includes any fees. Use LoanAmount instead.
                    'lngTotalLoanCompAmount = xmlGetMandatoryNodeAsLong(xmlLoanComp, "TOTALLOANCOMPONENTAMOUNT")
                    lngTotalLoanCompAmount = xmlGetMandatoryNodeAsLong(xmlLoanComp, "LOANAMOUNT")
                    lngLCUnallocated = lngTotalLoanCompAmount - lngAllocAmt
            
                    'SG 28/05/02 SYS4767 START
                    'OLD CODE:
                    ''SR 25/06/01 : SYS2412 - Get the CapitalAndInterestElement and InterestOnlyElement - they are populated
                    ''              when loan is Part & Part
                    'lngCapitalAndInterestAmount = xmlGetNodeAsLong(xmlLoanComp, "CAPITALANDINTERESTELEMENT")
                    'lngInterestOnlyAmount = xmlGetNodeAsLong(xmlLoanComp, "INTERESTONLYELEMENT")
                    'NEW CODE:
                    lngCapitalAndInterestAmount = xmlGetNodeAsLong(xmlLoanComp, "NETCAPANDINTELEMENT") 'JLD MSMS0022
                    lngInterestOnlyAmount = xmlGetNodeAsLong(xmlLoanComp, "NETINTONLYELEMENT") ' JLD MSMS0022
                    'SG 28/05/02 SYS4767 END
                    
                    strRepaymentMethod = xmlGetNodeText(xmlLoanComp, "REPAYMENTMETHOD")
                    If objComboDO.GetFirstComboValidation("RepaymentType", strRepaymentMethod) = "P" Then
                        blnPartAndPartPayment = True
                    Else
                        blnPartAndPartPayment = False
                    End If
                    
                    lngAllocAmt = GetAllocatedPRAmount(strAppNo, strAppFFNo, strPaySeqNo)
                    
                    'JD BMIDS735 if we already have a loancomponentpayment record this means
                    ' the Fees will already have been added.
                    'JD BMIDS735 fees are no longer held on the loancomponentpayment
                    'blnAddFees = True
'                    Set xmlElement = xmlDoc.createElement("LOANCOMPONENTPAYMENT")
'                    xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
'                    xmlElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
'                    xmlElement.setAttribute "MORTGAGESUBQUOTENUMBER", strMSQNo
'                    xmlTempRequest.appendChild xmlElement
'
'                    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
'                    blnRaiseError = False
'                    objPayProcDO.FindLoanComponentPaymentList xmlElement, xmlTempResponse
'                    blnRaiseError = True
'
'
'                    Dim xmlTmpList As IXMLDOMNodeList
'                    Set xmlTmpList = xmlTempResponse.selectNodes("LOANCOMPONENTPAYMENT")
'                    If Not xmlTmpList Is Nothing Then
'                        If xmlTmpList.length > 0 Then
'                            blnAddFees = False
'                        End If
'                        'Find the payment record associated with the fees.
'                        'This is the lowest payment record sequence number for loancomponent 1
'                        iLowestPaySeqNo = 999
'                        For Each XMLLoanCompPay In xmlTmpList
'                            If xmlGetAttributeText(XMLLoanCompPay, "LOANCOMPONENTSEQUENCENUMBER") = "1" Then
'                                iPaySeqNo = xmlGetAttributeAsInteger(XMLLoanCompPay, "PAYMENTSEQUENCENUMBER")
'                                If iLowestPaySeqNo > iPaySeqNo Then
'                                    iLowestPaySeqNo = iPaySeqNo
'                                End If
'                            End If
'                        Next
'                        If CStr(iLowestPaySeqNo) = strPaySeqNo And lngAllocAmt > 0 Then
'                            'GetAllocatedPRAmount will have returned an amount which included the fees
'                            'so remove them
'                            lngAllocAmt = lngAllocAmt - lngFees
'                        End If
'                    End If
                    
                    lngPRUnallocated = xmlGetMandatoryAttributeAsLong(xmlPayment, "AMOUNT") - lngAllocAmt
                    
                    Do Until lngPRUnallocated <= 0 Or lngLCUnallocated <= 0
                        If lngLCUnallocated >= lngPRUnallocated Then
                            lngAllocAmt = lngPRUnallocated
                        Else
                            lngAllocAmt = lngLCUnallocated
                        End If
                        
                        dteFirstPaymentDate = CheckFirstPaymentDate(dteIssueDate, strPreferedPaymentDay)
                                            
                        'Create a LOANCOMPONENTPAYMENT record allocating lngAllocAmt
                        'from the current PaymentRecord to the current LoanComponent
                        Set xmlElement = xmlDoc.createElement("LOANCOMPONENTPAYMENT")
                        xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
                        xmlElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
                        xmlElement.setAttribute "MORTGAGESUBQUOTENUMBER", strMSQNo
                        xmlElement.setAttribute "LOANCOMPONENTSEQUENCENUMBER", strLCSeqNo
                        xmlElement.setAttribute "PAYMENTSEQUENCENUMBER", strPaySeqNo
                        'If strLCSeqNo = "1" And blnAddFees = True Then
                        '    xmlElement.setAttribute "AMOUNT", CStr(lngAllocAmt + lngFees)
                        '    lngAllocedAmt = lngAllocAmt + lngFees
                        'Else
                            xmlElement.setAttribute "AMOUNT", CStr(lngAllocAmt)
                            lngAllocedAmt = lngAllocAmt
                        'End If
                        
                        'SR 05/07/01 : SYS2412 - assign values to CapitalAndInterestElement and InterstOnlyElement
                        If lngCapitalAndInterestAmount > lngAllocCapitalAndInterestAmount Then
                            If lngAllocedAmt > (lngCapitalAndInterestAmount - lngAllocCapitalAndInterestAmount) Then
                                xmlElement.setAttribute "CAPITALANDINTERESTELEMENT", _
                                        CStr(lngCapitalAndInterestAmount - lngAllocCapitalAndInterestAmount)
                                        
                                xmlElement.setAttribute "INTERESTONLYELEMENT", _
                                        CStr(lngAllocedAmt - (lngCapitalAndInterestAmount - lngAllocCapitalAndInterestAmount))
                            Else
                                xmlElement.setAttribute "CAPITALANDINTERESTELEMENT", CStr(lngAllocedAmt)
                            End If
                        Else
                            xmlElement.setAttribute "INTERESTONLYELEMENT", CStr(lngAllocedAmt)
                        End If
                        
                        ' SR 19/12/2001 : SYS2547
                        'xmlElement.setAttribute "FIRSTPAYMENTDATE", dteFirstPaymentDate
                        xmlElement.setAttribute "FIRSTPAYMENTDATE", CStr(dteFirstPaymentDate)
                        ' SR SYS2547 END
                        
                        xmlTempRequest.appendChild xmlElement
                        
                        Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
                        objPayProcDO.CreateLoanComponentPaymentDetails xmlElement, xmlTempResponse
                        
                        'Update the Unallocated Amounts
                        lngLCUnallocated = lngLCUnallocated - lngAllocAmt
                        lngPRUnallocated = lngPRUnallocated - lngAllocAmt
                        If lngLCUnallocated = 0 Then
                            'Don't look at any further payments as the Loan Component is now
                            'allocated in full. Move to the next Loan Component (if one exists)
                            Exit For
                        End If
                    Loop
                End If
                ' PSC 04/11/2002 BMIDS00721 - End
            Next
        End If
    Next
    
    gobjContext.SetComplete
    
AllocateDisbursementsToLoanComponentsExit:
    
    Set objPayProcDO = Nothing
    Set xmlDoc = Nothing
    Set xmlLoanCompList = Nothing
    Set xmlLoanComp = Nothing
    Set xmlPaymentList = Nothing
    Set xmlPayment = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set objMortSQBO = Nothing
    Set xmlTempNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
    Exit Sub
    
AllocateDisbursementsToLoanComponentsVbErr:

    'JD BMIDS735 check for record not found
    If errGetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If
    
    gobjContext.SetAbort
    GoTo AllocateDisbursementsToLoanComponentsExit

End Sub


Private Function GetAllocatedLCAmount(ByVal strAppNo As String, ByVal strAppFFNo As String, _
                                      ByVal strMSQNo As String, ByVal strLCSeqNo, _
                                      Optional ByRef lngAllocCapitalAndInterestAmount As Long, _
                                      Optional ByRef lngAllocInterestOnlyAmount As Long) As Long

On Error GoTo GetAllocatedLCAmountVbErr

Const strFunctionName As String = "GetAllocatedLCAmount"
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRequest As IXMLDOMElement
Dim xmlResponse As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim objPayProcDO As IPAymentProcessingDO

    'Create REQUEST
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("GETLCALLOCATEDAMOUNT")
    xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
    xmlElement.setAttribute "MORTGAGESUBQUOTENUMBER", strMSQNo
    xmlElement.setAttribute "LOANCOMPONENTSEQUENCENUMBER", strLCSeqNo
    xmlRequest.appendChild xmlElement
    
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    objPayProcDO.FindLoanComponentAllocatedAmount xmlRequest, xmlResponse
    
    'Extract the allocated amount from the Response
    Set xmlElement = xmlGetNode(xmlResponse, ".//GETLCALLOCATEDAMOUNT")
    If xmlElement Is Nothing Then
        GetAllocatedLCAmount = 0
        If Not IsMissing(lngAllocCapitalAndInterestAmount) Then lngAllocCapitalAndInterestAmount = 0
        If Not IsMissing(lngAllocInterestOnlyAmount) Then lngAllocInterestOnlyAmount = 0
    Else
        GetAllocatedLCAmount = xmlGetMandatoryAttributeAsLong(xmlElement, "LC_Allocated")
        
        'SR 26/06/01 : SYS2412 For Part&Part payment, get total amount allocated to CapitalAndInterestElement,
        '              InterestOnlyElement
        If Not IsMissing(lngAllocCapitalAndInterestAmount) Then
            lngAllocCapitalAndInterestAmount = xmlGetAttributeAsLong(xmlElement, "LC_ALLOCCANDIELEMENT")
        End If
        
        If Not IsMissing(lngAllocInterestOnlyAmount) Then
            lngAllocInterestOnlyAmount = xmlGetAttributeAsLong(xmlElement, "LC_ALLOCINTONLYELEMENT")
        End If
    End If

GetAllocatedLCAmountExit:
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlElement = Nothing
    Set objPayProcDO = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Function

GetAllocatedLCAmountVbErr:
    gobjContext.SetAbort
    GoTo GetAllocatedLCAmountExit

End Function

Private Function GetAllocatedPRAmount(ByVal strAppNo As String, ByVal strAppFFNo As String, _
                                        ByVal strPaySeqNo) As Long

On Error GoTo GetAllocatedPRAmountVbErr

Const strFunctionName As String = "GetAllocatedPRAmount"
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRequest As IXMLDOMElement
Dim xmlResponse As IXMLDOMElement
Dim xmlElement As IXMLDOMElement
Dim objPayProcDO As IPAymentProcessingDO

Dim lngRet As Long

    'Create REQUEST
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("GETPRALLOCATEDAMOUNT")
    xmlElement.setAttribute "APPLICATIONNUMBER", strAppNo
    'xmlElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
    xmlElement.setAttribute "PAYMENTSEQUENCENUMBER", strPaySeqNo
    xmlRequest.appendChild xmlElement
    
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    objPayProcDO.FindPaymentAllocatedAmount xmlRequest, xmlResponse

    'Extract the allocated amount from the Response
    Set xmlElement = xmlGetNode(xmlResponse, ".//GETPRALLOCATEDAMOUNT")
    If xmlElement Is Nothing Then
        GetAllocatedPRAmount = 0
    Else
        GetAllocatedPRAmount = xmlGetMandatoryAttributeAsLong(xmlElement, "PR_Allocated")
    End If

GetAllocatedPRAmountExit:
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlElement = Nothing
    Set objPayProcDO = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Function

GetAllocatedPRAmountVbErr:
    gobjContext.SetAbort
    GoTo GetAllocatedPRAmountExit

End Function

Private Function GetFactFindNumberForApplication(ByVal strAppNo As String) As String

On Error GoTo GetFactFindNumberForApplicationExit

Const strFunctionName As String = "GetFactFindNumberForApplication"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMElement
Dim xmlTempNode As IXMLDOMNode
Dim xmlElement As IXMLDOMElement
Dim objAppBO As ApplicationBO
Dim lngRet As Long
Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    
    'Get the ApplicationFactFindNumber for this application
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("APPLICATION")
    xmlTempRequest.appendChild xmlElement
    Set xmlTempNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlTempNode.Text = strAppNo
    xmlElement.appendChild xmlTempNode
    strResponse = objAppBO.GetApplicationData(xmlTempRequest.xml)
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        errThrowError strFunctionName, lngRet, "Error retrieving Application data"
    End If
    xmlDoc.loadXML strResponse
    
    GetFactFindNumberForApplication = xmlGetMandatoryNodeText(xmlDoc.documentElement, _
            "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/APPLICATIONFACTFINDNUMBER")

GetFactFindNumberForApplicationExit:
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempNode = Nothing
    Set xmlElement = Nothing
    Set objAppBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Sub IPaymentProcessingTxBO_UpdateFirstPaymentDateAndCreateBankBuildingSoc( _
        ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateFirstPaymentDateAndCreateBankBuildingSoc"
    
    Dim xmlNode As IXMLDOMNode
    Dim objAppTPBO As ApplicationThirdPartyBO
    
    Dim strResponse As String, strPrefPayDay As String
    
    Set objAppTPBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
        
     strPrefPayDay = xmlGetNodeText(vxmlRequestNode, "PREFEREDPAYMENTDAY")
     
    ' Update First Payment Date
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//LOANCOMPONENTPAYMENT")
    If Not xmlNode Is Nothing Then
        IPaymentProcessingTxBO_UpdateFirstPaymentDate vxmlRequestNode, vxmlResponseNode
    End If
    
    'Save ApplicationBankBuildingSoc details
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//APPLICATIONBANKBUILDINGSOC")
    If Not xmlNode Is Nothing Then
        strResponse = objAppTPBO.CreateBankBuildingSociety(vxmlRequestNode.xml)
        
        errCheckXMLResponse strResponse, True, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
UpdateVbExit:
    Set objAppTPBO = Nothing
         
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
UpdateVbErr:
    gobjContext.SetAbort
    
    GoTo UpdateVbExit
End Sub

Private Sub IPaymentProcessingTxBO_UpdateFirstPaymentDateAndBankBuildingSoc( _
            ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateFirstPaymentDateAndBankBuildingSoc"
    
    Dim xmlNode As IXMLDOMNode
    Dim objAppTPBO As ApplicationThirdPartyBO
    
    Dim strResponse As String, strPrefPayDay As String
    
    Set objAppTPBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
        
    strPrefPayDay = xmlGetNodeText(vxmlRequestNode, "PREFEREDPAYMENTDAY")
        
    ' Update First Payment Date
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//LOANCOMPONENTPAYMENT")
    If Not xmlNode Is Nothing Then
        IPaymentProcessingTxBO_UpdateFirstPaymentDate vxmlRequestNode, vxmlResponseNode, strPrefPayDay
    End If
    
    'Save ApplicationBankBuildingSoc details
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//APPLICATIONBANKBUILDINGSOC")
    If Not xmlNode Is Nothing Then
        strResponse = objAppTPBO.UpdateBankBuildingSociety(vxmlRequestNode.xml)
        
        errCheckXMLResponse strResponse, True, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
        
UpdateVbExit:
    Set objAppTPBO = Nothing
         
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
UpdateVbErr:
    gobjContext.SetAbort
    
    GoTo UpdateVbExit

End Sub

Private Sub IPaymentProcessingTxBO_UpdateFirstPaymentDate(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMNode, _
                                Optional ByVal strPrefPayDay As String = "")

On Error GoTo UpdateFirstPaymentDateExit

    Const strFunctionName As String = "IPaymentProcessingTxBO_UpdateFirstPaymentDate"

    Dim strPreferedPaymentDay As String
    Dim lngErrorNo As Long
    Dim strDescription As String
    Dim strSource As String

    Dim objGlobalParamDO As GlobalParameterDO
    Dim objPayProcDO As IPAymentProcessingDO
    Dim xmlPaymentNode As IXMLDOMNode
    Dim xmlInPaymentNode As IXMLDOMNode
    Dim xmlPaymentListNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList, xmlNode As IXMLDOMNode
    Dim xmlROFNodeList As IXMLDOMNodeList, xmlROFNode As IXMLDOMNode
    Dim xmlLCNode As IXMLDOMNode
    Dim xmlPreferedPaymentDay As IXMLDOMNode
    Dim lngTotalNoOfLCPayments As Long, lngUnPaidLCPayments As Long
    Dim lngActualPayment As Long, lngTotalROFPayment As Long
    Dim strPaidStatusId As String, strCondition As String, strPaymentSeqNo As String
    Dim blnCanDateBeChanged As Boolean

    Set xmlInPaymentNode = xmlGetMandatoryNode(vxmlRequestNode, "LOANCOMPONENTPAYMENT")
    Set xmlPaymentNode = vxmlRequestNode.ownerDocument.createElement("LOANCOMPONENTPPAYMENT")
    xmlCopyMandatoryAttribute xmlInPaymentNode, xmlPaymentNode, "APPLICATIONNUMBER"
    Set xmlPaymentListNode = vxmlRequestNode.ownerDocument.createElement("LOANCOMPONENTPPAYMENTLIST")
          
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    
    On Error Resume Next
    objPayProcDO.FindLoanComponentPaymentList xmlPaymentNode, xmlPaymentListNode
    
    lngErrorNo = Err.Number
    strDescription = Err.Description
    strSource = Err.Source
    
    On Error GoTo UpdateFirstPaymentDateExit
    
    If lngErrorNo <> 0 And errGetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
        Err.Raise lngErrorNo, strSource, strDescription
    End If
    
    '-----------------------------------------------------------------------------------------------
    'SR 24/08/01 :  Check whether all the loan component payments have status less than PAID
    '               (Unsanctioned, Sanctioned, cancelled, Error)
    '-----------------------------------------------------------------------------------------------
    Set xmlNodeList = xmlPaymentListNode.selectNodes(".//LOANCOMPONENTPAYMENT")
    lngTotalNoOfLCPayments = xmlNodeList.length
    
    strCondition = "LOANCOMPONENTPAYMENT[@PAYMENTSTATUS = 'U' or " & _
                                         " @PAYMENTSTATUS = 'S' or " & _
                                         " @PAYMENTSTATUS = 'C' or " & _
                                         " @PAYMENTSTATUS = 'E']"
    Set xmlNodeList = xmlPaymentListNode.selectNodes(strCondition)
    lngUnPaidLCPayments = xmlNodeList.length

    If lngTotalNoOfLCPayments = lngUnPaidLCPayments Then
        blnCanDateBeChanged = True
    End If
    
    
    If Not blnCanDateBeChanged Then
        '------------------------------------------------------------------------------------------
        'if any of the records have status more than PAID (Awaiting interface response, interfaced,
        'interfaced not paid), do not allow to edit the payment date.
        '------------------------------------------------------------------------------------------
        strCondition = "LOANCOMPONENTPAYMENT[@PAYMENTSTATUS = 'R' or " & _
                                         " @PAYMENTSTATUS = 'I' or " & _
                                         " @PAYMENTSTATUS = 'INP']"
        Set xmlNodeList = xmlPaymentListNode.selectNodes(strCondition)
        
        If xmlNodeList.length > 0 Then
            blnCanDateBeChanged = False
        Else
            '--------------------------------------------------------------------------------------
            'For each PAID record, find the corresponding total ROF amount. If ROF amount is less
            'than the actual payment amount, the date cannot be changed.
            '--------------------------------------------------------------------------------------
            strCondition = "LOANCOMPONENTPAYMENT[@PAYMENTSTATUS = 'P' ]"
            Set xmlNodeList = xmlPaymentListNode.selectNodes(strCondition)
        
            blnCanDateBeChanged = True
            
            For Each xmlNode In xmlNodeList
                lngActualPayment = 0: lngTotalROFPayment = 0
                strPaymentSeqNo = xmlGetAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
            
                strCondition = "LOANCOMPONENTPAYMENT[ASSOCPAYSEQNUMBER = " & strPaymentSeqNo & "]"
                Set xmlROFNodeList = xmlPaymentListNode.selectNodes(strCondition)
                If xmlROFNodeList.length > 0 Then
                    ' if the total ROF amount is less than the acual payment amount, payment date cannot be edited
                    For Each xmlROFNode In xmlROFNodeList
                        lngTotalROFPayment = lngTotalROFPayment + xmlGetAttributeAsLong(xmlROFNode, "AMOUNT")
                    Next xmlROFNode
            
                    If lngActualPayment > lngTotalROFPayment Then
                        blnCanDateBeChanged = False
                        Exit For
                    End If
                Else  ' No associated ROF records exist
                    blnCanDateBeChanged = False
                    Exit For
                End If
            Next xmlNode
        End If
    End If

    '-----------------------------------------------------------------------
    'If date is allowed to be changed, make the change
    '------------------------------------------------------------------------
    If blnCanDateBeChanged Then
        If strPrefPayDay <> "" Then
            strPreferedPaymentDay = strPrefPayDay
        Else
            'Make a call to find the prefered payment date
            Set objGlobalParamDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
            strPreferedPaymentDay = objGlobalParamDO.GetCurrentParameterByType("PreferredPaymentDay", "AMOUNT")
        End If
    Else
        errThrowError strFunctionName, 566
    End If
    
    UpdatePaymentDate xmlPaymentListNode, strPreferedPaymentDay
     
    gobjContext.SetComplete
    
UpdateFirstPaymentDateExit:
    
    Set objPayProcDO = Nothing
    Set xmlPaymentNode = Nothing
    Set xmlInPaymentNode = Nothing
    Set xmlPaymentListNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlLCNode = Nothing
    Set objGlobalParamDO = Nothing
    Set objPayProcDO = Nothing
    Set xmlROFNodeList = Nothing
    Set xmlROFNode = Nothing
         
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Function UpdatePaymentDate(ByVal xmlPaymentListNodeInput As IXMLDOMNode, Optional strInPreferedPaymentDaySelected As String)

On Error GoTo UpdatePaymentDateExit

Const strFunctionName As String = "UpdatePaymentDate"

Dim strInPreferedPaymentDay As String
Dim dteIssueDate As Date
Dim strIssueDay As String
Dim lngErrorNo As Long
Dim strDescription As String
Dim strSource As String
Dim dteRevDate As Date

Dim objPayProcDO As IPAymentProcessingDO
Dim xmlPaymentListNode As IXMLDOMNode
Dim xmlLoanCompPaymentList As IXMLDOMNodeList
Dim xmlLoanComponentPaymentNode As IXMLDOMElement
Dim xmlOutDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlResponseNode As IXMLDOMNode

    strInPreferedPaymentDay = strInPreferedPaymentDaySelected
        
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    Set xmlPaymentListNode = xmlPaymentListNodeInput
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlLoanCompPaymentList = xmlPaymentListNode.selectNodes("LOANCOMPONENTPAYMENT")
  
        For Each xmlLoanComponentPaymentNode In xmlLoanCompPaymentList
            dteIssueDate = xmlGetMandatoryAttributeText(xmlLoanComponentPaymentNode, "ISSUEDATE")
            strIssueDay = Format$(dteIssueDate, "dd")
            dteRevDate = CheckFirstPaymentDate(dteIssueDate, strInPreferedPaymentDay)
            xmlLoanComponentPaymentNode.setAttribute "FIRSTPAYMENTDATE", dteRevDate
            'Update the node
            objPayProcDO.UpdateLoanComponentPayment xmlLoanComponentPaymentNode, xmlResponseElem
        Next

UpdatePaymentDateExit:
    
    Set objPayProcDO = Nothing
    Set xmlPaymentListNode = Nothing
    Set xmlLoanCompPaymentList = Nothing
    Set xmlLoanComponentPaymentNode = Nothing
    Set xmlOutDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
         
         
    errCheckError strFunctionName, TypeName(Me)

End Function


Private Function CheckFirstPaymentDate(dteIssueDate As Date, strInPreferedPaymentDay As String) As Date
On Error GoTo CheckFirstPaymentDateExit

Const strFunctionName As String = "CheckFirstPaymentDate"

Dim dtePrefDate As Date
Dim dteNewDate As Date
Dim strIssueDay As String
Dim strMonth As String
Dim strYear As String

 'PSC 19/09/2002 SYS5349
 strIssueDay = (DatePart("d", Date))

 strMonth = (DatePart("m", Date))
 strYear = (DatePart("YYYY", Date))

    
    dtePrefDate = CDate(strInPreferedPaymentDay & "/" & strMonth & "/" & strYear)

    If strInPreferedPaymentDay = strIssueDay Then
        'Set FirstPaymentDate = IssueDate + 1 calendar month
        dteNewDate = Format(DateAdd("m", 1, dtePrefDate), "dd/mm/yyyy")
        'Passback the value
        CheckFirstPaymentDate = dteNewDate
    Else
        If dtePrefDate > dteIssueDate Then
        'Set FirstPaymentDate = strPrefDate + 1 calendar month
            dteNewDate = Format(DateAdd("m", 1, dtePrefDate), "dd/mm/yyyy")
            'Passback the value
            CheckFirstPaymentDate = dteNewDate
        Else
            'Set FirstPaymentDate = strPrefDate + 2 calendar months
            dteNewDate = Format(DateAdd("m", 2, dtePrefDate), "dd/mm/yyyy")
            'Update the node
            CheckFirstPaymentDate = dteNewDate
        End If
    End If
    
CheckFirstPaymentDateExit:
            
    errCheckError strFunctionName, TypeName(Me)
    

End Function


Private Sub IPaymentProcessingTxBO_CompleteInterfacing( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'DM     03/01/02    SYS3185 Implement CompleteInterfacing
'RF     03/01/02    Complete SYS3185 (Implement CompleteInterfacing)
'DM     14/03/02    SYS4274 CompleteInterfacing changes.
'DM     27/03/02    SYS4313 CompleteInterfacing changes
'------------------------------------------------------------------------------------------
On Error GoTo CompleteInterfacingVbErr

    Const cstrFunctionName As String = "IPaymentProcessingTxBO_CompleteInterfacing"

    Dim objAppManagerBO As ApplicationManagerBO
    Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
    Dim objBatchSchedBO As BatchScheduleBO
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim nodeTemp As IXMLDOMNode
    Dim nodeCBRequest As IXMLDOMNode
    Dim nodeCBRequestHeader As IXMLDOMNode
    Dim nodeRespFromNTTxBO As IXMLDOMNode
    Dim nodeNewRequest As IXMLDOMNode
    Dim nodeResponseFromGetLockData As IXMLDOMNode
    Dim elemTemp As IXMLDOMElement
    Dim nodeAppLock As IXMLDOMNode
    
    ' PSC 20/09/2002 BMIDS00480 - Start
    Dim xmlHeader As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    ' PSC 20/09/2002 BMIDS00480 - End
    
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strBatchAuditGuid As String
    
    Dim strResponse As String ' Return value for called methods
    Dim strApplicationNumber As String ' Application number from the request node
    Dim strEventMesssage As String
    
    Dim strUserId As String
    Dim strUnitID As String
    Dim strUserAuthorityLevel As String
    
    ' initialise some useful pointers
    ' DM 15/03/02 SYS4274
    ' PSC 20/09/2002 BMIDS00480 - Start
'    Set nodeCBRequest = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST")
    ' END DM 15/03/02 SYS4274
    
    Set xmlHeader = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/HEADER")
    Set xmlResponse = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/RESPONSE")

'    Set nodeCBRequestHeader = xmlGetMandatoryNode(nodeCBRequest, "HEADER")
    strApplicationNumber = xmlGetAttributeText(xmlHeader, "APPLICATIONNUMBER")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlHeader, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlHeader, "BATCHRUNNUMBER")
    strBatchAuditGuid = xmlGetMandatoryAttributeText(xmlHeader, "BATCHAUDITGUID")
    
    '------------------------------------------------------------------------------------------
    'DM     14/03/02    SYS4274 CompleteInterfacing changes.
    '   Set up the USERID UNITID from the USERIDENTIFICATION, USERUNITIDENTIFICATION tags
    '   that are returned from the admin system.
    '------------------------------------------------------------------------------------------
    strUserId = xmlGetAttributeText(xmlHeader, "USERID")
    strUnitID = xmlGetAttributeText(xmlHeader, "UNITID")
    ' PSC 20/09/2002 BMIDS00480 - End
    
    xmlSetAttributeValue vxmlRequestNode, "USERID", strUserId
    xmlSetAttributeValue vxmlRequestNode, "UNITID", strUnitID
    
    '--------------------------------------------------------------------------------
    'DM 27/03/02 SYS4313 set up the userauthoritylevel
    'DM 08/04/02 SYS4371 Temp fix for authority level.
    strUserAuthorityLevel = "99"
    'strUserAuthorityLevel = xmlGetAttributeText(nodeCBRequestHeader, "USERAUTHORITYLEVEL")
    'END DM 08/04/02 SYS4371 Temp fix for authority level.
    xmlSetAttributeValue vxmlRequestNode, "USERAUTHORITYLEVEL", strUserAuthorityLevel
    
    '------------------------------------------------------------------------------------------
    ' Check that the application is locked to the correct user
    '------------------------------------------------------------------------------------------
    
    Set objPayProcNTTxBO = _
        gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
    ' Need to pass an existing response node.
    Set nodeResponseFromGetLockData = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
    objPayProcNTTxBO.GetLockData vxmlRequestNode, nodeResponseFromGetLockData

    Dim lngErrNo As Long
    lngErrNo = errCheckXMLResponseNode(nodeResponseFromGetLockData, , False)
    
    Dim blnLockExists As Boolean
    
    If lngErrNo = 0 Then
        blnLockExists = True
    Else
        If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
            blnLockExists = False
        Else
            ' re-raise the error
            errCheckXMLResponseNode nodeResponseFromGetLockData, vxmlResponseNode, True
        End If
    End If
    
    If blnLockExists = True Then
    
        If (nodeResponseFromGetLockData.selectSingleNode("APPLICATIONLOCK/UNITID").Text = _
            strUnitID) And _
            (nodeResponseFromGetLockData.selectSingleNode("APPLICATIONLOCK/USERID").Text = _
            strUserId) Then
            '------------------------------------------------------------------------------------------
            ' Start a new transaction
            '------------------------------------------------------------------------------------------
            
            ' Pass the original request to the new transaction.
            Set nodeRespFromNTTxBO = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
            objPayProcNTTxBO.CompleteInterfacing vxmlRequestNode, nodeRespFromNTTxBO
                       
            ' check the response for Omiga errors
            errCheckXMLResponseNode nodeRespFromNTTxBO, vxmlResponseNode, True
            
            'BM0511 GHun 08/04/2003
            Set nodeRespFromNTTxBO = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
            objPayProcNTTxBO.UnlockApplication vxmlRequestNode, nodeRespFromNTTxBO
            'Errors in the response from UnlockApplication can be ignored
            'BM0511 End
                       
            ' handle the response
            Set objBatchSchedBO = gobjContext.CreateInstance( _
                    gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
            
            ' PSC 20/09/2002 BMIDS00480 - Start
            If xmlGetMandatoryNodeText( _
                vxmlRequestNode, "/REQUEST/RESPONSE/@TYPE") = "SUCCESS" Then
                'DM     14/03/02    SYS4274 CompleteInterfacing changes.
'                If xmlGetMandatoryNodeText( _
'                    vxmlRequestNode, "/REQUEST/BODY/STATUS/@VALUE") <> "FAILURE" Then

                    ' Complete success

                    ' Call BatchScheduleBO.IncrementBatchScheduleSuccess
                Set nodeNewRequest = Nothing
                Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue nodeNewRequest, _
                    "OPERATION", "IncrementBatchScheduleSuccess"
                Set nodeTemp = _
                    nodeNewRequest.ownerDocument.createElement("BATCHSCHEDULE")
                nodeNewRequest.appendChild nodeTemp
                xmlSetAttributeValue nodeTemp, "BATCHNUMBER", strBatchNumber
                xmlSetAttributeValue nodeTemp, "BATCHRUNNUMBER", strBatchRunNumber
                xmlSetAttributeValue nodeTemp, "BATCHAUDITGUID", strBatchAuditGuid

                strResponse = objBatchSchedBO.omBatchRequest(nodeNewRequest.xml)

                errCheckXMLResponse strResponse, True

            Else

                ' Error

                ' Write an event to the event log
                strEventMesssage = _
                    "A business error has occurred within Completions processing for Application " & strApplicationNumber
                App.LogEvent strEventMesssage, vbLogEventTypeError

                ' Call BatchScheduleBO.IncrementBatchScheduleFailure
                Set nodeNewRequest = Nothing
                Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue nodeNewRequest, _
                    "OPERATION", "IncrementBatchScheduleFailure"
                Set nodeTemp = _
                    nodeNewRequest.ownerDocument.createElement("BATCHSCHEDULE")
                nodeNewRequest.appendChild nodeTemp
                xmlSetAttributeValue nodeTemp, "BATCHNUMBER", strBatchNumber
                xmlSetAttributeValue nodeTemp, "BATCHRUNNUMBER", strBatchRunNumber
                xmlSetAttributeValue nodeTemp, "BATCHAUDITGUID", strBatchAuditGuid
                
                ' PSC 20/09/2002 BMIDS00480
                xmlSetAttributeValue nodeTemp, "ERRORNUMBER", 570
                xmlSetAttributeValue nodeTemp, "ERRORSOURCE", _
                GetGlobalParamString("AdminSystemName")
                    
                'SYS4835 - More meaningful message logged.
                ' DM SYS4274
                ' PSC 20/09/2002 BMIDS00480
                ' PSC 14/11/2005 MAR520 Get error form correct place
                ' PSC 29/11/2005 MAR733 Get DESCRIPTION rather than ERRORDESCRIPTION
                xmlSetAttributeValue nodeTemp, "ERRORDESCRIPTION", _
                    Mid$(xmlGetNodeText(xmlResponse, "ERROR/ERRORDESCRIPTION"), 1, 255)
                strResponse = objBatchSchedBO.omBatchRequest(nodeNewRequest.xml)

                errCheckXMLResponse strResponse, True

            End If

'            Else
'                ' Technical error
'
'                ' Write an event to the event log
'                ' DM SYS4274
'                strEventMesssage = _
'                    "A technical error has occurred within Completions processing for Application " & strApplicationNumber & _
'                    "." & vbLf & "Description: " & xmlGetNodeText(vxmlRequestNode, "//SYSTEMSTATUS/@DESCRIPTION")
'                App.LogEvent strEventMesssage, vbLogEventTypeError
'
'                ' Call BatchScheduleBO.IncrementBatchScheduleFailure
'                Set nodeNewRequest = Nothing
'                Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
'                xmlSetAttributeValue nodeNewRequest, _
'                    "OPERATION", "IncrementBatchScheduleFailure"
'                Set nodeTemp = _
'                    nodeNewRequest.ownerDocument.createElement("BATCHSCHEDULE")
'                nodeNewRequest.appendChild nodeTemp
'                xmlSetAttributeValue nodeTemp, "BATCHNUMBER", strBatchNumber
'                xmlSetAttributeValue nodeTemp, "BATCHRUNNUMBER", strBatchRunNumber
'                xmlSetAttributeValue nodeTemp, "BATCHAUDITGUID", strBatchAuditGuid
'                xmlSetAttributeValue nodeTemp, "ERRORNUMBER", 570
'                xmlSetAttributeValue nodeTemp, "ERRORSOURCE", _
'                    GetGlobalParamString("AdminSystemName")
'                ' DM SYS4274
'                xmlSetAttributeValue nodeTemp, "ERRORDESCRIPTION", _
'                    xmlGetNodeText(vxmlRequestNode, "REQUEST/SYSTEMSTATUS/DESCRIPTION")
'
'                strResponse = objBatchSchedBO.omBatchRequest(nodeNewRequest.xml)
'
'                errCheckXMLResponse strResponse, True
'
'            End If
            ' PSC 20/09/2002 BMIDS00480 - End



            HandleFinalTransaction _
                vxmlRequestNode, strResponse, _
                strBatchNumber, strBatchRunNumber
            
        Else
            ' "Cannot process Completions response as Application %s is locked to a different user"
            errThrowError cstrFunctionName, 573, , strApplicationNumber
        End If
    Else
        ' no lock
        
        ' "Cannot process Completions response as Application %s is not locked"
        errThrowError cstrFunctionName, 572, , strApplicationNumber
    End If
    
CompleteInterfacingExit:
    
    Set objAppManagerBO = Nothing
    Set objPayProcNTTxBO = Nothing
    Set objBatchSchedBO = Nothing
    Set xmlDoc = Nothing
    Set nodeTemp = Nothing
    Set nodeCBRequest = Nothing
    Set nodeRespFromNTTxBO = Nothing
    Set nodeAppLock = Nothing
    Set nodeResponseFromGetLockData = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CompleteInterfacingVbErr:
      
    ' Write an event to the event log
    strEventMesssage = _
        "An internal error has occurred within Completions processing for Application " & strApplicationNumber & _
        "." & vbLf & "Description: " & Err.Description
    App.LogEvent strEventMesssage, vbLogEventTypeError
    
    ' call BatchScheduleBO.IncrementBatchScheduleFailure
    
    Dim strSource As String             ' PSC 20/09/2002 BMIDS00480
    Dim strDescription As String        ' PSC 20/09/2002 BMIDS00480
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
    
    If objBatchSchedBO Is Nothing Then
        Set objBatchSchedBO = gobjContext.CreateInstance( _
            gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    End If
            
    Set nodeNewRequest = Nothing
    Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue nodeNewRequest, _
        "OPERATION", "IncrementBatchScheduleFailure"
    Set nodeTemp = _
        nodeNewRequest.ownerDocument.createElement("BATCHSCHEDULE")
    nodeNewRequest.appendChild nodeTemp
    xmlSetAttributeValue nodeTemp, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue nodeTemp, "BATCHRUNNUMBER", strBatchRunNumber
    xmlSetAttributeValue nodeTemp, "BATCHAUDITGUID", strBatchAuditGuid
    
    ' PSC 20/09/2002 BMIDS00480 - Start
    xmlSetAttributeValue nodeTemp, "ERRORNUMBER", CStr(lngErrNo)
    xmlSetAttributeValue nodeTemp, "ERRORSOURCE", strSource
    xmlSetAttributeValue nodeTemp, "ERRORDESCRIPTION", strDescription
    ' PSC 20/09/2002 BMIDS00480 - End
    
    strResponse = objBatchSchedBO.omBatchRequest(nodeNewRequest.xml)
    
    errCheckXMLResponse strResponse, True
    
    HandleFinalTransaction vxmlRequestNode, strResponse, strBatchNumber, strBatchRunNumber

    GoTo CompleteInterfacingExit

End Sub


Private Sub HandleFinalTransaction( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vstrBatchIncrementResponse As String, _
    ByVal vstrBatchNumber As String, _
    ByVal vstrBatchRunNumber As String)
' header ----------------------------------------------------------------------------------
' description:
'   Check for this being the final transaction in a batch job,
'   and if so call CompleteBatch
' pass:
' return:
' history:
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName As String = "HandleFinalTransaction"

    Dim objPayProcBO As PaymentProcessingBO
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim nodeRequest As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    Dim strResponse As String
    
    Set xmlDoc = xmlLoad(vstrBatchIncrementResponse, cstrFunctionName)
    If xmlGetMandatoryNodeText( _
        xmlDoc.documentElement, "//RESPONSE/BATCHSCHEDULE/@LASTTRANSACTION") = "1" Then
        
        ' call CompleteBatch
        
        Set objPayProcBO = _
            gobjContext.CreateInstance(App.Title & ".PaymentProcessingBO")
        
        Set nodeRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue nodeRequest, _
            "OPERATION", "CompleteBatch"
        Set nodeTemp = _
            nodeRequest.ownerDocument.createElement("BATCHSCHEDULE")
        nodeRequest.appendChild nodeTemp
        xmlSetAttributeValue nodeTemp, "BATCHNUMBER", vstrBatchNumber
        xmlSetAttributeValue nodeTemp, "BATCHRUNNUMBER", vstrBatchRunNumber
        
        strResponse = objPayProcBO.omPayProcRequest(nodeRequest.xml)
        
        errCheckXMLResponse strResponse, True
        
    ' Else do nothing
        
    End If

    Set xmlDoc = Nothing
    Set objPayProcBO = Nothing
    Set nodeRequest = Nothing
    Set nodeTemp = Nothing
    
End Sub

Private Sub IPaymentProcessingTxBO_CreateChequeAuditRecord(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateChequeAuditRecordVbErr

    Dim xmlChequeAuditNode As IXMLDOMNode
    Dim objPayProcDO As IPAymentProcessingDO
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlChequeAuditNode = vxmlRequestNode.selectSingleNode("PAYMENTCHEQUEAUDIT")
    If Not xmlChequeAuditNode Is Nothing Then
        objPayProcDO.CreateChequeAudit xmlChequeAuditNode, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
CreateChequeAuditRecordExit:
    Set objPayProcDO = Nothing
    Set xmlChequeAuditNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreateChequeAuditRecordVbErr:
    gobjContext.SetAbort
    GoTo CreateChequeAuditRecordExit

End Sub
'GD BM0198 START
Private Function GetNextFeeTypeSequenceNumber(ByVal strApplicationNumber As String, ByVal strFeeType As String) _
    As Integer
    
On Error GoTo GetNextFeeTypeSequenceNumberExit
    
    Dim xmlNode As IXMLDOMNode

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim param As ADODB.Parameter
    
    Dim strSQL As String, _
        strSQLWhere As String, _
        strPattern As String, _
        strDataSource As String, _
        strSequenceFieldName As String
        
    Dim intThisSequence As Integer
        
    Dim blnDbCmdOk As Boolean
    Dim strSQLNoLock As String
    
    strDataSource = "APPLICATIONFEETYPE"
    
    strSequenceFieldName = "FEETYPESEQUENCENUMBER"
    
    'If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
    'This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
    'If (genumDbProvider = omiga4DBPROVIDERSQLServer) And (xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK") = "TRUE") Then
    strSQLNoLock = " WITH (NOLOCK)"
    Set cmd = New ADODB.Command
    strSQLWhere = " APPLICATIONNUMBER = ? AND FEETYPE = ?"

    '2 Parameters
    '1 .. ApplicationNumber
    '2 .. FeeType
    Set param = New ADODB.Parameter
    param.Type = adBSTR
    param.Size = Len(strApplicationNumber)
    param.Direction = adParamInput
    param.Value = strApplicationNumber

    cmd.Parameters.Append param

    Set param = New ADODB.Parameter
    param.Type = adInteger
    param.Value = strFeeType
    cmd.Parameters.Append param
    strSQL = "SELECT MAX(" & strSequenceFieldName & ") FROM " & strDataSource & strSQLNoLock & " WHERE (" & strSQLWhere & ")"
    Debug.Print "GetNextFeeTypeSequenceNumber strSQL: " & strSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = executeGetRecordSet(cmd)
    If Not rst Is Nothing Then
        rst.MoveFirst
        If IsNull(rst.Fields.Item(0)) = False Then
            intThisSequence = rst.Fields.Item(0).Value
        End If
        rst.Close
    End If
    'GetNextFeeTypeSequenceNumber = intThisSequence + 1
    GetNextFeeTypeSequenceNumber = intThisSequence
    Err.Clear
GetNextFeeTypeSequenceNumberExit:

    Set rst = Nothing
    Set cmd = Nothing
    Set xmlNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

'GD BM0198 END
' PSC 15/11/2005 MAR140 - Start
Private Sub IPaymentProcessingTxBO_GetPostCompletionCaseTasksList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetPostCompletionCaseTasksListVbErr
    
    Const cstrFunctionName As String = "GetPostCompletionCaseTasksList"
    
    Dim xmlApplication As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlTask As IXMLDOMNode
    Dim xmlCompletedTask As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlAppPriority As IXMLDOMNode
    
    
    Dim xmlIncompleteTask As IXMLDOMNode
    Dim xmlCompletedTasks As IXMLDOMNodeList

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskResponse As FreeThreadedDOMDocument40
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    
    Dim objMsgTmBO As MsgTmBO
    Dim objomTMBO As OmTmBO
    Dim objomApp As ApplicationBO
    
    Dim strResponse As String
    Dim strPostCompValueId As String
    Dim strCompletedStatusSearch As String
    Dim strIncompleteStatusSearch As String
    Dim strTaskID As String
    Dim strAppPriority As String

    Dim intIndex As Integer
    Dim dblChasingPeriodDays As Double
    
    Dim dteDueDate As Date
    
    Dim varValue As Variant
    
    ' Raise error if no node exists
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    'Validate mandatory attributes are specified
    xmlGetMandatoryAttributeText xmlApplication, "APPLICATIONNUMBER"
    xmlGetMandatoryAttributeText xmlApplication, "APPLICATIONFACTFINDNUMBER"
    xmlGetMandatoryAttributeText xmlApplication, "STAGEID"
    xmlGetMandatoryAttributeText xmlApplication, "COMPLETIONDATE"
    
    ' Get combo values
    strPostCompValueId = GetFirstComboValueId("TaskType", "PC")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    ' Get currents stage tasks
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "FindCaseTaskList"
    
    Set xmlCaseTask = xmlDoc.createElement("CASETASK")
    xmlCopyAttributeValue xmlApplication, xmlCaseTask, "APPLICATIONNUMBER", "CASEID"
    xmlCopyAttributeValue xmlApplication, xmlCaseTask, "STAGEID", "STAGEID"               ' MAR1408
    xmlTempRequest.appendChild xmlCaseTask
    
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    Set objomTMBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
    Set objomApp = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    
    strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
    
    Set xmlCaseTaskResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlCaseTaskResponse.documentElement, , True
    
    Dim colIncompleteValues As Collection
    Set colIncompleteValues = New Collection
    
    ' Get the values for task status of incomplete
    GetValueIdsForValidationType "TaskStatus", "I", colIncompleteValues
    
    intIndex = 1
    
    ' Build Xpath search criteria
    For Each varValue In colIncompleteValues
        strCompletedStatusSearch = strCompletedStatusSearch & " and @TASKSTATUS !='" & varValue & "'"
        
        If intIndex = 1 Then
            strIncompleteStatusSearch = "("
        Else
            strIncompleteStatusSearch = strIncompleteStatusSearch & " or "
        End If
        
        strIncompleteStatusSearch = strIncompleteStatusSearch & "@TASKSTATUS ='" & varValue & "'"
        
        intIndex = intIndex + 1
    Next
    
    If Len(strIncompleteStatusSearch) > 0 Then
        strIncompleteStatusSearch = strIncompleteStatusSearch & ")"
    End If
    
    ' Find unique completed tasks
    Set xmlCompletedTasks = xmlCaseTaskResponse.selectNodes("RESPONSE/CASETASK[not(@TASKID=preceding-sibling::CASETASK/@TASKID) and @TASKTYPE='" & strPostCompValueId & "'" & strCompletedStatusSearch & "]")

    If xmlCompletedTasks.length > 0 Then
    
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        Set xmlAppPriority = xmlDoc.createElement("APPLICATIONPRIORITY")
        xmlTempRequest.appendChild xmlAppPriority
        Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlAppPriority.appendChild xmlNode
        xmlNode.Text = xmlGetAttributeText(xmlApplication, "APPLICATIONNUMBER")
        strResponse = objomApp.GetApplicationPriority(xmlTempRequest.xml)
        Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        strAppPriority = xmlGetNodeText(xmlTempResponse, "RESPONSE/APPLICATIONPRIORITY/APPLICATIONPRIORITYVALUE")
    
        ' Create request for case task creation
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "CreateAdhocCaseTask"
        Set xmlCaseTask = xmlDoc.createElement("CASETASK")
        xmlTempRequest.appendChild xmlCaseTask
        Set xmlNode = xmlDoc.createElement("APPLICATION")
        xmlTempRequest.appendChild xmlNode
        xmlSetAttributeValue xmlNode, "APPLICATIONPRIORITY", strAppPriority
        
        ' For each completed task check to see if there is another one that is incomplete.
        ' If not generate a new one
        For Each xmlCompletedTask In xmlCompletedTasks
            strTaskID = xmlGetAttributeText(xmlCompletedTask, "TASKID")
            Set xmlIncompleteTask = xmlCaseTaskResponse.selectSingleNode("RESPONSE/CASETASK[@TASKID='" & strTaskID & "' and " & strIncompleteStatusSearch & "]")
            
            ' Create new task
            If xmlIncompleteTask Is Nothing Then
            
                dblChasingPeriodDays = xmlGetAttributeAsDouble(xmlCompletedTask, "CHASINGPERIODDAYS")
                dteDueDate = xmlGetAttributeAsDate(xmlApplication, "COMPLETIONDATE")
                
                If dblChasingPeriodDays > 0 Then
                    dteDueDate = DateAdd("d", dblChasingPeriodDays, dteDueDate)
                End If
        
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "CASEID"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "ACTIVITYID"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "ACTIVITYINSTANCE"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "CASESTAGESEQUENCENO"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "STAGEID"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "TASKID"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "CASETASKNAME"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "MANDATORYINDICATOR"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "OWNINGUSERID"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "OWNINGUNITID"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "CONTEXT"
                xmlCopyAttribute xmlCompletedTask, xmlCaseTask, "CUSTOMERIDENTIFIER"
                xmlSetAttributeValue xmlCaseTask, "TASKDUEDATEANDTIME", Format(dteDueDate, "dd/mm/yyyy")
                
                strResponse = objomTMBO.OmTmRequest(xmlTempRequest.xml)
                
                errCheckXMLResponse strResponse, True
                
            End If
        Next
    End If
    
    gobjContext.SetComplete
    
GetPostCompletionCaseTasksListExit:

    Set xmlApplication = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTask = Nothing
    Set xmlTask = Nothing
    Set xmlCompletedTask = Nothing
    Set xmlNode = Nothing
    Set xmlAppPriority = Nothing
    Set xmlIncompleteTask = Nothing
    Set xmlCompletedTasks = Nothing
    Set xmlDoc = Nothing
    Set xmlCaseTaskResponse = Nothing
    Set xmlTempResponse = Nothing
    Set objMsgTmBO = Nothing
    Set objomTMBO = Nothing
    Set objomApp = Nothing
    Exit Sub
    
GetPostCompletionCaseTasksListVbErr:
    gobjContext.SetAbort
    GoTo GetPostCompletionCaseTasksListExit
End Sub
' PSC 15/11/2005 MAR140 - End


