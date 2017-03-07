VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAymentProcessingDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************
'   BMIDS
'
'   AW  23/08/02    BM029           Added CreateChequeAudit
'   PSC 20/11/2002  BMIDS01013      Added FindROFTotals
'   MC  16/04/2004  BMIDS763        ADDED GetProductSwitchFee() AND GetTTFeeAmount()
'************************************************************************************
'   MARS
'   HM  23/09/2005  WP13 MAR49      ADDED ValidateCompletionDate, ValidateSolicitor, SetAdvanceDate, SetPaymentMethod, CreateDelayedCompletionTask
'   PSC 15/11/2005  MAR140          Remove GetPostCompletionCaseTasksList
'   DRC 02/03/2006  MAR1334         Added FindPaymentsToCancel
'************************************************************************************

Option Explicit

Public Sub ValidateCompletionDate(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ValidateSolicitor(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub SetAdvanceDate(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub SetPaymentMethod(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ResetPostCompletionTasks(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub


Public Sub FindFeeTypeList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeleteApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeleteFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindFeePaymentList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateApplicationFeeType(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindPaymentHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindPayeeHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode, Optional ByVal blnErrorIfNotFound As Boolean = True)
End Sub

Public Sub FindPaymentsForPayeeList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdatePayeeHistoryDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindPaymentRecordList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreatePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdatePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeletePaymentRecord(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindLoanComponentPaymentList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindDisbursementPaymentList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateLoanComponentPaymentDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeleteLoanComponentPaymentDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindLoanComponentAllocatedAmount(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindPaymentAllocatedAmount(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetDisbursementPayments(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindApplicationThirdPartyList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindPaymentsForBatch(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindSanctioningList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindBatchPaymentList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetBatchParameters(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindPaymentAuditList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateLoanComponentPayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindDisbursementPaymentSummary(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeletePayeeHistory(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeleteFeePayments(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateChequeAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

' PSC 19/09/2002 SYS4863
Public Sub UpdateFeePaymentList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindInterfacedInitialAdvances(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

' PSC 20/11/2002 BMIDS01013
Public Sub FindROFTotals(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

'*=[MC]BMIDS763 - CC075
Public Sub GetProductSwitchFee(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

'*=[MC]BMIDS763 - CC075
Public Sub GetTTFeeAmount(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

'MAR28 Add new function
Public Sub GetValuationFee(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode, Optional ByVal blnRefund As Boolean = False)

End Sub
'*=SECTION END


Public Sub FindPaymentsToCancel(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

End Sub
