VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPaymentProcessingTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IPaymentProcessingTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   PaymentProcessing Business Object.
'
'Dependencies:
'-------------------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'PSC    15/11/2005  MAR140 Add GetPostCompletionCaseTasksList
'--------------------------------------------------------------------------------------------------------------

Option Explicit

Public Sub CreateApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, _
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

Public Sub DeleteFeeTypePayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeleteApplicationFeeTypes(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdatePaymentAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateFirstPaymentDate(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode, _
            Optional ByVal strPrefPayDay As String = "")
End Sub

Public Sub DeletePayeeHistory(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateReturnOfFunds(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateFirstPaymentDateAndCreateBankBuildingSoc(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateFirstPaymentDateAndBankBuildingSoc(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CompleteInterfacing(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateChequeAuditRecord(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Public Sub UpdateChequeDisbursement(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
' GD BM0198 START
Public Sub SaveCostsFeesDisbursements(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
' GD BM0198 END

' PSC 15/11/2005 MAR140 - Start
Public Sub GetPostCompletionCaseTasksList(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
' PSC 15/11/2005 MAR140 - End




