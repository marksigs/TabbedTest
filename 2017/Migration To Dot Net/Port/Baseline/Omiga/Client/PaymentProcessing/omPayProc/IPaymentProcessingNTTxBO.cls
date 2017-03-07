VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPaymentProcessingNTTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IPaymentProcessingNTTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   PaymentProcessing Business Object.
'
'Dependencies:
'-------------------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'PSC    25/11/2005  MAR140 Add GetPostCompletionCaseTasksList
'-------------------------------------------------------------------------------------------------------
Option Explicit

Public Sub PaySanctionedPayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub InterfacePaidPayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CompleteInterfacing(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ProcessPayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetLockData(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub SanctionPrintCheque(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode, ByVal bLocked As Boolean)
End Sub

Public Sub CheckLocks(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Public Sub SetPaidChequePayment(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMElement)
End Sub

'BM0339 MDC 07/03/2003
Public Sub UnlockApplication(ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'BM0339 MDC 07/03/2003 - End

' PSC 25/11/2005 MAR140 - Start
Public Function GetPostCompletionCaseTasksList(ByVal vxmlRequest As String) As String
End Function
' PSC 25/11/2005 MAR140 - End

