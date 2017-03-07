VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBuildingsAndContentsSubQuoteDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CreateSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function CreateValuablesOverLimit(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Sub CreateBCNotes(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub UpdateBCNotes(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub DeleteValuablesOverLimit(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub DeleteAll(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub DeleteAllValuables(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Function GetBCNotes(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetPaymentFrequency(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindValuablesOverLimitList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Sub UpdateSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub UpdateValuablesOverLimit(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub AddDerivedData(ByVal vxmlData As IXMLDOMNode)
End Sub
Public Function GetSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub GetInsuranceAdminFee(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponse As IXMLDOMNode)


End Sub

