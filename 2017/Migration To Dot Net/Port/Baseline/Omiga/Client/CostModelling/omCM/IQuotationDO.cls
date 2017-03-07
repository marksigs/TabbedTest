VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQuotationDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'ING History:
'
'Prog   Date        Description
'MV     14/09/2005  MARS35 - New Function GetAllQuotationsForApplication()
'------------------------------------------------------------------------------------------

Option Explicit

Public Function GetAllQuotationsForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetQuotationDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub Update(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub AddDerivedData(ByVal vxmlData As IXMLDOMNode)
End Sub

Public Function Create(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CountQuotationsForMortgageSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As Integer
End Function

Public Function StoredQuoteFindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetQuotationSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAcceptedBCSubQuoteNumber(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetQuotationLoanSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS624 GHun 22/10/2003
Public Function HaveRatesChanged(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS624 GHun End
