VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQuotationBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'ING History:
'
'Prog   Date        Description
'MV     14/09/2005  MARS35 - New Function UpdateLastOfferedQuotation() and GetAllQuotationsForApplication()
'------------------------------------------------------------------------------------------
Option Explicit

Public Function GetValidatedQuotationDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateFirstPPSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewPPSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function IsMortgageSubQuoteAttachedToOneQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function Update(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function Validate(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewQuotation(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function StoreQuotation(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateFirstMortgageLifeSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewMortgageLifeSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ReinstateSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function IsQuotationAffordable(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateFirstBCSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewBCSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindStoredQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateQuotationForReinstatement(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveOneOffCostDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetQuotationSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAcceptedBCSubQuoteNumber(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetQuotationLoanSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLinkedQuotations(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS624 GHun 22/10/2003
Public Function HaveRatesChanged(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS624 End

Public Function UpdateLastOfferedQuotation(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAllQuotationsForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
