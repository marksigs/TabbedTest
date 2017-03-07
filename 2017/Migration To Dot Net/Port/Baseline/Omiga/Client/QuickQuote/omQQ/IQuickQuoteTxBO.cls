VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQuickQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'MV     12/08/2002  BMIDS00311 - Core Ref : SYS1705 ;Modified CalculateMortgageCosts
'------------------------------------------------------------------------------------------
Public Function CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement, _
                                        xmlLoanCompositionNode As IXMLDOMNode, _
                                        xmlAppNo As IXMLDOMNode, _
                                        xmlAppFactFindNo As IXMLDOMNode, _
                                        xmlLifeSubQuoteNumber As IXMLDOMNode, _
                                        xmlMortgageSubQuoteNumber As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAndSaveBCSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

