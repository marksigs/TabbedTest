VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQuickQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'HISTORTY
'----------------------------------------------------------------
'SG     28/05/02    SYS4767 Added ValidateManualAdjustmentPercent
'----------------------------------------------------------------
Option Explicit
Public Function CalculateAndSaveBCSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMortgageProducts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindStoredQuoteDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetLoanCompositionDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAndSaveLifeCoverCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function BuildLifeCoverCalcInputInterface(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalcCostModelLTV(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ProcessAddedOneOffCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'SG 28/05/02 SYS4767
Public Function ValidateManualAdjustmentPercent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

