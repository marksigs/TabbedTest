VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'HISTORY
'-------------------------------------------------------------
'SG     28/05/02    SYS4746 MSMS to Core integration
'-------------------------------------------------------------

Option Explicit

Public Enum IApplicationQuoteBOMethod
    bomRemodelQuotation
    'Added DRC BMIDS767
    bomUpdateQuotationForAddedCosts
End Enum


'-------------------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date     Description
' GD    08/07/02 BMIDS00165 - Added new method GetAcceptedOrActiveQuoteData
' MV    03/02/03 BM0221 - Added new method ValidateUserMandateLevel
' DRC   20/05/04 BMIDS767 - Added new method UpdateQuotationForAddedCosts
' JD    04/04/2005 BMIDS977 - added CreateAdHocFeeAndRemodel
' HMA   12/04/2005 BMIDS977 - removed UpdateQuotationForAddedCosts
'-------------------------------------------------------------------------------------------------------
'Epsom History:
'
' Prog  Date     Description
' PSC   05/12/06 EP2_249 Added new method GetIntroducerProcFeeDetails
'-------------------------------------------------------------------------------------------------------
Public Function AcceptQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetLoanCompositionDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMortgageProducts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindStoredQuoteDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateFirstMortgageLifeSubQuotes(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAndSaveLifeCoverCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function BuildLifeCoverCalcInputInterface(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalcCostModelLTV(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecommendQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateUserMandateLevel(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAndSaveBCSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ProcessAddedOneOffCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetHighestEarners(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAffordability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateMortgageSubquote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetApplicationSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RemodelQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecostQuotationForRateChange(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAcceptedQuoteData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalcAllowableInc(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateManualAdjustmentPercent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetAcceptedOrActiveQuoteData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' BMIDS977
'Public Function UpdateQuotationForAddedCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
'End Function

Public Function RefreshMortgageAccountData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAndSavePortedStepAndPeriodFromMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateAdHocFeeAndRemodel(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateApproximateMonthlyCost(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

' PSC 05/12/2006 EP2_249 - Start
Public Function GetIntroducerProcFeeDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' PSC 05/12/2006 EP2_249 - End

