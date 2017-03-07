VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMortgageSubQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum IOmiga4BOMethod
    bomResetMortgageSubQuote
    bomReserveMortgageProduct
    bomSaveLoanComponentDetails
    bomCalcAndSaveMortgageCosts
    bomCreateOneOffCosts
    bomSaveIllustrationResults
    bomUpdateLoanComponent
End Enum

Public Function ValidateSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetOneOffCostDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetLoanCompositionDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetMortgageSubQuoteData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ResetMortgageSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteLoanComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ReserveMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateNewSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateComponentsFromExistingAcc(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function


Public Function CreateFirstSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveLoanComponentDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateCompulsoryProducts(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateOneOffCosts(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

'BMIDS782 GHun
'Public Function BuildCalcsEngineInputInterface(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
'End Function
'
'Public Function RunCalcsEngine(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
'End Function
'BMIDS782 End

Public Function SaveIllustrationResults(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveIncentives(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CalculateTotalIncentiveAmount(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetLoanComponentData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ProcessAddedOneOffCost(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function RecostComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveOneOffCostDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CalcAndSaveMortgageCosts(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CalcAndSaveRedemptionFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function Validate(ByVal vXMLRequest As IXMLDOMElement, _
                         ByVal vbomMethod As IOmiga4BOMethod) As IXMLDOMNode
End Function

Public Function Update(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindActiveLoanComponentsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function FindLoanComponentList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLoanComponent(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMortgageIncentiveList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindAvailableincentives(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS782 GHun
'Public Function BuildCompositeRateInterface(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
'End Function
'BMIDS782 End

Public Function RunCompositeRateCalc(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveCompositeRateResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RunCompositeRateCalcAndSaveResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateDrawDownCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetAPROneOffCosts(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

' PSC 30/11/2006 EP2_249 - Start
Public Function SaveIntroducerFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' PSC 30/11/2006 EP2_249 - End

' PSC 15/03/2007 EP2_1956 - Start
Public Function CopyIntroducerFees(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' PSC 15/03/2007 EP2_1956 - End

