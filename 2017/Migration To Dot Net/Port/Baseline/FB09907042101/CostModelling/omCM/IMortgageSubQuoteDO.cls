VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMortgageSubQuoteDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function FindLoanComponentList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function TotalIncentivesForAllComponents(ByVal vxmlTableElement As IXMLDOMElement) As Double

End Function

Public Sub DeleteAllMortgageIncentive(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateMortgageIncentive(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindMortgageIncentiveList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub SaveIncentiveAgainstLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

'BMIDS782 GHun
'Public Sub CreateLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'End Sub
'BMIDS782 End

Public Sub CreateLoanComponentRedemptionFee(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

'BMIDS782 GHun
'Public Sub DeleteAllLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'End Sub
'
'Public Sub DeleteLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'End Sub
'
'Public Function GetLoanComponentResultsData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
'End Function
'
'Public Sub UpdateLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'End Sub
'BMIDS782 End

Public Function GetNonPanelMortgageProductData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub DeleteAllNonPanelMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateNonPanelMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteNonPanelMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindNonPanelMortgageProductList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub DeleteAllLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetLoanComponentData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub CreateLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetLowestLoanComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetLoanComponentProductParameters(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub Create(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetPortedLender(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub ResetMortgageSubQuote(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub ReserveMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateOneOffCosts(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindMortgageOneOffCostList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub DeleteAllMortgageOneOffCost(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateMortgageOneOffCost(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub Update(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetDataForOneOffCostsDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetMortgageCalculatorDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub UpdateTotalLoanAmount(ByVal vxmlTableElement As IXMLDOMElement, ByVal lngLoanAmount As Long, _
                                    Optional ByVal blnNonPanel As Boolean = True)

End Sub

Public Function GetAmountRequested(ByVal vxmlTableElement As IXMLDOMElement) As Long
                        
End Function

Public Function FindActiveLoanComponentsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

'BMIDS736 GHun 23/02/2004
Public Sub DeleteAllLoanComponentPaymentSchedules(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteAllLoanComponentBalanceSchedules(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateLoanComponentPaymentSchedule(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateLoanComponentBalanceSchedule(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteAllLoanComponentRedemptionFees(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
'BMIDS736 End

'BMIDS782 GHun
Public Sub SaveIllustrationResults(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
'BMIDS782 End
