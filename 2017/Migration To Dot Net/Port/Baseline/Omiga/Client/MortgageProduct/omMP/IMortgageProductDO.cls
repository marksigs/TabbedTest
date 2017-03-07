VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMortgageProductDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'GD     10/09/02    BMIDS00313 - APWP2 - BM004 Added FindProductConditionsList
'MDC    01/11/2002  BMIDS00654 Added GetMortgageProductIncomeMultiplierSet
'JD     21/06/04    BMIDS765 CC076 added GetMortgageProductRentalIncomeRateSet
'------------------------------------------------------------------------------

Public Function FindProductNames(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetValidActiveProduct(ByVal vxmlTableElement As IXMLDOMElement) As Boolean
End Function

Public Function FindProductList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetFullProductDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetRateForBaseRateSet(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetMortgageProductRentalIncomeRateSet(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetCurrentRateForBaseRateSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetRateDifference(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetInterestRateType(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageProductLanguage(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindProductConditionsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

'BMIDS00654 MDC 01/11/2002
Public Function GetMortgageProductIncomeMultiplierSet(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function
'BMIDS00654 MDC 01/11/2002 - End


