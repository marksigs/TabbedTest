VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMortgageProductBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum IOmiga4BOMethod
    bomGetFullProductDetails
End Enum

'------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'GD     10/09/02    BMIDS00313 - APWP2 - BM004 Added FindProductConditionsList
'MDC    01/11/2002  BMIDS00654 Added GetMortgageProductIncomeMultiplierSet
'JD     21/06/04    BMIDS765 CC076 added GetMortgageProductRentalIncomeRateSet
'------------------------------------------------------------------------------

Public Function FindProductNames(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindProductList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetFullProductDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function Validate(ByVal vxmlRequest As IXMLDOMElement, _
                         ByVal vbomMethod As IOmiga4BOMethod) As IXMLDOMNode
End Function

Public Function GetRateForBaseRateSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageProductRentalIncomeRateSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetInterestRateForRateChange(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetMortgageProductLanguage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindProductConditionsList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

'BMIDS00654 MDC 01/11/2002
Public Function GetMortgageProductIncomeMultiplierSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
'BMIDS00654 MDC 01/11/2002 - End

