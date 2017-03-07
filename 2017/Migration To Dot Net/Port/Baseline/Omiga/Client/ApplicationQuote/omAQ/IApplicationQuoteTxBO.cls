VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
' MV    12/08/2002  BMIDS00311  Core Ref : SYS1705 Removed SafeRef ;
'                               Modified CalculateMortgageCosts
' DRC   20/05/2004              Added UpdateQuotationForAddedCosts
' JD    04/04/2005  BMIDS977    Added CreateAdHocFeeAndRemodel
' HMA   06/04/2005  BMIDS977    Added ProcessAllFeesForRemodelQuote
' HMA   12/04/2005  BMIDS977    Removed UpdateQuotationForAddedCosts
'------------------------------------------------------------------------------------------
Public Function CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement, _
                                        xmlLoanCompositionNode As IXMLDOMNode, _
                                        xmlAppNo As IXMLDOMNode, _
                                        xmlAppFactFindNo As IXMLDOMNode, _
                                        xmlLifeSubQuoteNumber As IXMLDOMNode, _
                                        xmlMortgageSubQuoteNumber As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function CreateNewQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecommendQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecostMortgageComponentAndLifeCover( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAndSaveBCSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateAffordability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RemodelQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function RecostQuotationForRateChange(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' BMIDS977
'Public Function UpdateQuotationForAddedCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
'End Function
'BMIDS815
Public Function RefreshMortgageAccountData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS815 - End

'BMIDS815
Public Function GetAndSavePortedStepAndPeriodFromMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS815 - End
Public Function CreateAdHocFeeAndRemodel(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ProcessAllFeesForRemodelQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

