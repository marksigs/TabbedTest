VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICostModelAssistBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------
'BMIDS Specific History :
'
'Prog   Date        AQR         Description
'MV     10/10/2002  BMIDS00590  Amended ProcessAddedOneOffCosts
'----------------------------------------------------------------------------------
'Epsom History:
'
' Prog  Date     Description
' PSC   01/12/06 EP2_249 Added new method GetIntroducerProcFeeDetails
'-------------------------------------------------------------------------------------------------------

Public Function GetBorrowingAndLenderName(ByVal vstrXMLRequest As String) As IXMLDOMNode
End Function

Public Function FindMortgageProducts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement, _
                                        xmlLoanCompositionNode As IXMLDOMNode, _
                                        xmlAppNo As IXMLDOMNode, _
                                        xmlAppFactFindNo As IXMLDOMNode, _
                                        xmlLifeSubQuoteNumber As IXMLDOMNode, _
                                        xmlMortgageSubQuoteNumber As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function GetLoansNotRedeemed(vxmlRequest As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function CalcCostModelLTV(ByVal vxmlRequest As IXMLDOMElement, _
                                 Optional ByVal vstrContext As String = "QUICK QUOTE") As IXMLDOMNode
End Function

Public Function BuildLifeCoverCalcInputInterface(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function BuildLifeCoverInterfaceXML(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CalcAndSaveLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement, _
                                   Optional ByVal strContext = "QUICK QUOTE") As IXMLDOMElement
End Function

Public Function RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
End Function

Public Function CalculateAndSaveBCSubQuote( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            ByVal strCustNumber1 As String, _
                            ByVal strCustNumber2 As String, _
                            ByVal strCustVersionNumber1 As String, _
                            ByVal strApplNumber As String, _
                            ByVal strAFFNumber As String, _
                            Optional ByVal strContext As String = "QUICK QUOTE", _
                            Optional ByVal blnBakruptcyApplicant1 As Boolean, _
                            Optional ByVal blnBakruptcyApplicant2 As Boolean, _
                            Optional ByVal strDeclarationValue As String, _
                            Optional xmlNewProperty As IXMLDOMElement) As IXMLDOMElement
            
End Function

Public Function ProcessAddedOneOffCosts( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            Optional ByVal vstrContext As String = "QUICK QUOTE", _
                            Optional ByVal vstrLocation As String, _
                            Optional ByVal vstrValuationType As String, _
                            Optional ByVal vvstrLoansNotRedeemed As String, _
                            Optional ByVal vvstrSharedPercentage As String) As IXMLDOMElement

End Function

Public Function FindStoredQuoteDetails( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function UpdateMortgageSubquote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetLTV(ByVal vxmlRequest As IXMLDOMElement, _
                                    Optional ByVal strContext = "QUICK QUOTE") As IXMLDOMNode
End Function

' PSC 01/12/2006 EP2_249 - Start
Public Function GetIntroducerProcFeeDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' PSC 01/12/2006 EP2_249 - End

