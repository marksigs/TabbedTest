VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' SR    01/09/00    Modified DeleteCustomerAddress - Added an additional input param
'SR     05/06/01    SYS2325 New method 'UpdateCustomer'
'------------------------------------------------------------------------------------------
'BMids Specific History:
'
'Prog   Date        Description
'MV     12/08/2002  BMIDS00311 - Core Ref : SYS1705 Removed SafeRef ;Modified SavePersonalDetailsAndAddress
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'HM     08/09/2005  MAR29   SaveWrapUpDetails were added
'HMA    20/10/2005  MAR23   Add CreateKYCCheck
'----------------------------------------------------------------------------------------------------------

Option Explicit


Public Function CreateAlias(ByVal vXMLRequest As IXMLDOMElement, _
                            ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateAlias(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteAlias(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteAllCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveCustomerTelephone(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveWrapUpDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerAddress(ByVal vXMLRequest As IXMLDOMElement, _
                                      ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateDependant(ByVal vXMLRequest As IXMLDOMElement, _
                                ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerDetailsAndEmpt( _
              ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdatePersonalDetailsAndEmpt( _
              ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteDependant(ByVal vXMLRequest As IXMLDOMElement, _
                                Optional ByVal rstrPersonGUID As Variant) As IXMLDOMNode

End Function

Public Function UpdateDependant(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerAddress(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteCustomerAddress(ByVal vXMLRequest As IXMLDOMElement, _
                                Optional ByVal blnDeleteMortgageAccount As Boolean = False) As IXMLDOMNode

End Function

Public Function UpdatePersonalDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerVersion(ByVal vXMLRequest As IXMLDOMElement, Optional ByVal bTaxDetailsOnly As Boolean = False) As IXMLDOMNode

End Function

Public Function SaveAlias(ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveTenancy(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveDependant(ByVal vXMLRequest As IXMLDOMElement, _
                              ByVal vxmlGeneratedKeys As IXMLDOMElement, _
                              Optional ByRef rbotOperationType As Variant) As IXMLDOMNode

End Function

Public Sub CreateCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, _
                                 ByVal vobjResponseNode As IXMLDOMNode)
                                
End Sub

Public Sub UpdateCustomerDetails(ByVal vobjXmlRequest As IXMLDOMNode, _
                                 ByVal vobjXmlResponse As IXMLDOMNode)
                        
End Sub

Public Function CreateFullCustomerDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function


Public Function SavePersonalDetailsAndAddress(ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function


Public Function CreatePropertyInsuranceDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdatePropertyInsuranceDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateContactHistory(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateContactHistory(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomer(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

'BMIDS00004 MDC 17/05/2002
Public Function CreateNewCustomerVersion(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00004 MDC 17/05/2002 - End

'MAR23
Public Function CreateKYCCheck(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
End Function

