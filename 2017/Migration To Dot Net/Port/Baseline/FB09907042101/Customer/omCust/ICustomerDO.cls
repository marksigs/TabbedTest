VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 'SR    01/09/00    Modified DeleteCustomerAddress - Added an additional input param
 'SR     05/06/01    SYS2325 New method 'UpdateCustomer'
'----------------------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    28/06/2005  MAR5 - Add new function GetCustomerAddressListAndRole
'HM     08/09/2005  MAR29  Create/Update/DeleteWrapUpDetails were added
'PSC    10/10/2005  MAR57 - Add new function FindCustomerCRSData
'HMA    20/10/2005  MAR23  Add GetCustomerNumbers and CreateKYCCheck
'------------------------------------------------------------------------------------------

Option Explicit



Public Function CreateAlias(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Sub UpdateAlias(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteAlias(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetAlias(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function FindAliasList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Sub CreateCustomerTelephone(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub CreateWrapUpDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub DeleteAllCustomerTelephone(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateCustomerTelephone(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteCustomerTelephone(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub UpdateWrapUpDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteWrapUpDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Function FindCustomerTelephoneList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function CreateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function FindCustomerAddressList(ByVal vxmlTableElement As IXMLDOMElement, _
                Optional ByVal blnOrderBySequenceNUmber As Boolean = False) As IXMLDOMElement

End Function

Public Function GetCustomerAddress(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function CreateDependant(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function GetDependant(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function DeleteDependant(ByVal vxmlRequest As IXMLDOMElement, _
                           Optional ByRef rstrPersonGUID As Variant) As IXMLDOMElement

End Function

Public Function FindDependantList(ByVal vxmRequest As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function UpdateDependant(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function UpdateCustomerVersion(ByVal vxmlTableElement As IXMLDOMElement, Optional ByVal bTaxDetailsOnly As Boolean = False) As IXMLDOMElement

End Function

Public Sub CreateTenancy(ByVal vxmlRequest As IXMLDOMElement)

End Sub

Public Sub UpdateTenancy(ByVal vxmlRequest As IXMLDOMElement)

End Sub

Public Sub DeleteTenancy(ByVal vxmlRequest As IXMLDOMElement)

End Sub

Public Function bGetKeysOfCustomerAddress(ByVal xmlElement As IXMLDOMElement, _
                                          ByRef strCustomerNumber As String, _
                                          ByRef strCustomerVersionNumber As String, _
                                          ByRef strCustomerAddressSequenceNumber As String) As Boolean

End Function

Public Function GetTenancy(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub UpdateCustomerAddress(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdatePersonalDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteCustomerAddress(ByVal vxmlTableElement As IXMLDOMElement, _
                                 Optional ByVal blnDeleteMortgageAccount As Boolean = False)

End Sub

Public Function GetPersonalDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function GetComponentData(ByVal vxmlRequest As IXMLDOMElement) _
                                            As IXMLDOMElement

End Function

Public Sub FindCustomer(ByVal vobjXmlRequest As IXMLDOMNode, _
                        ByVal vobjXmlResponse As IXMLDOMNode)
                        
End Sub

Public Sub GetCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, _
                              ByVal vobjResponseNode As IXMLDOMNode)
                                
End Sub

Public Sub CreateCustomerLock(ByVal vobjXmlRequestNode As IXMLDOMNode, _
                              ByVal vobjXmlResponseNode As IXMLDOMNode, _
                              Optional ByVal vstrLockType As String = "On")
                                
End Sub

Public Sub DeleteCustomerLock(ByVal vobjXmlRequest As IXMLDOMNode, _
                              ByVal vobjXmlResponse As IXMLDOMNode)
                        
End Sub

Public Function GetNextCustomerNumber() As Long

End Function

Public Sub CreateCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, _
                                 ByVal vobjResponseNode As IXMLDOMNode, _
                                 ByVal vstrCustomerNumber As String, _
                                 ByVal vstrCustomerVersionNumber As String, _
                                 Optional ByVal vblnCreateNewCustomer As Boolean = True)
End Sub

Public Sub UpdateCustomerDetails(ByVal vobjXmlRequest As IXMLDOMNode, _
                                 ByVal vobjXmlResponse As IXMLDOMNode)
                        
End Sub

Public Sub CustomerDetailsDerivedData(ByVal vobjXmlCustomerNode As IXMLDOMNode)
                                      
End Sub

Public Sub CreatePropertyInsuranceDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateCurrentProperty(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub CreateHomeInsurance(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateCurrentProperty(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateHomeInsurance(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdatePropertyInsuranceDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetPropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function FindContactHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetContactHistory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function CreateContactHistory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub UpdateContactHistory(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetCustomerFromLegacyNumber(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Sub UpdateCustomer(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

'BMIDS00005 CRWP2 BM065 GHun 29/05/2002
Public Function FindContactsPerApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00005 End

'BMIDS00026 IWP2 BM060 GHun 07/06/2002
Public Function FindCustomersPerApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00005 End

'BMIDS00443 GHun 04/10/2002
Public Function FindLatestCustomerVersion(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00443 End

'BMIDS758  Add function
Public Function GetRemovedToECustomers(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMElement
End Function

'PSC 28/06/2005 MAR5 - Start
Public Function GetCustomerAddressListAndRole(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
End Function
'PSC 28/06/2005 MAR5 - End
'PSC 10/10/2005 MAR57 - Start
Public Function FindCustomerCRSData(ByVal vxmlTableElement As IXMLDOMNode) As IXMLDOMNode
End Function
'PSC 10/10/2005 MAR57 - End

'MAR23
Public Function GetCustomerNumbers(ByVal vobjXmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub CreateKYCCheck(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

