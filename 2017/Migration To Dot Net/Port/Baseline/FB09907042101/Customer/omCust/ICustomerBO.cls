VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SR     05/06/01    SYS2325 New method 'UpdateCustomer'
'----------------------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    28/06/2005  MAR5 - Added GetNumberOfCopiesForKFI
'HM     08/09/2005  MAR29   SaveWrapUpDetails were added
'PSC    17/10/2005  MAR57 - Added ProcessLaunchCustomer
'HMA    20/10/2005  MAR23  Added GetCustomerNumbers. Added CreateKYCCheck.
'------------------------------------------------------------------------------------------

Option Explicit

Public Enum ICustomerBOMethod
    bomSavePersonalDetailsAndAddress
    bomCreatePropertyInsuranceDetails
    bomUpdatePropertyInsuranceDetails
    bomCreateContactHistory
    bomUpdateContactHistory
End Enum

Public Function FindContactHistoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetContactHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateContactHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateContactHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateAlias(ByVal vxmlRequest As IXMLDOMElement, _
                            ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateCustomerDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetCustomerFromLegacyNumber(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetAlias(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindAliasList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerDetailsAndEmpt(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdatePersonalDetailsAndEmpt(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateCustomerVersion(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveWrapUpDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteAllCustomerTelephone(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement, _
                                      ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindCustomerAddressList(ByVal vxmlRequest As IXMLDOMElement, _
                    Optional ByVal blnOrderBySequenceNUmber As Boolean = False) As IXMLDOMNode

End Function

Public Function GetCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateDependant(ByVal vxmlRequest As IXMLDOMElement, _
                            ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateDependant(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetDependant(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteDependant(ByVal vxmlRequest As IXMLDOMElement, _
                                Optional ByVal rstrPersonGUID As Variant) As IXMLDOMNode

End Function

Public Function FindDependantList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateDependant(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveDependant(ByVal vxmlRequest As IXMLDOMElement, _
                              ByVal vxmlGeneratedKey As IXMLDOMElement, _
                              Optional ByRef rbotOperationType As Variant) As IXMLDOMNode

End Function

Public Function ValidateTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function DeleteTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteCustomerAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement

End Function

Public Function SaveCustomerAddress(ByVal vxmlRequest As IXMLDOMElement, _
                              ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetCustomerPersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerPersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdatePersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerVersion(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveAlias(ByVal vxmlRequest As IXMLDOMElement, _
                          ByVal xmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Sub FindCustomer(ByVal vobjXmlRequest As IXMLDOMNode, _
                        ByVal vobjXmlResponse As IXMLDOMNode)
                        
End Sub

Public Sub GetCustomerDetails(ByVal vobjRequestNode As IXMLDOMNode, _
                                ByVal vobjResponseNode As IXMLDOMNode)
                                
End Sub

Public Sub GetAndSynchroniseCustomerDetails _
    (ByVal vobjRequestNode As IXMLDOMNode, _
     ByVal vobjResponseNode As IXMLDOMNode)
                                
End Sub

Public Sub CreateCustomerLock(ByVal vobjXmlRequestNode As IXMLDOMNode, _
                              ByVal vobjXmlResponseNode As IXMLDOMNode, _
                              Optional ByVal strLockRequired As String = "")
                                
End Sub

Public Sub DeleteCustomerLock(ByVal vobjXmlRequest As IXMLDOMNode, _
                              ByVal vobjXmlResponse As IXMLDOMNode, _
                              Optional ByVal strLockRequired As String = "")
                        
End Sub

Public Function GetComponentData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateFullCustomerDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetPersonalDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SavePersonalDetailsAndAddress(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function Validate(ByVal vxmlRequest As IXMLDOMElement, ByVal veboMethod As ICustomerBOMethod) As IXMLDOMNode

End Function

Public Function SaveTenancy(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveTaxDetails(ByVal vxmlRequest As IXMLDOMElement, _
                               Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function GetPersonalDetailsAndAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreatePropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdatePropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetPropertyInsuranceDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

'BMIDS00004 MDC 13/05/2002
Public Function CreateNewCustomerVersion(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00004 MDC 13/05/2002 - End

'BMIDS00005 CRWP2 BM065 GHun 28/05/2002
Public Function GetCRSContactData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateCRSContactLog(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00005 End

'BMIDS00026 IWP2 BM060 GHun 07/06/2002
'BMIDS00026 IWP2 BM060 GHun 05/07/2002 Changed from Private to Public following code review
Public Function UpdateCRSCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00026 End

'MO     21/08/2002  BMIDS00218 - Start
Public Function GetNextCustomerNumber() As IXMLDOMNode
End Function
'MO     21/08/2002  BMIDS00218 - End
'BMIDS758  Add function
Public Function GetRemovedToECustomers(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'PSC 28/06/2005 MAR5 - Start
Public Function GetNumberOfCopiesForKFI(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'PSC 28/06/2005 MAR5 - End
'PSC 17/10/2005 MAR57 - Start
Public Function ProcessLaunchCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'PSC 17/10/2005 MAR57 - End

'MAR23
Public Function GetCustomerNumbers(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateKYCCheck(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

