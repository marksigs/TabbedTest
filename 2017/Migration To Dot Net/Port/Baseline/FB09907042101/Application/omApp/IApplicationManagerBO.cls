VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum IApplicationManagerBOMethod
    bomApplicationManagerDefault
    bomCreateApplicationLock
    bomUpdateApplicationLock
    bomCreateCustomerLockApplicationLock
    bomCreateCustomerRole
    bomUpdateCustomerRole
    bomAddCustomerToApplication
    bomDeleteCustomerFromApplication
    bomCreateFullApplicationDetails
    bomImportAccountsIntoApplication
End Enum

Public Function CreateLock(ByVal vxmlRequest As IXMLDOMElement, _
                           Optional ByVal strLockRequired As String = "") As IXMLDOMNode
' strLockRequired : Whether Lock is available for the system
End Function

Public Function DeleteLock(ByVal vxmlRequest As IXMLDOMElement, _
                           Optional ByVal strLockRequired As String = "") As IXMLDOMNode

End Function

Public Function Validate( _
            ByVal vxmlRequest As IXMLDOMElement, _
            Optional ByVal veboMethod As IApplicationManagerBOMethod _
                    = bomApplicationManagerDefault) As IXMLDOMNode


End Function

Public Function GetLockData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerLockApplicationLock( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            Optional ByVal strLockRequired As String = "") As IXMLDOMNode

End Function

Public Function DeleteCustomerLockApplicationLock( _
                        ByVal vxmlRequest As IXMLDOMElement, _
                        Optional ByVal strLockRequired As String = "") As IXMLDOMNode

End Function

Public Function FindCustomerLockApplicationLockList( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function MaintainCustomerRoles( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerRoles( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function AddCustomerToApplication( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteCustomerFromApplication( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function LockCustomersForApplication( _
                        ByVal vxmlRequest As IXMLDOMElement, _
                        Optional ByVal strLockRequired As String = "") As IXMLDOMNode

End Function

Public Function CreateCustomerRole( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteCustomerRole( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerRole( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetCustomerRoleData( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindCustomerRoleList( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UnlockApplicationAndCustomers( _
                        ByVal vxmlRequest As IXMLDOMElement, _
                        Optional ByVal strLockRequired As String = "") As IXMLDOMNode

End Function

Public Function DeleteAllCustomerLockApplicationLock( _
                            ByVal vxmlRequest As IXMLDOMElement, _
                            Optional ByVal strLockRequired As String = "") As IXMLDOMNode

End Function

Public Function CreateFullApplicationDetails( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindLegacyBusinessForCustomer( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveCustomerAndEmploymentDetails( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindApplicationList( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindApplicationOwnershipList( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindBusinessForCustomer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ImportAccountsIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetApplicationSummaryData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ImportCustomersIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
Public Function GetAdditionalRequestDataForAccDownload(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ImportAccountRefreshIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
