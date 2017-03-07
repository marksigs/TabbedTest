VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationManagerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function CreateLock(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteLock(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerLockApplicationLock( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteCustomerLockApplicationLock( _
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
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateCustomerRole(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteCustomerRole(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UpdateCustomerRole(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function UnlockApplicationAndCustomers( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function DeleteAllCustomerLockApplicationLock( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateFullApplicationDetails( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveCustomerAndEmploymentDetails( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ImportAccountsIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ImportCustomersIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ImportAccountRefreshIntoApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
