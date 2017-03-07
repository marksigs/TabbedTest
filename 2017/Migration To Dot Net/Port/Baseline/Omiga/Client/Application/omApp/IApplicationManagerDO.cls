VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplicationManagerDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub CreateLock(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteLock(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetLockData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Sub CreateCustomerLockApplicationLock(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteCustomerLockApplicationLock(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindCustomerLockApplicationLockList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Sub CreateCustomerRole(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub DeleteCustomerRole(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateCustomerRole(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetCustomerRoleData( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function FindCustomerRoleList( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function FindMainApplicants( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function FindOmigaBusinessForCustomer(ByVal vstrXMLRequest As String) As IXMLDOMElement

End Function

Public Sub DeleteAllCustomerLockApplicationLock( _
                    ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function FindLegacyBusinessForCustomer(ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMElement

End Function

Public Function FindApplicationList(ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMElement

End Function

Public Sub UnlockApplicationAndCustomers(ByVal vxmlRequest As IXMLDOMNode)
End Sub

Public Sub LockCustomersForApplication(ByVal vxmlRequest As IXMLDOMNode)
End Sub

'BMIDS758
Public Sub CreateRemovedToECustomer(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
