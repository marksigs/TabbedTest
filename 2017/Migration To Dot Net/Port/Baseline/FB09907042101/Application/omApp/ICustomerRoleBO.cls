VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerRoleBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ICustomerRoleBOMethod
    bomCreateCustomerRole
    bomUpdateCustomerRole
End Enum
Public Function Create(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
Public Function Update(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
Public Function Delete(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetData(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function Validate(ByVal vxmlRequest As IXMLDOMElement, _
                         ByVal vbomMethod As ICustomerRoleBOMethod) As IXMLDOMNode
End Function
Public Function FindList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

