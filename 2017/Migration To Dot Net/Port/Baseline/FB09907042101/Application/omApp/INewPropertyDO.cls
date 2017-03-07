VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INewPropertyDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      INewPropertyDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/00    Created
' SR    02/06/00    Added 'CreateNewPropertyDeposit' and 'UpdateNewPropertyDeposit'
' SR    15/06/00    Added GetFullNewPropertyDetails
'------------------------------------------------------------------------------------------

Option Explicit
Public Function GetNewPropertyAndAddressDetails( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetFullNewPropertyDetails( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNewPropertyGeneral( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetValuationTypeAndLocation( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNewPropertyDescription( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNewPropertyAddress( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub CreateNewPropertyAddress(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteNewPropertyAddress(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function GetVendorDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub DeleteVendorDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateVendorDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateVendorDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateNewProperty(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub CreateNewPropertyDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub
Public Sub CreateNewPropertyDeposit(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub UpdateNewProperty(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub UpdateNewPropertyDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub UpdateNewPropertyRoomDetails(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub UpdateNewPropertyAddress(ByVal vxmlTableElement As IXMLDOMElement)

End Sub
Public Sub UpdateNewPropertyDeposit(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Function GetPropertyUnknownDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

