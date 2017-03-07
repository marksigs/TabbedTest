VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INewPropertyTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      INewPropertyTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/00    Created
' BG    26/05/00    Added SaveNewPropertyAndAddress
' SR    02/06/00    Added SaveNewPropertyAndDeposit
' SR    13/06/00    New method 'SaveNewProperty'
' MC    13/07/00    New method SaveNewPropertyAddress
'------------------------------------------------------------------------------------------

Option Explicit

Public Function CreateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewPropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateNewProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateNewPropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewPropertyAndAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewPropertyAndDeposit(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS887 GHun
Public Function SaveNewPropertyAndUpdateFactFind(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS887 End
