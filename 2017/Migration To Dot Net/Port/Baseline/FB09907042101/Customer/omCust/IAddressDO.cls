VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAddressDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub CreateAddress(ByVal vxmlTableElement As IXMLDOMElement, _
                          ByRef rxmlKeysNode As IXMLDOMNode)
End Sub

Public Sub DeleteAddress(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function GetAddress(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function IsEmptyAddress(ByVal vxmlElement As IXMLDOMElement) As Boolean
End Function

Public Sub UpdateAddress(ByVal vxmlTableElement As IXMLDOMElement)
End Sub





