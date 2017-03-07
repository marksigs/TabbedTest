VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAddressBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DetermineAddressOperation(ByVal vxmlElement As IXMLDOMElement) As Integer
End Function

Public Function IsEmptyAddress(ByVal vstrXMLRequest As String) As Boolean
End Function

Public Function SaveAddress(ByVal vxmlRequest As IXMLDOMElement, _
                                        Optional ByRef rbotOperationType As Integer = 0) As IXMLDOMNode
End Function

Public Function CreateAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function


