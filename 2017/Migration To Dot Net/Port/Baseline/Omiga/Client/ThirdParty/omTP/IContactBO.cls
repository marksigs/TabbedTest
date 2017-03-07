VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContactBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DetermineContactOperation(ByVal vxmlElement As IXMLDOMElement) As Integer
End Function

Public Function SaveContact(ByVal vxmlElement As IXMLDOMElement, _
                                        Optional ByRef rbotOperationType As Integer = 0) As IXMLDOMNode
End Function

Public Function CreateContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateContact(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function



