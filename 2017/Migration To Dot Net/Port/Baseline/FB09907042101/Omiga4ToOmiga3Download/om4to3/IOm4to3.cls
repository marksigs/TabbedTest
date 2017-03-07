VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOmiga4ToOmiga3BO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MH 15/06/00 vxmlPerformance is used for Tuning and analysis.
' it is only available for use if compiled with the right flags'
' the node is an attribute-only node whose contents will vary
Public Function Download(ByVal vxmlRequest As IXMLDOMElement, _
                Optional ByVal vxmlPerformance As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub AsyncDownload(ByVal vxmlRequest As IXMLDOMElement, ByVal xmlDataNode As IXMLDOMNode, vxmlResponse As IXMLDOMElement, _
                Optional ByVal vxmlPerformance As IXMLDOMElement)
End Sub

Public Sub CallTargetSystem(ByVal vxmlRequest As IXMLDOMElement, ByVal xmlDataNode As IXMLDOMNode, vxmlResponse As IXMLDOMElement)

End Sub
