VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IThirdPartyDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'GD ADDED

Public Function FindPanelValuerList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Sub CreateDirectory(ByVal vxmlTableElement As IXMLDOMElement, _
                           ByRef rxmlKeysNode As IXMLDOMNode)
End Sub

Public Sub CreateMortgageLenderDirectory(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateThirdParty(ByVal vxmlTableElement As IXMLDOMElement, _
                            ByRef rxmlKeysNode As IXMLDOMNode, _
                            Optional ByVal blnCreateDirectory As Boolean = False)
End Sub

Public Sub DeleteThirdParty(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function FindDirectoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindThirdPartyList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDirectorySummaryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMainMortgageLenderList() As IXMLDOMNode
End Function

Public Function GetThirdParty(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetThirdPartyMinimal(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetDirectoryGUID(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDirectory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDirectoryMinimal(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetDirectorySummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Sub UpdateThirdParty(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function GetPaymentDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

' PSC 03/12/2002 BM0105
Public Function GetDirectoryDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function


