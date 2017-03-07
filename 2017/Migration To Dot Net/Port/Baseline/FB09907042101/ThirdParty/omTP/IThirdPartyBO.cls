VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IThirdPartyBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'GD ADDED

Public Function FindPanelValuerList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateMortgageLenderDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDirectoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindDirectorySummaryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindMainMortgageLenderList() As IXMLDOMNode
End Function

Public Function FindThirdPartyList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetDirectoryGUID(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetPaymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

' PSC 03/12/2002 BM0105
Public Function GetDirectoryDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function



