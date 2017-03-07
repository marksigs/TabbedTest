VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IExperianDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Function DelphiData(ByVal strUserId As String, ByVal strUnitId As String, ByVal strApplicationNumber As String) As Boolean
Public Function DelphiData(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function DelphiDemoData(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function RunCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function RunFullBureau(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FullBureauDemoData(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FullBureauData(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
End Function

'BMIDS748
Public Function RunUpgradeToFullAppCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function
