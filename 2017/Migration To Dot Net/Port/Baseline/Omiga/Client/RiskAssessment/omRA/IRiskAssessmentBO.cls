VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRiskAssessmentBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetRiskAssessment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetLatestRiskAssessment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function ValidateRuleOverride(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CreateRuleOverride(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetUserRiskAssessmentAuthority(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function
'DRC MAR33 response must be call by ref
Public Sub RunRiskAssessment( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByRef vobjXmlResponse As IXMLDOMNode)

End Sub

Public Function GetLatestRiskAssessmentScore(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
' SR 09/08/2004 : BMIDS815
Public Function CheckRiskAssessmentExistsForApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'MAR1315
Public Function GetLatestRANoStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

