VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRiskAssessmentDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetRiskAssessment(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function GetRiskAssessmentRule(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function

Public Function GetLatestRiskAssessment(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Sub CreateRuleOverride(ByVal vxmlTableElement As IXMLDOMElement)

End Sub

Public Sub UpdateRAUnderwriterDecision(ByVal vxmlTableElement As IXMLDOMElement)

End Sub


Public Function GetUserRiskAssessmentAuthority(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function


Public Function GetPreviousOverrideReasons(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function GetRAApplicationStages(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function StoreRiskAssResults( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByVal vobjXmlRulesResponse As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetRiskAssessmentDecisionParameter( _
    ByVal vintFailCount As Integer) _
    As String

End Function

Public Function GetLatestRiskAssessmentScore(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

' SR 09/08/2004 : BMIDS815
Public Function CheckRiskAssessmentExistsForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
