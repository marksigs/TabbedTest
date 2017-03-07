VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICreditCheckDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'INR 12/08/2004 BMIDS744 Added GetCreditCheckStatus
'HM  31/08/2004 MAR13 Added GetCurrentXMLCreditCheckResults
'SR  20/10/2005 MAR24 Added FindApplicationCreditCheckDetailsList
'INR  23/11/2005 MAR645 Added CreateRuleOverride, GetPreviousOverrideReasons

Option Explicit

Public Function GetCurrentCreditCheckResults(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetCurrentXMLCreditCheckResults(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetCurrentCAISSummary(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function UpdateDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function GetDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauVotersRollList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauPublicInfoList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauCIFASList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauCAISList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauCAPSList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauAssociationsList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function FindFullBureauCorrectionLinesList(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Function GetFullBureauDataHeaderData(ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function FindApplicationCreditCheckList(ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetFullBureauStandardHeaderData(ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function FindCurrentFullBureauResultsList(ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetFullBureauRequest(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

End Function
'BMIDS744
Function GetCreditCheckStatus(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function
'SR 18/10/2005 : MAR24
Public Function FindApplicationCreditCheckDetailsList(ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

End Function
'INR 23/11/2005 : MAR645
Public Function CreateRuleOverride(ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

End Function
