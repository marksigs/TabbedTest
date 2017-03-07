VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICreditCheckBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' TODO ************************************************************************
'   Amend 'Omiga4BO' with actual class name
'   Amend enums to be the Create/Update type methods
' *****************************************************************************
'INR 12/08/2004 BMIDS744 Added GetCreditCheckStatus
'HM  31/08/2004 MAR13 Added GetCurrentXMLCreditCheckResults
'SR  20/10/2005 MAR24 Added FindApplicationCreditCheckDetailsList
'INR 23/11/2005  MAR645 Added CreateRuleOverride

Public Enum ICreditCheckBOMethod
    bomCreate
    bomUpdate
End Enum

Public Function GetCurrentCreditCheckResults(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetCurrentXMLCreditCheckResults(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetCurrentCAISSummary(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function RunCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function RunFullBureau(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function GetDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function UpdateDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

Public Function CheckCreditCheckResponse(ByVal vxmlResponseToCheck As IXMLDOMElement, _
                              ByVal vxmlResponseToAddTo As IXMLDOMElement, _
                              Optional ByVal vblnRaiseError As Boolean = False) As Long

End Function

Public Function GetFullBureauDataHeaderData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function FindFullBureauVotersRollList(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function FindFullBureauPublicInfoList(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function FindFullBureauCIFASList(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function FindFullBureauCAPSList(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function FindFullBureauAssociationsList(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function FindFullBureauCorrectionLinesList(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentVotersRollSummary(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentPublicInfoSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentCIFASSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentCAPSSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentAliasSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentFullBureauCAISSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCurrentFullBureauResults(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function FindApplicationCreditCheckList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetFullBureauStandardHeaderData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetFullBureauRequest(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function ValidateFBMandatoryData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

Public Function GetCCAndFBStatus(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
                         
End Function

'BMIDS00336 MDC 14/08/2002
Public Function GetAndDeduplicateBureauRecords(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function RunCreditScore(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function
'BMIDS00336 MDC 14/08/2002 - End

'BMIDS744
Public Function GetCurrentCreditCheckStatus(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

'BMIDS748
Public Function RunUpgradeToFullAppCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

End Function

'SR 18/10/2006 : MAR24
Public Function FindApplicationCreditCheckDetailsList(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

'INR 23/11/2005  MAR645
Public Function CreateRuleOverride(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode

End Function

