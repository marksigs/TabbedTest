VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRateChangeDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub FindApplications(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateRateChangeAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindRateChangeAuditList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub SetBaseRateAppliedDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindLoanComponentsForApplication(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindBatchApplicationList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindBaseRateDataforLoanComponent(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

'BM0230 MDC 09/01/2003
Public Sub FindFailedApplications(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'BM0230 MDC 09/01/2003 - End

'BMIDS622 GHun 18/08/2003 No longer required
''BM0367 GHun 20/02/2003
'Public Sub GetDatabaseServerDate(ByVal vxmlResponseNode As IXMLDOMNode)
'End Sub
''BM0367
'BMIDS622 End

'BMIDS622 GHun 18/08/2003
Public Sub PrepareRateChange(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
'BMIDS622 End
