VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAppSummaryBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IAppSummaryBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Application Summary Business Object Interface.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    26/02/01    SYS1879 Created
'------------------------------------------------------------------------------------------

Option Explicit

Public Sub FindApplicationReviewHistoryList(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateApplicationReviewHistory(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlResponseNode As IXMLDOMNode)
End Sub


Public Sub UpdateApplicationReviewHistory(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetLoanAndRisksData(ByVal vxmlRequestNode As IXMLDOMNode, _
                              ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ValidateUserLogon(ByVal vxmlRequestNode As IXMLDOMNode, _
                             ByVal vxmlResponseNode As IXMLDOMNode)
End Sub



