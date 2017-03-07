VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOmTmNTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------
'MARS Specific History :
'
'Prog   Date        Desctription
'PSC    25/09/2005  MAR32 Task Automation
'PSC    09/05/2006  MAR1643 Add LockApplication and UnlockApplication
'----------------------------------------------------------------------------------

Public Sub ForwardBORequest(ByVal vxmlRequestNode As IXMLDOMNode, _
                            ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Public Sub RunCompositeRequest(ByVal vxmlRequestNodes As Collection)

End Sub

Public Sub RunTmRequest( _
    ByVal vxmlRequestDoc As FreeThreadedDOMDocument40, _
    Optional ByVal vxmlResponseDoc As FreeThreadedDOMDocument40 = Nothing)
    
End Sub

Public Sub GetApplicationData( _
    ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
    ByVal vxmlApplicationDestDoc As FreeThreadedDOMDocument40, _
    Optional ByVal vstrRBTemplateName As String)

End Sub

'BMIDS00336 MDC 22/08/2002
'BMIDS682 Return AddressTargeting info
Public Function RunExperianCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode) As String
                                    
End Function
'BMIDS748
Public Function RunUpgradeToFullAppCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode) As String
                                    
End Function

Public Sub ImportBureauData(ByVal vxmlRequestNode As IXMLDOMNode)
                                    
End Sub
'BMIDS00336 MDC 22/08/2002 - End

'BMIDS697 GHun 22/01/2004 No longer used
''PSC 06/01/2003 BM0089 - Start
'Public Function MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'End Function
'BMIDS697 End
'
'Public Function MoveToStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'End Function
'BMIDS697 End

Public Sub ProcessTask(ByVal vxmlRequest As IXMLDOMNode)
End Sub
'PSC 06/01/2003 BM0089 - End


'BM0376 GHun 10/03/2003
Public Sub RunIncomeCalcs(ByVal vxmlRequestNode As IXMLDOMNode)
End Sub
'BM0376 End

'BM0340 MDC 24/04/2003
' PSC 22/09/2005 MAR32
Public Sub MoveCaseToNewStage(ByVal vxmlApplicationRequestDoc As FreeThreadedDOMDocument40, _
                              ByVal vxmlCreateCaseStageDoc As FreeThreadedDOMDocument40, _
                              ByVal vxmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40, _
                              Optional ByVal vxmlSetSLAExpiryDateDoc As FreeThreadedDOMDocument = Nothing)

End Sub
'BM0340 MDC 24/04/2003 - End

'PSC 25/08/2005 MAR32 - start
Public Function CreateAdhocCaseTask(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function

Public Sub ProcessTASTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub PerformTASTaskProcessing(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub TASMoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateCaseTask(ByVal vxmlRequestNode As IXMLDOMNode)
End Sub

Public Sub SendMessageToQueue(ByVal vxmlRequestNode As IXMLDOMNode, _
                               ByVal vxmlMessageToSend As IXMLDOMNode, _
                               vstrQueueName As String, _
                               Optional ByVal lngDelayInterval As Long = 0)
End Sub
                               
Public Sub PerformTASMoveToNextStageProcessing(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
                               
'PSC 25/08/2005 MAR32 - End

'PSC 09/05/2006 MAR1643 - Start
Public Sub LockApplication(ByVal xmlTempRequest As IXMLDOMNode, ByVal strAppNo As String, ByVal strAppFFNo As String)
End Sub

Public Sub UnlockApplication(ByVal xmlTempRequest As IXMLDOMNode, ByVal strAppNo As String)
End Sub
'PSC 09/05/2006 MAR1643 - Start






