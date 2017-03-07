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
'BMIDS Specific History :
'
'Prog   Date        AQR         Desctription
'MV     14/08/2002  BMIDS00333  Core Ref AQR: SYS49419 ;Modified GetApplicationData()
'MDC    22/08/2002  BMIDS00336  New methods: RunExperianCreditCheck & ImportBureauData
'PSC    06/01/2003  BM0089      New Methods: MoveToStage, MoveToNextStage & ProcessAutomaticTasks
'GHun   10/03/2003  BM0376      CC026 Added method RunIncomeCalcs
'MDC    24/04/2003  BM0340      Added MoveCaseToNewStage
'SDS    10/12/2004  BBG1741     Task Management - Move Stage is taking a long time when performed against a large DB
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
Public Sub RunExperianCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode)
                                    
End Sub
Public Sub ImportBureauData(ByVal vxmlRequestNode As IXMLDOMNode)
                                    
End Sub
'BMIDS00336 MDC 22/08/2002 - End
'PSC 06/01/2003 BM0089 - Start
'SDS  BBG1741  10/12/2004__START
'Public Function MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'End Function
'Public Function MoveToStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'End Function
'SDS  BBG1741  10/12/2004__END
Public Sub ProcessAutomaticTasks(ByVal vxmlRequest As IXMLDOMNode)
End Sub
'PSC 06/01/2003 BM0089 - End
'BM0376 GHun 10/03/2003
Public Sub RunIncomeCalcs(ByVal vxmlRequestNode As IXMLDOMNode)
End Sub
'BM0376 End
'BM0340 MDC 24/04/2003
Public Sub MoveCaseToNewStage(ByVal vxmlApplicationRequestDoc As FreeThreadedDOMDocument40, _
                                ByVal vxmlCreateCaseStageDoc As FreeThreadedDOMDocument40, _
                                ByVal vxmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40)
End Sub
'BM0340 MDC 24/04/2003 - End
