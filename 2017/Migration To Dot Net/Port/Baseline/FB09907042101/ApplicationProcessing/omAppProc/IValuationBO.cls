VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IValuationBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     03/09/2002  BMIDS00353  Added New Method GetLatestValuationAmounts()
'----------------------------------------------------------------------------------------------
' INGDUK
'JD     28/10/2005  MAR342       Added new method CreateValuationReportNoInst
'----------------------------------------------------------------------------------------------

Public Sub GetLatestValuationAmounts(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Public Sub GetValuationReport( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetValuationReportForTemplate( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateValuationReport( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateValuationReportUsingTemplate( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateValuationReportNoInst( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateValuationReport( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ValidateValuationReport( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ValidateInvoiceAmount( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetValuationPropertyDetails( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindValuerInstructionList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub GetValuationFee( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindInstructionsForBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ProcessValuerBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateValuationAudit( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindFailedInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ProcessValuationInstruction( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub ProduceValuerOutput( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

