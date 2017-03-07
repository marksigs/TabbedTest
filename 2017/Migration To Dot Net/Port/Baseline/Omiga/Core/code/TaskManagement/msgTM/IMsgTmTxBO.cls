VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "IMsgTmTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        Description
'MV     31/01/2005  BBG1908 - Added CreateCaseTrackNode and UpdateCaseTrackNote
'MV     01/08/2005  BBG2146 - New method added - CreateTaskNoteArchive()
'------------------------------------------------------------------------------
Option Explicit
Public Sub CreateCaseActivity(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateCaseStage(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateCaseTrackNote(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateCaseTask(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub CreateTaskNoteArchive(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateTaskNote(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateCaseActivity(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseStage(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseTask(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub UpdateCaseTrackNote(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub BaseCreate(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Public Sub BaseUpdate(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
