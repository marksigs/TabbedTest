VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IomMQTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IomMQTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Issues:        Instancing:         Public not creatable
'               MTSTransactionMode: Not an MTS object
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SF&PD  09/10/01    OmiPlus30: Created
'AD     30/10/01    Restructed the component so that it conforms to the Omiga 4 Phase 2
'                   architecture. To maintain compatibility the message queue class is
'                   still called from the DO.
'RF     01/03/02    SYS4215 omMQ.SendToQueue is producing "method ~ of object ~ failed"
'                   in release. Various improvements identified.
'------------------------------------------------------------------------------------------
Option Explicit

Public Function SendToQueue(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
End Function







