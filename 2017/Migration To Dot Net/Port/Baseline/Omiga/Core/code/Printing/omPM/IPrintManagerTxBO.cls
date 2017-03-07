VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPrintManagerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IPrintManagerTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Created:       26/06/2001
'Author:        Mike Kennedy
'Description:   COM Interface for the TxBO of the Print Manager component
'Dependencies:  PrintManagerTxBO.cls
'
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MK     26/06/01    Created
'------------------------------------------------------------------------------------------
Option Explicit
' Creates new DocumentDetail and EventDetail records
Public Sub PutAuditTrail( _
    ByVal vxmlAuditTrailDataNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
' Creates a summary list pf documents printed for the requested ApplicationNumber and any
' additional search criteria
Public Sub GetDocumentHistoryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
' Creates a summary list of document events for the requested search criteria
Public Sub GetDocumentEventList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
