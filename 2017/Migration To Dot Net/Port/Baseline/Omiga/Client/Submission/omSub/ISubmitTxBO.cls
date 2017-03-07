VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubmitTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ISubmitTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     04/04/00    Created.
'------------------------------------------------------------------------------------------

Option Explicit

Public Function CreateSubmissionStatus( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateSubmissionStatus( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DoRiskAssessment( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DoCreditCheck( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DoOmiga4To3Download( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function


