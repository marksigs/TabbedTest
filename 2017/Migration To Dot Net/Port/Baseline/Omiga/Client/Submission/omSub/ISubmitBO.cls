VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubmitBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ISubmitBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'RF     27/03/00    Created
'MC     30/06/00    Added GetSubmissionStatusForApplication method
'------------------------------------------------------------------------------------------
Option Explicit

Public Enum ISubmitBOMethod
    bomCreateSubmissionStatus
    bomUpdateSubmissionStatus
End Enum

Public Function DoCreditCheck(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DoRiskAssessment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DoOmiga4To3Download(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateSubmissionStatus(ByVal vxmlRequest As IXMLDOMElement) _
    As IXMLDOMNode
End Function

Public Function Validate(ByVal vxmlRequest As IXMLDOMElement, _
                         ByVal vbomMethod As ISubmitBOMethod) As IXMLDOMNode
End Function

Public Function UpdateSubmissionStatus(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetSubmissionStatusForApplication( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function


