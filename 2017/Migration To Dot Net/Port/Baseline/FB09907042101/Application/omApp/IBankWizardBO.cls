VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBankWizardBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'HISTORY
'---------------------------------------------------
'SG     27/05/02    SYS4767 MSMS to Core integration
'

Option Explicit

Public Function GetBankDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateBankDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'SG 27/05/02 SYS4767
Public Function GetBankWizardDetails(ByVal vxmlRequestNode As IXMLDOMElement) As IXMLDOMNode
End Function

