VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAuditBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum IAuditBOMethod
    bomCreateAccessAudit
End Enum
Public Function CreateAccessAudit(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function Validate(ByVal vxmlRequest As IXMLDOMElement, _
                         ByVal vbomMethod As IAuditBOMethod) As IXMLDOMNode
End Function
