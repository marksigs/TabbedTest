VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IScriptInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function RunScript(ByVal vstrComponentName As String, _
                          ByVal vstrObjectName As String, _
                          ByVal vstrFunctionName As String, _
                          ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
