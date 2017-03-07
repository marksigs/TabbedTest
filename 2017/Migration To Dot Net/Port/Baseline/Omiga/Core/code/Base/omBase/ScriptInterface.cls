VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ScriptInterface.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'RF    20/10/99 RunScript now returns a string
'------------------------------------------------------------------------------------------
Option Explicit
Implements IScriptInterface
Public Function RunScript(ByVal vstrComponentName As String, ByVal vstrObjectName As String, _
    ByVal vstrFunctionName As String, ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    ' fixme
    RunScript = "<RESPONSE TYPE=""SUCCESS""></RESPONSE>"
End Function
Private Function IScriptInterface_RunScript(ByVal vstrComponentName As String, _
                                            ByVal vstrObjectName As String, _
                                            ByVal vstrFunctionName As String, _
                                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set IScriptInterface_RunScript = xmlResponseNode.cloneNode(True)
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
End Function
