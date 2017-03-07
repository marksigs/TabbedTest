VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "CompRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CompRulesBO.cls
'Copyright:     Copyright © 2006 Vertex Financial Services

'Description:   Completeness Checks Business Object.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     07/04/2006  EPSOM re-write as .xslt
'------------------------------------------------------------------------------------------
Option Explicit

Public Function RunRules(ByVal vxmlRequest As String) As String
    
    On Error GoTo RunRulesVbErr
    
    Const strFunctionName As String = "RunRules"

    Dim xmlRequestDoc As MSXML2.DOMDocument40
    Dim xmlResponseDoc As MSXML2.DOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlRequestDoc = New DOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    xmlRequestDoc.loadXML vxmlRequest
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
            
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlResponseDoc.appendChild(xmlResponseElem)
    
    'Run completeness check data rules (contained in xslt file) against XML request doc
    xmlTransform xmlRequestDoc, xmlResponseDoc, "ccRules.xslt"

RunRulesExit:
    
    RunRules = xmlResponseDoc.xml
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
   
    Exit Function
    
RunRulesVbErr:
    
    xmlResponseElem.setAttribute "TYPE", "APPERR"       'MAR1021 GHun
    
    Set xmlElem = xmlResponseDoc.createElement("ERROR")
    Set xmlNode = xmlResponseElem.appendChild(xmlElem)
    
    Set xmlElem = xmlResponseDoc.createElement("NUMBER")        'MAR1021 GHun
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem
        
    Set xmlElem = xmlResponseDoc.createElement("SOURCE")        'MAR1021 GHun
    xmlElem.Text = Err.Source
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlResponseDoc.createElement("DESCRIPTION")       'MAR1021 GHun
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem

    GoTo RunRulesExit

End Function

Private Sub xmlTransform(ByVal xmlRequestDoc As DOMDocument40, ByVal xmlResponseDoc As DOMDocument40, ByVal vstrXSLFileName)

    Dim xsldoc As DOMDocument40
    
    Dim strFileSpec As String
    
    strFileSpec = App.Path & "\" & vstrXSLFileName
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set xsldoc = New DOMDocument40
    xsldoc.async = False
    xsldoc.Load strFileSpec
    
    If xsldoc.parseError.errorCode <> 0 Then
        Err.Raise _
            Err.Number, "", _
                "error loading xslt document: " & _
                strFileSpec & _
                " " & xsldoc.parseError.reason
    End If
    
    xmlResponseDoc.loadXML xmlRequestDoc.transformNode(xsldoc)
    
    If xsldoc.parseError.errorCode <> 0 Then
        Err.Raise _
            Err.Number, "", _
                "error in transformation: " & _
                strFileSpec & _
                " " & xsldoc.parseError.reason
    End If

End Sub


