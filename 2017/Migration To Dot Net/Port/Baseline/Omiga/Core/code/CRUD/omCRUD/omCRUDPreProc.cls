VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omCRUDPreProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function OmRequest(ByVal vstrXmlIn As String) As String
    
    On Error GoTo omRequestExit
    
    Dim xmlRequestDoc As DOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlErrorNode As IXMLDOMNode
    
    GetObjectContext.SetComplete
    
    Set xmlRequestDoc = New DOMDocument40
    xmlRequestDoc.async = True
    xmlRequestDoc.setProperty "NewParser", True
    
    xmlRequestDoc.loadXML vstrXmlIn
    
    If Not xmlRequestDoc.selectSingleNode("REQUEST/@preProcRef") Is Nothing Then
        Select Case xmlRequestDoc.selectSingleNode("REQUEST/@preProcRef").Text

' add code based on @preProcRef here
' e.g.
'            Case "newCustomer"
'
'                If Not xmlRequestDoc.selectSingleNode("//CUSTOMERVERSION[@SURNAME='Windsor']") Is Nothing Then
'
'                    xmlRequestDoc.loadXML "<RESPONSE TYPE='APPERR'/>"
'
'                    Set xmlElem = xmlRequestDoc.createElement("ERROR")
'                    Set xmlErrorNode = xmlRequestDoc.documentElement.appendChild(xmlElem)
'
'                    Set xmlElem = xmlRequestDoc.createElement("NUMBER")
'                    xmlElem.Text = "9000"
'                    xmlErrorNode.appendChild xmlElem
'
'                    Set xmlElem = xmlRequestDoc.createElement("SOURCE")
'                    xmlElem.Text = "omCRUDPreProc"
'                    xmlErrorNode.appendChild xmlElem
'
'                    Set xmlElem = xmlRequestDoc.createElement("DESCRIPTION")
'                    xmlElem.Text = "Windsor family barred"
'                    xmlErrorNode.appendChild xmlElem
'
'                End If
        
        End Select
    End If
    
    OmRequest = xmlRequestDoc.xml

omRequestExit:
    
    Set xmlRequestDoc = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

