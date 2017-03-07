VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omCRUDPostProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function OmRequest( _
    ByVal vstrXmlRequest As String, _
    ByVal vstrXmlResponse As String _
    ) As String
    
    On Error GoTo omRequestExit
    
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    
    GetObjectContext.SetComplete
    
    Set xmlRequestDoc = New DOMDocument40
    xmlRequestDoc.async = True
    xmlRequestDoc.setProperty "NewParser", True
    
    xmlRequestDoc.loadXML vstrXmlRequest
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = True
    xmlResponseDoc.setProperty "NewParser", True
    
    xmlResponseDoc.loadXML vstrXmlResponse
    
    If Not xmlRequestDoc.selectSingleNode("REQUEST/@postProcRef") Is Nothing Then
        Select Case xmlRequestDoc.selectSingleNode("REQUEST/@postProcRef").Text
            ' add code based on @postProcRef here
        End Select
    End If
    
    OmRequest = xmlResponseDoc.xml
    
omRequestExit:
    
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

