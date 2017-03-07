VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PostprocessingBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <PostProcessingBO>
' File:         <PostprocessingBO.cls>
' Author:       Martin O'Hanlon
' Date:         17/7/2002
' Reference:    Business Object Specification--DataIngestionBO.doc
' Purpose:      To perform any post processing tasks for the ingestion engine, also
'               calls the client configurable post processing component
' History:
'   AQR         AUTHOR            DESCRIPTION                               DATE
'   ---         ------            -----------                               ------
'   IDWP1       Martin O'Hanlon   Created                                   17/07/2002
'   BMIDS00999  Martin O'Hanlon   Change to pass strings to rules component 20/11/2002
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Implements IPostprocessingBO

Private gobjContext As ObjectContext

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'------------------------------------------------------------------------
' Procedure CheckMTSIsRunning
' Author: APS
' Date: 29/08/01
' Purpose:  Checks to mkae sure that MTS is running fine
' Input parameters: vstrFunctionName
' Output parameters: None
'------------------------------------------------------------------------
Private Sub CheckMTSIsRunning(ByVal vstrFunctionName As String)
    If gobjContext Is Nothing Then
        errThrowError vstrFunctionName, oeMTSNotFound
    End If
End Sub

'------------------------------------------------------------------------
' Procedure SetComplete
' Author: APS
' Date: 29/08/01
' Purpose: A central place to set complete
' Input parameters: None
' Output parameters: None
'------------------------------------------------------------------------
Private Sub SetComplete()
    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If
End Sub

'------------------------------------------------------------------------
' Procedure SetAbort
' Author: APS
' Date: 29/08/01
' Purpose: A central place to set abort
' Input parameters: None
' Output parameters: None
'------------------------------------------------------------------------
Private Sub SetAbort()
    If Not gobjContext Is Nothing Then
        gobjContext.SetAbort
    End If
End Sub

'------------------------------------------------------------------------
' Procedure DoRequest
' Author: MO
' Date: 17/07/2002
' Purpose: The main operation request handler for the class
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlResponseNode holding blank response
' Output parameters: vxmlResponseNode which will hold the XML response
'------------------------------------------------------------------------
Private Sub IPostprocessingBO_DoRequest(ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRequestVbErr
    
    Const strFunctionName As String = "IPostprocessingBO_DoRequest"
    
    Dim objPostprocessingRules As Object
    Dim strPostprocessingResponse As String
      
    CheckMTSIsRunning strFunctionName
    
    Set objPostprocessingRules = gobjContext.CreateInstance("omDataIngestionRules.Postprocessing")
    'MO - 20/11/2002 - BMIDS00999
    'objPostprocessingRules.DoRequest vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
    strPostprocessingResponse = objPostprocessingRules.OmRequest(BuildPostprocessingRulesRequest(vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode))
    ExtractPostprocessingRulesResponse strPostprocessingResponse, vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
    
    SetComplete
    
DoRequestExit:
    
    'destroy objects
    Set objPostprocessingRules = Nothing
    
    Exit Sub
    
DoRequestVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume DoRequestExit

End Sub

'------------------------------------------------------------------------
' Procedure BuildPostprocessingRulesRequest
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the 3 main nodes of the operation and converts them into
'           the string request needed for the Rules component
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildPostprocessingRulesRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As String

    On Error GoTo BuildPostprocessingRulesRequestVbErr
    
    Const strFunctionName As String = "BuildPostprocessingRulesRequest"
    
    Dim xmlRulesRequestDOM As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOriginalRequestNode As IXMLDOMNode
    Dim xmlOmigaFormattedRequestNode As IXMLDOMNode
    Dim xmlCurrentResponseNode As IXMLDOMNode
    
    Set xmlRulesRequestDOM = New FreeThreadedDOMDocument40
    xmlRulesRequestDOM.validateOnParse = False
    xmlRulesRequestDOM.setProperty "NewParser", True
    'create the top level request node
    Set xmlRequestNode = xmlRulesRequestDOM.createElement("REQUEST")
    xmlRulesRequestDOM.appendChild xmlRequestNode
    
    'create the original request node
    Set xmlOriginalRequestNode = xmlRulesRequestDOM.createElement("ORIGINALREQUEST")
    xmlRequestNode.appendChild xmlOriginalRequestNode
    'append the original request node to it!
    xmlOriginalRequestNode.appendChild vxmlRequestNode.cloneNode(True)
    
    'create the omiga formatted request node
    Set xmlOmigaFormattedRequestNode = xmlRulesRequestDOM.createElement("OMIGAFORMATTEDREQUEST")
    xmlRequestNode.appendChild xmlOmigaFormattedRequestNode
    'append the omiga formatted request to it
    xmlOmigaFormattedRequestNode.appendChild vxmlOmigaFormatRequestNode.cloneNode(True)
    
    'create the current response node
    Set xmlCurrentResponseNode = xmlRulesRequestDOM.createElement("CURRENTRESPONSE")
    xmlRequestNode.appendChild xmlCurrentResponseNode
    'append the response node to it
    xmlCurrentResponseNode.appendChild vxmlResponseNode.cloneNode(True)
        
    'return the xml as a string
    BuildPostprocessingRulesRequest = xmlRulesRequestDOM.xml

BuildPostprocessingRulesRequestExit:
    
    'destroy objects
    Set xmlRulesRequestDOM = Nothing
    Set xmlRequestNode = Nothing
    Set xmlOriginalRequestNode = Nothing
    Set xmlOmigaFormattedRequestNode = Nothing
    Set xmlCurrentResponseNode = Nothing

    Exit Function
    
BuildPostprocessingRulesRequestVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume BuildPostprocessingRulesRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure ExtractPostprocessingRulesResponse
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the response from preprocessing and extracts the 3 main nodes
'           of the operation and converts
' Input parameters: vstrRulesResponse holding the response from the rules component
'                   vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Sub ExtractPostprocessingRulesResponse(ByVal strRulesResponse As String, ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ExtractPostprocessingRulesResponseVbErr
    
    Const strFunctionName As String = "ExtractPostprocessingRulesResponse"
    
    Dim xmlRulesResponseDOM As FreeThreadedDOMDocument40
    
    Set xmlRulesResponseDOM = xmlLoad(strRulesResponse, strFunctionName)
    
    'get the original request node
    Set vxmlRequestNode = xmlGetMandatoryNode(xmlRulesResponseDOM.documentElement, "ORIGINALREQUEST/REQUEST").cloneNode(True)
    
    'get the omiga formatted request node
    Set vxmlOmigaFormatRequestNode = xmlGetMandatoryNode(xmlRulesResponseDOM.documentElement, "OMIGAFORMATTEDREQUEST/REQUEST").cloneNode(True)
    
    'get the current response node
    Set vxmlResponseNode = xmlGetMandatoryNode(xmlRulesResponseDOM.documentElement, "CURRENTRESPONSE/RESPONSE").cloneNode(True)
    
ExtractPostprocessingRulesResponseExit:
    
    'destroy objects
    Set xmlRulesResponseDOM = Nothing
    
    Exit Sub
    
ExtractPostprocessingRulesResponseVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume ExtractPostprocessingRulesResponseExit
    
End Sub


