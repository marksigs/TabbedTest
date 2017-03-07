VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "DataIngestionBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <DataIngestionBO>
' File:         <DataIngestionBO.cls>
' Author:       Martin O'Hanlon
' Date:         17/7/2002
' Reference:    Business Object Specification--DataIngestionBO.doc
' Purpose:      To ingest data passed by external applications
' History:
'   AQR         AUTHOR            DESCRIPTION                            DATE
'   ---         ------            -----------                            ------
'   IDWP1       Martin O'Hanlon   Created                                17/07/2002
'   BMIDS00999  Martin O'Hanlon   Made changes to pass strings to the    20/11/2002
'                                 rules component
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

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
' Procedure OmRequest
' Author:           MO
' Date:             17/07/2002
' Purpose:          The only public method and therefore single entry point of the
'                   class, from here the requested operation is brokered.
' Input parameters: vstrXmlIn - an XML String
' Output parameters: OmRequest returns an XML response string
'------------------------------------------------------------------------
Public Function OmRequest(ByVal vstrXmlIn As String) As String

    On Error GoTo OmRequestVbErr
    
    Const strFunctionName As String = "OmRequest"
    
    'variable declaration
    Dim xmlInDOM As FreeThreadedDOMDocument40
    Dim xmlOutDOM As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    
    ' check to make sure we are running under MTS
    CheckMTSIsRunning strFunctionName
    
    Set xmlInDOM = xmlAssistEx.xmlLoad(vstrXmlIn, "OmRequest")
    
    Set xmlOutDOM = New FreeThreadedDOMDocument40
    xmlOutDOM.validateOnParse = False
    xmlOutDOM.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOutDOM.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDOM.appendChild(xmlResponseElem)
    xmlSetAttributeValue xmlResponseNode, "TYPE", "SUCCESS"
    
    DoRequest xmlInDOM.documentElement, xmlResponseNode
    
    OmRequest = xmlResponseNode.xml
    
    SetComplete

OmRequestExit:
    
    ' destroy objects
    Set xmlInDOM = Nothing
    Set xmlOutDOM = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
   
    Exit Function

OmRequestVbErr:
    
    SetAbort
    
    OmRequest = errCreateErrorResponse()
    'Debug.Print OmRequest
    
    Err.Clear
    
    Resume OmRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure DoRequest
' Author: MO
' Date: 17/07/2002
' Purpose: The main operation request handler for the class
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlResponseNode holding blank response
' Output parameters: vxmlResponseNode which will hold the XML response
'------------------------------------------------------------------------
Private Sub DoRequest(ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRequestVbErr
    
    Const strFunctionName As String = "DoRequest"
    
    'variable declaration
    Dim objDataIngestionTxBO As IDataIngestionTxBO
    
    Set objDataIngestionTxBO = CreateObject(App.Title & ".DataIngestionTxBO")
    
    objDataIngestionTxBO.DoRequest vxmlRequestNode, vxmlResponseNode
        
DoRequestExit:
    
    'destroy objects
    Set objDataIngestionTxBO = Nothing
    
    Exit Sub
    
DoRequestVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume DoRequestExit

End Sub

