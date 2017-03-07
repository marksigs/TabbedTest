VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "DataIngestionTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <DataIngestionTxBO>
' File:         <DataIngestionTxBO.cls>
' Author:       Martin O'Hanlon
' Date:         17/7/2002
' Reference:    Business Object Specification--DataIngestionBO.doc
' Purpose:      To ingest data passed by external applications, and host within a transaction
' History:
'   AQR                 AUTHOR                      DESCRIPTION           DATE
'   ---                 ------                      -----------          ------
'   IDWP1               Martin O'Hanlon             Created              17/07/2002
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Implements IDataIngestionTxBO

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
' Procedure IOMDataIngestionTxBO_DoRequest
' Author: MO
' Date: 17/07/2002
' Purpose: Main request to do the data ingestion, calls the relevant function to do the job
' Input parameters: None
' Output parameters: None
'------------------------------------------------------------------------
Private Sub IDataIngestionTxBO_DoRequest(ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRequestVbErr
    
    Const strFunctionName As String = "IDataIngestionTxBO_DoRequest"
    
    'variable declaration
    Dim xmlOmigaFormatRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim objPreprocessing As IPreprocessingBO
    Dim objPostprocessing As IPostprocessingBO
    Dim strOperationRequested As String
    Dim strXSLLocation As String
    
    Dim strDebugTimer As String
    
    'get the operation that has been requested, CreateApplication, CreateCustomer etc
    strOperationRequested = xmlGetMandatoryAttributeText(vxmlRequestNode, "OPERATION")
    
    'PUT THE CURRENT DATETIME INTO THE REQUEST NODE
    'this time can be used within the transformation for any current datetime values
    xmlSetAttributeValue vxmlRequestNode, "NOW", Now
    
    'TRANSFORMATION
    ' pick up XSL file from "...\Omiga 4\XML" directory
    strXSLLocation = App.Path & "\" & strOperationRequested & ".xsl"
    strXSLLocation = Replace(strXSLLocation, "\DLL", "\XML", 1, 1, vbTextCompare)
    
    Set xmlOmigaFormatRequestNode = TransformXML(strXSLLocation, vxmlRequestNode, vxmlResponseNode)
       
    'PRE_PROCESSING
    Set objPreprocessing = gobjContext.CreateInstance(App.Title & ".PreprocessingBO")
    
    objPreprocessing.DoRequest vxmlRequestNode, xmlOmigaFormatRequestNode, vxmlResponseNode
     
    'DATA WRITING
    
    ' do we have a CRUD OPERATION node
    Set xmlOperationNode = adoGetOperationNode(strOperationRequested)
    
    If Not xmlOperationNode Is Nothing Then
    
        If xmlAttributeValueExists(xmlOperationNode, "CRUD_OP") Then
            'DEBUG, write dom to file!
            'strDebugTimer = Timer
            'WriteXMLToFile xmlOmigaFormatRequestNode, "c:\temp\DataIngestion_IntoCRUD" & strDebugTimer & ".xml"
            adoCRUD xmlOperationNode, xmlOmigaFormatRequestNode, vxmlResponseNode
            'DEBUG, write dom to file!
            'WriteXMLToFile xmlOmigaFormatRequestNode, "c:\temp\DataIngestion_OutOfCRUD" & strDebugTimer & ".xml"
        Else
            errThrowError "adoCRUD", oeXMLMissingAttribute, "OM_SCHEMA OPERATION missing CRUD_OP: " & xmlOperationNode.xml
        End If
    
    Else
        errThrowError "OmRequest", oeNotImplemented, vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    End If
    
    'POST PROCESSING
    Set objPostprocessing = gobjContext.CreateInstance(App.Title & ".PostprocessingBO")
    
    objPostprocessing.DoRequest vxmlRequestNode, xmlOmigaFormatRequestNode, vxmlResponseNode
    
    SetComplete
    
DoRequestExit:
    
    'destory objects
    Set objPreprocessing = Nothing
    Set objPostprocessing = Nothing
    Set xmlOmigaFormatRequestNode = Nothing
    Set xmlOperationNode = Nothing
    
    Exit Sub

DoRequestVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume DoRequestExit

End Sub

'------------------------------------------------------------------------
' Procedure TransformXML
' Author: MO
' Date: 18/07/2002
' Purpose: Transforms the XML from the passed in request to an Omiga formatted
'           CRUD operation, uses an XSL file.  The XSL filename is the operation
'           + .xsl
' Input parameters: None
' Output parameters: None
'------------------------------------------------------------------------
Private Function TransformXML(ByRef strXSLFilename As String, ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo TransformXMLVbErr
    
    Const strFunctionName As String = "TransformXML"
    
    Dim xmlTransformXSL As FreeThreadedDOMDocument40
    Dim xmlTransformedXML As FreeThreadedDOMDocument40
    Dim xmlTempDOM As FreeThreadedDOMDocument40
    Dim strTranformedXML As String
    
    Set xmlTransformXSL = New FreeThreadedDOMDocument40
    xmlTransformXSL.validateOnParse = False
    xmlTransformXSL.setProperty "NewParser", True
    
    'Load the xsl file
    If xmlTransformXSL.Load(strXSLFilename) Then
        strTranformedXML = xmlLoad(vxmlRequestNode.xml, strFunctionName).transformNode(xmlTransformXSL.documentElement)
        
        Set xmlTransformedXML = New FreeThreadedDOMDocument40
        xmlTransformedXML.validateOnParse = False
        xmlTransformedXML.setProperty "NewParser", True
        xmlTransformedXML.loadXML strTranformedXML
        
        Set TransformXML = xmlTransformedXML.documentElement
        
    Else
        xmlParserError xmlTransformXSL.parseError, strFunctionName
    End If
        
    
TransformXMLExit:
    
    'destroy objects
    Set xmlTransformXSL = Nothing
    Set xmlTransformedXML = Nothing
    Set xmlTempDOM = Nothing
    
    Exit Function
    
TransformXMLVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume TransformXMLExit

End Function


'HERE ONLY FOR DEBUGGING
Private Sub WriteXMLToFile(ByVal objNode As IXMLDOMNode, ByVal strFileName As String)

    Dim objDOM As FreeThreadedDOMDocument40
    
    Set objDOM = New FreeThreadedDOMDocument40
    objDOM.validateOnParse = False
    objDOM.setProperty "NewParser", True
    
    objDOM.loadXML objNode.xml
    
    objDOM.save strFileName
    
    Set objDOM = Nothing

End Sub
