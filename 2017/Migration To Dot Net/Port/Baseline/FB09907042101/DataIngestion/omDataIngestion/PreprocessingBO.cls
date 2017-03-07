VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PreprocessingBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <PreProcessingBO>
' File:         <PreprocessingBO.cls>
' Author:       Martin O'Hanlon
' Date:         17/7/2002
' Reference:    Business Object Specification--DataIngestionBO.doc
' Purpose:      To perform any pre processing tasks for the ingestion engine, also
'               calls the client configurable pre processing component
' History:
'   AQR         AUTHOR            DESCRIPTION                               DATE
'   ---         ------            -----------                               ------
'   IDWP1       Martin O'Hanlon   Created                                   17/07/2002
'   BMIDS00999  Martin O'Hanlon   Change to pass strings to rules component 20/11/2002
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Implements IPreprocessingBO

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
Private Sub IPreprocessingBO_DoRequest(ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRequestVbErr
    
    Const strFunctionName As String = "IPreprocessingBO_DoRequest"
    
    Dim objPreprocessingRules As Object
    Dim strOperationRequested As String
    Dim strPreprocessingResponse As String
    
    CheckMTSIsRunning strFunctionName
    
    'get the operation that has been requested, CreateApplication, CreateCustomer etc
    strOperationRequested = xmlGetMandatoryAttributeText(vxmlRequestNode, "OPERATION")
    
    'call the client configurable object, preprocessing rules
    Set objPreprocessingRules = gobjContext.CreateInstance("omDataIngestionRules.Preprocessing")
    
    'MO - 20/11/2002 - BMIDS00999
    'objPreprocessingRules.DoRequest vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
    strPreprocessingResponse = objPreprocessingRules.OmRequest(BuildPreprocessingRulesRequest(vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode))
    ExtractPreprocessingRulesResponse strPreprocessingResponse, vxmlRequestNode, vxmlOmigaFormatRequestNode, vxmlResponseNode
        
    'do the value conversion
    ValueConversion strOperationRequested, vxmlOmigaFormatRequestNode, vxmlResponseNode
        
DoRequestExit:
    
    'destroy objects
    Set objPreprocessingRules = Nothing
    
    Exit Sub
    
DoRequestVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume DoRequestExit
    
End Sub

'------------------------------------------------------------------------
' Procedure ValueConversion
' Author: MO
' Date: 18/07/2002
' Purpose: Converts values that are passed in by the calling application to
'           the correct value required by omiga.  Uses a xml config file.
' Example Config xml:
'<VALUECONVERSION>
'   <VALUETOCONVERTLIST @OPERATION="CREATEAPPLICATION">
'       <VALUETOCONVERT XPATH="CUSTOMER\CUSTOMERVERSION\@TITLE" DEFAULTVALUE="99">
'           <VALUE DESCRIPTION="Mr" INPUTVALUE="MR" OUTPUTVALUE="1"/>
'           <VALUE DESCRIPTION="Mrs" INPUTVALUE="MRS" OUTPUTVALUE="2"/>
'       </VALUETOCONVERT >
'       <VALUETOCONVERT XPATH="APPLICATION\@TYPEOFBUYER">
'           <VALUE DESCRIPTION="First Time" INPUTVALUE="10" OUTPUTVALUE="1"/>
'           <VALUE DESCRIPTION="Subsequent"INPUTVALUE ="5" OUTPUTVALUE="2"/>
'           <VALUE DESCRIPTION="Concurrent" INPUTVALUE="15" OUTPUTVALUE="3"/>
'       </VALUETOCONVERT >
'   </VALUETOCONVERT>
'<VALUECONVERSION>
' Input parameters: OmigaFormattedXML, ResponseXML
' Output parameters: None
'------------------------------------------------------------------------
Private Sub ValueConversion(ByRef strOperationRequest As String, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValueConversionVbErr
    
    Const strFunctionName As String = "ValueConversion"
    Dim xmlValueConfig As FreeThreadedDOMDocument40
    Dim xmlValueFoundNodeList As IXMLDOMNodeList
    Dim xmlValueToConvert As IXMLDOMNode
    Dim xmlNodeXPath As IXMLDOMNode
    Dim xmlValueFoundNode As IXMLDOMNode
    Dim xmlNewValueAttribute As IXMLDOMNode
    Dim xmlDefaultValueNode As IXMLDOMNode
    Dim strXMLLocation As String
    
    Set xmlValueConfig = New FreeThreadedDOMDocument40
    xmlValueConfig.validateOnParse = False
    xmlValueConfig.setProperty "NewParser", True
    
    ' pick up XML file from "...\Omiga 4\XML" directory
    strXMLLocation = App.Path & "\ValueConversion.xml"
    strXMLLocation = Replace(strXMLLocation, "\DLL", "\XML", 1, 1, vbTextCompare)
    
    'Load the combovalueconversion file
    If xmlValueConfig.Load(strXMLLocation) Then
        'loop through all the combo value
        For Each xmlValueToConvert In xmlValueConfig.selectNodes("VALUECONVERSION/VALUETOCONVERTLIST[@OPERATION = """ & strOperationRequest & """]/VALUETOCONVERT")
            ' get the default value attribute
            Set xmlDefaultValueNode = xmlValueToConvert.selectSingleNode("@DEFAULTVALUE")
            ' get the xpath from the config xml, that Im going to modify the results for
            Set xmlNodeXPath = xmlGetMandatoryNode(xmlValueToConvert, "@XPATH")
            Set xmlValueFoundNodeList = vxmlOmigaFormatRequestNode.selectNodes(xmlNodeXPath.Text)
            For Each xmlValueFoundNode In xmlValueFoundNodeList
                ' see if there is a new combo value for this value
                Set xmlNewValueAttribute = xmlValueToConvert.selectSingleNode("VALUE/@OUTPUTVALUE[../@INPUTVALUE='" & UCase(xmlValueFoundNode.Text) & "']")
                If Not xmlNewValueAttribute Is Nothing Then
                    'set the new value
                    xmlValueFoundNode.Text = xmlNewValueAttribute.Text
                Else
                    'is there a default value?
                    If Not xmlDefaultValueNode Is Nothing Then
                        xmlValueFoundNode.Text = xmlDefaultValueNode.Text
                    End If
                End If
            Next
        Next
    Else
        xmlParserError xmlValueConfig.parseError, strFunctionName
    End If
    
ValueConversionExit:
    
    'destroy objects
    Set xmlNewValueAttribute = Nothing
    Set xmlValueFoundNodeList = Nothing
    Set xmlValueFoundNode = Nothing
    Set xmlNewValueAttribute = Nothing
    Set xmlValueToConvert = Nothing
    Set xmlValueConfig = Nothing
    
    Exit Sub
    
ValueConversionVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume ValueConversionExit
    
End Sub

'------------------------------------------------------------------------
' Procedure BuildPreprocessingRulesRequest
' Author: MO
' Date: 20/11/2002
' Purpose: Takes the 3 main nodes of the operation and converts them into
'           the string request needed for the Rules component
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlOmigaFormatRequestNode holding the omiga xml request
'                   vxmlResponseNode holding blank response
' Output parameters: String
'------------------------------------------------------------------------
Private Function BuildPreprocessingRulesRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlOmigaFormatRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As String

    On Error GoTo BuildPreprocessingRulesRequestVbErr
    
    Const strFunctionName As String = "BuildPreprocessingRulesRequest"
    
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
    BuildPreprocessingRulesRequest = xmlRulesRequestDOM.xml

BuildPreprocessingRulesRequestExit:
    
    'destroy objects
    Set xmlRulesRequestDOM = Nothing
    Set xmlRequestNode = Nothing
    Set xmlOriginalRequestNode = Nothing
    Set xmlOmigaFormattedRequestNode = Nothing
    Set xmlCurrentResponseNode = Nothing

    Exit Function
    
BuildPreprocessingRulesRequestVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume BuildPreprocessingRulesRequestExit
    
End Function

'------------------------------------------------------------------------
' Procedure ExtractPreprocessingRulesResponse
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
Private Sub ExtractPreprocessingRulesResponse(ByVal strRulesResponse As String, ByRef vxmlRequestNode As IXMLDOMNode, ByRef vxmlOmigaFormatRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ExtractPreprocessingRulesResponseVbErr
    
    Const strFunctionName As String = "ExtractPreprocessingRulesResponse"
    
    Dim xmlRulesResponseDOM As FreeThreadedDOMDocument40
    
    Set xmlRulesResponseDOM = xmlLoad(strRulesResponse, strFunctionName)
    
    'get the original request node
    Set vxmlRequestNode = xmlGetMandatoryNode(xmlRulesResponseDOM.documentElement, "ORIGINALREQUEST/REQUEST").cloneNode(True)
    
    'get the omiga formatted request node
    Set vxmlOmigaFormatRequestNode = xmlGetMandatoryNode(xmlRulesResponseDOM.documentElement, "OMIGAFORMATTEDREQUEST/REQUEST").cloneNode(True)
    
    'get the current response node
    Set vxmlResponseNode = xmlGetMandatoryNode(xmlRulesResponseDOM.documentElement, "CURRENTRESPONSE/RESPONSE").cloneNode(True)
    
ExtractPreprocessingRulesResponseExit:
    
    'destroy objects
    Set xmlRulesResponseDOM = Nothing
    
    Exit Sub
    
ExtractPreprocessingRulesResponseVbErr:

    SetAbort
    
    errCheckError strFunctionName, TypeName(Me)
    
    Resume ExtractPreprocessingRulesResponseExit
    
End Sub

