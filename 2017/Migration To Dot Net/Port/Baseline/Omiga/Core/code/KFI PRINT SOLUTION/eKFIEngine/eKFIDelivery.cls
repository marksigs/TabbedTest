VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "eKFIDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl


Private Sub ApplyAlphaCalcsResultsToNode(xmlParentNode As IXMLDOMNode, xmlDataDoc As FreeThreadedDOMDocument40, xmlAlphaNode As IXMLDOMNode)
Const cFunction = "eKFIEngine.eKFIDelivery.ApplyAlphaCalcsResultsToNode"

Dim arrPaths
Dim strFormat As String
Dim xmlAlphaAttribute As IXMLDOMAttribute
Dim xmlAlphaNodeList As IXMLDOMNodeList
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlDataAttribute As IXMLDOMAttribute
Dim xmlDataNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
Dim X As Integer

    On Error GoTo ApplyAlphaCalcsResultsToNodeError:
    
    If xmlParentNode.hasChildNodes Then
        For Each xmlNode In xmlParentNode.childNodes
'Build path for Data Item
            Call ApplyAlphaCalcsResultsToNode(xmlNode, xmlDataDoc, xmlAlphaNode)
        Next
        Exit Sub
    End If

'Check if Data Node(s) exist in the Data documnent
    Set xmlDataNodeList = xmlDataDoc.selectNodes(GetNodePath(xmlParentNode))
    If xmlDataNodeList.length = 0 Then
        Exit Sub
    End If
    
    For Each xmlAttribute In xmlParentNode.Attributes
        Select Case xmlAttribute.nodeName
            Case "AlphaCalcs"
            
'Check if AlphaCalcs path is contains a valid attribute definition
                arrPaths = Split(xmlAttribute.Text, "/@")
                If UBound(arrPaths, 1) <> 1 Then
                    Err.Raise 1
                End If
                
'Check if AlphaCalcs path refers to a node in the AlphaCalcs results
                Set xmlAlphaNodeList = xmlAlphaNode.selectNodes(arrPaths(0))
                If xmlAlphaNodeList.length = 0 Then
                    Err.Raise 2
                End If
            Case "Format"
                strFormat = xmlAttribute.Text
        End Select
    Next
    
'AlphaCalcs path not set for this item
    If xmlAlphaNodeList Is Nothing Then
        Err.Raise 3
    End If
    
    X = 0
    For Each xmlNode In xmlDataNodeList
        Set xmlDataAttribute = xmlNode.Attributes.getNamedItem(xmlParentNode.nodeName)
        If xmlDataAttribute Is Nothing Then
            Exit Sub
        End If
        If X > xmlAlphaNodeList.length - 1 Then
            Err.Raise 4
        End If
        Set xmlAlphaAttribute = xmlAlphaNodeList.Item(X).Attributes.getNamedItem(arrPaths(1))
        X = X + 1
        If xmlAlphaAttribute Is Nothing Then
            Err.Raise 4
        End If
        xmlDataAttribute.Text = Format$(xmlAlphaAttribute.Text, strFormat)
TryNext:
    Next

TidyUp:
    On Error GoTo 0
    Set xmlAlphaAttribute = Nothing
    Set xmlAlphaNodeList = Nothing
    Set xmlAttribute = Nothing
    Set xmlDataAttribute = Nothing
    Set xmlDataNodeList = Nothing
    Set xmlNode = Nothing
    Exit Sub
    
ApplyAlphaCalcsResultsToNodeError:
    Select Case Err.Number
        Case 1
            xmlDataAttribute.Text = "No AlphaCalcs attribute in path !!"
        Case 2
            xmlDataAttribute.Text = "Not found in AlphaCalcs results !!"
        Case 3
            xmlDataAttribute.Text = "No AlphaCalcs path !!"
        Case 4
            xmlDataAttribute.Text = "Not found in AlphaCalcs results !!"
            Resume TryNext:
    End Select
    Resume TidyUp:
End Sub





Private Function GetNodePath(xmlNode As IXMLDOMNode) As String
Dim strPath As String
Dim xmlParentNode As IXMLDOMNode
    
    Set xmlParentNode = xmlNode.parentNode
    strPath = xmlParentNode.nodeName

'Exit from loop when the top of the tree has been reached
    Do Until xmlParentNode Is xmlNode.ownerDocument.documentElement
        Set xmlParentNode = xmlParentNode.parentNode
        strPath = xmlParentNode.nodeName & "/" & strPath
    Loop
    GetNodePath = strPath
End Function




Private Function RunAlphaCalcs(xmlAlphaNode As IXMLDOMNode, strInput As String, strXSL As String, strAlphaXMLCalcs As String, lngError As Long, strErrorSource As String, strErrorDescription As String) As String
Const cFunction = "eKFIEngine.eKFIDelivery.RunAlphaCalcs"

Dim AlphaCalcs As New Alpha

Dim strXML As String

'Dim xmlAlphaNode As IXMLDOMNode
Dim xmlDataDoc As New FreeThreadedDOMDocument40
Dim xmlResultDoc As New FreeThreadedDOMDocument40
Dim xmlTransformDoc As New FreeThreadedDOMDocument40
Dim xmlTransformedDataDoc As New FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
    
Dim xslProcessor As Object
Dim xslSheet As New FreeThreadedDOMDocument40
Dim xslTemplate As New XSLTemplate40
Dim X As Long

    On Error GoTo RunAlphaCalcsError:
    xmlDataDoc.setProperty "NewParser", True
    
'Ensure the format of the input data is correct for AlphaCalcs
    strXML = Replace(strInput, ",", "")
    xmlDataDoc.loadXML strXML
    
'Use the xsl transformation to create the input to AlphCalcs from the input data
    xslSheet.loadXML strXSL

    Set xslTemplate.stylesheet = xslSheet
    Set xslProcessor = xslTemplate.createProcessor()

    xslProcessor.input = xmlDataDoc
    X = xslProcessor.Transform()
    
    xmlTransformedDataDoc.setProperty "NewParser", True
    
'Ensure the date format of the input data is correct for AlphaCalcs
    strXML = Replace(xslProcessor.output, "%Today%", Format$(Now(), "yyyy-mm-dd"))
    xmlTransformedDataDoc.loadXML strXML
    
    If xmlTransformedDataDoc.xml = "" Or Not xmlTransformedDataDoc.parsed Then
        Err.Raise 60007, cFunction, "Error in transforming input data"
    End If
    
'Perform some final tweaks of the data format
    For Each xmlNode In xmlTransformedDataDoc.childNodes
        Call TransformNode(xmlNode)
    Next
    
'Call AlphaCalcs
    xmlResultDoc.setProperty "NewParser", True
    xmlResultDoc.loadXML AlphaCalcs.aceRequest(xmlTransformedDataDoc.xml)

'Check the AlphaCalcs Response node
    If xmlResultDoc.xml = "" Then
        Err.Raise 60008, cFunction, "AlphaCalcs.aceRequest response block empty"
    End If
    
    Set xmlNode = xmlResultDoc.selectSingleNode(".//Errors")
    If xmlNode Is Nothing Then
        Err.Raise 60009, cFunction, "AlphaCalcs.aceRequest response block 'Errors' node missing"
    End If
    
    If xmlNode.childNodes.length = 0 Then
        Set xmlNode = xmlResultDoc.selectSingleNode("//Response/Outputs")
        If xmlNode Is Nothing Then
            Err.Raise 60010, cFunction, "AlphaCalcs.aceRequest response block 'Response/Outputs' node missing"
        End If
        Set xmlAlphaNode = xmlResultDoc.selectSingleNode("//Outputs/Mortgage")
        If xmlAlphaNode Is Nothing Then
            Err.Raise 60011, cFunction, "AlphaCalcs.aceRequest response block 'Outputs/Mortgage' node missing"
        End If

'Merge the results from AlphaCalcs
        xmlTransformDoc.setProperty "NewParser", True
        xmlTransformDoc.loadXML strAlphaXMLCalcs
        For Each xmlNode In xmlTransformDoc.childNodes
            Call ApplyAlphaCalcsResultsToNode(xmlNode, xmlDataDoc, xmlAlphaNode)
        Next
        
        RunAlphaCalcs = xmlDataDoc.xml
    Else
        xmlTransformedDataDoc.save App.Path & "/AlphaCalcs_Input.xml"
        xmlResultDoc.save App.Path & "/AlphaCalcs_Results.xml"
        RunAlphaCalcs = xmlNode.xml
        Err.Raise 60012, cFunction, "AlphaCalcs.aceRequest returned error(s)"
    End If
    
TidyUp:
    Set xmlDataDoc = Nothing
    Set xmlResultDoc = Nothing
    Set xslProcessor = Nothing
    Set xslTemplate = Nothing
    Set xmlNode = Nothing
    Set AlphaCalcs = Nothing
'    Set xmlAlphaNode = Nothing
    Set xmlTransformedDataDoc = Nothing
    Set xslSheet = Nothing
    
    Exit Function
    
RunAlphaCalcsError:
    lngError = Err.Number
    strErrorDescription = Err.Description
    strErrorSource = Err.Source
    GoTo TidyUp:
End Function

Private Sub TransformNode(xmlParentNode As IXMLDOMNode)
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlNode As IXMLDOMNode

'This routine ensures that the data supplied to AlphaCalcs is in the correct format
    For Each xmlAttribute In xmlParentNode.Attributes
'Eliminate embedded spaces
        xmlAttribute.Text = Replace(xmlAttribute.Text, " ", "")
'Replace any occurrence of '/' with '-'
        xmlAttribute.Text = Replace(xmlAttribute.Text, "/", "-")
        If Len(xmlAttribute.Text) > 7 Then
            If IsDate(xmlAttribute.Text) Then
'Format date as 'yyyy-mm-dd'
                xmlAttribute.Text = Format$(xmlAttribute.Text, "yyyy-mm-dd")
            End If
        End If
    Next
    For Each xmlNode In xmlParentNode.childNodes
        Call TransformNode(xmlNode)
    Next
End Sub

Public Function CreateKFI(xmlRequestDoc As FreeThreadedDOMDocument40) As FreeThreadedDOMDocument40
' was Public Function CreateKFI(xmlRequestDoc As FreeThreadedDOMDocument40, strTemplate As String, strInput As String, strXSL As String, strAlphaCalcsXML As String) As FreeThreadedDOMDocument40
'This function performs all the processing required to create a
'KFI document from the supplied template and input data.
'AlphaCalcs is called if required and the results of this call are
'merged into the input data prior to the creation of the final document.
'CreateKFI returns a FreeThreadedDOMDocument40

'Parameters :
'   xmlRequestDoc is a FreeThreadedDOMDocument40 containing the parameters
'   strTemplate is a String containing the template in rtf format
'   strInput is a String containing the input data in xml format
'   strXSL is a String containing the xsl transformation required to transform
'       the input data to the format required for AlphaCalcs
'   strAlphaCalcsXML is a String containing xml defining where AlphaCalcs results
'       are to be merged into the input data

Const cFunction = "eKFIEngine.eKFIDelivery.CreateKFI"

Dim arrBytes() As Byte

Dim dmsCompression As New dmsCompression1
'Dim eKFICompression As New eKFICompression.eKFICompress
Dim Delivery As New DeliveryProcessor.Delivery
Dim Processor As New TemplateProcessor.Processor

Dim lngError As Long
Dim lngOriginalSize As Long
Dim lngTemplateSize As Long

Dim strAlphaCalcs As String
Dim strAlphaCalcsXML As String
Dim strCompression As String
Dim strErrorDescription As String
Dim strErrorSource As String
Dim strInputContent As String
Dim strInputData As String
Dim strOutputContent As String
Dim strTemplateData As String
Dim strWork As String
Dim strXSL As String

'Delivery parameters
Dim intPDFFontMode As Integer
Dim intPDFSecurity As Integer
Dim strPDFUserPassword As String
Dim strPDFOwnerPassword As String

'TemplateProcessor parameters
Dim strIntegerFormat As String
Dim strNumericFormat As String
Dim strDateFormat As String

Dim xmlAlphaCalcsResultsNode As IXMLDOMNode
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlDocumentContentsNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseNode As IXMLDOMNode
Dim xmlResponseDoc As New FreeThreadedDOMDocument40
Dim xmlSizeNode As IXMLDOMNode
Dim xmlTempAttribute As IXMLDOMAttribute
Dim X As Long

    On Error GoTo CreateKFIError:
    
'Set up RESPONSE node
    Set xmlResponseNode = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlResponseNode
    Set xmlDocumentContentsNode = xmlResponseDoc.createElement("DOCUMENTCONTENTS")
    xmlResponseNode.appendChild xmlDocumentContentsNode
    
'Check for REQUEST node
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("//REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise 60001, cFunction, "REQUEST node not found"
    End If
    
'Check input parameters
    Set xmlNode = xmlRequestDoc.selectSingleNode("//TEMPLATE")
    If xmlNode Is Nothing Then
        Err.Raise 60002, cFunction, "No Template Data supplied"
    End If
    
    Set xmlSizeNode = xmlRequestDoc.selectSingleNode("//TEMPLATESIZE")
    If xmlSizeNode Is Nothing Then
        strTemplateData = xmlNode.Text 'Template was not compressed
    Else
        arrBytes = xmlNode.nodeTypedValue
        dmsCompression.CompressionMethod = "ZLIB"
        arrBytes = dmsCompression.SafeArrayDecompressToSafeArray(arrBytes)
        strTemplateData = arrBytes
    End If
    
    Set xmlNode = xmlRequestDoc.selectSingleNode("//INPUT")
    If xmlNode Is Nothing Then
        Err.Raise 60003, cFunction, "No Input Data supplied"
    End If
    strInputData = xmlNode.Text
    
    Set xmlNode = xmlRequestDoc.selectSingleNode("//XSL")
    If Not xmlNode Is Nothing Then
        strXSL = xmlNode.Text
    End If
    
    Set xmlNode = xmlRequestDoc.selectSingleNode("//XML")
    
    If Not xmlNode Is Nothing Then
        strAlphaCalcsXML = xmlNode.Text
    End If
    
'Set up defaults
    strAlphaCalcs = "no"
    strInputContent = "rtf"
    strOutputContent = "rtf"
    strIntegerFormat = "#,##0"
    strNumericFormat = "#,##0.00"
    strDateFormat = "dd mmm yyyy"
    strCompression = ""

    intPDFFontMode = 0 'Use TrueType fonts - not embedded (default)
    intPDFSecurity = 3 'Printing allowed - no other access (default)
    strPDFUserPassword = ""
    strPDFOwnerPassword = "R3NrstsEs4EdeC7rtff4"

'Get the supplied request attributes
    For Each xmlAttribute In xmlRequestNode.Attributes
        strWork = xmlAttribute.Text
        Select Case xmlAttribute.nodeName
            Case "InputType"
                Select Case LCase$(strWork)
                    Case "rtf", "htm", "txt"
                        strInputContent = strWork
                    Case Else
                        Err.Raise 60015, cFunction, "REQUEST attribute InputType must be 'rtf', 'htm' or 'txt'"
                End Select
            Case "OutputType"
                Select Case LCase$(strWork)
                    Case "pdf", "doc", "rtf", "htm"
                        strOutputContent = strWork
                    Case Else
                        Err.Raise 60004, cFunction, "REQUEST attribute OutputType must be 'pdf', 'doc' or 'rtf'"
                End Select
            Case "AlphaCalcs"
                Select Case LCase$(strWork)
                    Case "yes", "no"
                        strAlphaCalcs = strWork
                    Case Else
                        Err.Raise 60005, cFunction, "REQUEST attribute AlphaCalcs must be 'yes' or 'no'"
                End Select
            Case "DateFormat"
                strDateFormat = strWork
            Case "IntegerFormat"
                strIntegerFormat = strWork
            Case "NumericFormat"
                strNumericFormat = strWork
            Case "PDFFontMode"
                intPDFFontMode = Val(strWork)
            Case "PDFSecurity"
                intPDFSecurity = Val(strWork)
            Case "PDFUserPassword"
                strPDFUserPassword = strWork
            Case "PDFOwnerPassword"
                strPDFOwnerPassword = strWork
            Case "Compression"
                Select Case strWork
                    Case "DMS", "KFI", ""
                        strCompression = strWork
                    Case Else
                        Err.Raise 60013, cFunction, "REQUEST attribute Compression must be 'DMS' or 'KFI'"
                End Select
        End Select
    Next
    
'Convert Template to XML
'    strTemplateData = Processor.ConvertRTFTemplateToXML(Replace(strTemplate, vbCrLf, ""))
    Processor.DateFormat = strDateFormat
    Processor.IntegerFormat = strIntegerFormat
    Processor.NumericFormat = strNumericFormat
    
    Select Case strInputContent
        Case "rtf"
            strTemplateData = Processor.ConvertRTFTemplateToXML(strTemplateData)
        Case "htm"
            strTemplateData = Processor.ConvertHTMLTemplateToXML(strTemplateData)
            strOutputContent = "htm"
        Case "txt"
            Err.Raise 60016, cFunction, "Text format Template not currently supported"
    End Select
    If Processor.ErrorState <> 0 Then
        Err.Raise Processor.ErrorState, cFunction, Processor.ErrorResponse
    End If
    
'Run AlphaCalcs if required
    If strAlphaCalcs = "yes" Then
        If strXSL = "" Then
            Err.Raise 60006, cFunction, "No XSL Transformation data supplied for AlphaCalcs"
        End If

'Set up Element to receive AlphaCalcs results
        Set xmlAlphaCalcsResultsNode = xmlResponseDoc.createElement("ALPHACALCSRESULTS")
        
        strInputData = RunAlphaCalcs(xmlAlphaCalcsResultsNode, strInputData, strXSL, strAlphaCalcsXML, lngError, strErrorSource, strErrorDescription)
        If lngError <> 0 Then
            Err.Raise lngError, strErrorSource, strErrorDescription
        End If
        
'Add AlphaCalcs results to response
        xmlResponseNode.appendChild xmlAlphaCalcsResultsNode
    
    End If
    
'Process the Template against the Input Data to produce an output document string
    strWork = Processor.ProcessXMLTemplate(strInputData, strTemplateData)
    
    If Processor.ErrorState <> 0 Then
        Err.Raise Processor.ErrorState, cFunction, Processor.ErrorResponse
    End If
    
'Do any required conversion of output document string
    Select Case strOutputContent
        Case "pdf"
            Delivery.SecurityLevel = intPDFSecurity
            Delivery.FontMode = intPDFFontMode
            Delivery.OwnerPassword = strPDFOwnerPassword
            Delivery.UserPassword = strPDFUserPassword
            
            arrBytes = Delivery.DeliverAsPDF(strWork)
        Case "doc"
            arrBytes = Processor.DeliverAsWordDoc(strWork)
        Case Else
            arrBytes = strWork
    End Select
    
    Select Case strOutputContent
        Case "pdf"
            If Delivery.ErrorState <> 0 Then
                Err.Raise Delivery.ErrorState, cFunction, Delivery.ErrorResponse
            End If
        Case "doc"
            If Processor.ErrorState <> 0 Then
                Err.Raise Processor.ErrorState, cFunction, Processor.ErrorResponse
            End If
    End Select
    
    If strCompression <> "" Then
        lngOriginalSize = UBound(arrBytes, 1)
        Select Case strCompression
            Case "KFI"
                dmsCompression.CompressionMethod = "ZLIB"
            Case "DMS"
                dmsCompression.CompressionMethod = "COMPAPI"
        End Select
        arrBytes = dmsCompression.SafeArrayCompressToSafeArray(arrBytes())
    End If
    
'Set up the (processed) output document string as XML
    Set xmlNode = Processor.DeliverAsXML(arrBytes)
    For Each xmlAttribute In xmlNode.Attributes
        Set xmlTempAttribute = xmlResponseDoc.createAttribute(xmlAttribute.nodeName)
        If IsNull(xmlAttribute.dataType) Then
            xmlTempAttribute.Text = xmlAttribute.Text
        Else
            xmlTempAttribute.dataType = xmlAttribute.dataType
            xmlTempAttribute.nodeTypedValue = xmlAttribute.nodeTypedValue
        End If
        xmlDocumentContentsNode.Attributes.setNamedItem xmlTempAttribute
    Next
    
    Set xmlTempAttribute = xmlResponseDoc.createAttribute("FILECOMPRESSION")
    xmlTempAttribute.Text = strCompression
    If strCompression <> "" Then
        xmlDocumentContentsNode.Attributes.setNamedItem xmlTempAttribute
        Set xmlTempAttribute = xmlResponseDoc.selectSingleNode("//DOCUMENTCONTENTS").Attributes.getNamedItem("FILESIZE")
        xmlTempAttribute.Text = lngOriginalSize
    End If
    
    Set xmlTempAttribute = xmlResponseDoc.createAttribute("CONTENTTYPE")
    xmlTempAttribute.Text = strOutputContent
    xmlDocumentContentsNode.Attributes.setNamedItem xmlTempAttribute
    
'Set RESPONSE node Type attribute to "SUCCESS"
    Set xmlAttribute = xmlResponseDoc.createAttribute("Type")
    xmlResponseNode.Attributes.setNamedItem xmlAttribute
    xmlAttribute.Text = "SUCCESS"

    If strAlphaCalcs = "yes" Then
        
    End If

TidyUp:
'Tidy up before exit
    Set CreateKFI = xmlResponseDoc
    
    Set Delivery = Nothing
    Set Processor = Nothing
    
    Set xmlAttribute = Nothing
    Set xmlDocumentContentsNode = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlResponseNode = Nothing
    Set xmlSizeNode = Nothing
    Set xmlTempAttribute = Nothing
    
    Erase arrBytes
    Exit Function

CreateKFIError:
    App.LogEvent cFunction & " " & Err.Description, vbLogEventTypeWarning
    
'Set RESPONSE node Type attribute to "ERROR"
    Set xmlAttribute = xmlResponseDoc.createAttribute("Type")
    xmlResponseNode.Attributes.setNamedItem xmlAttribute
    xmlAttribute.Text = "ERROR"
    
'Set RESPONSE node ErrorNumber attribute to the appropriate error number
    Set xmlAttribute = xmlResponseDoc.createAttribute("ErrorNumber")
    xmlAttribute.Text = Err.Number
    xmlResponseNode.Attributes.setNamedItem xmlAttribute
    
'Set RESPONSE node ErrorMessage attribute to the appropriate error description
    Set xmlAttribute = xmlResponseDoc.createAttribute("ErrorMessage")
    xmlAttribute.Text = Err.Description
    xmlResponseNode.Attributes.setNamedItem xmlAttribute
    GoTo TidyUp:

End Function




Private Sub ObjectControl_Activate()

End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function


Private Sub ObjectControl_Deactivate()

End Sub


