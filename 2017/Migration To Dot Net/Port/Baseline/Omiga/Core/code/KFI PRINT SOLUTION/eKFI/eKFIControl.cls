VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "eKFIControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const cFunction = "eKFI.eKFIControl.ProcessKFI"

Implements ObjectControl

Private Function GetFileData(strPath) As String
Dim F As Integer
Dim strWork As String

    strWork = Space$(FileLen(strPath))
    F = FreeFile()
    Open strPath For Binary Access Read As F
    Get F, , strWork
    Close F
    GetFileData = strWork
    strWork = ""
End Function


Private Function getXML(strSource As String, intSourceType As Integer) As String
Dim blnFile As Boolean
Dim fs As New Scripting.FileSystemObject
Dim strData As String
    
    If intSourceType = 1 Then 'rtf Template
        If Left$(strSource, 1) <> "{" Then
            blnFile = True
        End If
    Else
        If InStr(1, strSource, "<") = 0 Then
            blnFile = True
        End If
    End If
    
    If blnFile Then
        If fs.FileExists(strSource) Then
            strData = GetFileData(strSource)
        Else
            Exit Function
        End If
    Else
        strData = strSource
    End If
    getXML = strData
End Function

Public Function ProcessKFI(ByVal strXMLRequest, Optional ByVal ReturnDOMDocument = False)

Dim arrBytes() As Byte
Dim fs As New Scripting.FileSystemObject
Dim kfiObject As New eKFIEngine.eKFIDelivery
Dim intPDFFontMode As Integer
Dim intPDFSecurity As Integer
Dim lngOriginalSize As Long
'Dim objKFICompression As New eKFICompression.eKFICompress
Dim objDMSCompression As New dmsCompression1
Dim blnReturnDOMDocument As Boolean
Dim strAlphaCalcs As String
Dim strCompression As String
Dim strDateFormat As String
Dim strKFIInputType As String
Dim strIntegerFormat As String
Dim strKFIDataLocation As String
Dim strKFIDataContent As String
Dim strKFIOutputLocation As String
Dim strKFIOutputType As String
Dim strKFITemplateContent As String
Dim strKFITemplateLocation As String
Dim strNumericFormat As String
Dim strPDFOwnerPassword As String
Dim strPDFUserPassword As String
Dim strXMLLocation As String
Dim strXMLContent As String
Dim strXSLLocation As String
Dim strXSLContent As String
Dim strXMLResponse As String
Dim strPath As String
Dim strWork As String
Dim X As Integer
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlDocResponse As New FreeThreadedDOMDocument40
Dim xmlInputDoc As New FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlRequestDoc As New FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode

' Temporary for pdf
Dim xmlTempRequestDoc As New FreeThreadedDOMDocument40
Dim xmlTempDocResponse As New FreeThreadedDOMDocument40

'Sample XML Request block layout

'<REQUEST Template="C:\KFI Print Solution\KFI Templates\KFI Single and Multi Component.rtf"
'     Data = "C:\KFI Print Solution\KFI Templates\KFI Multi Component Data.xml"
'     Output = "C:\KFI Print Solution\KFI Templates\KFI Single and Multi Component_Output.rtf"
'     AlphaCalcs = "yes"
'     XSL = "C:\KFI Print Solution\Template Processing\Alpha_Calcs_Request_Block.xsl"
'     XML = "C:\KFI Print Solution\Template Processing\AlphaCalcs_Results_Processor.xml"
'     OutputType="rtf"/>

    On Error GoTo ProcessKFIError:
'Set Defaults

    If IsMissing(ReturnDOMDocument) Then
        blnReturnDOMDocument = False
    Else
        blnReturnDOMDocument = ReturnDOMDocument
    End If
    strAlphaCalcs = "no"
    strKFIOutputType = "rtf"
    strKFIInputType = "rtf"
    strDateFormat = "dd mmm yyyy"
    strIntegerFormat = "#,##0"
    strNumericFormat = "#,##0.00"
    If blnReturnDOMDocument Then
        Set ProcessKFI = Nothing
    Else
        ProcessKFI = ""
    End If
    
'Validate Request
    If strXMLRequest = "" Then
        Err.Raise 1, cFunction, "No Request xml supplied"
    End If
    xmlInputDoc.setProperty "NewParser", True
    xmlInputDoc.loadXML strXMLRequest
    If xmlInputDoc.parseError.errorCode <> 0 Then
        Err.Raise xmlInputDoc.parseError.errorCode, cFunction, xmlInputDoc.parseError.reason
    End If
    
    Set xmlNode = xmlInputDoc.selectSingleNode("REQUEST")
    If xmlNode Is Nothing Then
        Err.Raise 2, cFunction, "Request xml does not contain a REQUEST node"
    End If
    
    For Each xmlAttribute In xmlNode.Attributes
        strWork = xmlAttribute.Text
        Select Case xmlAttribute.nodeName
            Case "Template"
                strKFITemplateLocation = strWork
            Case "Data"
                strKFIDataLocation = strWork
            Case "Output"
                strKFIOutputLocation = strWork
            Case "AlphaCalcs"
                strAlphaCalcs = strWork
            Case "XSL"
                strXSLLocation = strWork
            Case "XML"
                strXMLLocation = strWork
            Case "InputType"
                strKFIInputType = strWork
            Case "OutputType"
                strKFIOutputType = strWork
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
                strCompression = strWork
        End Select
    Next
    
    If strKFITemplateLocation = "" Then
        Err.Raise 3, cFunction, "No KFI template details supplied"
    End If
    If strKFIDataLocation = "" Then
        Err.Raise 4, cFunction, "No KFI data details supplied"
    End If
    
    Select Case strKFIOutputType
        Case "rtf", "doc", "pdf", "htm"
        Case Else
            Err.Raise 6, cFunction, "KFI output type invalid"
    End Select
    
    strKFITemplateContent = getXML(strKFITemplateLocation, 1)
    strKFIDataContent = getXML(strKFIDataLocation, 0)
    
    If strKFITemplateContent = "" Then
        Err.Raise 7, cFunction, "KFI template not found"
    End If
    
    If strKFIDataContent = "" Then
        Err.Raise 8, cFunction, "KFI data not found"
    End If
    
    If strKFIOutputLocation <> "" Then
        strPath = strKFIOutputLocation
        X = InStrRev(strKFIOutputLocation, "\")
        If X > 0 Then
            strPath = Left$(strPath, X - 1)
        End If
        If Not fs.FolderExists(strPath) Then
            Err.Raise 9, cFunction, "Output location folder does not exist"
        End If
    End If
    
    If strAlphaCalcs = "yes" Then
        If strXMLLocation = "" Then
            Err.Raise 10, cFunction, "No AlphaCalcs XML details supplied"
        End If
        If strXSLLocation = "" Then
            Err.Raise 11, cFunction, "No AlphaCalcs XSL details supplied"
        End If
        strXMLContent = getXML(strXMLLocation, 0)
        strXSLContent = getXML(strXSLLocation, 0)
        
        If strXMLContent = "" Then
            Err.Raise 12, cFunction, "AlphaCalcs XML data not found"
        End If
        
        If strXSLContent = "" Then
            Err.Raise 13, cFunction, "AlphaCalcs XSL data not found"
        End If
    
    End If
    
    Select Case strCompression
        Case "KFI", "DMS", ""
        Case Else
            Err.Raise 14, cFunction, "Compression type must be KFI or DMS"
    End Select

    Set xmlRequestDoc = New FreeThreadedDOMDocument40

    Set xmlRequestNode = xmlRequestDoc.createElement("REQUEST")
    xmlRequestDoc.appendChild xmlRequestNode
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("InputContent")
    xmlAttribute.Text = strKFIInputType
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("OutputType")
    xmlAttribute.Text = strKFIOutputType
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlNode = xmlRequestDoc.createElement("TEMPLATE")
    xmlNode.dataType = "bin.base64"
    arrBytes = Replace(strKFITemplateContent, vbCrLf, "")
    
    lngOriginalSize = UBound(arrBytes, 1)
    objDMSCompression.CompressionMethod = "ZLIB"
    arrBytes = objDMSCompression.SafeArrayCompressToSafeArray(arrBytes())
    xmlNode.nodeTypedValue = arrBytes
    xmlRequestNode.appendChild xmlNode
    
    Set xmlNode = xmlRequestDoc.createElement("TEMPLATESIZE")
    xmlNode.Text = lngOriginalSize
    xmlRequestNode.appendChild xmlNode
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("AlphaCalcs")
    xmlAttribute.Text = strAlphaCalcs
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("DateFormat")
    xmlAttribute.Text = strDateFormat
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("IntegerFormat")
    xmlAttribute.Text = strIntegerFormat
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("NumericFormat")
    xmlAttribute.Text = strNumericFormat
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("PDFFontMode")
    xmlAttribute.Text = intPDFFontMode
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("PDFSecurity")
    xmlAttribute.Text = intPDFSecurity
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("PDFUserPassword")
    xmlAttribute.Text = strPDFUserPassword
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("PDFOwnerPassword")
    xmlAttribute.Text = strPDFOwnerPassword
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
    Set xmlNode = xmlRequestDoc.createElement("INPUT")
    xmlNode.Text = strKFIDataContent
    xmlRequestNode.appendChild xmlNode
    
    Set xmlNode = xmlRequestDoc.createElement("XML")
    xmlNode.Text = strXMLContent
    xmlRequestNode.appendChild xmlNode
    
    Set xmlNode = xmlRequestDoc.createElement("XSL")
    xmlNode.Text = strXSLContent
    xmlRequestNode.appendChild xmlNode
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("Compression")
    xmlAttribute.Text = strCompression
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    Set xmlDocResponse = kfiObject.CreateKFI(xmlRequestDoc)
    If xmlDocResponse Is Nothing Then
        App.LogEvent cFunction & " No Response Document", vbLogEventTypeWarning
    Else
        Set xmlNode = xmlDocResponse.selectSingleNode("//RESPONSE")
        If Not xmlNode Is Nothing Then
            Set xmlAttribute = xmlNode.Attributes.getNamedItem("Type")
            If Not xmlAttribute Is Nothing Then
                If xmlAttribute.Text = "SUCCESS" Then
                    If strKFIOutputLocation <> "" Then
'Extract the processed data from the RESPONSE block
                        lngOriginalSize = Val(xmlDocResponse.selectSingleNode("//RESPONSE/DOCUMENTCONTENTS/@FILESIZE").Text)
                        Set xmlNode = xmlDocResponse.selectSingleNode("//RESPONSE/DOCUMENTCONTENTS/@FILECONTENTS")
                        xmlNode.dataType = "bin.base64"
                        If strKFIOutputType = "rtf" Or strKFIOutputType = "htm" Then
                            Select Case strCompression
                                Case ""
                                    strWork = xmlNode.nodeTypedValue
                                Case "KFI"
                                    arrBytes = xmlNode.nodeTypedValue
                                    objDMSCompression.CompressionMethod = "ZLIB"
                                    strWork = objDMSCompression.SafeArrayDecompressToBSTR(arrBytes())
                                Case "DMS"
                                    objDMSCompression.CompressionMethod = "COMPAPI"
                                    arrBytes = xmlNode.nodeTypedValue
                                    strWork = objDMSCompression.SafeArrayDecompressToBSTR(arrBytes)
                            End Select
                        Else
                            arrBytes = xmlNode.nodeTypedValue
                            Select Case strCompression
                                Case ""
                                Case "KFI"
                                    objDMSCompression.CompressionMethod = "ZLIB"
                                    arrBytes = objDMSCompression.SafeArrayDecompressToSafeArray(arrBytes())
                                Case "DMS"
                                    objDMSCompression.CompressionMethod = "COMPAPI"
                                    arrBytes = objDMSCompression.SafeArrayDecompressToBSTR(arrBytes)
                            End Select
                        End If
                        If strKFIOutputType = "rtf" Or strKFIOutputType = "htm" Then
                            Call WriteFileData(strKFIOutputLocation, strWork)
                        Else
                            Call WriteFileData(strKFIOutputLocation, arrBytes)
                        End If
                    
                    End If
                End If
            End If
        End If
    End If

TidyUp:
    On Error Resume Next
    If Not xmlDocResponse Is Nothing Then
        If strKFIOutputLocation <> "" Then
            Set xmlNode = xmlDocResponse.selectSingleNode("//RESPONSE")
            Set xmlTempNode = xmlNode.selectSingleNode("DOCUMENTCONTENTS")
            If Not xmlTempNode Is Nothing Then
                xmlNode.removeChild xmlTempNode
            End If
        End If
        If blnReturnDOMDocument Then
            Set ProcessKFI = xmlDocResponse
        Else
            ProcessKFI = xmlDocResponse.xml
        End If
    End If
    Set fs = Nothing
    Set kfiObject = Nothing
    Set xmlAttribute = Nothing
    Set xmlDocResponse = Nothing
    Set xmlInputDoc = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlTempNode = Nothing
    
    Exit Function
ProcessKFIError:
    App.LogEvent cFunction & " " & Err.Description, vbLogEventTypeWarning
        
    Set xmlDocResponse = New FreeThreadedDOMDocument40
        
    Set xmlNode = xmlDocResponse.createElement("RESPONSE")
    xmlDocResponse.appendChild xmlNode
    
    Set xmlAttribute = xmlDocResponse.createAttribute("TYPE")
    xmlAttribute.Text = "ERROR"
    xmlNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlTempNode = xmlDocResponse.createElement("ERROR")
    xmlNode.appendChild xmlTempNode
    
    Set xmlNode = xmlTempNode
    
    Set xmlTempNode = xmlDocResponse.createElement("NUMBER")
    xmlTempNode.Text = Err.Number
    xmlNode.appendChild xmlTempNode
    
    Set xmlTempNode = xmlDocResponse.createElement("DESCRIPTION")
    xmlTempNode.Text = Err.Description
    xmlNode.appendChild xmlTempNode
    
    Set xmlTempNode = xmlDocResponse.createElement("SOURCE")
    xmlTempNode.Text = cFunction
    xmlNode.appendChild xmlTempNode
    
    Resume TidyUp:
End Function

Private Sub WriteFileData(strPath As String, strData)
Dim arrBytes() As Byte
Dim strWork As String
Dim F As Integer

    F = FreeFile()
    On Error GoTo WriteFileDataError:
    Open strPath For Output Access Write As F
    Close F
    Open strPath For Binary Access Write As F
    
    If VarType(strData) = vbString Then
        strWork = strData
        Put F, , strWork
    Else
        arrBytes = strData
        Put F, , arrBytes
    End If

TidyUp:
    On Error Resume Next
    Close F
    On Error GoTo 0
    Exit Sub

WriteFileDataError:
    App.LogEvent cFunction & " " & Err.Description & " WriteFileData Can't write to file '" & strPath & "'"
    Resume TidyUp:
End Sub


Private Sub ObjectControl_Activate()

End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function


Private Sub ObjectControl_Deactivate()

End Sub


