VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'Workfile:          clsTemplateProcessor.cls

'Prog   Date        Description
'IK     19/09/2006  CORE300 - EP1146
'                   Wrap rtf data with a CDATA tag as RTF data is not always XML compatable.
'TW     25/10/2006  CORE307 Deal with Word 2003 problem

'---------------------------------------------------------------------------------------------

Option Explicit

Dim lngErrorState As Long
Dim strErrorResponse As String
Dim strIntegerFormat As String
Dim strNumericFormat As String
Dim strDateFormat As String

Private Type AutoNumberStruct
    Name As String
    Number As Integer
    Step As Integer
End Type

Implements ObjectControl

Private Function BuildxmlSubDataStructure(xmlDataNode As IXMLDOMNode) As FreeThreadedDOMDocument40
Dim xmlSubData As New FreeThreadedDOMDocument40
Dim arrElements() As String
Dim xmlParentNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim X As Integer

    arrElements = Split(CreateXMLPath(xmlDataNode), "/")
    xmlSubData.preserveWhiteSpace = True
    Set xmlParentNode = xmlSubData.createElement(arrElements(0))
    
    xmlSubData.appendChild xmlParentNode
    For X = 1 To UBound(arrElements, 1) - 1
        Set xmlNode = xmlSubData.createElement(arrElements(X))
        xmlParentNode.appendChild xmlNode
        Set xmlParentNode = xmlNode
    Next X
    xmlParentNode.appendChild xmlDataNode
    Set BuildxmlSubDataStructure = xmlSubData
    
'Tidy up
    Erase arrElements
    Set xmlNode = Nothing
    Set xmlParentNode = Nothing
    Set xmlSubData = Nothing
End Function

Public Function ConvertTextTemplateToXML(strTemplateData As String) As String
Const cFunction = "TemplateProcessor.Processor.ConvertTextTemplateToXML"
'Not used yet
End Function

Public Function DeliverAsWordDoc(strRTF As String) As Variant
Const cFunction = "TemplateProcessor.Processor.DeliverAsWordDoc"

Dim arrBytes() As Byte
Dim F As Integer
Dim fs As New Scripting.FileSystemObject
Dim L As Long
Dim objWord As Object
Dim objWordDoc As Object
Dim DOCFileName As String
Dim RTFFileName As String

    On Error GoTo DeliverAsPDFError:
'Create Temporary Files
    RTFFileName = App.Path & "\" & fs.GetTempName
    DOCFileName = App.Path & "\" & fs.GetTempName
    
    F = FreeFile()
    Open RTFFileName For Binary Access Write As F
    Put F, , strRTF
    Close F
    
    Set objWord = CreateObject("Word.Application")
    Set objWordDoc = objWord.Documents.Open(RTFFileName, False)
    
    objWordDoc.SaveAs DOCFileName, 0 'Save as document format
    objWordDoc.Close
    objWord.Quit

    L = FileLen(DOCFileName)
    ReDim arrBytes(L - 1)
    Open DOCFileName For Binary Access Read As F
    Get F, , arrBytes
    Close F

    DeliverAsWordDoc = arrBytes
    
DeliverAsPDFError:
    
    lngErrorState = Err.Number
    strErrorResponse = Err.Description
    Set objWord = Nothing
    Set objWordDoc = Nothing

'Delete Temporary Files
    If fs.FileExists(RTFFileName) Then
        Kill RTFFileName
    End If
    If fs.FileExists(DOCFileName) Then
        Kill DOCFileName
    End If
    
    Set fs = Nothing
End Function

Public Function DeliverAsXML(strData() As Byte) As IXMLDOMNode
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlNode As IXMLDOMNode
Dim xmlTempDoc As New FreeThreadedDOMDocument40
    Set xmlNode = xmlTempDoc.createElement("DOCUMENTCONTENTS")
    
    Set xmlAttribute = xmlTempDoc.createAttribute("FILEVERSION")
    xmlAttribute.Text = "1"
    xmlNode.Attributes.setNamedItem xmlAttribute

    Set xmlAttribute = xmlTempDoc.createAttribute("FILESIZE")
    xmlNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlTempDoc.createAttribute("FILECONTENTS_TYPE")
    xmlAttribute.Text = "BIN.BASE64"
    xmlNode.Attributes.setNamedItem xmlAttribute

    Set xmlAttribute = xmlTempDoc.createAttribute("FILECONTENTS")
    xmlAttribute.dataType = "bin.base64"
    xmlAttribute.nodeTypedValue = strData
    xmlNode.Attributes.setNamedItem xmlAttribute
    
    Set DeliverAsXML = xmlNode
    
    Set xmlAttribute = Nothing
    Set xmlNode = Nothing
    Set xmlTempDoc = Nothing
End Function


Public Function ConvertHTMLTemplateToXML(strInputTemplateData As String) As String
Const cFunction = "TemplateProcessor.Processor.ConvertHTMLTemplateToXML"
Dim arrElements() As String
Dim arrWords() As String
Dim arrXML() As String
Dim C As Integer
Dim strRTF As String
Dim strTemplateData As String
Dim strTag As String
Dim strTagName As String
Dim strWork As String
Dim strXML As String
Dim X As Long
Dim Y As Long
    
'This routine converts  the input html string to xml
    On Error GoTo ConvertHTMLTemplateToXML:
    strWork = strInputTemplateData

'Remove html [if ... lines
    Do
        X = InStr(1, strWork, "<![if")
        If X = 0 Then
            Exit Do
        End If
        Y = InStr(X, strWork, "<![endif]>")
        strWork = Left$(strWork, X - 1) & Mid$(strWork, Y + 10)
    Loop

'Remove html comment lines
    Do
        X = InStr(1, strWork, "<!")
        If X = 0 Then
            Exit Do
        End If
        Y = InStr(X, strWork, ">")
        strWork = Left$(strWork, X - 1) & Mid$(strWork, Y + 1)
    Loop
    
    
'Remove unuseable references
    strWork = Replace(strWork, "&nbsp;", "")
    strWork = Replace(strWork, "ü", "") 'Don't know what this is but it causes the xml parser to fail !
    
'Remove unwanted paragraphs
'    Do
'        X = InStr(1, strWork, ">[*")
'        If X = 0 Then
'            X = InStr(1, strWork, ">[/")
'            If X = 0 Then
'                Exit Do
'            End If
'        End If
'        Y = InStrRev(strWork, "<p class", X)
'        strWork = Left$(strWork, Y - 1) & Mid$(strWork, X + 1)
'    Loop
    
'Remove paragraph following 'Tag'
    strWork = Replace(strWork, "]</span></p>", "]</span>")
    
'Convert html tags to xml equivalents
    strWork = Replace(strWork, "<", "&lt;")
    strWork = Replace(strWork, ">", "&gt;")
    
    ReDim arrXML(0 To 0)
    
    arrXML(0) = "<TEMPLATEDATA>"
    strTemplateData = Replace(strWork, vbLf, "")
    X = 1
        
    Y = InStr(X, strTemplateData, "[") + 1
    Do While Y > 1
        strRTF = Mid$(strTemplateData, X, Y - X - 1)
        X = InStr(Y, strTemplateData, "]") + 1
        If X = 1 Then
            Err.Raise 65006, cFunction, "Badly formed tag in rtf"
        End If
        strTag = Mid$(strTemplateData, Y, X - Y - 1)
        
        strTagName = Mid$(strTag, 2)
        
        If Left$(strTag, 1) = "@" Then
            strXML = "<" & strTagName & "/>"                    'Empty tag for attribute data
        Else
'            If Mid$(strTemplateData, X, 5) = "\par " Then
'                X = X + 5
'            End If
            If Mid$(strTemplateData, X, 4) = "</p>" Then
                X = X + 4
            End If
            Select Case Left$(strTag, 1)
                Case "*"
                    strXML = "<" & strTagName & ">"             'Start of block tag
                Case "/"
                    strXML = "</" & Mid$(strTagName, 2) & ">"   'End of block tag
                Case "+"
                    strXML = "<AUTONUMBERING>" & strTagName & "</AUTONUMBERING>"    'Auto numbering identity
                Case "="
                    strXML = "<_" & strTagName & "/>"           'Empty tag for total of named attribute data
                Case Else
                    Err.Raise 65007, cFunction, "Unknown tag type in rtf - '" & strTagName
            End Select
        End If
        C = C + 1
        ReDim Preserve arrXML(0 To C)
        arrXML(C) = "<RTF>" & strRTF & "</RTF>" & strXML & vbCrLf
        Y = InStr(X, strTemplateData, "[") + 1
    Loop
    strXML = Mid$(strTemplateData, X)
    C = C + 1
    ReDim Preserve arrXML(0 To C)
    arrXML(C) = "<RTF>" & strXML & "</RTF></TEMPLATEDATA>"
    ConvertHTMLTemplateToXML = Join(arrXML, "")
    
ConvertHTMLTemplateToXML:
    
'Tidy up
    Erase arrElements
    Erase arrWords
    Erase arrXML
    
    lngErrorState = Err.Number
    strErrorResponse = Err.Description

End Function


Private Function CreateXMLPath(xmlNode As IXMLDOMNode) As String
Dim xmlParentNode As IXMLDOMNode
Dim strPath As String

    Set xmlParentNode = xmlNode.parentNode
    strPath = "@" & xmlNode.nodeName
    Do While xmlParentNode.nodeName <> "#document"
        strPath = xmlParentNode.nodeName & "/" & strPath
        Set xmlParentNode = xmlParentNode.parentNode
    Loop
    CreateXMLPath = strPath
    
'Tidy up
    Set xmlParentNode = Nothing
End Function

Private Function DealWithSpecialCharacters(strData As String) As String
Dim strWork As String

    strWork = Replace(strData, vbCrLf, "")
    strWork = Replace(strWork, "&", "&amp;")
    strWork = Replace(strWork, "<", "&lt;")
    strWork = Replace(strWork, "'", "&apos;")
    strWork = Replace(strWork, Chr$(34), "&quot;")
    DealWithSpecialCharacters = Replace(strWork, ">", "&gt;")
End Function

Public Property Get ErrorResponse() As String
    ErrorResponse = strErrorResponse
End Property

Public Property Get ErrorState() As Long
    ErrorState = lngErrorState
End Property


Private Function FollowHeirarchy(xmlNode As IXMLDOMNode, strXMLPath As String, strNodeName As String) As Boolean
Dim strPath As String
Dim strTempPath As String
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlChildNode As IXMLDOMNode
Dim xmlSubNode As IXMLDOMNode
Dim xmlTempAttribute As IXMLDOMAttribute
Dim xmlTempNode As IXMLDOMNode
    
    For Each xmlChildNode In xmlNode.childNodes
        strPath = strXMLPath & "/" & xmlChildNode.nodeName
        Set xmlTempNode = xmlNode.ownerDocument.selectSingleNode(strPath)
        If Not xmlTempNode Is Nothing Then
            Set xmlTempAttribute = xmlTempNode.Attributes.getNamedItem(strNodeName)
            If Not xmlTempAttribute Is Nothing Then
                strXMLPath = "sum(" & strPath & "/@" & strNodeName & ")"
                
                Set xmlTempNode = xmlTempNode.parentNode
                strTempPath = Replace(CreateXMLPath(xmlTempNode), "@", "")
                For Each xmlSubNode In xmlTempNode.childNodes
                    If xmlSubNode.nodeName <> xmlChildNode.nodeName Then
                        strPath = strTempPath & "/" & xmlSubNode.nodeName
                        Set xmlTempNode = xmlNode.ownerDocument.selectSingleNode(strPath)
                        If Not xmlTempNode Is Nothing Then
                            Set xmlTempAttribute = xmlTempNode.Attributes.getNamedItem(strNodeName)
                            If Not xmlTempAttribute Is Nothing Then
                                strXMLPath = strXMLPath & " + sum(" & strPath & "/@" & strNodeName & ")"
                            End If
                        End If
                    End If
                Next
                FollowHeirarchy = True
                Exit Function
            End If
            If xmlTempNode.hasChildNodes Then
                strXMLPath = strPath
                If FollowHeirarchy(xmlTempNode, strXMLPath, strNodeName) Then
                    Exit Function
                End If
            End If
        End If
    Next
    
End Function


Private Function FormatField(strData As String, blnCurrency As Boolean) As String
Dim strWork As String
    strWork = Replace(strData, ",", "")
    If IsNumeric(strWork) Then
        If InStr(1, strWork, ".") = 0 Then
            FormatField = Format$(strWork, strIntegerFormat)
        Else
            If blnCurrency Then
                FormatField = Format$(strWork, "#,##0.00")
            Else
                FormatField = Format$(strWork, strNumericFormat)
            End If
        End If
    Else
        If IsDate(strWork) Then
            FormatField = Format$(strWork, strDateFormat)
        Else
            FormatField = strData
        End If
    End If

End Function

Public Property Let NumericFormat(strFormat As String)
    strNumericFormat = strFormat
End Property

Public Property Let DateFormat(strFormat As String)
    strDateFormat = strFormat
End Property

Public Property Let IntegerFormat(strFormat As String)
    strIntegerFormat = strFormat
End Property



Public Function ProcessXMLTemplate(strInputData As String, strTemplateData As String) As String
Const cFunction = "TemplateProcessor.Processor.ProcessXMLTemplate"

Dim arrRTF() As String

Dim blnWork As Boolean
Dim AutoNumbering() As AutoNumberStruct
Dim xmlData As New FreeThreadedDOMDocument40
Dim xmlMainDoc As New FreeThreadedDOMDocument40
Dim xmlTemplate As New FreeThreadedDOMDocument40

'This routine controls the merging of the input data into the input template
    On Error GoTo ProcessXMLTemplateError:
    Debug.Print "ProcessXMLTemplate Start " & Timer
    
    ReDim arrRTF(0 To 0)

'Check that InputData is not empty
    If strInputData = "" Then
        Err.Raise 65001, cFunction, "No input data supplied"
    End If

'Check that TemplateData is not empty
    If strTemplateData = "" Then
        Err.Raise 65002, cFunction, "No template data supplied"
    End If

    xmlData.setProperty "NewParser", True
    xmlData.loadXML strInputData

'Check that InputData contains a valid xml block
    If xmlData.parseError.errorCode <> 0 Then
        Err.Raise 65003, cFunction, "Input data is not a valid xml block - Line: " & xmlData.parseError.Line & " Pos: " & xmlData.parseError.linepos & " Reason: " & xmlData.parseError.reason
    End If
    
    xmlMainDoc.setProperty "NewParser", True
    xmlMainDoc.loadXML strInputData
    
'The template is in Element Format, therefore the whitespace needs to be presrved on loading
    xmlTemplate.setProperty "NewParser", True
    xmlTemplate.preserveWhiteSpace = True
    xmlTemplate.loadXML strTemplateData
    
'Ensure leading and trailing spaces in data are preserved
    xmlData.preserveWhiteSpace = True
    xmlMainDoc.preserveWhiteSpace = True
    
    If xmlTemplate.parseError.errorCode <> 0 Then
        Call WriteFileData(App.Path & "\xmlDebug.xml", strTemplateData)
        Err.Raise 65005, cFunction, "RTF Template data is not a valid xml block - Line: " & xmlTemplate.parseError.Line & " Pos: " & xmlTemplate.parseError.linepos & " Reason: " & xmlTemplate.parseError.reason
    End If
    
    Call SetUpAutoNumbering(AutoNumbering, xmlData)
    
    Call ProcessXMLTemplateNode(arrRTF, xmlTemplate.documentElement, xmlData, AutoNumbering, xmlMainDoc, blnWork)
    
    ProcessXMLTemplate = Join(arrRTF, "")
    
    Debug.Print "ProcessXMLTemplate End " & Timer
    
ProcessXMLTemplateError:
    lngErrorState = Err.Number
    strErrorResponse = Err.Description

'Tidy up
    Set xmlMainDoc = Nothing
    Set xmlTemplate = Nothing
    Set xmlData = Nothing
    
    Erase AutoNumbering
    Erase arrRTF
End Function

Private Function ValidateTag(strTagName As String) As Boolean
Dim X As Integer

'Check that the tag supplied conforms to xml standards
    ValidateTag = True
    If Left$(strTagName, 1) = "*" Then
        X = 2
    Else
        X = 1
    End If
    For X = X To Len(strTagName)
        Select Case Mid$(strTagName, X, 1)
            Case "A" To "Z", "0" To "9", "_"
            Case Else
                ValidateTag = False
                Exit Function
        End Select
    Next X
End Function

Private Sub WriteFileData(strPath As String, strData As String)
Dim F As Integer

    F = FreeFile()
    Open strPath For Output Access Write As F
    Close F
    Open strPath For Binary Access Write As F
    Put F, , strData
    Close F
End Sub


Private Sub ProcessXMLTemplateNode(arrRTF() As String, xmlNode As IXMLDOMNode, xmlData As FreeThreadedDOMDocument40, AutoNumbering() As AutoNumberStruct, xmlMainDoc As FreeThreadedDOMDocument40, blnCurrency As Boolean)
Dim C As Integer
Dim xmlChildNode As IXMLDOMNode
Dim xmlDataNode As IXMLDOMNode
Dim xmlDataNodeList As IXMLDOMNodeList
Dim xslSheet As New FreeThreadedDOMDocument40
Dim xmlSubData As New FreeThreadedDOMDocument40
Dim strWork As String
Dim strXMLPath As String

'This routine process a template node and merges the relevant data
    Select Case xmlNode.nodeName
        Case "RTF"
            blnCurrency = False
            strWork = xmlNode.Text
            If Len(strWork) > 0 Then
                blnCurrency = (InStrRev(strWork, "\'a3") > Len(strWork) - 5)
                Call WriteRTFElement(arrRTF, strWork)
            End If
            Exit Sub
        Case "AUTONUMBERING"
            strWork = xmlNode.Text
            For C = 1 To UBound(AutoNumbering, 1)
                If AutoNumbering(C).Name = strWork Then
                    strWork = AutoNumbering(C).Number
                    AutoNumbering(C).Number = AutoNumbering(C).Number + AutoNumbering(C).Step
                    Call WriteRTFElement(arrRTF, strWork)
                    Exit Sub
                End If
            Next C
        Case "TODAY"
            strWork = FormatField(Now(), False)
            Call WriteRTFElement(arrRTF, strWork)
            Exit Sub
        Case Else
            If Left$(xmlNode.nodeName, 1) = "_" Then
'This routine handles the totalling function
                strXMLPath = CreateXMLPath(xmlNode.parentNode)                             'Build the full path for the current node as an attribute
'                strXMLPath = "//" & Replace(strXMLPath, "@", "")
                strXMLPath = Replace(strXMLPath, "@", "")
                Set xmlDataNode = xmlMainDoc.selectSingleNode(strXMLPath)
                Call FollowHeirarchy(xmlDataNode, strXMLPath, Mid$(xmlNode.nodeName, 2))
                
                xslSheet.setProperty "NewParser", True
                xslSheet.loadXML CreateXSLTemplate(strXMLPath)
                strWork = GetResponseValueUsingXSL(xmlMainDoc, xslSheet, blnCurrency)
'                strWork = "{\field{\*\fldinst { =SUM(ABOVE) \\# ""#,##0.00"" }}{\fldrslt {}}}"
                Call WriteRTFElement(arrRTF, strWork)
                Exit Sub
            End If
    
'1. Treat as attribute in current node
'            strXMLPath = "//" & CreateXMLPath(xmlNode)                             'Build the full path for the current node as an attribute
            strXMLPath = CreateXMLPath(xmlNode)                             'Build the full path for the current node as an attribute
            Set xmlDataNode = xmlData.selectSingleNode(strXMLPath)          'Look in current node
            If xmlDataNode Is Nothing Then
'2. Treat as child of current node
                strXMLPath = Replace(strXMLPath, "@", "")
                Set xmlDataNodeList = xmlData.selectNodes(strXMLPath)   'Look for node(s)
                Select Case xmlDataNodeList.length
                    Case 0      'Node not found
'3. Treat as attribute in root node (TEMPLATEDATA)
                        Set xmlDataNode = xmlMainDoc.selectSingleNode("//TEMPLATEDATA/@" & xmlNode.nodeName)          'Look for attribute at top level
                        If xmlDataNode Is Nothing Then
                            Exit Sub
                        End If
                        strWork = FormatField(xmlDataNode.Text, blnCurrency)
                        Call WriteRTFElement(arrRTF, strWork)
                    Case 1      'Single node found
                    Case Else   'Multiple nodes found
                        For Each xmlDataNode In xmlDataNodeList
                            xmlSubData.preserveWhiteSpace = True
                            Set xmlSubData = BuildxmlSubDataStructure(xmlDataNode)
                            Call ProcessXMLTemplateNode(arrRTF, xmlNode, xmlSubData, AutoNumbering, xmlMainDoc, blnCurrency)
                        Next
                        Exit Sub
                End Select
            Else
                strWork = FormatField(xmlDataNode.Text, blnCurrency)
                Call WriteRTFElement(arrRTF, strWork)
            End If
            
    End Select
    
    If xmlNode.hasChildNodes Then
        For Each xmlChildNode In xmlNode.childNodes
            If xmlChildNode.nodeName <> "#text" Then
                Call ProcessXMLTemplateNode(arrRTF, xmlChildNode, xmlData, AutoNumbering, xmlMainDoc, blnCurrency)
            End If
        Next
    End If

'Tidy up
    Set xmlChildNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlDataNodeList = Nothing
    Set xmlSubData = Nothing
    Set xslSheet = Nothing
End Sub


Private Function GetResponseValueUsingXSL(xmlIn As FreeThreadedDOMDocument40, xslSheet As FreeThreadedDOMDocument40, blnCurrency As Boolean) As String
Dim strXML As String
Dim xmlTemp As New FreeThreadedDOMDocument40
Dim X As Long
Dim xslProcessor As Object
Dim xslTemplate As New XSLTemplate40
    
    On Error GoTo GetResponseValueUsingXSLError:
    Set xslTemplate.stylesheet = xslSheet
    Set xslProcessor = xslTemplate.createProcessor()
    strXML = Replace(xmlIn.xml, ",", "")
    xmlTemp.setProperty "NewParser", True
    xmlTemp.loadXML strXML
    xslProcessor.input = xmlTemp
    X = xslProcessor.Transform()
    xmlTemp.loadXML xslProcessor.output
'    GetResponseValueUsingXSL = Format$(xmlTemp.selectSingleNode("RESPONSE").Text, "#,##0.00")
    GetResponseValueUsingXSL = FormatField(xmlTemp.selectSingleNode("RESPONSE").Text, blnCurrency)

TidyUp:
    Set xslProcessor = Nothing
    Set xslTemplate = Nothing
    Set xmlTemp = Nothing
    Exit Function

GetResponseValueUsingXSLError:
    GetResponseValueUsingXSL = "XSL Error!"
    GoTo TidyUp:
End Function


Private Function CreateXSLTemplate(strArgument)
Dim xslResponse
Dim xslSheet
Dim xslStyleSheet
Dim xslTemplate
Dim xslValueOf
    
    Set xslSheet = CreateObject("MSXML2.FreeThreadedDOMDocument.4.0")
    Set xslStyleSheet = xslSheet.createElement("xsl:stylesheet")
    xslStyleSheet.setAttribute "version", "1.0"
    xslStyleSheet.setAttribute "xmlns:xsl", "http://www.w3.org/1999/XSL/Transform"
    
    Set xslTemplate = xslSheet.createElement("xsl:template")
    xslTemplate.setAttribute "match", "/"
    
    Set xslResponse = xslSheet.createElement("RESPONSE")
    
    Set xslValueOf = xslSheet.createElement("xsl:value-of")
    
    xslValueOf.setAttribute "select", strArgument
    
    xslTemplate.appendChild xslResponse
    xslResponse.appendChild xslValueOf
    xslStyleSheet.appendChild xslTemplate
    CreateXSLTemplate = xslStyleSheet.xml
End Function

Public Function ConvertRTFTemplateToXML(strInputTemplateData As String) As String
Const cFunction = "TemplateProcessor.Processor.ConvertRTFTemplateToXML"

Dim arrXML() As String
Dim C As Long
Dim strRTF As String
Dim strTag As String
Dim strTagName As String
Dim strTemplateData As String
Dim strXML As String
Dim X As Long
Dim Y As Long
    
'This routine converts the input rtf string to xml

    On Error GoTo ConvertRTFTemplateToXMLError:
    Debug.Print "ConvertRTFTemplateToXML Start " & Timer
    ReDim arrXML(0 To 0)
    
    arrXML(0) = "<TEMPLATEDATA>"
'    strTemplateData = Replace(strInputTemplateData, vbLf, "")
    strTemplateData = Replace(strInputTemplateData, vbCrLf, "")
'    strTemplateData = Replace(strInputTemplateData, vbLf, "")

' TW 25/10/2006 CORE307 - Deal with Word 2003 problem
    strTemplateData = Replace(strTemplateData, "{\*\xmlopen\xmlns1{\factoidname stockticker}}", "")
    strTemplateData = Replace(strTemplateData, "{\*\xmlclose}", "")

    X = 1
        
    Y = InStr(X, strTemplateData, "[") + 1
    Do While Y > 1
        strRTF = Mid$(strTemplateData, X, Y - X - 1)
        X = InStr(Y, strTemplateData, "]") + 1
        If X = 1 Then
            Err.Raise 65006, cFunction, "Badly formed tag in rtf"
        End If
        strTag = Mid$(strTemplateData, Y, X - Y - 1)
        
        strTagName = Mid$(strTag, 2)
        
        If Left$(strTag, 1) = "@" Then
            strXML = "<" & strTagName & "/>"                    'Empty tag for attribute data
        Else
            If Mid$(strTemplateData, X, 5) = "\par " Then
                X = X + 5
            End If
            Select Case Left$(strTag, 1)
                Case "*"
                    strXML = "<" & strTagName & ">"             'Start of block tag
                Case "/"
                    strXML = "</" & Mid$(strTagName, 2) & ">"   'End of block tag
                Case "+"
                    strXML = "<AUTONUMBERING>" & strTagName & "</AUTONUMBERING>"    'Auto numbering identity
                Case "="
                    strXML = "<_" & strTagName & "/>"           'Empty tag for total of named attribute data
                Case Else
                    Err.Raise 65007, cFunction, "Unknown tag type in rtf - '" & strTagName
            End Select
        End If
        C = C + 1
        ReDim Preserve arrXML(0 To C)
        ' IK - 19/09/2006 CORE300
        ' Peter Edney - 15/09/2006 - EP1146
        ' Wrap rtf data with a CDATA tag as RTF data is not always XML compatable.
        arrXML(C) = "<RTF><![CDATA[" & strRTF & "]]></RTF>" & strXML & vbCrLf
        Y = InStr(X, strTemplateData, "[") + 1
    Loop
    strXML = Mid$(strTemplateData, X)
    C = C + 1
    ReDim Preserve arrXML(0 To C)
    
    ' IK - 19/09/2006 CORE300
    ' Peter Edney - 15/09/2006 - EP1146
    ' Wrap rtf data with a CDATA tag as RTF data is not always XML compatable.
    arrXML(C) = "<RTF><![CDATA[" & strXML & "]]></RTF></TEMPLATEDATA>"
    ConvertRTFTemplateToXML = Join(arrXML, "")
    
    Debug.Print "ConvertRTFTemplateToXML End " & Timer
    
ConvertRTFTemplateToXMLError:
    lngErrorState = Err.Number
    strErrorResponse = Err.Description
    
'Tidy up
    Erase arrXML
    
End Function

Private Sub SetUpAutoNumbering(AutoNumbering() As AutoNumberStruct, xmlTemplate As FreeThreadedDOMDocument40)
Dim C As Integer
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode

'Set up the Auto-numbering specifications
    ReDim AutoNumbering(0 To 0)
    Set xmlNodeList = xmlTemplate.selectNodes("//AUTONUMBERING")
    For Each xmlNode In xmlNodeList
        C = C + 1
        ReDim Preserve AutoNumbering(0 To C)
        
        AutoNumbering(C).Number = 1                  'Default start value
        AutoNumbering(C).Step = 1                    'Default step value
        For Each xmlAttribute In xmlNode.Attributes
            Select Case UCase$(xmlAttribute.Name)
                Case "NAME"
                    AutoNumbering(C).Name = xmlAttribute.Text
                Case "START"
                    AutoNumbering(C).Number = Val(xmlAttribute.Text)
                Case "STEP"
                    AutoNumbering(C).Step = Val(xmlAttribute.Text)
            End Select
        Next
        xmlNode.parentNode.removeChild xmlNode
    Next
    
'Tidy up
    Set xmlAttribute = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
End Sub


Private Sub WriteRTFElement(arrRTF() As String, strWork As String)
Dim C As Integer

    C = UBound(arrRTF, 1) + 1
    ReDim Preserve arrRTF(0 To C)
    arrRTF(C) = strWork
End Sub

'Private Sub WriteXML(strXML As String, C As Integer)
'    arrXML(C) = strXML & vbCrLf
'    C = C + 1
'End Sub
'
Private Sub WriteXMLElement(arrXML() As String, strXML As String)
Dim C As Integer
    C = UBound(arrXML, 1) + 1
    ReDim Preserve arrXML(0 To C)
    arrXML(C) = strXML & vbCrLf
End Sub

Private Sub Class_Initialize()
    strDateFormat = "dd mmm yyyy"
    strIntegerFormat = "#,##0"
    strNumericFormat = "#,##0.0#"
End Sub

Private Sub ObjectControl_Activate()

End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function


Private Sub ObjectControl_Deactivate()

End Sub


