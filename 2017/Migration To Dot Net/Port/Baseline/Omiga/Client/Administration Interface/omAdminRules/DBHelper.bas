Attribute VB_Name = "DBHelper"
Option Explicit

Private Enum GUIDSTYLE
    guidBinary = 0
    guidHyphen = 1
    guidLiteral = 2
End Enum

Public Enum DBPROVIDER
    omiga4DBPROVIDERUnknown
    omiga4DBPROVIDEROracle
    omiga4DBPROVIDERSQLServer
End Enum

Private Const m_cstrAppName = "Omiga4"
Private Const m_cstrREGISTRY_SECTION = "Database Connection"
Private Const m_cstrPROVIDER_KEY = "Provider"
Private Const m_cstrSERVER_KEY = "Server"
Private Const m_cstrDATABASE_KEY = "Database Name"
Private Const m_cstrUID_KEY = "User ID"
Private Const m_cstrPASSWORD_KEY = "Password"
Private Const m_cstrDATA_SOURCE_KEY = "Data Source"
Private Const m_cstrRETRIES_KEY = "Retries"

Private Const m_lngENTRYNOTFOUND As Long = -2147024894
Private m_xmldocSchemas As FreeThreadedDOMDocument40
Private m_eDbProvider As DBPROVIDER
Private m_intDBRetries As Integer
Private m_strDbConnectionString As String



Public Sub ConvertRecordSetToXML( _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vrst As ADODB.Recordset, _
    ByVal vblnDoComboLookUp As Boolean)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlThisSchemaNode As IXMLDOMNode
    
    Dim fld As ADODB.Field
    
    Dim strNodeName As String
        
    If Not vxmlSchemaNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
        strNodeName = vxmlSchemaNode.Attributes.getNamedItem("OUTNAME").Text
    Else
        strNodeName = vxmlSchemaNode.nodeName
    End If
    
    Do While Not vrst.EOF
    
        If vxmlResponseNode.nodeName = strNodeName Then
            Set xmlElem = vxmlResponseNode
        Else
            Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
        End If
        
        For Each xmlThisSchemaNode In vxmlSchemaNode.childNodes
        
            Set fld = Nothing
            If GetAttributeText(vxmlSchemaNode, "ENTITYTYPE") = "PROCEDURE" Then
                If GetAttributeText(xmlThisSchemaNode, "PARAMMODE") <> "IN" Then
                    Set fld = vrst.Fields.Item(xmlThisSchemaNode.nodeName)
                End If
            Else
                Set fld = vrst.Fields.Item(xmlThisSchemaNode.nodeName)
            End If
            
            If Not fld Is Nothing Then
    
                If Not IsNull(fld) Then
            
                    If Not IsNull(fld.Value) Then
                    
                        FieldToXml fld, xmlElem, xmlThisSchemaNode, vblnDoComboLookUp
            
                    End If
                
                End If
                
            End If
        
        Next
        
        If xmlElem.Attributes.length > 0 Then
        
            vxmlResponseNode.appendChild xmlElem
        
        End If
            
        vrst.MoveNext
        
    Loop

    vrst.Close
    
    Set fld = Nothing
    Set xmlElem = Nothing
    Set xmlThisSchemaNode = Nothing
    
End Sub

Private Function getParam( _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlAttrib As IXMLDOMAttribute, _
    Optional ByVal blnIsCreate As Boolean = False) _
    As ADODB.Parameter
    
    Dim param As ADODB.Parameter
    Dim strDataType As String, _
        strDataValue As String
    
    Set param = New ADODB.Parameter
        
    strDataType = vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
    
    Select Case strDataType
    
        Case "STRING"
            param.Type = adBSTR
            param.Size = Len(vxmlAttrib.Text)
    
        Case "GUID"
            param.Type = adBinary
            param.Size = 16
        
        Case "SHORT", "BOOLEAN", "COMBO", "LONG"
            param.Type = adInteger
    
        Case "DOUBLE", "CURRENCY"
            param.Type = adDouble
    
        Case "DATE", "DATETIME"
            param.Type = adDBTimeStamp
    
    End Select
    
    param.Direction = adParamInput
    param.Value = Null
    
    If Not vxmlAttrib Is Nothing Then
        If Len(vxmlAttrib.Text) > 0 Then
            If strDataType <> "GUID" Then
                param.Value = vxmlAttrib.Text
            Else
                param.Value = GuidStringToByteArray(vxmlAttrib.Text)
            End If
        End If
    End If
    
    Set getParam = param
    
    Set param = Nothing
    
End Function

Public Function executeGetRecordSet( _
    ByVal cmd As ADODB.Command) _
    As ADODB.Recordset
    
On Error GoTo executeGetRecordSetExit
    
    Const strFunctionName As String = "executeGetRecordSet"
    
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    conn.ConnectionString = adoGetDbConnectString
    conn.Open
    
    cmd.ActiveConnection = conn
    
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    
    rst.Open cmd

    ' disconnect RecordSet
    Set rst.ActiveConnection = Nothing
    
    conn.Close

    If Not rst.EOF Then
        rst.MoveFirst
        Set executeGetRecordSet = rst
    End If
    
executeGetRecordSetExit:
    
    Set rst = Nothing
    Set cmd = Nothing
    Set conn = Nothing
        
    CheckError strFunctionName

End Function

Private Sub FieldToXml( _
    ByVal vfld As ADODB.Field, _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean)
    
    Dim xmlAttrib As IXMLDOMAttribute

    Dim strElementValue As String, _
        strComboGroup As String, _
        strComboValue As String, _
        strFormatMask As String, _
        strFormatted As String
            
    Select Case vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text

        Case "GUID"

            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                GuidToString(vfld.Value)

        Case "CURRENCY"
            
            strFormatted = Format(vfld, "0.0000000")
            strFormatted = Left(strFormatted, InStr(1, strFormatted, ".") + 2)
            
            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                strFormatted

        Case "DOUBLE"
            
            strFormatMask = Empty
            
            If Not vxmlSchemaNode.Attributes.getNamedItem("FORMATMASK") Is Nothing Then
                strFormatMask = vxmlSchemaNode.Attributes.getNamedItem("FORMATMASK").Text
            End If
            
            If Len(strFormatMask) <> 0 Then
                
                vxmlOutElem.setAttribute _
                    vxmlSchemaNode.nodeName & "_RAW", _
                    vfld.Value
                
                vxmlOutElem.setAttribute _
                    vxmlSchemaNode.nodeName, _
                    Format(vfld.Value, strFormatMask)
                
            Else
                
                vxmlOutElem.setAttribute _
                    vxmlSchemaNode.nodeName, _
                    vfld.Value
            
            End If

        Case "COMBO"

            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                vfld.Value
                
            If vblnDoComboLookUp = True Then

                If Not vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP") Is Nothing Then
    
                    strComboGroup = vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP").Text
                    strComboValue = GetComboText(strComboGroup, vfld.Value)
                
                    vxmlOutElem.setAttribute _
                        vxmlSchemaNode.nodeName & "_TEXT", strComboValue
    
                End If
                
            End If

        Case "DATE"

            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                DateToString(vfld.Value)
        
        Case "DATETIME"

            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                DateTimeToString(vfld.Value)

        Case Else

            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                vfld.Value

    End Select
    
End Sub

Public Function GuidStringToByteArray(ByVal strGuid As String) As Byte()
    
    On Error GoTo GuidStringToByteArrayVbErr
    Const strFunctionName As String = "GuidStringToByteArray"

    Dim rbytGuid(15) As Byte
    
    If Len(strGuid) = 38 Then
        ' Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
        strGuid = FormatGuid(strGuid, guidBinary)
    End If
    
    If Len(strGuid) = 32 Then
        ' Convert from "DA6DA163412311D4B5FA00105ABB1680" to byte array.
        Dim intIndex As Integer
        For intIndex = 0 To UBound(rbytGuid)
            rbytGuid(intIndex) = CByte("&H" & Mid(strGuid, (intIndex * 2) + 1, 2))
        Next intIndex
    Else
        Debug.Assert 0
        Err.Raise eINVALIDPARAMETER, strFunctionName, "strGuid length must be 32"
    End If
    
    GuidStringToByteArray = rbytGuid

    Exit Function

GuidStringToByteArrayVbErr:

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function adoGetDbConnectString() As String
    adoGetDbConnectString = m_strDbConnectionString
End Function

Public Function GuidToString(bytArray() As Byte) As String
    Dim i As Integer
    Dim strGuid As String
    
    strGuid = ""
    
    For i = 0 To 15
        strGuid = strGuid & Right$(Hex$(bytArray(i) + &H100), 2)
    Next i
    
    GuidToString = strGuid

End Function

Private Function DateToString(ByVal vvarDate As Variant) As String
    
    DateToString = Right(Str(DatePart("d", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("m", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("yyyy", vvarDate)), 4)

End Function

Private Function DateTimeToString(ByVal vvarDate As Variant) As String
    
    DateTimeToString = DateToString(vvarDate) & " " & _
        Right(Str(DatePart("h", vvarDate) + 100), 2) & ":" & _
        Right(Str(DatePart("n", vvarDate) + 100), 2) & ":" & _
        Right(Str(DatePart("s", vvarDate) + 100), 2)

End Function

Private Function FormatGuid(ByVal strSrcGuid As String, Optional ByVal eGuidStyle As GUIDSTYLE = guidLiteral) As String
    On Error GoTo FormatGuidVbErr
    
    Const strFunctionName As String = "FormatGuid"
    Dim strTgtGuid As String
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
        
    ' By default return string unchanged, i.e., if style is guidHyphen but the source string
    ' is currently not in binary format, then it will be returned unchanged (perhaps it is
    ' already in hyphenated format?).
    strTgtGuid = strSrcGuid
    
    If eGuidStyle = guidHyphen Then
        ' Convert from "DA6DA163412311D4B5FA00105ABB1680" to "{63A16DDA-2341-D411-B5FA-00105ABB1680}"
        
        Debug.Assert Len(strSrcGuid) = 32
        
        If Len(strSrcGuid) = 32 Then
            strTgtGuid = "{"
            For intIndex1 = 0 To 15
                intIndex2 = ConvertGuidIndex(intIndex1)
                strTgtGuid = strTgtGuid & Mid$(strSrcGuid, (intIndex2 * 2) + 1, 2)
                If intIndex1 = 3 Or intIndex1 = 5 Or intIndex1 = 7 Or intIndex1 = 9 Then
                    strTgtGuid = strTgtGuid & "-"
                End If
            Next intIndex1
            strTgtGuid = strTgtGuid & "}"
        Else
            Err.Raise eINVALIDPARAMETER, strFunctionName, "strSrcGuid length must be 32"
        End If
        
    ElseIf eGuidStyle = guidBinary Then
        ' Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
        
        Debug.Assert Len(strSrcGuid) = 38
        
        If Len(strSrcGuid) = 38 Then
            Dim intOffset As Integer
            strTgtGuid = ""
            intOffset = 2
            For intIndex1 = 0 To 15
                intIndex2 = ConvertGuidIndex(intIndex1)
                strTgtGuid = strTgtGuid & Mid$(strSrcGuid, (intIndex2 * 2) + intOffset, 2)
                If intIndex1 = 3 Or intIndex1 = 5 Or intIndex1 = 7 Or intIndex1 = 9 Then
                    intOffset = intOffset + 1
                End If
            Next intIndex1
        Else
            Err.Raise eINVALIDPARAMETER, strFunctionName, "strSrcGuid length must be 38"
        End If
    
    ElseIf eGuidStyle = guidLiteral Then
        ' Convert guid into a format that can be used as a literal input parameter to the database.
        ' This assumes that the database type is raw for Oracle, or binary/varbinary for SQL Server.
        
        If Len(strSrcGuid) = 38 Then
            ' Guid is in hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}",
            ' so convert to binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
            strSrcGuid = FormatGuid(strSrcGuid, guidBinary)
        End If
        
        Debug.Assert Len(strSrcGuid) = 32
        
        If Len(strSrcGuid) = 32 Then
            Select Case GetDbProvider()
            Case omiga4DBPROVIDERSQLServer
                ' e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
                strTgtGuid = "0x" & strSrcGuid
            Case omiga4DBPROVIDEROracle
                ' e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
                strTgtGuid = "HEXTORAW('" & strSrcGuid & "')"
            End Select
        End If
        
    Else
        Debug.Assert 0
        Err.Raise eINVALIDPARAMETER, strFunctionName, "strGuid length must be 32"
    End If
    
    FormatGuid = strTgtGuid

    Exit Function

FormatGuidVbErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ConvertGuidIndex(ByVal intIndex1 As Integer)
    Dim intIndex2 As Integer
    Select Case intIndex1
    Case 0
        intIndex2 = 3
    Case 1
        intIndex2 = 2
    Case 2
        intIndex2 = 1
    Case 3
        intIndex2 = 0
    Case 4
        intIndex2 = 5
    Case 5
        intIndex2 = 4
    Case 6
        intIndex2 = 7
    Case 7
        intIndex2 = 6
    Case Else
        intIndex2 = intIndex1
    End Select
    ConvertGuidIndex = intIndex2
End Function

Private Function GetDbProvider() As DBPROVIDER
    GetDbProvider = m_eDbProvider
End Function


Public Sub DBAssistBuildDbConnectionString()
On Error GoTo DBAssistBuildDbConnectionStringVbErr

    Dim objWshShell As Object
    
    Dim strConnection As String, _
        strProvider As String, _
        strRegSection As String, _
        strRetries As String
        
    Dim strUserId As String
    Dim strPassword As String
    
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    strRegSection = "HKLM\SOFTWARE\" & m_cstrAppName & "\" & App.Title & "\" & m_cstrREGISTRY_SECTION & "\"
    
On Error Resume Next
    
    strProvider = objWshShell.RegRead(strRegSection & m_cstrPROVIDER_KEY)

On Error GoTo DBAssistBuildDbConnectionStringVbErr
    
    If Len(strProvider) = 0 Then
        strRegSection = "HKLM\SOFTWARE\" & m_cstrAppName & "\" & m_cstrREGISTRY_SECTION & "\"
        strProvider = objWshShell.RegRead(strRegSection & m_cstrPROVIDER_KEY)
    End If
    
    m_eDbProvider = omiga4DBPROVIDERUnknown
    
    If strProvider = "MSDAORA" Then
        m_eDbProvider = omiga4DBPROVIDEROracle
        strConnection = _
            "Provider=MSDAORA;Data Source=" & objWshShell.RegRead(strRegSection & m_cstrDATA_SOURCE_KEY) & ";" & _
            "User ID=" & objWshShell.RegRead(strRegSection & m_cstrUID_KEY) & ";" & _
            "Password=" & objWshShell.RegRead(strRegSection & m_cstrPASSWORD_KEY) & ";"
    ElseIf strProvider = "SQLOLEDB" Then
        m_eDbProvider = omiga4DBPROVIDERSQLServer
        
        ' PSC 17/10/01 SYS2815 - Start
        strUserId = ""
        strPassword = ""
        
        On Error Resume Next
                
        strUserId = objWshShell.RegRead(strRegSection & m_cstrUID_KEY)
        
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
                
        On Error GoTo DBAssistBuildDbConnectionStringVbErr

        If Err.Number <> m_lngENTRYNOTFOUND And Err.Number <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
    
        On Error Resume Next
                
        strPassword = objWshShell.RegRead(strRegSection & m_cstrPASSWORD_KEY)
        
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
                
        On Error GoTo DBAssistBuildDbConnectionStringVbErr

        If Err.Number <> m_lngENTRYNOTFOUND And Err.Number <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
        
        strConnection = _
            "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & m_cstrSERVER_KEY) & ";" & _
            "database=" & objWshShell.RegRead(strRegSection & m_cstrDATABASE_KEY) & ";"

        ' If User Id is present use SQL Server Authentication else
        ' use integrated security
        If Len(strUserId) > 0 Then
            strConnection = strConnection & "UID=" & strUserId & ";" & _
                "pwd=" & strPassword & ";"
        Else
            strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
        End If
        ' PSC 17/10/01 SYS2815 - End
    End If
           
    m_strDbConnectionString = strConnection
    
    strRetries = objWshShell.RegRead(strRegSection & m_cstrRETRIES_KEY)
    
    If Len(strRetries) > 0 Then
        m_intDBRetries = CInt(strRetries)
    End If
    
    Set objWshShell = Nothing
    
    Debug.Print strConnection
    
    Exit Sub

DBAssistBuildDbConnectionStringVbErr:
    
    Set objWshShell = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub DBAssistLoadSchema()
    
    Dim strFileSpec As String
    
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\" & App.Title & ".xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set m_xmldocSchemas = New FreeThreadedDOMDocument40
    m_xmldocSchemas.validateOnParse = False
    m_xmldocSchemas.setProperty "NewParser", True
    
    m_xmldocSchemas.async = False
    m_xmldocSchemas.Load strFileSpec

End Sub



Public Sub GetRecordSetAsXML( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    Optional vstrFilter As String, _
    Optional vstrOrderBy As String)
    
    On Error GoTo GetRecordSetAsXMLExit
    
    Const strFunctionName As String = "GetRecordSetAsXML"
        
    Dim rst As ADODB.Recordset
    Set rst = GetRecordSet(vxmlDataNode, vxmlSchemaNode, vstrFilter, vstrOrderBy)
    
    If Not rst Is Nothing Then
        
        ConvertRecordSetToXML _
            vxmlSchemaNode, _
            vxmlResponseNode, _
            rst, _
            IsComboLookUpRequired(vxmlDataNode)
            
    
    End If
    
GetRecordSetAsXMLExit:
    
    Set rst = Nothing
        
    CheckError strFunctionName
    
End Sub

Public Function GetRecordSet( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    Optional vstrFilter As String, _
    Optional vstrOrderBy As String) _
    As ADODB.Recordset
    
On Error GoTo GetRecordSetExit
    
    Const strFunctionName As String = "GetRecordSet"
    
    
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQL As String, _
        strSQLWhere As String, _
        strPattern As String, _
        strDataSource As String
        
    Dim intLoop As Integer
                
    Dim blnDbCmdOk As Boolean
    
    strDataSource = GetAttributeText(vxmlSchemaNode, "DATASRCE")
    
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaNode.nodeName
    End If
        
    Set cmd = New ADODB.Command
    
    If Not vxmlDataNode Is Nothing Then
        
        For Each xmlDataAttrib In vxmlDataNode.Attributes
        
            strPattern = xmlDataAttrib.Name & "[@DATATYPE]"
        
            Set xmlSchemaNode = vxmlSchemaNode.selectSingleNode(strPattern)
            
            If Not xmlSchemaNode Is Nothing Then
            
                If Len(strSQLWhere) <> 0 Then
                    strSQLWhere = strSQLWhere & " AND "
                End If
            
                strSQLWhere = strSQLWhere & xmlSchemaNode.nodeName & "=?"
            
                Set param = getParam(xmlSchemaNode, xmlDataAttrib)
            
                cmd.Parameters.Append param
                
            End If
        
        Next
        
    End If
    
    If Len(vstrFilter) > 0 And UCase(Left(vstrFilter, 6)) = "SELECT" Then
        
        strSQL = vstrFilter
    
    Else
        
        If Len(vstrFilter) > 0 Then
            If Len(strSQLWhere) <> 0 Then
                strSQLWhere = strSQLWhere & " AND "
            End If
            strSQLWhere = strSQLWhere & vstrFilter
        End If
        
        If cmd.Parameters.Count <> 0 Then
            strSQL = _
                "SELECT * FROM " & strDataSource & _
                " WHERE (" & strSQLWhere & ")"
        Else
            strSQL = "SELECT * FROM " & strDataSource
            If Len(strSQLWhere) <> 0 Then
                strSQL = strSQL & " WHERE (" & strSQLWhere & ")"       'JLD SYS1774
            End If
        End If
        
        If Len(vstrOrderBy) > 0 Then
            strSQL = strSQL & " ORDER BY " & vstrOrderBy
        End If
        
    End If
    
    Debug.Print "GetRecordSet strSQL: " & strSQL
    
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    
    Set GetRecordSet = executeGetRecordSet(cmd)
        
    If GetRecordSet Is Nothing Then
        Debug.Print "GetRecordSet records retrieved: 0"
    Else
        Debug.Print "GetRecordSet records retrieved: "; GetRecordSet.RecordCount
    End If
    
GetRecordSetExit:
    
    Set cmd = Nothing
    Set xmlSchemaNode = Nothing
    Set xmlDataAttrib = Nothing
    
    
    CheckError strFunctionName

End Function

Private Function IsComboLookUpRequired(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean

    If Not vxmlRequestNode.Attributes.getNamedItem("_COMBOLOOKUP_") Is Nothing Then
        IsComboLookUpRequired = _
            GetAttributeAsBoolean(vxmlRequestNode, "_COMBOLOOKUP_")
    Else
        If Not vxmlRequestNode.ownerDocument.firstChild Is Nothing Then
            IsComboLookUpRequired = _
                GetAttributeAsBoolean(vxmlRequestNode.ownerDocument.firstChild, "_COMBOLOOKUP_")
        Else
            IsComboLookUpRequired = False
        End If
    End If

End Function




