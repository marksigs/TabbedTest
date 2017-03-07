Attribute VB_Name = "adoAssistEx"
'-----------------------------------------------------------------------------
'Prog   Date        Description
'SR     05/03/2001  New public function 'adoGetOmigaNumberForDatabaseError'
'AW     20/03/2001  Added 'PARAMMODE' check to adoConvertRecordSetToXML
'AS     31/05/2001  Added adoGetValidRecordset for SQL Server Port
'AS     11/06/2001  Fixed compile error with GENERIC_SQL=1
'LD     12/06/2001  SYS2368 Modify code to get the connection string for Oracle and SQLServer
'LD     11/06/2001  SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     15/06/2001  SYS2368 SQL Server Port - Guids to be generated as type adBinary in function getParam
'APS    20/08/2001  New CRUD method enhancements
'IK     05/11/2001  add ATTRIBUTE parameter DB_REF
'                   maps ATTRIBUTE NAME value to db column
'                   add ENTITY parameter COMPONENT_REF
'                   used to retrieve component level connection string
'AL     08/02/2002  NodeType checking to enable comments on schema file
'SR     11/12/2002  BM0074 - Use NOLOCK while using MAX()
'PSC    06/01/2003  BM232 - Allow integrated security
'AS     13/11/03    CORE1 Removed GENERIC_SQL.
'SDS    20/01/2004  LIVE00009653 - NOLOCK is used only for SQLServer sql statements
'SDS    22/01/2004  LIVE00009659  Added pieces of code to support both Microsoft and Oracle OLEDB Drivers
'HMA    17/05/2004  BBG512 Add the changes which Andy Magggs has implemented in the Baseline code for KFIHelper.
'-------------------------------------------------------------------------------
Option Explicit
' ik_wip_20030430
Private gcolSchemaDocs As Collection
Private gstrActiveSchema As String
' ik_wip_20030430_ends
Private gxmldocSchemas As FreeThreadedDOMDocument40
Private gxmldocConnections As FreeThreadedDOMDocument40
Private gstrDbConnectionString As String
Private gintDbRetries As Integer
Private genumDbProvider As DBPROVIDER
' application name
Private Const gstrAppName = "Omiga4"
' registry section for database connection info
Private Const gstrREGISTRY_SECTION = "Database Connection"
' registry keys for database connection info
Private Const gstrPROVIDER_KEY = "Provider"
Private Const gstrSERVER_KEY = "Server"
Private Const gstrDATABASE_KEY = "Database Name"
Private Const gstrUID_KEY = "User ID"
Private Const gstrPASSWORD_KEY = "Password"
Private Const gstrDATA_SOURCE_KEY = "Data Source"
' registry keys for other database info
Private Const gstrRETRIES_KEY = "Retries"
 
 '=============================================
 'Enumeration Declaration Section
 '=============================================
Public Enum DBPROVIDER
    omiga4DBPROVIDERUnknown
    omiga4DBPROVIDEROracle
    omiga4DBPROVIDERMSOracle    'SDS LIVE00009659  22/01/2004 Support for MS and Oracle OLEDB Drivers
    omiga4DBPROVIDERSQLServer
End Enum
Private Enum GUIDSTYLE
    guidBinary = 0
    guidHyphen = 1
    guidLiteral = 2
End Enum
Private Const gstrISO8601 As String = "ISO8601"
Private Enum DATESTYLE
    e_dateUK = 0          'dd/mm/yyyy
    e_dateISO8601 = 1     'yyyy-mm-yy
End Enum
Private geDateStyle As DATESTYLE
Private gstrBooleanTrue As String
Private gstrBooleanFalse As String
Public Const gstrSchemaNameId As String = "_SCHEMA_"
Public Const gstrExtractTypeId As String = "_EXTRACTTYPE_"
Public Const gstrExtractTypeNode As String = "NODE"
Public Const gstrOutNodeId As String = "_OUTNODENAME_"
Public Const gstrOrderById As String = "_ORDERBY_"
Public Const gstrWhereId As String = "_WHERE_"
Public Const gstrDataTypeId As String = "DATATYPE"
Public Const gstrDataTypeDerivedId As String = "DERIVED"
Public Const gstrDataSourceId As String = "DATASRCE"
Public Const gstrDbRefId As String = "DB_REF"
Public Const gstrComponentRefId As String = "COMPONENT_REF"
'APS - Implemented IK's CRUD changes 20/08/01
Public Const gstrSchemaRefId = "SCHEMA_REF"
Private Const gstrMODULEPREFIX As String = "adoAssist."
Private Const glngENTRYNOTFOUND As Long = -2147024894       ' PSC 06/01/03 BM232
' IK_WIP (incorporate omRb style schema + processing)
Public Enum SCHEMASTYLE
    omiga4SCHEMASTYLEUndefined
    omiga4SCHEMASTYLEOriginal
    omiga4SCHEMASTYLEPlus
End Enum
Private genumSchemaStyle As SCHEMASTYLE
' IK_WIP ends

' ik_wip_20040331
Public Sub adoEmptySchemaCache()
    Dim xmlDoc As FreeThreadedDOMDocument40
'   ik_debug
'   App.LogEvent "adoEmptySchemaCache, gcolSchemaDocs.Count: " & gcolSchemaDocs.Count, vbLogEventTypeInformation
'   ik_debug_ends
    Do While gcolSchemaDocs.Count > 0
        Set xmlDoc = gcolSchemaDocs.Item(1)
        Set xmlDoc = Nothing
        gcolSchemaDocs.Remove 1
    Loop
End Sub

' ik_wip_20040331_ends

' IK_13/04/2002
' support functions for (original) omRB schema
Public Function adoLegacyGetSchemaNode(ByVal vstrSearchPattern) As IXMLDOMNode
    Set adoLegacyGetSchemaNode = gxmldocSchemas.selectSingleNode(vstrSearchPattern)
End Function
Public Sub adoLegacyGetDatasrce( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean)
    If xmlGetAttributeText(vxmlSchemaNode, "DATASRCE") = "_NONE_" Then
        adoLegacyAddLogicalGroupNode _
            vxmlRequestNode, vxmlResponseNode, vxmlSchemaNode, vblnDoComboLookUp
    Else
        adoLegacyGetDbData _
            vxmlRequestNode, vxmlResponseNode, vxmlSchemaNode, vblnDoComboLookUp
    End If
End Sub
Private Sub adoLegacyAddLogicalGroupNode( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean)
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim strNodeName As String
    strNodeName = "UNDEFINED"
        
    If xmlAttributeValueExists(vxmlSchemaNode, "OUTNAME") Then
        strNodeName = xmlGetAttributeText(vxmlSchemaNode, "OUTNAME")
    Else
        strNodeName = vxmlSchemaNode.nodeName
    End If
            
    Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
        
    Set xmlThisResponseNode = vxmlResponseNode.appendChild(xmlElem)
    adoLegacyGetChildNodes _
        vxmlRequestNode, _
        xmlThisResponseNode, _
        vxmlSchemaNode, _
        vblnDoComboLookUp
End Sub
Private Sub adoLegacyGetDbData( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean)
    Const cstrFunctionName As String = "adoLegacyGetDbData"
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaAttribNode As IXMLDOMNode
    Dim xmlSchemaDataItemList As IXMLDOMNodeList
    Dim xmlRequestAttribNode As IXMLDOMAttribute
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim xmlParamSrceNode As IXMLDOMNode
    Dim xmlDataSrceAttrib As IXMLDOMAttribute
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim param As ADODB.Parameter
    Dim strSQL As String, _
        strSQLWhere As String, _
        strSQLParamClause As String, _
        strPattern As String, _
        strDataSource As String, _
        strDbColumnName As String, _
        strKeySrceNodeName As String, _
        strParamSrceAttribName As String
    Dim blnGetViaSP As Boolean
    blnGetViaSP = xmlGetAttributeText(vxmlSchemaNode, "ENTITYTYPE") = "PROCEDURE"
    Set cmd = New ADODB.Command
    If xmlAttributeValueExists(vxmlRequestNode, gstrDataSourceId) Then
        ' allows REQUEST (template) to override DATASRCE
        ' e.g. ACTIVEQUOTATION view
        strDataSource = xmlGetAttributeText(vxmlRequestNode, gstrDataSourceId)
    Else
        strDataSource = xmlGetAttributeText(vxmlSchemaNode, gstrDataSourceId)
    End If
    ' get primary key values
    Set xmlThisResponseNode = vxmlResponseNode
    ' if response source has no attributes then is logical grouping
    ' i.e. DATASRCE = _NONE_
    Do While (xmlThisResponseNode.Attributes.length = 0) _
        And (xmlThisResponseNode.nodeName <> "RESPONSE")
        Set xmlThisResponseNode = xmlThisResponseNode.parentNode
    Loop
    If xmlThisResponseNode.nodeName = "RESPONSE" Then
        ' no data values from RESPONSE node
        Set xmlThisResponseNode = Nothing
    End If
    ' IK_TODO - confirm as valid
    ' Set xmlSchemaDataItemList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']")
    Set xmlSchemaDataItemList = vxmlSchemaNode.selectNodes("*[@KEYTYPE]")
    For Each xmlSchemaAttribNode In xmlSchemaDataItemList
        Set xmlDataSrceAttrib = Nothing
        strParamSrceAttribName = xmlSchemaAttribNode.nodeName
        If xmlAttributeValueExists(xmlSchemaAttribNode, "KEYSRCEITEM") Then
            strParamSrceAttribName = xmlGetAttributeText(xmlSchemaAttribNode, "KEYSRCEITEM")
        End If
        ' does REQUEST have dataitemname_REF entry
        If Not vxmlRequestNode Is Nothing Then
            If xmlAttributeValueExists(vxmlRequestNode, xmlSchemaAttribNode.nodeName & "_REF") Then
                strParamSrceAttribName = _
                    xmlGetAttributeText(vxmlRequestNode, xmlSchemaAttribNode.nodeName & "_REF")
            End If
        End If
        ' source attribute will mostly be on RESPONSE node
        If Not xmlThisResponseNode Is Nothing Then
            Set xmlDataSrceAttrib = _
                xmlThisResponseNode.Attributes.getNamedItem(strParamSrceAttribName)
        End If
            
        ' but could be on REQUEST
        If Not vxmlRequestNode Is Nothing Then
            If xmlAttributeValueExists(vxmlRequestNode, strParamSrceAttribName) Then
                Set xmlDataSrceAttrib = vxmlRequestNode.Attributes.getNamedItem(strParamSrceAttribName)
            End If
        End If
        If Not xmlDataSrceAttrib Is Nothing Then
            If blnGetViaSP Then
                
                If Len(strSQLParamClause) = 0 Then
                    strSQLParamClause = "?"
                Else
                    strSQLParamClause = strSQLParamClause & ",?"
                End If
            Else
                strDbColumnName = xmlSchemaAttribNode.nodeName
                If xmlAttributeValueExists(xmlSchemaAttribNode, gstrDbRefId) Then
                    strDbColumnName = _
                        xmlGetAttributeText(xmlSchemaAttribNode, gstrDbRefId)
                End If
                If Len(strSQLWhere) <> 0 Then
                    strSQLWhere = strSQLWhere & " AND "
                End If
                strSQLWhere = strSQLWhere & strDbColumnName & "=?"
            End If
            Set param = getParam(xmlSchemaAttribNode, xmlDataSrceAttrib)
            cmd.Parameters.Append param
        End If
    Next
    ' add any (non-key) values from REQUEST
    If Not blnGetViaSP Then
        If Not vxmlRequestNode Is Nothing Then
            For Each xmlRequestAttribNode In vxmlRequestNode.Attributes
                Set xmlSchemaAttribNode = _
                    vxmlSchemaNode.selectSingleNode( _
                        xmlRequestAttribNode.nodeName & "[@DATATYPE]")
                ' IK_TODO (?)
                ' add (non-key field) _REF lookups
                If Not xmlSchemaAttribNode Is Nothing Then
                    ' IK_TODO - confirm as valid
                    ' If xmlGetAttributeText(xmlSchemaAttribNode, "KEYTYPE") <> "PRIMARY" Then
                    If Not xmlAttributeValueExists(xmlSchemaAttribNode, "KEYTYPE") Then
                        strDbColumnName = xmlSchemaAttribNode.nodeName
                        If xmlAttributeValueExists(xmlSchemaAttribNode, gstrDbRefId) Then
                            strDbColumnName = _
                                xmlGetAttributeText(xmlSchemaAttribNode, gstrDbRefId)
                        End If
                        If Len(strSQLWhere) <> 0 Then
                            strSQLWhere = strSQLWhere & " AND "
                        End If
                        strSQLWhere = strSQLWhere & strDbColumnName & "=?"
                        Set param = getParam(xmlSchemaAttribNode, xmlRequestAttribNode)
                        cmd.Parameters.Append param
                    End If
                End If
            Next
        End If
    End If
    ' add any SQL additions values from REQUEST
    If Not blnGetViaSP Then
        If Not vxmlRequestNode Is Nothing Then
            If xmlAttributeValueExists(vxmlRequestNode, gstrWhereId) Then
                If strSQLWhere = Empty Then
                    strSQLWhere = _
                        " WHERE " & _
                        xmlGetAttributeText(vxmlRequestNode, gstrWhereId)
                Else
                    strSQLWhere = _
                        strSQLWhere & _
                        " AND " & _
                        xmlGetAttributeText(vxmlRequestNode, gstrWhereId)
                End If
            End If
        End If
    End If
    If blnGetViaSP Then
        If cmd.Parameters.Count = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']").length _
        Then
            strSQL = "{CALL " & strDataSource & "(" & strSQLParamClause & ")}"
        End If
    Else
        
        ' IK SYSMCP0795
        ' _TP / _DIR FRIG
        If Right(strDataSource, 3) = "_TP" Then
            strDataSource = Left(strDataSource, Len(strDataSource) - 3) & "_DIR"
        End If
        If Len(strSQLWhere) > 0 Then
' ik_20030221
'            strSQL = _
'                "SELECT * FROM " & strDataSource & _
'                " WHERE (" & strSQLWhere & ")"
                
' SDS LIVE00009653 / 20/01/2004 _ STARTS
'            strSQL = _
'                "SELECT * FROM " & strDataSource & " WITH (NOLOCK)" & _
'                " WHERE (" & strSQLWhere & ")"
                
            strSQL = "SELECT * FROM " & strDataSource
            If genumDbProvider = omiga4DBPROVIDERSQLServer Then
                strSQL = strSQL & " WITH (NOLOCK)"
            End If
            strSQL = strSQL & " WHERE (" & strSQLWhere & ")"
                
' ik_20030221_ends
' SDS LIVE00009653 / 20/01/2004_ ENDS
        Else
            ' IK 19/10/2001
            ' select * only available at top of hierarchy
            If vxmlResponseNode.nodeName = "RESPONSE" Then
                strSQL = "SELECT * FROM " & strDataSource
            End If
        End If
        ' add any SQL ORDER BY clause from REQUEST
        If Len(strSQL) > 0 Then
            If Not vxmlRequestNode Is Nothing Then
                If xmlAttributeValueExists(vxmlRequestNode, gstrOrderById) Then
                    strSQL = _
                        strSQL & _
                        " ORDER BY " & _
                        xmlGetAttributeText(vxmlRequestNode, gstrOrderById)
                End If
            End If
        End If
    End If
    Debug.Print cstrFunctionName & " strSQL: " & strSQL
    If Len(strSQL) > 0 Then
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
        Set rst = executeGetRecordSet(cmd)
        ' _TP / _DIR FRIG
        If Not blnGetViaSP Then
            If rst Is Nothing Then
                    
                If Right(strDataSource, 4) = "_DIR" Then
                    
                    strDataSource = Left(strDataSource, Len(strDataSource) - 4) & "_TP"
                    strSQL = _
                        "SELECT * FROM " & _
                        strDataSource & _
                        Right(strSQL, Len(strSQL) - (InStr(1, strSQL, "_DIR") + 3))
                    Debug.Print cstrFunctionName & " strSQL: " & strSQL
                    cmd.CommandText = strSQL
                    Set rst = executeGetRecordSet(cmd)
                End If
            End If
        End If
        If Not rst Is Nothing Then
' ik_20030221_debug
'        App.LogEvent strSQL & " records: " & rst.RecordCount
' ik_20030221_debug
            Debug.Print cstrFunctionName & " rst.RecordCount: " & rst.RecordCount
            adoLegacyConvertRecordSetToXML _
                vxmlRequestNode, vxmlResponseNode, vxmlSchemaNode, rst, vblnDoComboLookUp
' ik_20030221_debug
'        App.LogEvent strSQL & " records: 0"
' ik_20030221_debug
        End If
    End If
End Sub
Private Sub adoLegacyConvertRecordSetToXML( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vrst As ADODB.Recordset, _
    ByVal vblnDoComboLookUp As Boolean)
    Dim xmlSchemaAttributeList As IXMLDOMNodeList
    Dim xmlSchemaAttribNode As IXMLDOMNode
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim fld As ADODB.Field
    Dim strNodeName As String, _
        strAttribName As String
    strNodeName = vxmlSchemaNode.nodeName
        
    If xmlAttributeValueExists(vxmlSchemaNode, "OUTNAME") Then
        strNodeName = xmlGetAttributeText(vxmlSchemaNode, "OUTNAME")
    End If
        
    If xmlAttributeValueExists(vxmlRequestNode, "OUTNAME") Then
        strNodeName = xmlGetAttributeText(vxmlRequestNode, "OUTNAME")
    End If
'    Debug.Print "adoLegacyConvertRecordSetToXML SCHEMA NODE: " & vxmlSchemaNode.nodeName
    
    Do While Not vrst.EOF
            
        Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
        Set xmlSchemaAttributeList = vxmlSchemaNode.selectNodes("*[@DATATYPE][not(@DATATYPE = 'DERIVED')]")
        For Each xmlSchemaAttribNode In xmlSchemaAttributeList
'            Debug.Print "adoLegacyConvertRecordSetToXML ATTRIBUTE: " & xmlSchemaAttribNode.nodeName
            ' IK 05/11/2001 DB_REF support
            If xmlAttributeValueExists(xmlSchemaAttribNode, gstrDbRefId) Then
                strAttribName = _
                    xmlGetAttributeText(xmlSchemaAttribNode, gstrDbRefId)
            Else
                strAttribName = xmlSchemaAttribNode.nodeName
            End If
            Set fld = Nothing
            On Error Resume Next
            If xmlGetAttributeText(xmlSchemaAttribNode, "PARAM_TYPE") <> "IN" Then
                Set fld = vrst.Fields.Item(strAttribName)
            End If
            On Error GoTo 0
            If Not fld Is Nothing Then
                If Not IsNull(fld) Then
                    If Not IsNull(fld.Value) Then
                        FieldToXml fld, xmlElem, xmlSchemaAttribNode, vblnDoComboLookUp
                    End If
                End If
            End If
        Next
        Set xmlSchemaAttributeList = vxmlSchemaNode.selectNodes("*[@DATATYPE='DERIVED']")
        For Each xmlSchemaAttribNode In xmlSchemaAttributeList
'            Debug.Print "adoLegacyConvertRecordSetToXML ATTRIBUTE: " & xmlSchemaAttribNode.nodeName
            
            adoLegacyDerivedFieldToXml xmlElem, xmlSchemaAttribNode
        Next
        If xmlElem.Attributes.length > 0 Then
            Set xmlThisResponseNode = vxmlResponseNode.appendChild(xmlElem)
            adoLegacyGetChildNodes _
                vxmlRequestNode, _
                xmlThisResponseNode, _
                vxmlSchemaNode, _
                vblnDoComboLookUp
        End If
            
        vrst.MoveNext
    Loop
End Sub
Private Sub adoLegacyDerivedFieldToXml( _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    Dim xmlElem As IXMLDOMElement
    Dim objScriptControl As Object
    Dim objScriptModule As Object
    Dim strCode As String, _
        strDerivedValue As String
    Dim intIndex As Integer
    Set objScriptControl = CreateObject("MSScriptControl.ScriptControl")
    objScriptControl.Language = "VBScript"
    objScriptControl.Timeout = -1
    Set objScriptModule = objScriptControl.Modules.Add("omiga4")
    strCode = vxmlSchemaNode.Text
    objScriptModule.AddCode strCode
    For intIndex = 1 To objScriptModule.Procedures.Count
        If objScriptModule.Procedures.Item(intIndex).Name = "DerivedValue" Then
            strDerivedValue = _
                objScriptModule.Run("DerivedValue", vxmlOutElem)
            If Not strDerivedValue = Empty Then
                
                vxmlOutElem.setAttribute _
                    vxmlSchemaNode.nodeName, _
                    strDerivedValue
            End If
        End If
    Next
    Set objScriptModule = Nothing
    Set objScriptControl = Nothing
    Set xmlElem = Nothing
            
End Sub
Private Sub adoLegacyGetChildNodes( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean)
    Dim xmlRequestNodeList As IXMLDOMNodeList
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    Dim xmlThisSchemaNode As IXMLDOMNode
    Dim strSchemaChildNodeName As String
    If Not vxmlRequestNode Is Nothing Then
        If xmlGetAttributeText(vxmlRequestNode, gstrExtractTypeId) = gstrExtractTypeNode Then
            ' REQUEST node terminates schema hierarchy
            Exit Sub
        End If
        If vxmlRequestNode.nodeName = "DATASRCE" Then
            Set xmlRequestNodeList = vxmlRequestNode.selectNodes("DATASRCE")
        Else
            Set xmlRequestNodeList = vxmlRequestNode.childNodes
        End If
    End If
    If Not xmlRequestNodeList Is Nothing Then
        If xmlRequestNodeList.length = 0 Then
            Set xmlRequestNodeList = Nothing
        End If
    End If
    If Not xmlRequestNodeList Is Nothing Then
        For Each xmlThisRequestNode In xmlRequestNodeList
            
            If xmlThisRequestNode.nodeName = "DATASRCE" Then
                strSchemaChildNodeName = xmlGetAttributeText(xmlThisRequestNode, "NAME")
            Else
                strSchemaChildNodeName = xmlThisRequestNode.nodeName
            End If
            Set xmlThisSchemaNode = vxmlSchemaNode.selectSingleNode(strSchemaChildNodeName)
            If xmlThisSchemaNode Is Nothing Then
                
                errThrowError _
                    "adoLegacyGetChildNodes", _
                    oeXMLMissingElement, _
                    "missing SCHEMA node,  NODE name: " & strSchemaChildNodeName
            End If
            adoLegacyGetDatasrce _
                xmlThisRequestNode, vxmlResponseNode, xmlThisSchemaNode, vblnDoComboLookUp
        Next
    Else
        Set xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@DATASRCE]")
        For Each xmlThisSchemaNode In xmlSchemaNodeList
            
            adoLegacyGetDatasrce _
                xmlThisRequestNode, vxmlResponseNode, xmlThisSchemaNode, vblnDoComboLookUp
        Next
    End If
End Sub
Public Function executeGetRecordSet( _
    ByVal cmd As ADODB.Command, _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As ADODB.Recordset
On Error GoTo executeGetRecordSetExit
    
    Const strFunctionName As String = "executeGetRecordSet"
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    conn.ConnectionString = adoGetDbConnectString(vstrComponentRef)
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
    errCheckError strFunctionName
End Function
' APS 29/08/01 Removed Optional Parameter
Private Function getParam( _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlAttrib As IXMLDOMAttribute) _
    As ADODB.Parameter
    Dim param As ADODB.Parameter
    Dim strDataType As String, _
        strDataValue As String
    Dim LngStringLen As Long
    Set param = New ADODB.Parameter
        
    strDataType = vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
    If Not vxmlAttrib Is Nothing Then
        strDataValue = vxmlAttrib.Text
    End If
    Select Case strDataType
        Case "STRING"
            
            param.Type = adBSTR
            LngStringLen = xmlGetAttributeAsLong(vxmlSchemaNode, "LENGTH")
            If LngStringLen > 0 And Len(strDataValue) > LngStringLen Then
                strDataValue = Left(strDataValue, LngStringLen)
            End If
                        
            If Not vxmlAttrib Is Nothing Then
                param.Size = Len(strDataValue)
            End If
        Case "GUID"
            
            param.Type = adBinary
            param.Size = 16
            If Len(strDataValue) > 0 Then
                strDataValue = GuidStringToByteArray(strDataValue)
            End If
        Case "SHORT", "COMBO", "LONG"
            
            param.Type = adInteger
        Case "BOOLEAN"
            
            strDataValue = UCase(strDataValue)
            If strDataValue = gstrBooleanTrue Or _
                Left(strDataValue, 1) = "Y" Or _
                Left(strDataValue, 1) = "T" Or _
                strDataValue = "1" _
            Then
                strDataValue = "1"
            Else
                strDataValue = "0"
            End If
            param.Type = adInteger
        Case "DOUBLE", "CURRENCY"
            param.Type = adDouble
        Case "DATE", "DATETIME"
            param.Type = adDBTimeStamp
    End Select
    param.Direction = adParamInput
    param.Value = Null
    If Not vxmlAttrib Is Nothing Then
        If Len(strDataValue) > 0 Then
            param.Value = strDataValue
        End If
    End If
    Set getParam = param
    Set param = Nothing
End Function
Private Function adoGuidToString(bytArray() As Byte) As String
    Dim i As Integer
    Dim strGuid As String
    strGuid = ""
    For i = 0 To 15
        strGuid = strGuid & Right$(Hex$(bytArray(i) + &H100), 2)
    Next i
    adoGuidToString = strGuid
End Function
'
Private Function adoDateToString(ByVal vvarDate As Variant) As String
    If geDateStyle = e_dateISO8601 Then
        adoDateToString = _
             Right(Str(DatePart("yyyy", vvarDate)), 4) & "-" & Right(Str(DatePart("m", vvarDate) + 100), 2) & "-" & Right(Str(DatePart("d", vvarDate) + 100), 2)
    Else
        adoDateToString = _
            Right(Str(DatePart("d", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("m", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("yyyy", vvarDate)), 4)
    End If
End Function
Private Function adoDateTimeToString(ByVal vvarDate As Variant) As String
    adoDateTimeToString = adoDateToString(vvarDate) & " " & _
               Right(Str(DatePart("h", vvarDate) + 100), 2) & ":" & _
               Right(Str(DatePart("n", vvarDate) + 100), 2) & ":" & _
               Right(Str(DatePart("s", vvarDate) + 100), 2)
End Function
Private Sub FieldToXml( _
    ByVal vfld As ADODB.Field, _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean)
    Dim xmlAttrib As IXMLDOMAttribute
    'BBG512  Add ComboValidationID
    Dim strElementValue As String, _
               strComboGroup As String, _
               strComboValue As String, _
               strComboValidationID As String, _
               strFormatMask As String, _
               strFormatted As String, _
               strAttribName As String
    If vxmlSchemaNode.nodeName = "ATTRIBUTE" Then
        strAttribName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
    Else
        strAttribName = vxmlSchemaNode.nodeName
    End If
    Select Case vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
        Case "GUID"
            vxmlOutElem.setAttribute strAttribName, adoGuidToString(vfld.Value)
        Case "CURRENCY"
            
            strFormatted = Format(vfld, "0.0000000")
            strFormatted = Left(strFormatted, InStr(1, strFormatted, ".") + 2)
            vxmlOutElem.setAttribute strAttribName, strFormatted
        Case "DOUBLE"
            
            strFormatMask = Empty
            If Not vxmlSchemaNode.Attributes.getNamedItem("FORMATMASK") Is Nothing Then
                strFormatMask = vxmlSchemaNode.Attributes.getNamedItem("FORMATMASK").Text
            End If
            If Len(strFormatMask) <> 0 Then
                
                vxmlOutElem.setAttribute _
                           strAttribName & "_RAW", _
                           vfld.Value
                vxmlOutElem.setAttribute _
                           strAttribName, _
                           Format(vfld.Value, strFormatMask)
            Else
                
                vxmlOutElem.setAttribute strAttribName, vfld.Value
            End If
        'BBG512
        Case "COMBO"
            vxmlOutElem.setAttribute strAttribName, vfld.Value
            If vblnDoComboLookUp = True Then
                strComboGroup = xmlGetAttributeText(vxmlSchemaNode, "COMBOGROUP")
                If Len(strComboGroup) > 0 Then
                    If xmlGetAttributeAsBoolean(vxmlSchemaNode, "GETCOMBOVALIDID") Then
                        Call GetComboData(strComboGroup, vfld.Value, strComboValue, _
                                strComboValidationID)
                        If Len(strComboValue) > 0 Then
                            vxmlOutElem.setAttribute _
                                strAttribName & "_TEXT", strComboValue
                        End If
                        If Len(strComboValidationID) > 0 Then
                            vxmlOutElem.setAttribute _
                                strAttribName & "_VALIDID", strComboValidationID
                        End If
                    Else
                        strComboValue = GetComboText(strComboGroup, vfld.Value)
                        If Len(strComboValue) > 0 Then
                            vxmlOutElem.setAttribute _
                                strAttribName & "_TEXT", strComboValue
                        End If
                    End If
                End If
            End If
        Case "DATE"
            vxmlOutElem.setAttribute _
                       strAttribName, _
                       adoDateToString(vfld.Value)
        Case "DATETIME"
            vxmlOutElem.setAttribute _
                       strAttribName, _
                       adoDateTimeToString(vfld.Value)
        Case "BOOLEAN"
            If vfld.Value = 1 Then
                vxmlOutElem.setAttribute strAttribName, gstrBooleanTrue
            Else
                vxmlOutElem.setAttribute strAttribName, gstrBooleanFalse
            End If
        Case Else
            vxmlOutElem.setAttribute strAttribName, vfld.Value
    End Select
End Sub
Public Sub adoBuildDbConnectionString( _
    Optional ByVal vstrComponentRef As String = "DEFAULT")
On Error GoTo BuildDbConnectionStringVbErr
    Dim objWshShell As Object
    Dim strThisComponentRef As String, _
        strConnection As String, _
        strProvider As String, _
        strRegSection As String, _
        strRetries As String, _
        strUserId As String, _
        strPassword As String, _
        strDescription As String, _
        strSource As String
    Dim lngErrNo As Long
        
    If vstrComponentRef = "DEFAULT" Then
        strThisComponentRef = App.Title
    Else
        strThisComponentRef = vstrComponentRef
    End If
    Set objWshShell = CreateObject("WScript.Shell")
    strRegSection = "HKLM\SOFTWARE\" & gstrAppName & "\" & strThisComponentRef & "\" & gstrREGISTRY_SECTION & "\"
On Error Resume Next
    
    strProvider = objWshShell.RegRead(strRegSection & gstrPROVIDER_KEY)
On Error GoTo BuildDbConnectionStringVbErr
    
    If Len(strProvider) = 0 Then
        strRegSection = "HKLM\SOFTWARE\" & gstrAppName & "\" & gstrREGISTRY_SECTION & "\"
        strProvider = objWshShell.RegRead(strRegSection & gstrPROVIDER_KEY)
    End If
    genumDbProvider = omiga4DBPROVIDERUnknown
    'If strProvider = "MSDAORA" Then
    'SDS LIVE00009659  22/01/2004 STARTS
    'Support for both MS and Oracle OLEDB Drivers
    If (strProvider = "MSDAORA" Or strProvider = "ORAOLEDB.ORACLE") Then
        Select Case strProvider
            Case "MSDAORA"
                genumDbProvider = omiga4DBPROVIDERMSOracle
            Case "ORAOLEDB.ORACLE"
                genumDbProvider = omiga4DBPROVIDEROracle
        End Select
        'genumDbProvider = omiga4DBPROVIDEROracle
        strConnection = _
            "Provider=" & strProvider & ";Data Source=" & objWshShell.RegRead(strRegSection & gstrDATA_SOURCE_KEY) & ";" & _
            "User ID=" & objWshShell.RegRead(strRegSection & gstrUID_KEY) & ";" & _
            "Password=" & objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY) & ";"
    'SDS LIVE00009659  22/01/2004 ENDS
    ElseIf strProvider = "SQLOLEDB" Then
        genumDbProvider = omiga4DBPROVIDERSQLServer
        ' PSC 06/01/03 BM232 - Start
'        strConnection = _
'                   "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & gstrSERVER_KEY) & ";" & _
'                   "database=" & objWshShell.RegRead(strRegSection & gstrDATABASE_KEY) & ";" & _
'                   "UID=" & objWshShell.RegRead(strRegSection & gstrUID_KEY) & ";" & _
'                   "pwd=" & objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY) & ";"
        
        strUserId = ""
        strPassword = ""
        On Error Resume Next
                
        strUserId = objWshShell.RegRead(strRegSection & gstrUID_KEY)
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
                
        On Error GoTo BuildDbConnectionStringVbErr
        If lngErrNo <> glngENTRYNOTFOUND And lngErrNo <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
        On Error Resume Next
                
        strPassword = objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY)
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
                
        On Error GoTo BuildDbConnectionStringVbErr
        If lngErrNo <> glngENTRYNOTFOUND And lngErrNo <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
        strConnection = _
            "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & gstrSERVER_KEY) & ";" & _
            "database=" & objWshShell.RegRead(strRegSection & gstrDATABASE_KEY) & ";"
        ' If User Id is present use SQL Server Authentication else
        ' use integrated security
        If Len(strUserId) > 0 Then
            strConnection = strConnection & "UID=" & strUserId & ";" & _
                "pwd=" & strPassword & ";"
        Else
            strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
        End If
        ' PSC 06/01/03 BM232 - End
    End If
           
    gstrDbConnectionString = strConnection
    strRetries = objWshShell.RegRead(strRegSection & gstrRETRIES_KEY)
    If Len(strRetries) > 0 Then
        gintDbRetries = CInt(strRetries)
    End If
    Set objWshShell = Nothing
    putConnectString vstrComponentRef, strConnection
'     Debug.Print strConnection
    
    Exit Sub
BuildDbConnectionStringVbErr:
    
    Set objWshShell = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub putConnectString( _
    ByVal vstrRef As String, _
    ByVal vstrConn As String)
    Dim xmlElem As IXMLDOMElement
    If gxmldocConnections Is Nothing Then
        Set gxmldocConnections = New FreeThreadedDOMDocument40
        gxmldocConnections.async = False
        Set xmlElem = gxmldocConnections.createElement("CONN_ROOT")
        gxmldocConnections.appendChild xmlElem
    End If
    Set xmlElem = gxmldocConnections.createElement("CONN")
    xmlElem.setAttribute vstrRef, vstrConn
    gxmldocConnections.firstChild.appendChild xmlElem
    Debug.Print gxmldocConnections.xml
End Sub
Public Function adoGetDbConnectString( _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As String
    If gxmldocConnections.selectSingleNode("CONN_ROOT/CONN/@" & vstrComponentRef) Is Nothing _
    Then
        adoBuildDbConnectionString vstrComponentRef
    End If
    adoGetDbConnectString = _
        gxmldocConnections.selectSingleNode("CONN_ROOT/CONN/@" & vstrComponentRef).Text
'    Debug.Print adoGetDbConnectString
End Function
Public Function adoGetDbRetries() As Integer
    adoGetDbRetries = gintDbRetries
End Function
Public Function adoGetDbProvider() As DBPROVIDER
    adoGetDbProvider = genumDbProvider
End Function
Public Sub adoLoadSchema()
    Dim xmlNode As IXMLDOMNode
    Dim strFileSpec As String
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\" & App.Title & ".xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    Set gxmldocSchemas = New FreeThreadedDOMDocument40
    gxmldocSchemas.async = False
    gxmldocSchemas.setProperty "NewParser", True
    gxmldocSchemas.validateOnParse = False
    gxmldocSchemas.Load strFileSpec
    geDateStyle = e_dateUK
    gstrBooleanTrue = "1"
    gstrBooleanFalse = "0"
    
    If gxmldocSchemas.parseError.errorCode = 0 Then
        
        ' IK_WIP (incorporate omRb style schema + processing)
        genumSchemaStyle = omiga4SCHEMASTYLEPlus
        ' IK_WIP ends
        Set xmlNode = gxmldocSchemas.selectSingleNode("OM_SCHEMA")
        If Not xmlNode Is Nothing Then
            If xmlGetAttributeText(xmlNode, "DATEFORMAT") = gstrISO8601 Then
                geDateStyle = e_dateISO8601
            End If
            If xmlAttributeValueExists(xmlNode, "BOOLEANTRUE") Then
                gstrBooleanTrue = UCase(xmlGetAttributeText(xmlNode, "BOOLEANTRUE"))
            End If
            If xmlAttributeValueExists(xmlNode, "BOOLEANFALSE") Then
                gstrBooleanFalse = UCase(xmlGetAttributeText(xmlNode, "BOOLEANFALSE"))
            End If
            Set xmlNode = Nothing
        End If
    ' IK_WIP (incorporate omRb style schema + processing)
    Else
        
        ' do we have an 'original' omRB style schema
        strFileSpec = App.Path & "\" & "omSchema.xml"
        strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
        gxmldocSchemas.setProperty "NewParser", True
        gxmldocSchemas.validateOnParse = False
        gxmldocSchemas.Load strFileSpec
        If gxmldocSchemas.parseError.errorCode = 0 Then
            genumSchemaStyle = omiga4SCHEMASTYLEOriginal
        End If
    ' IK_WIP ends
    End If

' ik_wip_20030430
    If gxmldocSchemas.parseError.errorCode = 0 Then
        Set gcolSchemaDocs = New Collection
        gcolSchemaDocs.Add gxmldocSchemas, "default"
        gstrActiveSchema = "default"
'   ik_debug
'       App.LogEvent "adoLoadSchema, gcolSchemaDocs.Count: " & gcolSchemaDocs.Count, vbLogEventTypeInformation
'   ik_debug_ends
    End If
' ik_wip_20030430_ends
    
End Sub

' ik_wip_20030430
Public Sub adoLoadNamedSchema(ByVal vstrSchemaName As String)
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim strFileSpec As String
    
    If gstrActiveSchema = vstrSchemaName Then
        Exit Sub
    End If
    
    geDateStyle = e_dateUK
    gstrBooleanTrue = "1"
    gstrBooleanFalse = "0"
    
    Set gxmldocSchemas = Nothing
    On Error Resume Next
    Set gxmldocSchemas = gcolSchemaDocs.Item(vstrSchemaName)
    On Error GoTo 0
    
    If Not gxmldocSchemas Is Nothing Then
        gstrActiveSchema = vstrSchemaName
    Else
        ' pick up XML map from "...\Omiga 4\XML" directory
        ' Only do the subsitution once to change DLL -> XML
        strFileSpec = App.Path & "\" & vstrSchemaName & ".xml"
        strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
        xmlDoc.validateOnParse = False
        xmlDoc.Load strFileSpec
        
        If xmlDoc.parseError.errorCode = 0 Then
            gcolSchemaDocs.Add xmlDoc, vstrSchemaName
            Set gxmldocSchemas = xmlDoc
            gstrActiveSchema = vstrSchemaName
'   ik_debug
'           App.LogEvent "adoLoadNamedSchema, gcolSchemaDocs.Count: " & gcolSchemaDocs.Count, vbLogEventTypeInformation
'   ik_debug_ends
        End If
    
    End If

End Sub
' ik_wip_20030430_ends

Public Function adoGetSchema(ByVal vstrSchemaName As String) As IXMLDOMNode
    
    Dim strPattern As String
    If gxmldocSchemas Is Nothing Then
        errThrowError _
            "adoGetSchema", _
            oeSchemaNotLoaded
    End If
    If gxmldocSchemas.parseError.errorCode <> 0 Then
        errThrowError _
            "adoGetSchema", _
            oeSchemaParseError, _
            gxmldocSchemas.parseError.reason
    End If
    strPattern = "//" & vstrSchemaName & "[@DATASRCE]"
    Set adoGetSchema = gxmldocSchemas.selectSingleNode(strPattern)
End Function
Private Function GetErrorNumber(ByVal strErrDesc) As Long
On Error GoTo GetErrorNumberVbErr
    Const cstrFunctionName = "GetErrorNumber"
    Dim strErr As String
    Select Case genumDbProvider
        Case omiga4DBPROVIDEROracle, omiga4DBPROVIDERMSOracle 'SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
            ' oracle errors have format "ORA-nnnnn"
            If Len(strErrDesc) > 10 Then
                strErr = Mid(strErrDesc, 5, 5)
            Else
                ' "Cannot interpret database error"
                errThrowError cstrFunctionName, 108
            End If
        Case omiga4DBPROVIDERSQLServer  'SQL Server
            errThrowError cstrFunctionName, oeNotImplemented
                
        Case omiga4DBPROVIDERUnknown
            errThrowError cstrFunctionName, oeNotImplemented
    End Select
    GetErrorNumber = strErr
    Exit Function
GetErrorNumberVbErr:
    Err.Raise Err.Number, cstrFunctionName, Err.Description
End Function
Public Function adoGetOmigaNumberForDatabaseError(ByVal strErrDesc As String) As Long
'-----------------------------------------------------------------------------------
'Description : Find the omiga equivalent number for a database error. This is used
'              to trap specific errors. Add to the list below, if you want trap a
'              a new error
'Pass        : strErrDesc : Description of the Error Message (from database )
'------------------------------------------------------------------------------------
On Error GoTo GetOmigaNumberForDatabaseErrorVbErr
    Const cstrFunctioName = "adoGetOmigaNumberForDatabaseError"
    Dim lngErrNo As Long, lngOmigaErrorNo As Long
    lngErrNo = GetErrorNumber(strErrDesc)
    'SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
    If genumDbProvider = omiga4DBPROVIDEROracle Or genumDbProvider = omiga4DBPROVIDERMSOracle Then
        Select Case lngErrNo
            Case 1
                lngOmigaErrorNo = oeDuplicateKey
            Case 2292
                lngOmigaErrorNo = oeChildRecordsFound
            Case Else
                lngOmigaErrorNo = oeUnspecifiedError
        End Select
    Else
        errThrowError cstrFunctioName, oeNotImplemented, "Error trapping for this database engine"
    End If
    adoGetOmigaNumberForDatabaseError = lngOmigaErrorNo
    Exit Function
GetOmigaNumberForDatabaseErrorVbErr:
    
    Err.Raise Err.Number, cstrFunctioName, Err.Description
End Function
' AS 31/05/01 CC012 Start
' -----------------------------------------------------------------------------------------
' Description:  Gets a valid recordset returned by executing a stored procedure.
'               Useful when a stored procedure may return multiple recordsets,
'               only one of which is the one you want. This is particularly
'               true of SQL Server, where inserts, creates etc generated
'               recordsets.
' Pass:
' vrstRecordSet The initial recordset returned by the stored procedure. Note: This is
'               passed by reference, which is OK as we're not marshalling across
'               process boundaries.
' nRecordSet    The n'th valid recordset to return. Valid means the recordset is open
'               and contains records. Default to 1 to return first recordset.
' Return:       True if successful.
' AS            31/05/01    First version
'------------------------------------------------------------------------------------------
Public Function adoGetValidRecordset(ByRef vrstRecordSet As ADODB.Recordset, Optional ByVal nRecordSet As Integer = 1) As Boolean
On Error GoTo adoGetValidRecordsetErr
    Const cstrFunctionName = "adoGetValidRecordset"
    Dim bSuccess As Boolean
    bSuccess = False
    If genumDbProvider = omiga4DBPROVIDERSQLServer Then
        Do Until vrstRecordSet Is Nothing
            If Not vrstRecordSet.State = adStateClosed Then
                If Not (vrstRecordSet.BOF And vrstRecordSet.EOF) Then
                    If nRecordSet > 0 Then
                        nRecordSet = nRecordSet - 1
                    End If
                    bSuccess = True
                    If nRecordSet = 0 Then
                        Exit Do
                    End If
                End If
            End If
            Set vrstRecordSet = vrstRecordSet.NextRecordset
        Loop
    ElseIf Not vrstRecordSet Is Nothing And Not vrstRecordSet.State = adStateClosed And Not (vrstRecordSet.BOF And vrstRecordSet.EOF) Then
        ' Found an open recordset with records.
        bSuccess = True
    End If
    adoGetValidRecordset = bSuccess
    Exit Function
adoGetValidRecordsetErr:
    Err.Raise Err.Number, cstrFunctionName, Err.Description
End Function
' AS 31/05/01 CC012 End
' private methods ======================================================================
Private Function IsComboLookUpRequired(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
    If Not vxmlRequestNode.Attributes.getNamedItem("_COMBOLOOKUP_") Is Nothing Then
        IsComboLookUpRequired = _
                   xmlGetAttributeAsBoolean(vxmlRequestNode, "_COMBOLOOKUP_")
    Else
        If Not vxmlRequestNode.ownerDocument.firstChild Is Nothing Then
            IsComboLookUpRequired = _
                       xmlGetAttributeAsBoolean(vxmlRequestNode.ownerDocument.firstChild, "_COMBOLOOKUP_")
        Else
            IsComboLookUpRequired = False
        End If
    End If
End Function
Private Function FormatGuid(ByVal strSrcGuid As String, Optional ByVal eGuidStyle As GUIDSTYLE = guidLiteral) As String
' header ----------------------------------------------------------------------------------
' description:  Converts a guid string from one format to another.
' pass:
'   strSrcGuid  The guid string to be converted.
'   eGuidStyle  The style of the return string.
'               guidHyphen converts from "DA6DA163412311D4B5FA00105ABB1680" to
'               "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format.
'               guidBinary converts from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to
'               "DA6DA163412311D4B5FA00105ABB1680" format.
'               guidLiteral converts from either guidBinary or guidHyphen format to
'               a format that can be used as a literal input parameter for the
'               database. This is the default to maintain backwards compatibility with
'               Omiga 4 Phase 2.
' return:       The converted guid string.
' AS            05/03/01    First version
'------------------------------------------------------------------------------------------
    On Error GoTo FormatGuidVbErr
    Dim strFunctionName As String
    Dim strTgtGuid As String
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    strFunctionName = "FormatGuid"
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
            errThrowError strFunctionName, oeInvalidParameter, "strSrcGuid length must be 32"
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
            errThrowError strFunctionName, oeInvalidParameter, "strSrcGuid length must be 38"
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
            Select Case adoGetDbProvider()
            Case omiga4DBPROVIDERSQLServer
                ' e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
                strTgtGuid = "0x" & strSrcGuid
            Case omiga4DBPROVIDEROracle, omiga4DBPROVIDERMSOracle 'SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
                ' e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
                strTgtGuid = "HEXTORAW('" & strSrcGuid & "')"
            End Select
        End If
    Else
        Debug.Assert 0
        errThrowError strFunctionName, oeInvalidParameter, "strGuid length must be 32"
    End If
    FormatGuid = strTgtGuid
    Exit Function
FormatGuidVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GuidStringToByteArray(ByVal strGuid As String) As Byte()
' header ----------------------------------------------------------------------------------
' description:      Converts a guid string into a byte array.
' pass:
'   strGuid         The guid string to be converted.
'                   Can be in either binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680",
'                   or hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
' return:           The byte array.
' AS                05/03/01    First version
'------------------------------------------------------------------------------------------
    On Error GoTo GuidStringToByteArrayVbErr
    Dim strFunctionName As String
    strFunctionName = "GuidStringToByteArray"
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
        errThrowError strFunctionName, oeInvalidParameter, "strGuid length must be 32"
    End If
    GuidStringToByteArray = rbytGuid
    Exit Function
GuidStringToByteArrayVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function ConvertGuidIndex(ByVal intIndex1 As Integer)
' header ----------------------------------------------------------------------------------
' description:  Helper function for ByteArrayToGuidString and FormatGuid.
' pass:
' return:
' AS            31/01/01    First version
'------------------------------------------------------------------------------------------
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

'BBG512  *** Start
'********************************************************************************
'** Function:       GetComboText
'** Created by:     Unknown
'** Date:           Unknown
'** Description:    Gets the appropriate value text for the supplied combo group
'**                 and ID value.
'** Parameters:     vstrGroupName - the name of the combo group.
'**                 vintValueID - the ID value of the combo item to get the name
'**                 for.
'** Returns:        The name of the combo item.
'** Errors:         None Expected
'********************************************************************************
Private Function GetComboText(ByVal vstrGroupName As String, _
        ByVal vintValueID As Integer) As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Set cmd = BuildComboExtractCommand(vstrGroupName, vintValueID)
    cmd.CommandText = "SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME=? And VALUEID=?"
    Set rst = GetDisconnectedRecordset(cmd)
    If Not rst.EOF Then
        rst.MoveFirst
        GetComboText = rst.Fields(0).Value
    End If
    rst.Close

End Function

'********************************************************************************
'** Function:       GetComboData
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Gets the appropriate value text and validation type ID for
'**                 the specified combo group and item.
'** Parameters:     vstrGroupName - the name of the combo group.
'**                 vintValueID - the ID value of the combo item to get the name
'**                 for.
'**                 rstrText - the text of the combo item.
'**                 rstrValidationID - the comma separated list of validation
'**                 ID's
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub GetComboData(ByVal vstrGroupName As String, ByVal vintValueID As Integer, _
        ByRef rstrText As String, ByRef rstrValidationID As String)
    Dim cmd As ADODB.Command

    Dim rst As ADODB.Recordset
    
    '*-ensure return parameters are empty before we start
    rstrText = vbNullString
    rstrValidationID = vbNullString
    
    '*-firstly, get the combo value text from the combovalue table
    Set cmd = BuildComboExtractCommand(vstrGroupName, vintValueID)
    cmd.CommandText = "SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME=? And VALUEID=?"
    Set rst = GetDisconnectedRecordset(cmd)
    If Not rst.EOF Then
        rst.MoveFirst
        rstrText = rst.Fields(0).Value
    End If
    rst.Close
    
    '*-next, get the combo validation ID's from the combovalidation table
    cmd.CommandText = "SELECT VALIDATIONTYPE FROM COMBOVALIDATION WHERE GROUPNAME=? And VALUEID=?"
    Set rst = GetDisconnectedRecordset(cmd)
    If Not rst.EOF Then
        rst.MoveFirst
        Do While Not rst.EOF
            If Len(rstrValidationID) = 0 Then
                rstrValidationID = rst.Fields(0).Value
            Else
                rstrValidationID = rstrValidationID & "," & rst.Fields(0).Value
            End If
            rst.MoveNext
        Loop
    End If
    rst.Close
    
End Sub

'********************************************************************************
'** Function:       BuildComboExtractCommand
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Creates a command object and sets the appropriate properties
'**                 to be used to extract data from the combo related tables.
'** Parameters:     vstrGroupName - the name of the combo group.
'**                 vintValueID - the ID of the appropriate combo item.
'** Returns:        The constructed command object.
'** Errors:         None Expected
'********************************************************************************
Private Function BuildComboExtractCommand(ByVal vstrGroupName As String, _
        ByVal vintValueID As Integer) As ADODB.Command
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Set cmd = New ADODB.Command

    Set param = New ADODB.Parameter
    param.Type = adBSTR
    param.Size = Len(vstrGroupName)
    param.Direction = adParamInput
    param.Value = vstrGroupName
    cmd.Parameters.Append param

    Set param = New ADODB.Parameter
    param.Type = adInteger
    param.Direction = adParamInput
    param.Value = vintValueID
    cmd.Parameters.Append param
    cmd.CommandType = adCmdText
    
    Set BuildComboExtractCommand = cmd
    
End Function

'********************************************************************************
'** Function:       GetDisconnectedRecordset
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Generic method to get and return a disconnected recordset
'**                 from a command object.
'** Parameters:     vcmd - the command object to create the recordset with.
'** Returns:        The disconnected recordset.
'** Errors:         None Expected
'********************************************************************************
Private Function GetDisconnectedRecordset(ByVal vcmd As ADODB.Command) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = adoGetDbConnectString()
    conn.Open
    vcmd.ActiveConnection = conn

    Set rst = New ADODB.Recordset

    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.Open vcmd

    ' disconnect RecordSet and close the connection
    Set rst.ActiveConnection = Nothing
    conn.Close
    
    '*-return the recordset
    Set GetDisconnectedRecordset = rst
    
End Function
'BBG512 *** End


