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
'LD     19/09/2001  SYS2722 SQL Server Port - Make function adoGuidToString public
'IK     12/10/2001  SYS2803 Work Arounds for MSXML3 IXMLDOMNodeList bug
'PSC    17/10/01    SYS2815 Allow integrated security with SQL Server
'SG     18/06/02    SYS4889 Code error in executeGetRecordSet
'SDS    22/01/2004   LIVE00009659  Added pieces of code to support both Microsoft and Oracle OLEDB Drivers
'-------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        AQR     Description
'TK     22/11/2004  BBG1821 Performance related fixes
'-------------------------------------------------------------------------------

Option Explicit
Private gxmldocSchemas As FreeThreadedDOMDocument40
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
 
Public Const gstrSchemaNameId As String = "_SCHEMA_"
Public Const gstrExtractTypeId As String = "_EXTRACTTYPE_"
Public Const gstrOutNodeId As String = "_OUTNODENAME_"
Public Const gstrOrderById As String = "_ORDERBY_"
Public Const gstrWhereId As String = "_WHERE_"
Public Const gstrDataTypeId As String = "DATATYPE"
Public Const gstrDataTypeDerivedId As String = "DERIVED"
Public Const gstrDataSourceId As String = "DATASRCE"
Private Const gstrMODULEPREFIX As String = "adoAssist."
Private Const glngENTRYNOTFOUND As Long = -2147024894
'-------------------------------------------------------------------------------
'BMIDS History
'RF     18/11/02    BMIDS00935 Applied Core Change SYS4752 (SYSMCP0734)
'-------------------------------------------------------------------------------
Public Sub adoCreateFromNodeList( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vstrSchemaName As String)
On Error GoTo adoCreateFromNodeListExit
    
    Const strFunctionName As String = "adoCreateFromNodeList"
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
' IK AQR SYS2803
' fix for MSXML bug
    Dim intIndex As Integer
    For intIndex = 1 To vxmlRequestNodeList.length
        adoCreate vxmlRequestNodeList.Item(intIndex - 1), xmlSchemaNode
    Next
'    For Each xmlDataNode In vxmlRequestNodeList
'        adoCreate xmlDataNode, xmlSchemaNode
'    Next
    
adoCreateFromNodeListExit:
    
    Set xmlSchemaNode = Nothing
    Set xmlDataNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoCreateFromNode( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vstrSchemaName As String)
On Error GoTo adoCreateFromNodeExit
    
    Const strFunctionName As String = "adoCreateFromNode"
    Dim xmlSchemaNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    adoCreate vxmlRequestNode, xmlSchemaNode
adoCreateFromNodeExit:
    
    Set xmlSchemaNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoUpdateFromNodeList( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vstrSchemaName As String)
On Error GoTo adoUpdateFromNodeListExit
    
    Const strFunctionName As String = "adoUpdateFromNodeList"
    Dim xmlSchemaNode As IXMLDOMNode
    'Dim xmlDataNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
' IK AQR SYS2803
' fix for MSXML bug
    Dim intIndex As Integer
    For intIndex = 1 To vxmlRequestNodeList.length
        adoUpdate vxmlRequestNodeList.Item(intIndex - 1), xmlSchemaNode
    Next
'    For Each xmlDataNode In vxmlRequestNodeList
'        adoUpdate xmlDataNode, xmlSchemaNode
'    Next
    
adoUpdateFromNodeListExit:
    
    Set xmlSchemaNode = Nothing
    'Set xmlDataNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoUpdateFromNode( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vstrSchemaName As String)
On Error GoTo adoUpdateFromNodeExit
    
    Const strFunctionName As String = "adoUpdateFromNode"
    Dim xmlSchemaNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    adoUpdate vxmlRequestNode, xmlSchemaNode
adoUpdateFromNodeExit:
    
    Set xmlSchemaNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoUpdateMultipleInstancesFromNode( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vstrSchemaName As String)
On Error GoTo adoUpdateMultipleInstancesFromNodeExit
    
    Const strFunctionName As String = "adoUpdateMultipleInstancesFromNode"
    Dim xmlSchemaNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    adoUpdate vxmlRequestNode, xmlSchemaNode, True
adoUpdateMultipleInstancesFromNodeExit:
    
    Set xmlSchemaNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoGetAsXML( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vstrSchemaName As String, _
    Optional vstrFilter As String, _
    Optional vstrOrderBy As String)
    On Error GoTo adoGetAsXMLExit
    Const strFunctionName As String = "adoGetAsXML"
    Dim strOrderBy As String
    strOrderBy = vstrOrderBy
    If xmlAttributeValueExists(vxmlRequestNode, gstrOrderById) = True Then
        If Len(strOrderBy) > 0 Then
            strOrderBy = strOrderBy & ","
        End If
        strOrderBy = strOrderBy & xmlGetAttributeText(vxmlRequestNode, gstrOrderById)
    End If
    Dim xmlSchemaNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    adoCloneBaseEntityRefs xmlSchemaNode
    If xmlGetAttributeText(xmlSchemaNode, "ENTITYTYPE") = "PROCEDURE" Then
        adoGetStoredProcAsXML vxmlRequestNode, xmlSchemaNode, vxmlResponseNode
    Else
        adoGetRecordSetAsXML _
            vxmlRequestNode, xmlSchemaNode, vxmlResponseNode, vstrFilter, strOrderBy
    End If
        
adoGetAsXMLExit:
    
    Set xmlSchemaNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoCreate( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
On Error GoTo adoCreateExit
    
    Const strFunctionName As String = "adoCreate"
    Dim xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    Dim strSQL As String, _
        strDataType As String, _
        strDataSource As String, _
        strColumns As String, _
        strValues As String
    Dim intRetryCount As Integer, _
        intRetryMax As Integer
    Dim lngRecordsAffected As Long
    Dim blnDbCmdOk As Boolean
    strDataSource = xmlGetAttributeText(vxmlSchemaNode, "DATASRCE")
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaNode.nodeName
    End If
    ' get any generated keys & default values
    For Each xmlSchemaFieldNode In vxmlSchemaNode.childNodes
        If vxmlDataNode.Attributes.getNamedItem(xmlSchemaFieldNode.nodeName) Is Nothing Then
            If Not xmlSchemaFieldNode.Attributes.getNamedItem("DEFAULT") Is Nothing Then
                Set xmlAttrib = _
                    vxmlDataNode.ownerDocument.createAttribute(xmlSchemaFieldNode.nodeName)
                xmlAttrib.Text = _
                    xmlSchemaFieldNode.Attributes.getNamedItem("DEFAULT").Text
                vxmlDataNode.Attributes.setNamedItem xmlAttrib
            End If
            If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
                If (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "PRIMARY") Or _
                    (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "SECONDARY") _
                Then
                    If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYSOURCE") Is Nothing Then
                        If xmlSchemaFieldNode.Attributes.getNamedItem("KEYSOURCE").Text = "GENERATED" Then
                            GetGeneratedKey vxmlDataNode, xmlSchemaFieldNode
                        End If
                    End If
                End If
            End If
        End If
    Next
    intRetryMax = adoGetDbRetries
    Set cmd = New ADODB.Command
    For Each xmlDataAttrib In vxmlDataNode.Attributes
        
        Set xmlSchemaFieldNode = vxmlSchemaNode.selectSingleNode(xmlDataAttrib.Name)
        If Not xmlSchemaFieldNode Is Nothing Then
            If Len(strColumns) > 0 Then
                strColumns = strColumns & ","
            End If
            strColumns = strColumns & xmlDataAttrib.Name
            If Len(strValues) > 0 Then
                strValues = strValues & ","
            End If
            strValues = strValues & "?"
            Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib, True)
            cmd.Parameters.Append param
      End If
    Next
    strSQL = "INSERT INTO " & strDataSource & " (" & strColumns & ") VALUES (" & strValues & ")"
    Debug.Print "adoCreate strSQL: " & strSQL
    If Len(strValues) > 0 Then
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
        lngRecordsAffected = executeSQLCommand(cmd)
    End If
'    If Len(strValues) > 0 Then
'
'        cmd.CommandType = adCmdText
'        cmd.CommandText = strSQL
'
'        On Error Resume Next
'
'        Do While blnDbCmdOk = False
'
'            lngRecordsAffected = executeSQLCommand(cmd)
'
'            If Err.Number = 0 Then
'
'                blnDbCmdOk = True
'
'            Else
'
'                If IsRetryError() = False Then
'
'                    On Error GoTo adoCreateExit
'
'                Else
'
'                    intRetryCount = intRetryCount + 1
'
'                    Dim nTimer1 As Single, _
'                        nTimer2 As Single
'
'                    ' DEBUG
'                    App.LogEvent "adoCreate - retry [" & intRetryCount & "]: " & strSQL
'
'                    If (intRetryCount = intRetryMax) Then
'
'                        On Error GoTo adoCreateExit
'
'                    End If
'
'                End If
'
'            End If
'
'        Loop
'
'    End If
    
    Set cmd = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing
    Set param = Nothing

    Debug.Print "adoCreate records created: " & lngRecordsAffected
adoCreateExit:
    
    
    errCheckError strFunctionName
End Sub
Public Sub adoUpdate( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    Optional ByVal blnNonUniqueInstanceAllowed As Boolean = False)
On Error GoTo adoUpdateExit
    
    Const strFunctionName As String = "adoUpdate"
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    Dim xmlSchemaAttrib As IXMLDOMAttribute
    Dim xmlDataAttrib As IXMLDOMAttribute
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    Dim strSQL As String, _
        strSQLSet As String, _
        strSQLWhere As String, _
        strDataType As String, _
        strDataSource As String
    Dim intLoop As Integer, _
        intSetCount As Integer, _
        intKeys As Integer, _
        intKeyCount As Integer
    Dim intRetryCount As Integer, _
        intRetryMax As Integer
    Dim lngRecordsAffected As Long
    Dim blnDbCmdOk As Boolean
    strDataSource = xmlGetAttributeText(vxmlSchemaNode, "DATASRCE")
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaNode.nodeName
    End If
    intRetryMax = adoGetDbRetries
    Dim blnDoSet As Boolean
    Set cmd = New ADODB.Command
    strSQL = "UPDATE " & strDataSource
    Set xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@DATATYPE]")
    For Each xmlSchemaNode In xmlSchemaNodeList
        blnDoSet = True
        Set xmlSchemaAttrib = xmlSchemaNode.Attributes.getNamedItem("KEYTYPE")
        If Not xmlSchemaAttrib Is Nothing Then
            
            If (xmlSchemaAttrib.Value = "PRIMARY") Or _
                (xmlSchemaAttrib.Value = "SECONDARY") _
            Then
                blnDoSet = False
            End If
        End If
        If blnDoSet Then
            Set xmlDataAttrib = _
                vxmlDataNode.Attributes.getNamedItem(xmlSchemaNode.nodeName)
            If Not xmlDataAttrib Is Nothing Then
                If Len(strSQLSet) <> 0 Then
                    strSQLSet = strSQLSet & ", "
                End If
                strSQLSet = strSQLSet & xmlSchemaNode.nodeName & "=?"
                Set param = getParam(xmlSchemaNode, xmlDataAttrib)
                cmd.Parameters.Append param
                intSetCount = intSetCount + 1
            End If
        End If
    Next
    Set xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']")
    intKeys = xmlSchemaNodeList.length
    For Each xmlSchemaNode In xmlSchemaNodeList
        
        Set xmlDataAttrib = _
            vxmlDataNode.Attributes.getNamedItem(xmlSchemaNode.nodeName)
        If Not xmlDataAttrib Is Nothing Then
            If xmlDataAttrib.Value <> "" Then
                If Len(strSQLWhere) <> 0 Then
                    strSQLWhere = strSQLWhere & " AND "
                End If
                strSQLWhere = strSQLWhere & xmlSchemaNode.nodeName & "=?"
                Set param = getParam(xmlSchemaNode, xmlDataAttrib)
                cmd.Parameters.Append param
                intKeyCount = intKeyCount + 1
            End If
        End If
    Next
    strSQL = strSQL & " SET " & strSQLSet
    strSQL = strSQL & " WHERE " & strSQLWhere
    Debug.Print "adoUpdate strSQL: " & strSQL
    If (blnNonUniqueInstanceAllowed = False) And (intKeys <> intKeyCount) Then
        errThrowError "adoUpdate", oeXMLMissingAttribute
    End If
    If intSetCount > 0 Then
        
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
        lngRecordsAffected = executeSQLCommand(cmd)
    End If
'On Error Resume Next
'
'    Do While blnDbCmdOk = False
'
'        cmd.CommandType = adCmdText
'        cmd.CommandText = strSQL
'
'        lngRecordsAffected = executeSQLCommand(cmd)
'
'        If Err.Number = 0 Then
'
'            blnDbCmdOk = True
'
'        Else
'
'            If IsRetryError() = False Then
'
'                On Error GoTo adoUpdateExit
'
'            Else
'
'                intRetryCount = intRetryCount + 1
'
'                ' DEBUG
'                App.LogEvent "adoUpdate - retry [" & intRetryCount & "]: " & strSQL
'
'                If (intRetryCount = intRetryMax) Then
'
'                    On Error GoTo adoUpdateExit
'
'                End If
'
'            End If
'
'        End If
'
'    Loop
'
'    Set cmd = Nothing
    
    Debug.Print "adoUpdate records updated: " & lngRecordsAffected
adoUpdateExit:
    
    Set cmd = Nothing
    Set xmlSchemaNode = Nothing
    Set xmlSchemaNodeList = Nothing
    Set xmlSchemaAttrib = Nothing
    Set xmlDataAttrib = Nothing
    Set param = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoGetStoredProcAsXML( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo adoGetStoredProcAsXMLExit
    Const strFunctionName As String = "adoGetRecordSetAsXML"
    Dim xmlSchemaKeyNodeList As IXMLDOMNodeList
    Dim xmlSchemaKeyNode As IXMLDOMNode
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim strSQLParamClause As String
    Set cmd = New ADODB.Command
    Set xmlSchemaKeyNodeList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']")
    For Each xmlSchemaKeyNode In xmlSchemaKeyNodeList
        adoAddParam _
            xmlSchemaKeyNode, _
            xmlGetAttributeNode(vxmlDataNode, xmlSchemaKeyNode.nodeName), _
            cmd
        If Len(strSQLParamClause) = 0 Then
            strSQLParamClause = "?"
        Else
            strSQLParamClause = strSQLParamClause & ",?"
        End If
            
    Next
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "{CALL " & _
        vxmlSchemaNode.Attributes.getNamedItem("DATASRCE").Text & _
        "(" & strSQLParamClause & ")}"
    Set rst = executeGetRecordSet(cmd)
    If Not rst Is Nothing Then
        
        adoConvertRecordSetToXML _
            vxmlSchemaNode, _
            vxmlResponseNode, _
            rst, _
            IsComboLookUpRequired(vxmlDataNode)
    End If
adoGetStoredProcAsXMLExit:
    Set rst = Nothing
    Set cmd = Nothing
    Set xmlSchemaKeyNodeList = Nothing
    Set xmlSchemaKeyNode = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoGetRecordSetAsXML( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    Optional vstrFilter As String, _
    Optional vstrOrderBy As String)
    On Error GoTo adoGetRecordSetAsXMLExit
    Const strFunctionName As String = "adoGetRecordSetAsXML"
    Dim rst As ADODB.Recordset
    Set rst = adoGetRecordSet(vxmlDataNode, vxmlSchemaNode, vstrFilter, vstrOrderBy)
    If Not rst Is Nothing Then
        
        adoConvertRecordSetToXML _
            vxmlSchemaNode, _
            vxmlResponseNode, _
            rst, _
            IsComboLookUpRequired(vxmlDataNode)
    End If
adoGetRecordSetAsXMLExit:
    
    Set rst = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoCloneBaseEntityRefs(ByVal vxmlSchemaNode As IXMLDOMNode)
    
    Dim xmlSchemaElement As IXMLDOMNode
    Dim xmlBaseSchemaEntityNode As IXMLDOMNode
    Dim xmlBaseSchemaNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim strEntityRef As String
    For Each xmlSchemaElement In vxmlSchemaNode.childNodes
        If (xmlSchemaElement.Attributes.getNamedItem("DATASRCE") Is Nothing) And _
            (Not xmlSchemaElement.Attributes.getNamedItem("ENTITYREF") Is Nothing) _
            Then
            strEntityRef = xmlSchemaElement.Attributes.getNamedItem("ENTITYREF").Text
            If (xmlBaseSchemaEntityNode Is Nothing) Then
                Set xmlBaseSchemaEntityNode = adoGetSchema(strEntityRef)
            Else
                If xmlBaseSchemaEntityNode.nodeName <> strEntityRef Then
                    Set xmlBaseSchemaEntityNode = adoGetSchema(strEntityRef)
                End If
            End If
            If Not (xmlBaseSchemaEntityNode Is Nothing) Then
                Set xmlBaseSchemaNode = _
                    xmlBaseSchemaEntityNode.selectSingleNode(xmlSchemaElement.nodeName)
                If Not xmlBaseSchemaNode Is Nothing Then
                    For Each xmlAttrib In xmlBaseSchemaNode.Attributes
                        If Left(xmlAttrib.Name, 3) <> "KEY" Then
                            If xmlSchemaElement.Attributes.getNamedItem(xmlAttrib.Name) _
                                Is Nothing _
                            Then
                                xmlSchemaElement.Attributes.setNamedItem _
                                    xmlAttrib.cloneNode(True)
                            End If
                        End If
                    Next
                    xmlSchemaElement.Attributes.removeNamedItem "ENTITYREF"
'                    vxmlSchemaNode.replaceChild _
'                        xmlBaseSchemaNode.cloneNode(False), _
'                        xmlSchemaElement
                
                End If
            End If
        End If
    Next
    Set xmlSchemaElement = Nothing
    Set xmlBaseSchemaEntityNode = Nothing
    Set xmlBaseSchemaNode = Nothing
    Set xmlAttrib = Nothing
End Sub
Public Function adoGetRecordSet( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    Optional vstrFilter As String, _
    Optional vstrOrderBy As String) _
    As ADODB.Recordset
On Error GoTo adoGetRecordSetExit
    
    Const strFunctionName As String = "adoGetRecordSet"
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    Dim strSQL As String, _
        strSQLWhere As String, _
        strPattern As String, _
        strDataSource As String
    Dim intLoop As Integer
        
    Dim intRetryCount As Integer, _
        intRetryMax As Integer
    Dim blnDbCmdOk As Boolean
    Dim strSQLNoLock As String 'RF BMIDS00935 (SYS4752)
    strDataSource = xmlGetAttributeText(vxmlSchemaNode, "DATASRCE")
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaNode.nodeName
    End If
    'RF BMIDS00935 Start
    'SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
    'This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
    If (genumDbProvider = omiga4DBPROVIDERSQLServer) And (xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK") = "TRUE") Then
        strSQLNoLock = " WITH (NOLOCK)"
    End If
    'RF BMIDS00935 End
    intRetryMax = adoGetDbRetries
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
    ' used by comboAssist
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
            'RF BMIDS00935 Start
            'SYS4752 - allow schema to specify the SQL-Server (NOLOCK) hint.
            strSQL = "SELECT * FROM " & strDataSource & strSQLNoLock & " WHERE (" & strSQLWhere & ")"
            'RF BMIDS00935 End
        Else
            'RF BMIDS00935 Start
            'SYS4752 - allow schema to specify the SQL-Server (NOLOCK) hint.
            strSQL = "SELECT * FROM " & strDataSource & strSQLNoLock
            'RF BMIDS00935 End
            If Len(strSQLWhere) <> 0 Then
                strSQL = strSQL & " WHERE (" & strSQLWhere & ")"       'JLD SYS1774
            End If
        End If
        If Len(vstrOrderBy) > 0 Then
            strSQL = strSQL & " ORDER BY " & vstrOrderBy
        End If
    End If
    Debug.Print "adoGetRecordSet strSQL: " & strSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set adoGetRecordSet = executeGetRecordSet(cmd)
' retries never work !!!
'On Error Resume Next
'
'    Do While blnDbCmdOk = False
'
'        Set adoGetRecordSet = executeGetRecordSet(cmd)
'
'        If Err.Number = 0 Then
'
'            blnDbCmdOk = True
'
'        Else
'
'            If IsRetryError() = False Then
'
'                On Error GoTo adoGetRecordSetExit
'
'            Else
'
'                intRetryCount = intRetryCount + 1
'
'                ' DEBUG
'                App.LogEvent "adoGetRecordSet - retry [" & intRetryCount & "]: " & strSQL
'
'                If (intRetryCount = intRetryMax) Then
'
'                    On Error GoTo adoGetRecordSetExit
'
'                End If
'
'            End If
'
'        End If
'
'    Loop
    
    If adoGetRecordSet Is Nothing Then
        Debug.Print "adoGetRecordSet records retrieved: 0"
    Else
        Debug.Print "adoGetRecordSet records retrieved: "; adoGetRecordSet.RecordCount
    End If
adoGetRecordSetExit:
    
    Set cmd = Nothing
    Set xmlSchemaNode = Nothing
    Set xmlDataAttrib = Nothing
    Set param = Nothing
    errCheckError strFunctionName
End Function
Public Sub adoConvertRecordSetToXML( _
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
            If xmlGetAttributeText(vxmlSchemaNode, "ENTITYTYPE") = "PROCEDURE" Then
                If xmlGetAttributeText(xmlThisSchemaNode, "PARAMMODE") <> "IN" Then
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
    'SG 18/06/02 SYS4889
    'adoGetValidRecordset rst
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
Private Function executeSQLCommand( _
    ByVal cmd As ADODB.Command) _
    As Long
On Error GoTo executeSQLCommandExit
    
    Const strFunctionName As String = "executeSQLCommand"
        
    Dim lngRecordsAffected As Long
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = adoGetDbConnectString
    conn.Open
    cmd.ActiveConnection = conn
    cmd.Execute lngRecordsAffected, , adExecuteNoRecords
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    executeSQLCommand = lngRecordsAffected
executeSQLCommandExit:
    
    
    errCheckError strFunctionName
End Function
Public Sub adoAddParam( _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlAttrib As IXMLDOMAttribute, _
    ByVal vcmd As ADODB.Command)
    Dim param As ADODB.Parameter
    Set param = getParam(vxmlSchemaNode, vxmlAttrib)
    vcmd.Parameters.Append param
    Set param = Nothing
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
Private Sub GetGeneratedKey( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    Dim xmlAttrib As IXMLDOMAttribute
    Dim strDataType As String, _
        strGuid As String
    strDataType = xmlGetAttributeText(vxmlSchemaNode, "DATATYPE")
    Select Case strDataType
        Case "GUID"
            
            Set xmlAttrib = _
                vxmlDataNode.ownerDocument.createAttribute(vxmlSchemaNode.nodeName)
            xmlAttrib.Value = CreateGUID()
            vxmlDataNode.Attributes.setNamedItem xmlAttrib
            Set xmlAttrib = Nothing
        Case "SHORT"
            
            Set xmlAttrib = _
                vxmlDataNode.ownerDocument.createAttribute(vxmlSchemaNode.nodeName)
            xmlAttrib.Value = _
                GetNextKeySequence(vxmlDataNode, vxmlSchemaNode, vxmlSchemaNode.parentNode)
            vxmlDataNode.Attributes.setNamedItem xmlAttrib
            Set xmlAttrib = Nothing
    End Select
End Sub
Private Function GetNextKeySequence( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlSchemaParentNode As IXMLDOMNode) _
    As Integer
On Error GoTo GetNextKeySequenceExit
    
    Dim xmlNode As IXMLDOMNode
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim param As ADODB.Parameter
    Dim strSQL As String, _
        strSQLWhere As String, _
        strPattern As String, _
        strDataSource As String, _
        strSequenceFieldName
    Dim intThisSequence As Integer
        
    Dim blnDbCmdOk As Boolean
    Dim strSQLNoLock As String 'RF BMIDS00935 (SYS4572)
    strDataSource = xmlGetAttributeText(vxmlSchemaParentNode, "DATASRCE")
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaParentNode.nodeName
    End If
    strSequenceFieldName = vxmlSchemaNode.nodeName
    'RF BMIDS00935 Start
    'SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
    'This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
    If (genumDbProvider = omiga4DBPROVIDERSQLServer) And (xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK") = "TRUE") Then
        strSQLNoLock = " WITH (NOLOCK)"
    End If
    'RF BMIDS00935 End
    Set cmd = New ADODB.Command
    For Each xmlNode In vxmlSchemaParentNode.childNodes
        
        If Not xmlNode.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
            
            If xmlNode.Attributes.getNamedItem("KEYSOURCE") Is Nothing Then
                
                If Not vxmlDataNode.Attributes.getNamedItem(xmlNode.nodeName) Is Nothing Then
                    If Len(strSQLWhere) <> 0 Then
                        strSQLWhere = strSQLWhere & " AND "
                    End If
                    strSQLWhere = strSQLWhere & xmlNode.nodeName & "=?"
                    Set param = _
                        getParam(xmlNode, _
                            vxmlDataNode.Attributes.getNamedItem(xmlNode.nodeName))
                    cmd.Parameters.Append param
                End If
            End If
        End If
    Next
    'RF BMIDS00935 Start
    'SYS4752 - Allow schema to specify the SQL-Server (NOLOCK) hint.
    strSQL = "SELECT MAX(" & strSequenceFieldName & ") FROM " & strDataSource & strSQLNoLock & " WHERE (" & strSQLWhere & ")"
    'RF BMIDS00935 End
    Debug.Print "adoAssist.GetNextKeySequence strSQL: " & strSQL
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = executeGetRecordSet(cmd)
    If Not rst Is Nothing Then
        rst.MoveFirst
        If IsNull(rst.Fields.Item(0)) = False Then
            intThisSequence = rst.Fields.Item(0).Value
        End If
        rst.Close
    End If
    GetNextKeySequence = intThisSequence + 1
    Err.Clear
GetNextKeySequenceExit:
    Set rst = Nothing
    Set cmd = Nothing
    Set xmlNode = Nothing
    Set param = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
Public Function adoGuidToString(bytArray() As Byte) As String
    Dim i As Integer
    Dim strGuid As String
    strGuid = ""
    For i = 0 To 15
        strGuid = strGuid & Right$(Hex$(bytArray(i) + &H100), 2)
    Next i
    adoGuidToString = strGuid
End Function
Private Function adoDateToString(ByVal vvarDate As Variant) As String
    
    adoDateToString = Right(Str(DatePart("d", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("m", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("yyyy", vvarDate)), 4)
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
    'Dim xmlAttrib As IXMLDOMAttribute
    Dim strElementValue As String, _
        strComboGroup As String, _
        strComboValue As String, _
        strFormatMask As String, _
        strFormatted As String
            
    Select Case vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
        Case "GUID"
            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                adoGuidToString(vfld.Value)
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
                adoDateToString(vfld.Value)
        Case "DATETIME"
            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                adoDateTimeToString(vfld.Value)
        Case Else
            vxmlOutElem.setAttribute _
                vxmlSchemaNode.nodeName, _
                vfld.Value
    End Select
End Sub
Public Function adoDeleteFromNode( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vstrSchemaName As String, _
    Optional ByVal blnOnlySingleRecord As Boolean = True) As Long
On Error GoTo adoDeleteFromNodeExit
    
    Const strFunctionName As String = "adoDeleteFromNode"
    Dim lngRecordsAffected As Long
    Dim xmlSchemaNode As IXMLDOMNode
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    lngRecordsAffected = adoDelete(vxmlRequestNode, xmlSchemaNode, blnOnlySingleRecord)
    adoDeleteFromNode = lngRecordsAffected
adoDeleteFromNodeExit:
    
    Set xmlSchemaNode = Nothing
    errCheckError strFunctionName
End Function
Public Function adoDelete(ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    Optional ByVal blnOnlySingleRecord As Boolean = True) As Long
On Error GoTo adoDeleteExit
    
    Const strFunctionName As String = "adoDelete"
    Dim strDataSource As String
    Dim strSQL As String
    Dim strSQLWhere As String
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    Dim xmlDataAttrib As IXMLDOMAttribute
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    Dim lngRecordsAffected As Long
    Dim blnDbCmdOk As Boolean
    Dim intRetryMax As Integer
    Dim intRetryCount As Integer
    Set cmd = New ADODB.Command
    'Get the table name from the schema
    strDataSource = xmlGetAttributeText(vxmlSchemaNode, "DATASRCE")
    If Len(Trim$(strDataSource)) = 0 Then
        errThrowError strFunctionName, oeXMLMissingAttribute, "DATASRCE"
    End If
    'If specified that only a single record may be deleted, ensure that
    'the whole primary key has been provided. Otherwise raise an error.
    If blnOnlySingleRecord Then
        Set xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']")
        For Each xmlSchemaNode In xmlSchemaNodeList
            Set xmlDataAttrib = vxmlDataNode.Attributes.getNamedItem(xmlSchemaNode.nodeName)
            If Not xmlDataAttrib Is Nothing Then
                If xmlDataAttrib.Value = "" Then
                    errThrowError strFunctionName, oeXMLInvalidAttributeValue, xmlSchemaNode.nodeName
                End If
            Else
                errThrowError strFunctionName, oeXMLMissingAttribute, xmlSchemaNode.nodeName
            End If
        Next
    End If
    'Get each condition in the WHERE clause
    For Each xmlSchemaNode In vxmlSchemaNode.childNodes
        Set xmlDataAttrib = vxmlDataNode.Attributes.getNamedItem(xmlSchemaNode.nodeName)
        If Not xmlDataAttrib Is Nothing Then
            If xmlDataAttrib.Value <> "" Then
                If Len(strSQLWhere) <> 0 Then
                    strSQLWhere = strSQLWhere & " AND "
                End If
                strSQLWhere = strSQLWhere & xmlSchemaNode.nodeName & "=?"
                Set param = getParam(xmlSchemaNode, xmlDataAttrib)
                cmd.Parameters.Append param
            End If
        End If
    Next
    'Run the SQL command
    If Len(strSQLWhere) > 0 Then
        strSQL = "DELETE FROM " & strDataSource & " WHERE " & strSQLWhere
        Debug.Print "adoDelete strSQL: " & strSQL
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
'        intRetryMax = adoGetDbRetries
'        On Error Resume Next
'        Do While blnDbCmdOk = False
            lngRecordsAffected = executeSQLCommand(cmd)
'            If Err.Number = 0 Then
'                blnDbCmdOk = True
'            Else
'                If IsRetryError() = False Then
'                    On Error GoTo adoDeleteExit
'                Else
'                    intRetryCount = intRetryCount + 1
'                    ' DEBUG
'                    App.LogEvent "adoDelete - retry [" & intRetryCount & "]: " & strSQL
'                    If (intRetryCount = intRetryMax) Then
'                        On Error GoTo adoDeleteExit
'                    End If
'                End If
'            End If
'        Loop
    End If
    Debug.Print "adoDelete records deleted: " & lngRecordsAffected
    adoDelete = lngRecordsAffected
adoDeleteExit:
    Set cmd = Nothing
    Set param = Nothing
    Set xmlSchemaNode = Nothing
    Set xmlDataAttrib = Nothing
    Set param = Nothing
    errCheckError strFunctionName
End Function
Public Sub adoPopulateChildKeys( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode)
    Const strFunctionName As String = "adoPopulateChildKeys"
    On Error GoTo adoPopulateChildKeysExit
    Dim xmlSrceSchemaNode As IXMLDOMNode
    Dim xmlDestSchemaNode As IXMLDOMNode
    Dim xmlPrimaryKeyNodeList As IXMLDOMNodeList
    Set xmlSrceSchemaNode = adoGetSchema(vxmlSrceNode.nodeName)
    If xmlSrceSchemaNode Is Nothing Then
        Exit Sub
    End If
    Set xmlDestSchemaNode = adoGetSchema(vxmlDestNode.nodeName)
    If xmlDestSchemaNode Is Nothing Then
        Set xmlSrceSchemaNode = Nothing
        Exit Sub
    End If
    Set xmlPrimaryKeyNodeList = _
        xmlDestSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']")
    For Each xmlSrceSchemaNode In xmlPrimaryKeyNodeList
        xmlCopyAttribute vxmlSrceNode, vxmlDestNode, xmlSrceSchemaNode.nodeName
    Next
adoPopulateChildKeysExit:
    
    Set xmlSrceSchemaNode = Nothing
    Set xmlDestSchemaNode = Nothing
    Set xmlPrimaryKeyNodeList = Nothing
    errCheckError strFunctionName
End Sub
Public Sub adoBuildDbConnectionString()
On Error GoTo BuildDbConnectionStringVbErr
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
    strRegSection = "HKLM\SOFTWARE\" & gstrAppName & "\" & App.Title & "\" & gstrREGISTRY_SECTION & "\"
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
        ' PSC 17/10/01 SYS2815 - Start
        strUserId = ""
        strPassword = ""
        On Error Resume Next
                
        strUserId = objWshShell.RegRead(strRegSection & gstrUID_KEY)
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
                
        On Error GoTo BuildDbConnectionStringVbErr
        If Err.Number <> glngENTRYNOTFOUND And Err.Number <> 0 Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
        On Error Resume Next
                
        strPassword = objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY)
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
                
        On Error GoTo BuildDbConnectionStringVbErr
        If Err.Number <> glngENTRYNOTFOUND And Err.Number <> 0 Then
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
        ' PSC 17/10/01 SYS2815 - End
    End If
           
    gstrDbConnectionString = strConnection
    strRetries = objWshShell.RegRead(strRegSection & gstrRETRIES_KEY)
    If Len(strRetries) > 0 Then
        gintDbRetries = CInt(strRetries)
    End If
    Set objWshShell = Nothing
    Debug.Print strConnection
    Exit Sub
BuildDbConnectionStringVbErr:
    
    Set objWshShell = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function adoGetDbConnectString() As String
    adoGetDbConnectString = gstrDbConnectionString
End Function
Public Function adoGetDbRetries() As Integer
    adoGetDbRetries = gintDbRetries
End Function
Public Function adoGetDbProvider() As DBPROVIDER
    adoGetDbProvider = genumDbProvider
End Function
Public Sub adoLoadSchema()
    
    Dim strFileSpec As String
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\" & App.Title & ".xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    Set gxmldocSchemas = New FreeThreadedDOMDocument40
    gxmldocSchemas.async = False
    gxmldocSchemas.setProperty "NewParser", True
    gxmldocSchemas.validateOnParse = False
    gxmldocSchemas.load strFileSpec
End Sub
Public Function adoGetSchema(ByVal vstrSchemaName As String) As IXMLDOMNode
    
    Dim strPattern As String
    strPattern = "//" & vstrSchemaName & "[@DATASRCE]"
    Set adoGetSchema = gxmldocSchemas.selectSingleNode(strPattern)
End Function
Private Function GetErrorNumber(ByVal strErrDesc) As Long
On Error GoTo GetErrorNumberVbErr
    Const cstrFunctionName = "GetErrorNumber"
    Dim strErr As String
    Select Case genumDbProvider
        'SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
        Case omiga4DBPROVIDEROracle, omiga4DBPROVIDERMSOracle
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
    ''SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
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
