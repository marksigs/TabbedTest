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
'-------------------------------------------------------------------------------
'BMIDS Specific Change History
'Prog   Date        Description
'MO     25/07/2002  BMIDS00218 : Modified adoCRUDPopulateChildKeys so that keys are taken not
'                   only from the parent node but any parent node recursively,
'                   created method adoCRUDGetFirstNodeWhereAttributeExistsInRecursiveParent
'                   (sorry about the name!)
'MO     2/08/2002   BMIDS00218 : Modified
'MO     9/8/2002    BMIDS00218 : Fixed bug in adoCRUDGetNextKeySequence, which ment the secondary keys
'                       where being used as well as the primary keys to get the next number
'MO     15/08/2002  BMIDS00218 : Modified adoCRUDPopulateChildKeys so that a new method of attaining
'                       keys can be used which will get them from an xpath.
'MO     12/09/2002  BMIDS00435 : Modified adoCRUDGetNextKeySequence so run the SQL with nolocks
'MO     20/11/2002  BMIDS00999 : Fixed bug in adoCRUDGetFirstNodeWhereAttributeExistsInRecursiveParent where
'                                 there is no document_node (root), it fails with an object variable or with
'                                 block variable not set error
'SR     11/12/2002  BM0074   Apply NOLOCK while using MAX() for generating new sequence numbers.
'-----------------------------------------------------------------------------------------------------
Option Explicit

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

'APS - Implemented IK's CRUD changes 20/08/01
Public Function adoGetOperationNode( _
    ByVal vstrOperation As String) _
    As IXMLDOMNode
    
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
    
    Set adoGetOperationNode = _
               gxmldocSchemas.selectSingleNode("OM_SCHEMA/OPERATION[@NAME='" & _
                   vstrOperation & "']")
    
End Function

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUD( _
    ByVal vxmlOperationNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
   
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlSchemaStartNode As IXMLDOMNode
    Dim xmlRequestDataNode As IXMLDOMNode
   
    Dim strNodeName As String, _
        strPattern As String
        
    Dim intNodeIndex As Long
    
    ' check OPERATION details
    
    ' do we have a valid CRUD_OP
    Select Case UCase(xmlGetAttributeText(vxmlOperationNode, "CRUD_OP"))
    
        Case "CREATE", "READ", "UPDATE", "DELETE"
        
        Case Else
    
            errThrowError _
                "adoCRUD", _
                oeXMLInvalidAttributeValue, _
                "OM_SCHEMA OPERATION invalid CRUD_OP: " & vxmlOperationNode.xml
    
    End Select
    
    If Not xmlAttributeValueExists(vxmlOperationNode, "ENTITY_REF") Then
        errThrowError _
            "adoCRUD", _
            oeXMLMissingAttribute, _
            "schema OPERATION missing ENTITY_REF: " & vxmlOperationNode.xml
    End If
    
    strPattern = _
        "OM_SCHEMA/ENTITY[@NAME='" & _
        xmlGetAttributeText(vxmlOperationNode, "ENTITY_REF") & _
        "']"
    
    ' get schema node for ENTITY_REF
    Set xmlSchemaNode = gxmldocSchemas.selectSingleNode(strPattern)
    
    If xmlSchemaNode Is Nothing Then
        errThrowError _
            "adoCRUD", _
            oeXMLMissingElement, _
            "invalid schema reference: " & _
                xmlGetAttributeText(vxmlOperationNode, "ENTITY_REF")
    End If
        
    ' process each child node on the request
    For intNodeIndex = 0 To (vxmlRequestNode.childNodes.length - 1)
    
        Set xmlRequestDataNode = vxmlRequestNode.childNodes.Item(intNodeIndex)
    
        ' find start node within ENTITY for ENTITY_REF
        Set xmlSchemaStartNode = _
            adoCRUDGetSchemaStartNode(xmlRequestDataNode.nodeName, xmlSchemaNode)
    
        adoCRUDProcessRequestChildNode _
            xmlRequestDataNode, _
            vxmlOperationNode, _
            xmlSchemaStartNode, _
            vxmlResponseNode
            
        Set xmlSchemaStartNode = Nothing
                
    Next
    
    Set xmlSchemaNode = Nothing
    Set xmlSchemaStartNode = Nothing
    Set xmlRequestDataNode = Nothing
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDProcessRequestChildNode( _
    ByVal vxmlRequestDataNode As IXMLDOMNode, _
    ByVal vxmlOperationNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlSchemaRefNode As IXMLDOMNode
    
    Dim strCrudOp As String
        
    ' does this entity node contain a list of attributes with entity refs?
    If vxmlSchemaNode.selectNodes("ATTRIBUTE[@ENTITY_REF]").length <> 0 _
    Then
        
        adoCRUDPrepareAttributeEntityRefs vxmlSchemaNode
        
    End If
    
    If xmlAttributeValueExists(vxmlSchemaNode, "ENTITY_REF") Then
        
        Set xmlSchemaRefNode = adoCRUDGetSchemaRefNode(vxmlSchemaNode)
        
    End If
    
    strCrudOp = UCase(xmlGetAttributeText(vxmlOperationNode, "CRUD_OP"))
    
    Select Case strCrudOp
    
        Case "CREATE"
            adoCRUDCreate _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode
                
        Case "READ"
        
            adoCRUDRead _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode, _
                vxmlResponseNode
    
        Case "UPDATE"
            adoCRUDUpdate _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode
    
        Case "DELETE"
            adoCRUDDelete _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode
    
    End Select
    
    Set xmlSchemaRefNode = Nothing
    
End Sub

Private Function adoCRUDGetSchemaStartNode( _
    ByVal vstrRequestDataNodeName As String, _
    ByVal vxmlSchemaNode As IXMLDOMNode) _
    As IXMLDOMNode
    
    Dim strPattern As String
    
    strPattern = _
        ".[@NAME='" & vstrRequestDataNodeName & "' or " & _
        "@DATA_REF='" & vstrRequestDataNodeName & "' or " & _
        "@ENTITY_REF='" & vstrRequestDataNodeName & "']"
    
    Set adoCRUDGetSchemaStartNode = _
        vxmlSchemaNode.selectSingleNode(strPattern)
        
    If adoCRUDGetSchemaStartNode Is Nothing Then
    
        strPattern = _
            ".//*[@NAME='" & vstrRequestDataNodeName & "' or " & _
            "@DATA_REF='" & vstrRequestDataNodeName & "' or " & _
            "@ENTITY_REF='" & vstrRequestDataNodeName & "']"
        
        Set adoCRUDGetSchemaStartNode = _
            vxmlSchemaNode.selectSingleNode(strPattern)
    
    End If
        
    If adoCRUDGetSchemaStartNode Is Nothing Then
        
        errThrowError _
            "CRUD_OP", _
            oeXMLMissingElement, _
            "REQUEST data node: " & vstrRequestDataNodeName & _
                " not found in ENTITY: " & _
                xmlGetAttributeText(vxmlSchemaNode, "NAME")
    
    End If

End Function

Private Function adoCRUDGetSchemaChildNode( _
    ByVal vstrRequestDataNodeName As String, _
    ByVal vxmlSchemaNode As IXMLDOMNode) _
    As IXMLDOMNode
    
    Dim xmlThisSchemaNode As IXMLDOMNode
    
    Dim strPattern As String
    
    strPattern = _
        "*[@NAME='" & vstrRequestDataNodeName & "' or " & _
        "@DATA_REF='" & vstrRequestDataNodeName & "' or " & _
        "@ENTITY_REF='" & vstrRequestDataNodeName & "']"
    
    Set xmlThisSchemaNode = _
        vxmlSchemaNode.selectSingleNode(strPattern)
        
    If xmlThisSchemaNode Is Nothing Then
        
        errThrowError _
            "CRUD_OP", _
            oeXMLMissingElement, _
            "REQUEST data node: " & vstrRequestDataNodeName & _
                " not found in ENTITY: " & _
                xmlGetAttributeText(vxmlSchemaNode, "NAME")
                
    Else
        
        ' does this entity ref. contain a list of attributes refs?
        If xmlThisSchemaNode.selectNodes("ATTRIBUTE[@ENTITY_REF]").length <> 0 _
        Then
            
            adoCRUDPrepareAttributeEntityRefs xmlThisSchemaNode
            
        End If
    
    End If
    
    Set adoCRUDGetSchemaChildNode = xmlThisSchemaNode
    Set xmlThisSchemaNode = Nothing
    
End Function

Private Function adoCRUDGetSchemaRefNode( _
    ByVal vxmlSchemaNode As IXMLDOMNode) _
    As IXMLDOMNode
    
    Dim strPattern As String
    
    strPattern = _
               "OM_SCHEMA/ENTITY[@NAME='" & _
               xmlGetAttributeText(vxmlSchemaNode, "ENTITY_REF") & _
               "']"

    Set adoCRUDGetSchemaRefNode = gxmldocSchemas.selectSingleNode(strPattern)
    
    If adoCRUDGetSchemaRefNode Is Nothing Then
        errThrowError _
            "CRUD_OP", _
            oeXMLMissingElement, _
            "invalid schema ENTITY_REF on node: " & _
                vxmlSchemaNode.cloneNode(False).xml
    End If
        
    ' does this entity ref. contain a list of attributes refs?
    If adoCRUDGetSchemaRefNode.selectNodes("ATTRIBUTE[@ENTITY_REF]").length <> 0 _
    Then
        
        adoCRUDPrepareAttributeEntityRefs adoCRUDGetSchemaRefNode
        
    End If
    
End Function

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDPrepareAttributeEntityRefs( _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    
    Dim xmlAttribList As IXMLDOMNodeList
    
    Dim xmlSchemaSrceNode As IXMLDOMNode, _
        xmlAttribNode As IXMLDOMNode
        
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim strPattern As String
    
    Set xmlAttribList = vxmlSchemaNode.selectNodes("ATTRIBUTE[@ENTITY_REF]")
    
    For Each xmlAttribNode In xmlAttribList
    
        strPattern = _
            "OM_SCHEMA/ENTITY[@NAME='" & _
            xmlGetAttributeText(xmlAttribNode, "ENTITY_REF") & _
            "']/ATTRIBUTE[@NAME='" & _
            xmlGetAttributeText(xmlAttribNode, "NAME") & _
            "']"

        Set xmlSchemaSrceNode = gxmldocSchemas.selectSingleNode(strPattern)
        
        If xmlSchemaSrceNode Is Nothing Then
            errThrowError _
                "CRUD_OP", _
                oeXMLMissingElement, _
                "invalid schema ENTITY_REF on node: " & _
                    xmlAttribNode.cloneNode(False).xml
        End If
        
        ' don't copy KEY... attributes
        For Each xmlAttrib In xmlSchemaSrceNode.Attributes
            If (xmlAttrib.baseName <> "KEYTYPE") And (xmlAttrib.baseName <> "KEYSOURCE") Then
                If xmlAttribNode.Attributes.getNamedItem(xmlAttrib.baseName) Is Nothing Then
                    xmlAttribNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
                End If
            End If
        Next
        
        xmlAttribNode.Attributes.removeNamedItem ("ENTITY_REF")
    
    Next
    
    Set xmlSchemaSrceNode = Nothing
    Set xmlAttribNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlAttribList = Nothing
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreate( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Const strFunctionName As String = "adoCRUDCreate"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
        
    Dim strAttribName As String
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    ' get any generated keys & default values
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE")
        strAttribName = xmlGetMandatoryAttributeText(xmlSchemaFieldNode, "NAME")
        If vxmlRequestNode.Attributes.getNamedItem(strAttribName) Is Nothing Then
            If Not xmlSchemaFieldNode.Attributes.getNamedItem("DEFAULT") Is Nothing Then
                Set xmlAttrib = _
                    vxmlRequestNode.ownerDocument.createAttribute(strAttribName)
                xmlAttrib.Text = _
                    xmlSchemaFieldNode.Attributes.getNamedItem("DEFAULT").Text
                vxmlRequestNode.Attributes.setNamedItem xmlAttrib
            End If
            If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
                If (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "PRIMARY") Or _
                    (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "SECONDARY") _
                Then
                    If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYSOURCE") Is Nothing Then
                        If xmlSchemaFieldNode.Attributes.getNamedItem("KEYSOURCE").Text = "GENERATED" Then
                            adoCRUDGetGeneratedKey vxmlRequestNode, xmlSchemaFieldNode
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "CREATE_PROC") Then
    
        adoCRUDCreateViaSP _
            vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
    
    ElseIf xmlAttributeValueExists(xmlSchemaEntityNode, "CREATE_REF") Then
    
        adoCRUDCreateViaSQL _
            vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
            
    Else
    
        If vxmlSchemaRefNode Is Nothing Then
        
            If vxmlSchemaMasterNode.selectNodes("ENTITY").length = 0 Then
            
                errThrowError _
                    "CRUDOP", _
                    oeXMLMissingAttribute, _
                    "schema missing CREATE operation: " & _
                    vxmlSchemaMasterNode.cloneNode(False).xml
            
            End If
        
        Else
    
            If xmlAttributeValueExists(vxmlSchemaRefNode, "CREATE_PROC") Then
            
                adoCRUDCreateViaSP _
                    vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
            
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "CREATE_REF") Then
            
                adoCRUDCreateViaSQL _
                    vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
                    
            Else
            
                errThrowError _
                    "CRUDOP", _
                    oeXMLMissingAttribute, _
                    "schema missing CREATE operation: " & _
                    vxmlSchemaRefNode.cloneNode(False).xml
                    
            End If
        
        End If
    
    End If
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing

End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDGetGeneratedKey( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim strDataType As String, _
               strGuid As String, _
               strAttribName As String
    
    strAttribName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
    strDataType = xmlGetAttributeText(vxmlSchemaNode, "DATATYPE")
    
    Select Case strDataType
    
        Case "GUID"
            
            Set xmlAttrib = _
                vxmlDataNode.ownerDocument.createAttribute(strAttribName)
            
            xmlAttrib.Value = CreateGUID()
            
            vxmlDataNode.Attributes.setNamedItem xmlAttrib
            
            Set xmlAttrib = Nothing
    
        Case "SHORT"
            
            Set xmlAttrib = _
                vxmlDataNode.ownerDocument.createAttribute(strAttribName)
            
            xmlAttrib.Value = _
                adoCRUDGetNextKeySequence(vxmlDataNode, vxmlSchemaNode)
            
            vxmlDataNode.Attributes.setNamedItem xmlAttrib
            
            Set xmlAttrib = Nothing
    
    End Select
    
    Set xmlAttrib = Nothing
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Function adoCRUDGetNextKeySequence( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode) _
    As Integer
    
On Error GoTo adoCRUDGetNextKeySequenceExit
    
    Dim xmlNode As IXMLDOMNode

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim param As ADODB.Parameter
    
    Dim strSQL As String, _
        strSQLWhere As String, _
        strPattern As String, _
        strDataSource As String, _
        strSequenceFieldName As String, _
        strKeyAttribName As String, _
        strKeyDbRef As String
        
    Dim intThisSequence As Integer
        
    Dim blnDbCmdOk As Boolean
    Dim strSQLNoLock As String 'SR 11/12/2002 - BM0074
    
    ' parent node must be ENTITY with CREATE_REF
    
    strDataSource = xmlGetMandatoryAttributeText(vxmlSchemaNode.parentNode, "CREATE_REF")
    strSequenceFieldName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
       
    Set cmd = New ADODB.Command
    
    For Each xmlNode In vxmlSchemaNode.parentNode.childNodes
        
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlNode.nodeType <> NODE_CDATA_SECTION And _
            xmlNode.nodeType <> NODE_COMMENT And _
            xmlNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlNode.nodeType <> NODE_DOCUMENT_TYPE) Then
        
            If Not xmlNode.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
                
                'MO     9/8/2002    BMIDS00218 - Fixed bug - Start
                'Only need to use the primary keys to generate the next sequence number
                If xmlNode.Attributes.getNamedItem("KEYTYPE").Text = "PRIMARY" Then
                'MO     9/8/2002    BMIDS00218 - Fixed bug - End
                
                    If xmlNode.Attributes.getNamedItem("KEYSOURCE") Is Nothing Then
                    
                        strKeyAttribName = xmlGetAttributeText(xmlNode, "NAME")
                        
                        If Not vxmlDataNode.Attributes.getNamedItem(strKeyAttribName) Is Nothing Then
                        
                            If Len(strSQLWhere) <> 0 Then
                                strSQLWhere = strSQLWhere & " AND "
                            End If
                        
                            ' IK 05/11/2001 DB_REF support
                            If xmlAttributeValueExists(xmlNode, gstrDbRefId) Then
                                strKeyDbRef = xmlGetAttributeText(xmlNode, gstrDbRefId)
                            Else
                                strKeyDbRef = strKeyAttribName
                            End If
                        
                            strSQLWhere = strSQLWhere & strKeyDbRef & "=?"
                        
                            Set param = _
                                       getParam( _
                                           xmlNode, _
                                           vxmlDataNode.Attributes.getNamedItem(strKeyAttribName))
                        
                            cmd.Parameters.Append param
                            
                        End If
                    
                    End If
                
                'MO     9/8/2002    BMIDS00218 - Fixed bug - Start
                End If
                'MO     9/8/2002    BMIDS00218 - Fixed bug - End
                                
            End If
        
        End If

    Next
    
    'MO     12/09/2002  BMIDS00435 - Start
    'strSQL = _
    '    "SELECT MAX(" & strSequenceFieldName & _
    '    ") FROM " & strDataSource & _
    '    " WHERE (" & strSQLWhere & ")"
    strSQL = _
        "SELECT MAX(" & strSequenceFieldName & _
        ") FROM " & strDataSource & " WITH (NOLOCK) " & _
        " WHERE (" & strSQLWhere & ")"
    'MO     12/09/2002  BMIDS00435 - End
    
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

    adoCRUDGetNextKeySequence = intThisSequence + 1
    
    Err.Clear
    
adoCRUDGetNextKeySequenceExit:

    Set rst = Nothing
    Set cmd = Nothing
    Set xmlNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreateViaSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDCreateViaSPExit
    
    Const strFunctionName As String = "adoCRUDCreateViaSP"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQLParamClause As String, _
        strProcRef As String
        
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.childNodes
    
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlSchemaFieldNode.nodeType <> NODE_CDATA_SECTION And _
            xmlSchemaFieldNode.nodeType <> NODE_COMMENT And _
            xmlSchemaFieldNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlSchemaFieldNode.nodeType <> NODE_DOCUMENT_TYPE) Then
    
            If Not xmlGetAttributeText(xmlSchemaFieldNode, "KEYSOURCE") = "PROC" _
            Then
                
                If Len(strSQLParamClause) = 0 Then
                    strSQLParamClause = "?"
                Else
                    strSQLParamClause = strSQLParamClause & ",?"
                End If
            
                Set xmlDataAttrib = _
                    vxmlRequestNode.Attributes.getNamedItem( _
                       xmlGetAttributeText(xmlSchemaFieldNode, "NAME"))
            
                Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
                
                cmd.Parameters.Append param
                
            End If
        
        End If
    
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "CREATE_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "CREATE_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "CREATE_PROC")
    End If
    
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "{CALL " & strProcRef & "(" & strSQLParamClause & ")}"

    lngRecordsAffected = adoCRUDExecuteSP(cmd)
    
    If lngRecordsAffected = 0 Then
        errThrowError "CRUD_OP", oeNoRowsAffected
    End If
    
    Set cmd = Nothing
    
    Debug.Print "adoCRUDCreateViaSP records created: " & lngRecordsAffected
    
    adoCRUDCreateRequestOffspring _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode
    
adoCRUDCreateViaSPExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing
    
    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreateViaSQL( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
On Error GoTo adoCRUDCreateViaSQLExit
    
    Const strFunctionName As String = "adoCRUDCreateViaSQL"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQL As String, _
        strPattern As String, _
        strDataType As String, _
        strDataSource As String, _
        strColumns As String, _
        strValues As String
        
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "CREATE_REF") Then
        strDataSource = xmlGetAttributeText(vxmlSchemaMasterNode, "CREATE_REF")
    Else
        strDataSource = xmlGetAttributeText(vxmlSchemaRefNode, "CREATE_REF")
    End If
    
    Set cmd = New ADODB.Command
    
    For Each xmlDataAttrib In vxmlRequestNode.Attributes
    
        strPattern = "ATTRIBUTE[@NAME='" & xmlDataAttrib.Name & "']"
        
        Set xmlSchemaFieldNode = xmlSchemaEntityNode.selectSingleNode(strPattern)
        
        If Not xmlSchemaFieldNode Is Nothing Then
    
            If Len(strColumns) > 0 Then
                strColumns = strColumns & ","
            End If
                
            ' IK 05/11/2001 DB_REF support
            If xmlAttributeValueExists(xmlSchemaFieldNode, gstrDbRefId) Then
                strColumns = strColumns & _
                    xmlGetAttributeText(xmlSchemaFieldNode, gstrDbRefId)
            Else
                strColumns = strColumns & xmlDataAttrib.Name
            End If
        
            If Len(strValues) > 0 Then
                strValues = strValues & ","
            End If
            
            strValues = strValues & "?"
            
            Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
            
            cmd.Parameters.Append param
            
      End If
    
    Next
    
    strSQL = "INSERT INTO " & strDataSource & " (" & strColumns & ") VALUES (" & strValues & ")"
    Debug.Print "adoCreate strSQL: " & strSQL
     
    If Len(strValues) > 0 Then

        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
            
        lngRecordsAffected = executeSQLCommand(cmd)
        
        If lngRecordsAffected = 0 Then
            errThrowError "CRUD_OP", oeNoRowsAffected
        End If
        
    Else
    
        errThrowError "CRUD_OP", oeNoDataForCreate
        
    End If
    
    Set cmd = Nothing
    
    Debug.Print "adoCRUDCreateViaSQL records created: " & lngRecordsAffected
    
    adoCRUDCreateRequestOffspring _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode
    
adoCRUDCreateViaSQLExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing
    
    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDCreateRequestOffspring( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisSchemaMasterNode As IXMLDOMNode
    Dim xmlThisSchemaRefNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    
    Dim strPattern As String
    
    Dim intNodeIndex As Long
    
    If vxmlRequestNode.hasChildNodes Then
    
        For intNodeIndex = 0 To (vxmlRequestNode.childNodes.length - 1)
        
            Set xmlThisRequestNode = _
                vxmlRequestNode.childNodes.Item(intNodeIndex)
                
            Set xmlThisSchemaMasterNode = _
                adoCRUDGetSchemaChildNode( _
                    xmlThisRequestNode.nodeName, vxmlSchemaMasterNode)
                    
            If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                
                Set xmlThisSchemaRefNode = _
                    adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                
                adoCRUDPopulateChildKeys _
                    xmlThisRequestNode, _
                    xmlThisSchemaRefNode, _
                    vxmlRequestNode, _
                    "CREATE" 'MO - BMIDS00218 - 25/07/2002
                
            Else
            
                Set xmlThisSchemaRefNode = Nothing
                
                adoCRUDPopulateChildKeys _
                    xmlThisRequestNode, _
                    xmlThisSchemaMasterNode, _
                    vxmlRequestNode, _
                    "CREATE" 'MO - BMIDS00218 - 25/07/2002

            End If
            
            adoCRUDCreate _
                xmlThisRequestNode, _
                xmlThisSchemaMasterNode, _
                xmlThisSchemaRefNode

        Next
    
    End If
    
    Set xmlThisRequestNode = Nothing
    Set xmlThisSchemaMasterNode = Nothing
    Set xmlThisSchemaRefNode = Nothing
    Set xmlSchemaNodeList = Nothing
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDRead( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlThisResponseNode As IXMLDOMNode
    
    Dim strNodeName As String
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_PROC") Then
    
        adoCRUDReadViaSP _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vxmlResponseNode
    
    ElseIf xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_REF") Then
    
        adoCRUDReadViaSQL _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vxmlResponseNode
            
    Else
    
        If vxmlSchemaRefNode Is Nothing Then
    
            ' IK 26/10/2001
            ' allow schema node with no READ_REF or READ_PROC
            ' implies logical grouping
            
            If Not vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
                strNodeName = _
                    vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME").Text
            Else
                strNodeName = _
                    vxmlSchemaMasterNode.Attributes.getNamedItem("NAME").Text
            End If
            
            Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
            
            Set xmlThisResponseNode = vxmlResponseNode.appendChild(xmlElem)
            
            Set xmlElem = Nothing
            
            adoCRUDReadGetOffspring _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                xmlThisResponseNode
            
        Else
    
            If xmlAttributeValueExists(vxmlSchemaRefNode, "READ_PROC") Then
            
                adoCRUDReadViaSP _
                    vxmlRequestNode, _
                    vxmlSchemaMasterNode, _
                    vxmlSchemaRefNode, _
                    vxmlResponseNode
            
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "READ_REF") Then
            
                adoCRUDReadViaSQL _
                    vxmlRequestNode, _
                    vxmlSchemaMasterNode, _
                    vxmlSchemaRefNode, _
                    vxmlResponseNode
                    
            End If
            
        End If
    
    End If
    
    Set xmlElem = Nothing
    Set xmlThisResponseNode = Nothing

End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDReadViaSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const strFunctionName As String = "adoCRUDReadViaSP"
    
    On Error GoTo adoCRUDReadViaSPExit
    
    Dim xmlSchemaKeyNodeList As IXMLDOMNodeList
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaKeyNode As IXMLDOMNode
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim strSQLParamClause As String, _
        strAttribName As String, _
        strProcRef As String
    
    Set cmd = New ADODB.Command
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "READ_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "READ_PROC")
    End If
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set xmlSchemaKeyNodeList = xmlSchemaEntityNode.selectNodes("*[@KEYTYPE='PRIMARY']")
    
    For Each xmlSchemaKeyNode In xmlSchemaKeyNodeList
    
        strAttribName = xmlGetMandatoryAttributeText(xmlSchemaKeyNode, "NAME")
    
        adoAddParam _
            xmlSchemaKeyNode, _
            xmlGetAttributeNode(vxmlRequestNode, strAttribName), _
            cmd
            
        If Len(strSQLParamClause) = 0 Then
            strSQLParamClause = "?"
        Else
            strSQLParamClause = strSQLParamClause & ",?"
        End If
            
    Next
    
    cmd.CommandType = adCmdText
    cmd.CommandText = _
               "{CALL " & strProcRef & "(" & strSQLParamClause & ")}"
    
    Debug.Print strFunctionName & " strSQL: " & cmd.CommandText
        
    '   COMPONENT_REF only implemented for CRUD_OP="READ"
    If xmlAttributeValueExists(xmlSchemaEntityNode, gstrComponentRefId) Then
        Set rst = _
            executeGetRecordSet( _
                cmd, _
                UCase(xmlGetAttributeText(xmlSchemaEntityNode, gstrComponentRefId)))
    Else
        Set rst = executeGetRecordSet(cmd)
    End If
    
    If Not rst Is Nothing Then
        
        adoCRUDConvertRecordSetToXML _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vxmlResponseNode, _
            rst
    
    End If
    
adoCRUDReadViaSPExit:

    Set rst = Nothing
    Set cmd = Nothing
    
    Set xmlSchemaKeyNodeList = Nothing
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaKeyNode = Nothing

    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDReadViaSQL( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDReadViaSQLExit
    
    Const strFunctionName As String = "adoCRUDReadViaSQL"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaAttribNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim param As ADODB.Parameter
    
    Dim strSQL As String, _
        strSQLWhere As String, _
        strPattern As String, _
        strDataSource As String, _
        strDbColumnName As String
        
    Dim blnDbCmdOk As Boolean
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_REF") Then
        strDataSource = xmlGetAttributeText(vxmlSchemaMasterNode, "READ_REF")
    Else
        strDataSource = xmlGetAttributeText(vxmlSchemaRefNode, "READ_REF")
    End If
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
        
    For Each xmlDataAttrib In vxmlRequestNode.Attributes
    
        strPattern = _
                   "ATTRIBUTE[@NAME='" & xmlDataAttrib.Name & "'][@DATATYPE]"
    
        Set xmlSchemaAttribNode = _
            xmlSchemaEntityNode.selectSingleNode(strPattern)
        
        If Not xmlSchemaAttribNode Is Nothing Then
        
            If Len(strSQLWhere) <> 0 Then
                strSQLWhere = strSQLWhere & " AND "
            End If
            
            ' IK 05/11/2001 DB_REF support
            If xmlAttributeValueExists(xmlSchemaAttribNode, gstrDbRefId) Then
                strDbColumnName = _
                    xmlGetAttributeText(xmlSchemaAttribNode, gstrDbRefId)
            Else
                strDbColumnName = xmlDataAttrib.Name
            End If
        
            ' APS 03/09/01 - Explicit handling of null columns
            If Len(xmlDataAttrib.Text) = 0 Then
                strSQLWhere = strSQLWhere & strDbColumnName & " is NULL"
            Else
            
                strSQLWhere = strSQLWhere & strDbColumnName & "=?"
                
                Set param = getParam(xmlSchemaAttribNode, xmlDataAttrib)
        
                cmd.Parameters.Append param
                
            End If
        
        End If
    
    Next
        
    If cmd.Parameters.Count <> 0 Then
        
        strSQL = _
            "SELECT * FROM " & strDataSource & _
            " WHERE (" & strSQLWhere & ")"
            
    Else
    
        ' IK 19/10/2001
        ' select * only available at top of hierarchy
        
        If vxmlSchemaMasterNode.parentNode.nodeName <> "ENTITY" Then
            strSQL = "SELECT * FROM " & strDataSource
        End If
        
    End If
    
    Debug.Print strFunctionName & " strSQL: " & strSQL
    
    If Len(strSQL) > 0 Then
    
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
        
'       COMPONENT_REF only implemented for CRUD_OP="READ"
        If xmlAttributeValueExists(xmlSchemaEntityNode, gstrComponentRefId) Then
            Set rst = _
                executeGetRecordSet( _
                    cmd, _
                    UCase(xmlGetAttributeText( _
                        xmlSchemaEntityNode, gstrComponentRefId)))
        Else
            Set rst = executeGetRecordSet(cmd)
        End If
    
        If Not rst Is Nothing Then

            adoCRUDConvertRecordSetToXML _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                vxmlResponseNode, _
                rst
        
        End If
        
    End If
    
adoCRUDReadViaSQLExit:

    Set rst = Nothing
    Set cmd = Nothing
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaAttribNode = Nothing
    Set xmlDataAttrib = Nothing

    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDConvertRecordSetToXML( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vrst As ADODB.Recordset)
    
    Dim xmlSchemaAttributeList As IXMLDOMNodeList
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaAttribNode As IXMLDOMNode
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    Dim fld As ADODB.Field
    
    Dim strNodeName As String, _
        strAttribName As String
        
    Dim blnDoComboLookUp As Boolean
    
    Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    
    If vxmlSchemaMasterNode.selectNodes("ATTRIBUTE").length = 0 Then
        If Not vxmlSchemaRefNode Is Nothing Then
            Set xmlSchemaEntityNode = vxmlSchemaRefNode
        End If
    End If
    
    strNodeName = "UNDEFINED"
        
    If Not vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
        strNodeName = _
            vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME").Text
    Else
        If Not vxmlSchemaRefNode Is Nothing Then
            If Not vxmlSchemaRefNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
                strNodeName = _
                    vxmlSchemaRefNode.Attributes.getNamedItem("OUTNAME").Text
            Else
                strNodeName = _
                    vxmlSchemaRefNode.Attributes.getNamedItem("NAME").Text
            End If
        Else
            If Not vxmlSchemaMasterNode.Attributes.getNamedItem("NAME") Is Nothing Then
                strNodeName = _
                    vxmlSchemaMasterNode.Attributes.getNamedItem("NAME").Text
            End If
        End If
    End If
    
    blnDoComboLookUp = IsComboLookUpRequired(vxmlRequestNode)
    
    Do While Not vrst.EOF
    
        If vxmlResponseNode.nodeName = strNodeName Then
            Set xmlElem = vxmlResponseNode
        Else
            Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
        End If
        
        ' Get the attribute list again to avoid COM+/XML3 bug
        Set xmlSchemaAttributeList = xmlSchemaEntityNode.selectNodes("ATTRIBUTE")
        
        For Each xmlSchemaAttribNode In xmlSchemaAttributeList
                
            ' IK 05/11/2001 DB_REF support
            If xmlAttributeValueExists(xmlSchemaAttribNode, gstrDbRefId) Then
                strAttribName = _
                    xmlGetAttributeText(xmlSchemaAttribNode, gstrDbRefId)
            Else
                strAttribName = xmlGetAttributeText(xmlSchemaAttribNode, "NAME")
            End If
        
            Set fld = Nothing
            If xmlGetAttributeText(xmlSchemaAttribNode, "PARAM_TYPE") <> "IN" Then
                Set fld = vrst.Fields.Item(strAttribName)
            End If
            
            If Not fld Is Nothing Then
    
                If Not IsNull(fld) Then
            
                    If Not IsNull(fld.Value) Then
                    
                        FieldToXml fld, xmlElem, xmlSchemaAttribNode, blnDoComboLookUp
            
                    End If
                
                End If
                
            End If
        
        Next
        
        ' APS - 03/09/01 Added in extra check to make sure that we
        ' do not append the node when it already exists
        If xmlElem.Attributes.length > 0 And _
            vxmlResponseNode.nodeName <> strNodeName _
        Then
        
            Set xmlThisResponseNode = vxmlResponseNode.appendChild(xmlElem)
            
            adoCRUDReadGetOffspring _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                xmlThisResponseNode
        
        End If
            
        vrst.MoveNext
        
    Loop

    vrst.Close
    
    Set fld = Nothing
    
    Set xmlElem = Nothing
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaAttribNode = Nothing
    Set xmlThisResponseNode = Nothing
    Set xmlSchemaAttributeList = Nothing
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDReadGetOffspring( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisSchemaMasterNode As IXMLDOMNode
    Dim xmlThisSchemaRefNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    
    Dim strPattern As String
    
    Dim intNodeIndex As Long
    
    If xmlGetAttributeText(vxmlRequestNode, gstrExtractTypeId) = "NODE" Then
        ' stop at this (request) node
        Exit Sub
    End If
    
    If xmlGetAttributeText(vxmlSchemaMasterNode, gstrExtractTypeId) = "NODE" Then
        ' stop at this (schema) node
        Exit Sub
    End If
    
    If vxmlRequestNode.hasChildNodes Then
    
        For intNodeIndex = 0 To (vxmlRequestNode.childNodes.length - 1)
        
            Set xmlThisRequestNode = _
                vxmlRequestNode.childNodes.Item(intNodeIndex)
                
            Set xmlThisSchemaMasterNode = _
                adoCRUDGetSchemaChildNode( _
                    xmlThisRequestNode.nodeName, vxmlSchemaMasterNode)
                    
            If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                
                Set xmlThisSchemaRefNode = _
                    adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                
                adoCRUDPopulateChildKeys _
                    xmlThisRequestNode, _
                    xmlThisSchemaRefNode, _
                    vxmlResponseNode, _
                    "READ" 'MO - BMIDS00218 - 25/07/2002
                
            Else
            
                Set xmlThisSchemaRefNode = Nothing
                
                adoCRUDPopulateChildKeys _
                    xmlThisRequestNode, _
                    xmlThisSchemaMasterNode, _
                    vxmlResponseNode, _
                    "READ" 'MO - BMIDS00218 - 25/07/2002
            
            End If
            
            adoCRUDRead _
                xmlThisRequestNode, _
                xmlThisSchemaMasterNode, _
                xmlThisSchemaRefNode, _
                vxmlResponseNode

        Next
    
    Else
    
        Set xmlSchemaNodeList = vxmlSchemaMasterNode.selectNodes("ENTITY")
    
        If xmlSchemaNodeList.length > 0 Then
    
            For intNodeIndex = 0 To (xmlSchemaNodeList.length - 1)
                
                Set xmlThisSchemaMasterNode = _
                    xmlSchemaNodeList.Item(intNodeIndex)
        
                ' does this entity ref. contain a list of attributes refs?
                If xmlThisSchemaMasterNode.selectNodes("ATTRIBUTE[@ENTITY_REF]").length _
                    <> 0 _
                Then
                    
                    adoCRUDPrepareAttributeEntityRefs xmlThisSchemaMasterNode
                    
                End If
                
                If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                    
                    Set xmlThisSchemaRefNode = _
                        adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                    
                    adoCRUDPopulateChildKeys _
                        vxmlRequestNode, _
                        xmlThisSchemaRefNode, _
                        vxmlResponseNode, _
                        "READ" 'MO - BMIDS00218 - 25/07/2002
                    
                Else
                
                    Set xmlThisSchemaRefNode = Nothing
                    
                    adoCRUDPopulateChildKeys _
                        vxmlRequestNode, _
                        xmlThisSchemaMasterNode, _
                        vxmlResponseNode, _
                        "READ" 'MO - BMIDS00218 - 25/07/2002
                
                End If
                
                adoCRUDRead _
                    vxmlRequestNode, _
                    xmlThisSchemaMasterNode, _
                    xmlThisSchemaRefNode, _
                    vxmlResponseNode
    
            Next
            
        Else
        
            If Not vxmlSchemaRefNode Is Nothing Then
    
                Set xmlSchemaNodeList = vxmlSchemaRefNode.selectNodes("ENTITY")
            
                If xmlSchemaNodeList.length > 0 Then
            
                    For intNodeIndex = 0 To (xmlSchemaNodeList.length - 1)
                        
                        Set xmlThisSchemaMasterNode = _
                            xmlSchemaNodeList.Item(intNodeIndex)
        
                        ' does this entity ref. contain a list of attributes refs?
                        If xmlThisSchemaMasterNode.selectNodes("ATTRIBUTE[@ENTITY_REF]").length _
                            <> 0 _
                        Then
                            
                            adoCRUDPrepareAttributeEntityRefs xmlThisSchemaMasterNode
                            
                        End If
                
                        If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                            
                            Set xmlThisSchemaRefNode = _
                                adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                            
                            adoCRUDPopulateChildKeys _
                                vxmlRequestNode, _
                                xmlThisSchemaRefNode, _
                                vxmlResponseNode, _
                                "READ" 'MO - BMIDS00218 - 25/07/2002
                                
                            
                        Else
                        
                            Set xmlThisSchemaRefNode = Nothing
                            
                            adoCRUDPopulateChildKeys _
                                vxmlRequestNode, _
                                xmlThisSchemaMasterNode, _
                                vxmlResponseNode, _
                                "READ" 'MO - BMIDS00218 - 25/07/2002
                        
                        End If
                        
                        adoCRUDRead _
                            vxmlRequestNode, _
                            xmlThisSchemaMasterNode, _
                            xmlThisSchemaRefNode, _
                            vxmlResponseNode
            
                    Next
                    
                End If
                
            End If
            
        End If
    
    End If
    
    Set xmlThisRequestNode = Nothing
    Set xmlThisSchemaMasterNode = Nothing
    Set xmlThisSchemaRefNode = Nothing
    Set xmlSchemaNodeList = Nothing
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDUpdate( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Const strFunctionName As String = "adoCRUDUpdate"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
        
    Dim strAttribName As String
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "UPDATE_PROC") Then
    
        adoCRUDUpdateViaSP _
            vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
    
    ElseIf xmlAttributeValueExists(xmlSchemaEntityNode, "UPDATE_REF") Then
    
        adoCRUDUpdateViaSQL _
            vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
            
    Else
    
        If vxmlSchemaRefNode Is Nothing Then
        
            If vxmlSchemaMasterNode.selectNodes("ENTITY").length = 0 Then
            
                errThrowError _
                    "CRUDOP", _
                    oeXMLMissingAttribute, _
                    "schema missing UPDATE operation: " & _
                    vxmlSchemaMasterNode.cloneNode(False).xml
            
            End If
        
        Else
    
            If xmlAttributeValueExists(vxmlSchemaRefNode, "UPDATE_PROC") Then
            
                adoCRUDUpdateViaSP _
                    vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
            
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "UPDATE_REF") Then
            
                adoCRUDUpdateViaSQL _
                    vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
                    
            Else
            
                errThrowError _
                    "CRUDOP", _
                    oeXMLMissingAttribute, _
                    "schema missing UPDATE operation: " & _
                    vxmlSchemaRefNode.cloneNode(False).xml
                    
            End If
            
        End If
    
    End If

End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDUpdateViaSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDUpdateViaSPExit
    
    Const strFunctionName As String = "adoCRUDUpdateViaSP"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQLParamClause As String, _
        strProcRef As String
        
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.childNodes
    
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlSchemaFieldNode.nodeType <> NODE_CDATA_SECTION And _
            xmlSchemaFieldNode.nodeType <> NODE_COMMENT And _
            xmlSchemaFieldNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlSchemaFieldNode.nodeType <> NODE_DOCUMENT_TYPE) Then
    
            If Not xmlGetAttributeText(xmlSchemaFieldNode, "KEYSOURCE") = "PROC" _
            Then
                
                If Len(strSQLParamClause) = 0 Then
                    strSQLParamClause = "?"
                Else
                    strSQLParamClause = strSQLParamClause & ",?"
                End If
            
                Set xmlDataAttrib = _
                    vxmlRequestNode.Attributes.getNamedItem( _
                       xmlGetAttributeText(xmlSchemaFieldNode, "NAME"))
            
                Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
                
                cmd.Parameters.Append param
                
            End If
        
        End If
    
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "UPDATE_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "UPDATE_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "UPDATE_PROC")
    End If
    
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "{CALL " & strProcRef & "(" & strSQLParamClause & ")}"

    lngRecordsAffected = adoCRUDExecuteSP(cmd)
    
    If lngRecordsAffected = 0 Then
        errThrowError "CRUD_OP", oeNoRowsAffected
    End If
    
    Set cmd = Nothing
    
    Debug.Print "adoCRUDUpdateViaSP records updated: " & lngRecordsAffected
    
    adoCRUDUpdateRequestOffspring _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode
    
adoCRUDUpdateViaSPExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing
    
    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDUpdateViaSQL( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
On Error GoTo adoCRUDUpdateViaSQLExit
    
    Const strFunctionName As String = "adoCRUDUpdateViaSQL"
    
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlSchemaAttrib As IXMLDOMAttribute
    Dim xmlDataAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQL As String, _
               strSQLSet As String, _
               strSQLWhere As String, _
               strDataType As String, _
               strDataSource As String, _
               strAttribName As String
        
    Dim intLoop As Integer, _
               intSetCount As Integer, _
               intKeys As Integer, _
               intKeyCount As Integer
        
    Dim lngRecordsAffected As Long
    
    Dim blnNonUniqueInstanceAllowed As Boolean
    Dim blnDoSet As Boolean
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "UPDATE_REF") Then
        strDataSource = xmlGetAttributeText(vxmlSchemaMasterNode, "UPDATE_REF")
    Else
        strDataSource = xmlGetAttributeText(vxmlSchemaRefNode, "UPDATE_REF")
    End If
    
    Set cmd = New ADODB.Command
    
    strSQL = "UPDATE " & strDataSource
    
    Set xmlSchemaNodeList = xmlSchemaEntityNode.selectNodes("*[@DATATYPE]")
    
    For Each xmlSchemaFieldNode In xmlSchemaNodeList
    
        blnDoSet = True
    
        Set xmlSchemaAttrib = xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE")
        
        If Not xmlSchemaAttrib Is Nothing Then
            
            If (xmlSchemaAttrib.Value = "PRIMARY") Or _
                (xmlSchemaAttrib.Value = "SECONDARY") _
            Then
                blnDoSet = False
            End If
            
        End If
        
        If blnDoSet Then
        
            strAttribName = xmlGetAttributeText(xmlSchemaFieldNode, "NAME")
        
            Set xmlDataAttrib = _
                vxmlRequestNode.Attributes.getNamedItem(strAttribName)
                
            If Not xmlDataAttrib Is Nothing Then
        
                If Len(strSQLSet) <> 0 Then
                    strSQLSet = strSQLSet & ", "
                End If
                    
                ' IK 05/11/2001 DB_REF support
                If xmlAttributeValueExists(xmlSchemaFieldNode, gstrDbRefId) Then
                    strAttribName = xmlGetAttributeText(xmlSchemaFieldNode, gstrDbRefId)
                End If
            
                strSQLSet = strSQLSet & strAttribName & "=?"
                
                Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
        
                cmd.Parameters.Append param
                
                intSetCount = intSetCount + 1
                
            End If
            
        End If
    
    Next
    
    Set xmlSchemaNodeList = xmlSchemaEntityNode.selectNodes("*[@KEYTYPE='PRIMARY']")
    intKeys = xmlSchemaNodeList.length
    
    For Each xmlSchemaFieldNode In xmlSchemaNodeList
    
        strAttribName = xmlGetAttributeText(xmlSchemaFieldNode, "NAME")
        
        Set xmlDataAttrib = _
                   vxmlRequestNode.Attributes.getNamedItem(strAttribName)
            
        If Not xmlDataAttrib Is Nothing Then
        
            If xmlDataAttrib.Value <> "" Then
    
                If Len(strSQLWhere) <> 0 Then
                    strSQLWhere = strSQLWhere & " AND "
                End If
            
                strSQLWhere = strSQLWhere & strAttribName & "=?"
                
                Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
        
                cmd.Parameters.Append param
                
                intKeyCount = intKeyCount + 1
                
            End If
            
        End If
    
    Next
    
    strSQL = strSQL & " SET " & strSQLSet
    strSQL = strSQL & " WHERE " & strSQLWhere
    
    Debug.Print "adoUpdate strSQL: " & strSQL
    
    ' TODO
    ' modify schema to allow non-unique updates
    ' (schema specific?)
    
    If (blnNonUniqueInstanceAllowed = False) And (intKeys <> intKeyCount) Then
        errThrowError "adoUpdate", oeXMLMissingAttribute
    End If

    If intSetCount > 0 Then
        
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
            
        lngRecordsAffected = executeSQLCommand(cmd)
        
        If lngRecordsAffected = 0 Then
            errThrowError "CRUD_OP", oeNoRowsAffected
        End If
    
    Else
    
        errThrowError "CRUD_OP", oeNoDataForUpdate
    
    End If
    
    Debug.Print "adoUpdate records updated: " & lngRecordsAffected
    
    adoCRUDUpdateRequestOffspring _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode
    
adoCRUDUpdateViaSQLExit:
    
    Set xmlSchemaNodeList = Nothing
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlSchemaAttrib = Nothing
    Set xmlDataAttrib = Nothing
    
    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Public Sub adoCRUDUpdateRequestOffspring( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisSchemaMasterNode As IXMLDOMNode
    Dim xmlThisSchemaRefNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    
    Dim strPattern As String
    
    Dim intNodeIndex As Long
    
    If vxmlRequestNode.hasChildNodes Then
    
        For intNodeIndex = 0 To (vxmlRequestNode.childNodes.length - 1)
        
            Set xmlThisRequestNode = _
                vxmlRequestNode.childNodes.Item(intNodeIndex)
                
            Set xmlThisSchemaMasterNode = _
                adoCRUDGetSchemaChildNode( _
                    xmlThisRequestNode.nodeName, vxmlSchemaMasterNode)
                    
            If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                
                Set xmlThisSchemaRefNode = _
                    adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                
                adoCRUDPopulateChildKeys _
                    xmlThisRequestNode, _
                    xmlThisSchemaRefNode, _
                    vxmlRequestNode, _
                    "UPDATE" 'MO - BMIDS00218 - 25/07/2002
            Else
            
                Set xmlThisSchemaRefNode = Nothing
                
                adoCRUDPopulateChildKeys _
                    xmlThisRequestNode, _
                    xmlThisSchemaMasterNode, _
                    vxmlRequestNode, _
                    "UPDATE" 'MO - BMIDS00218 - 25/07/2002
                
            End If
            
            adoCRUDUpdate _
                xmlThisRequestNode, _
                xmlThisSchemaMasterNode, _
                xmlThisSchemaRefNode

        Next
    
    End If
    
End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDDelete( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Const strFunctionName As String = "adoCRUDDelete"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "DELETE_PROC") Then
    
        adoCRUDDeleteViaSP _
            vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
    
    ElseIf xmlAttributeValueExists(xmlSchemaEntityNode, "DELETE_REF") Then
    
        errThrowError "CRUD_OP", oeNotImplemented, "DELETE_REF"
            
    Else
    
        If vxmlSchemaRefNode Is Nothing Then
        
            If vxmlSchemaMasterNode.selectNodes("ENTITY").length = 0 Then
            
                errThrowError _
                    "CRUDOP", _
                    oeXMLMissingAttribute, _
                    "schema missing DELETE_PROC operation: " & _
                    vxmlSchemaMasterNode.cloneNode(False).xml
            
            End If
        
        Else
    
            If xmlAttributeValueExists(vxmlSchemaRefNode, "DELETE_PROC") Then
            
                adoCRUDDeleteViaSP _
                    vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
            
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "DELETE_REF") Then
                
                errThrowError "CRUD_OP", oeNotImplemented, "DELETE_REF"
            
            Else
            
                errThrowError _
                    "CRUDOP", _
                    oeXMLMissingAttribute, _
                    "schema missing DELETE_PROC operation: " & _
                    vxmlSchemaRefNode.cloneNode(False).xml
        
            End If
            
        End If
    
    End If

End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDDeleteViaSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDDeleteViaSPExit
    
    Const strFunctionName As String = "adoCRUDDeleteViaSP"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute

    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQLParamClause As String, _
        strProcRef As String
        
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.childNodes
    
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlSchemaFieldNode.nodeType <> NODE_CDATA_SECTION And _
            xmlSchemaFieldNode.nodeType <> NODE_COMMENT And _
            xmlSchemaFieldNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlSchemaFieldNode.nodeType <> NODE_DOCUMENT_TYPE) Then
             
            If Not xmlGetAttributeText(xmlSchemaFieldNode, "KEYSOURCE") = "PROC" _
            Then
                
                If Len(strSQLParamClause) = 0 Then
                    strSQLParamClause = "?"
                Else
                    strSQLParamClause = strSQLParamClause & ",?"
                End If
            
                Set xmlDataAttrib = _
                    vxmlRequestNode.Attributes.getNamedItem( _
                       xmlGetAttributeText(xmlSchemaFieldNode, "NAME"))
            
                Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
                
                cmd.Parameters.Append param
                
            End If
    
        End If
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "DELETE_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "DELETE_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "DELETE_PROC")
    End If
    
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "{CALL " & strProcRef & "(" & strSQLParamClause & ")}"

    lngRecordsAffected = adoCRUDExecuteSP(cmd)
    
    If lngRecordsAffected = 0 Then
        errThrowError "CRUD_OP", oeNoRowsAffected
    End If
    
    Set cmd = Nothing
    
    Debug.Print "adoCRUDDeleteViaSP records deleted: " & lngRecordsAffected
    
adoCRUDDeleteViaSPExit:
    
    errCheckError strFunctionName
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDPopulateChildKeys( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrCRUDOperation As String)
    
    Dim xmlKeyNodes As IXMLDOMNodeList
    Dim xmlKeyNode As IXMLDOMNode
    Dim xmlParentNode As IXMLDOMNode
    
    Dim strAttribName As String, _
        strAttribValue As String
        
    Dim strXPathToKey As String
    Dim strValueGainedFromXPath As String
    
    'MO - BMIDS00218 - 25/07/2002 - Start
    'Set xmlParentNode = vxmlParentNode
    'Do While xmlParentNode.Attributes.length = 0
    '
    '    Set xmlParentNode = xmlParentNode.parentNode
    '
    '    If xmlParentNode.parentNode Is Nothing Then
    '
    '        Set xmlParentNode = Nothing
    '        Exit Do
    '
    '    End If
    '
    'Loop
    
    'If xmlParentNode Is Nothing Then
    '    Exit Sub
    'End If
    'MO - BMIDS00218 - 25/07/2002 - End
    
    Set xmlKeyNodes = vxmlSchemaNode.selectNodes("ATTRIBUTE[@KEYTYPE]")
    
    For Each xmlKeyNode In xmlKeyNodes
        
        'MO - BMIDS00218 - 25/07/2002 - Start
        ' If the attribute INHERITONCREATE has been set to false, dont inherit the key from
        ' any parent, if an INHERITONCREATE is not specified TRUE is default
        If vstrCRUDOperation <> "CREATE" Or xmlGetAttributeText(xmlKeyNode, "INHERITKEYONCREATE") <> "FALSE" Then
            
            strAttribName = _
                    xmlGetMandatoryAttributeText(xmlKeyNode, "NAME")
            
            'MO     15/08/2002  BMIDS00218 - Start
            'If the inheritkeyfromxpath = true, then read the xpath out of the value in the node and replace it with the value gained from this xpath
            If xmlGetAttributeText(xmlKeyNode, "INHERITKEYFROMXPATH") = "TRUE" Then
                
                'get the xpath that has been passed in as the value of the attribute
                strXPathToKey = xmlGetAttributeText(vxmlDataNode, strAttribName)
                
                'get the value of the xpath from the datanode
                strValueGainedFromXPath = xmlGetMandatoryNode(vxmlDataNode, strXPathToKey).Text
                
                'get the key value
                xmlSetAttributeValue vxmlDataNode, strAttribName, strValueGainedFromXPath
                                
            'MO     15/08/2002  BMIDS00218 - End
            Else
                
                
                'if the value already exists in the datanode there is no need to inherit it
                If xmlAttributeValueExists(vxmlDataNode, strAttribName) = False Then
                           
                    Set xmlParentNode = adoCRUDGetFirstNodeWhereAttributeExistsInRecursiveParent(vxmlDataNode, strAttribName, vxmlDataNode.nodeName)
                    
                    If Not xmlParentNode Is Nothing Then
                        xmlCopyAttribute xmlParentNode, vxmlDataNode, strAttribName
                    End If
                                    
                End If
                                    
                'If xmlAttributeValueExists(xmlParentNode, strAttribName) Then
                '
                '    xmlCopyAttribute _
                '        xmlParentNode, vxmlDataNode, strAttribName
                '
                'End If
            End If
            
        End If
        'MO - BMIDS00218 - 25/07/2002 - End
        
    Next
    
End Sub

'MO - BMIDS00218 - 25/07/2002 - New method to cope with Primary keys not in the parent node, but in any recursive parent node
Private Function adoCRUDGetFirstNodeWhereAttributeExistsInRecursiveParent(objNode As IXMLDOMNode, ByVal strAttributeName As String, ByVal strOriginalNodeName As String) As IXMLDOMNode

    Dim objParentNode As IXMLDOMNode
    Dim blnFoundAttribute As Boolean
    
    'get this nodes parent
    Set objParentNode = objNode.parentNode
    'MO - 20/11/2002 - BMIDS00999
    'If the parent node is the document node (or root) dont go any further
    If (Not objParentNode Is Nothing) Then
        If objParentNode.nodeType = NODE_DOCUMENT Then
            Set objParentNode = Nothing
        End If
    End If
    
    blnFoundAttribute = False
    
    'loop though until we get to the document node (root) or we find an attribute
    While (Not objParentNode Is Nothing) And (blnFoundAttribute = False)
        
        'does this node have the same name as the node we are trying to find,
        ' it is has dismiss it, as we will end up with duplicate keys,
        ' ADDRESS/ADDRESSGUID for example is the address table and the primary key
        If objParentNode.nodeName <> strOriginalNodeName Then
            
            'does the attribute exist in this parent?
            blnFoundAttribute = xmlAttributeValueExists(objParentNode, strAttributeName)

        End If
                
        'if it doesnt move onto the next
        If blnFoundAttribute = False Then
            'get this parent's parent
            Set objParentNode = objParentNode.parentNode
            'MO - 20/11/2002 - BMIDS00999
            'If the parent node is the document node (or root) dont go any further
            If (Not objParentNode Is Nothing) Then
                If objParentNode.nodeType = NODE_DOCUMENT Then
                    Set objParentNode = Nothing
                End If
            End If
        Else
            'return the node
            Set adoCRUDGetFirstNodeWhereAttributeExistsInRecursiveParent = objParentNode
        End If
    Wend
    
    Set objParentNode = Nothing
    
End Function

'APS - Implemented IK's CRUD changes 20/08/01
Private Function adoCRUDExecuteSP( _
    ByVal cmd As ADODB.Command, _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As Long
    
On Error GoTo adoCRUDExecuteSPExit
    
    Const strFunctionName As String = "adoCRUDExecuteSP"
    
    Dim lngRecordsAffected As Long
    
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    
    conn.ConnectionString = adoGetDbConnectString(vstrComponentRef)
    conn.Open
    
    cmd.ActiveConnection = conn

    cmd.Execute lngRecordsAffected
    
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    
    adoCRUDExecuteSP = lngRecordsAffected

adoCRUDExecuteSPExit:
    
    errCheckError strFunctionName

End Function

Public Sub adoCreateFromNodeList( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vstrSchemaName As String)
    
On Error GoTo adoCreateFromNodeListExit
    
    Const strFunctionName As String = "adoCreateFromNodeList"
    
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    
    For Each xmlDataNode In vxmlRequestNodeList
        adoCreate xmlDataNode, xmlSchemaNode
    Next
    
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
    Dim xmlDataNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema(vstrSchemaName)
    
    For Each xmlDataNode In vxmlRequestNodeList
        adoUpdate xmlDataNode, xmlSchemaNode
    Next
    
adoUpdateFromNodeListExit:
    
    Set xmlSchemaNode = Nothing
    Set xmlDataNode = Nothing
    
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
    
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlSchemaFieldNode.nodeType <> NODE_CDATA_SECTION And _
            xmlSchemaFieldNode.nodeType <> NODE_COMMENT And _
            xmlSchemaFieldNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlSchemaFieldNode.nodeType <> NODE_DOCUMENT_TYPE) Then
    
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
            
            Set param = getParam(xmlSchemaFieldNode, xmlDataAttrib)
            
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
    
    errCheckError strFunctionName
    
End Sub

Public Sub adoGetStoredProcAsXML( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo adoGetStoredProcAsXMLExit
    
    Const strFunctionName As String = "adoGetStoredProcAsXML"
    
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
    
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlSchemaElement.nodeType <> NODE_CDATA_SECTION And _
            xmlSchemaElement.nodeType <> NODE_COMMENT And _
            xmlSchemaElement.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlSchemaElement.nodeType <> NODE_DOCUMENT_TYPE) Then
    
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
    
        End If
    
    Next
    
    Set xmlSchemaElement = Nothing
    Set xmlBaseSchemaEntityNode = Nothing
    Set xmlBaseSchemaNode = Nothing

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
    
    Dim strSQLNoLock As String 'SR 11/11/2002 : BM0074
    
    strDataSource = xmlGetAttributeText(vxmlSchemaNode, "DATASRCE")
    
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaNode.nodeName
    End If
    
    'SR 11/12/2002 : BM0074
    'SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
    'This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
    If (genumDbProvider = omiga4DBPROVIDERSQLServer) And (xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK") = "TRUE") Then
        strSQLNoLock = " WITH (NOLOCK)"
    End If
    'SR 11/12/2002 : BM0074 - End
    
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
            strSQL = "SELECT * FROM " & strDataSource & strSQLNoLock & " WHERE (" & strSQLWhere & ")" ' SR 11/12/2002 : BM0074
        Else
            strSQL = "SELECT * FROM " & strDataSource
            strSQL = "SELECT * FROM " & strDataSource & strSQLNoLock ' SR 11/12/2002 : BM0074
            
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
        
            'AL - NodeType checking to enable comments on schema file 08/02/2002
            If (xmlThisSchemaNode.nodeType <> NODE_CDATA_SECTION And _
                xmlThisSchemaNode.nodeType <> NODE_COMMENT And _
                xmlThisSchemaNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
                xmlThisSchemaNode.nodeType <> NODE_DOCUMENT_TYPE) Then
        
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

Private Function executeSQLCommand( _
    ByVal cmd As ADODB.Command, _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As Long
    
On Error GoTo executeSQLCommandExit
    
    Const strFunctionName As String = "executeSQLCommand"
    
    Dim lngRecordsAffected As Long
    
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    
    conn.ConnectionString = adoGetDbConnectString(vstrComponentRef)
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
            
            strDataValue = GuidStringToByteArray(strDataValue)
            param.Type = adBinary
            param.Size = 16
        
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
    Dim strSQLNoLock As String 'SR 11/12/2002 : BM0074
    
    strDataSource = xmlGetAttributeText(vxmlSchemaParentNode, "DATASRCE")
    
    If Len(strDataSource) = 0 Then
        strDataSource = vxmlSchemaParentNode.nodeName
    End If
    
    strSequenceFieldName = vxmlSchemaNode.nodeName
        
    'SR 11/12/2002: BM0074
    If (genumDbProvider = omiga4DBPROVIDERSQLServer) And (xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK") = "TRUE") Then
        strSQLNoLock = " WITH (NOLOCK)"
    End If
    'SR 11/12/2002: BM0074 - End

    Set cmd = New ADODB.Command
    
    For Each xmlNode In vxmlSchemaParentNode.childNodes
        
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlNode.nodeType <> NODE_CDATA_SECTION And _
            xmlNode.nodeType <> NODE_COMMENT And _
            xmlNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlNode.nodeType <> NODE_DOCUMENT_TYPE) Then
        
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
    
        End If
    
    Next
    
    ' SR 11/12/2002 :BM0074
    strSQL = "SELECT MAX(" & strSequenceFieldName & ") FROM " & strDataSource & strSQLNoLock & " WHERE (" & strSQLWhere & ")"
    ' SR 11/12/2002 :BM0074 - End
    
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
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

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

    Dim strElementValue As String, _
               strComboGroup As String, _
               strComboValue As String, _
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

        Case "COMBO"

            vxmlOutElem.setAttribute strAttribName, vfld.Value
                
            If vblnDoComboLookUp = True Then

                If Not vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP") Is Nothing Then
    
                    strComboGroup = vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP").Text
                    strComboValue = GetComboText(strComboGroup, vfld.Value)
                
                    vxmlOutElem.setAttribute _
                               strAttribName & "_TEXT", strComboValue
    
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
    
        'AL - NodeType checking to enable comments on schema file 08/02/2002
        If (xmlSchemaNode.nodeType <> NODE_CDATA_SECTION And _
            xmlSchemaNode.nodeType <> NODE_COMMENT And _
            xmlSchemaNode.nodeType <> NODE_PROCESSING_INSTRUCTION And _
            xmlSchemaNode.nodeType <> NODE_DOCUMENT_TYPE) Then
    
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

Public Sub adoBuildDbConnectionString( _
    Optional ByVal vstrComponentRef As String = "DEFAULT")

On Error GoTo BuildDbConnectionStringVbErr

    Dim objWshShell As Object
    
    Dim strThisComponentRef As String, _
        strConnection As String, _
        strProvider As String, _
        strRegSection As String, _
        strRetries As String
        
    If vstrComponentRef = "DEFAULT" Then
        strThisComponentRef = App.Title
    Else
        strThisComponentRef = vstrComponentRef
    End If
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    strRegSection = "HKLM\SOFTWARE\" & gstrAppName & "\" & strThisComponentRef & "\" & gstrREGISTRY_SECTION & "\"
    
'On Error Resume Next
    
'    strProvider = objWshShell.RegRead(strRegSection & gstrPROVIDER_KEY)

On Error GoTo BuildDbConnectionStringVbErr
    
    If Len(strProvider) = 0 Then
        strRegSection = "HKLM\SOFTWARE\" & gstrAppName & "\" & gstrREGISTRY_SECTION & "\"
        strProvider = objWshShell.RegRead(strRegSection & gstrPROVIDER_KEY)
    End If
    
    genumDbProvider = omiga4DBPROVIDERUnknown
    
    If strProvider = "MSDAORA" Then
        genumDbProvider = omiga4DBPROVIDEROracle
        strConnection = _
                   "Provider=MSDAORA;Data Source=" & objWshShell.RegRead(strRegSection & gstrDATA_SOURCE_KEY) & ";" & _
                   "User ID=" & objWshShell.RegRead(strRegSection & gstrUID_KEY) & ";" & _
                   "Password=" & objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY) & ";"
    ElseIf strProvider = "SQLOLEDB" Then
        genumDbProvider = omiga4DBPROVIDERSQLServer
        strConnection = _
                   "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & gstrSERVER_KEY) & ";" & _
                   "database=" & objWshShell.RegRead(strRegSection & gstrDATABASE_KEY) & ";" & _
                   "UID=" & objWshShell.RegRead(strRegSection & gstrUID_KEY) & ";" & _
                   "pwd=" & objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY) & ";"
    End If
           
    gstrDbConnectionString = strConnection
    
    strRetries = objWshShell.RegRead(strRegSection & gstrRETRIES_KEY)
    
    If Len(strRetries) > 0 Then
        gintDbRetries = CInt(strRetries)
    End If
    
    Set objWshShell = Nothing
    
    putConnectString vstrComponentRef, strConnection
    
    Debug.Print strConnection
    
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
        gxmldocConnections.validateOnParse = False
        gxmldocConnections.setProperty "NewParser", True
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
        
    Debug.Print adoGetDbConnectString

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
    gxmldocSchemas.validateOnParse = False
    gxmldocSchemas.setProperty "NewParser", True
    
    gxmldocSchemas.async = False
    gxmldocSchemas.Load strFileSpec

    geDateStyle = e_dateUK
    gstrBooleanTrue = "1"
    gstrBooleanFalse = "0"
    
    If gxmldocSchemas.parseError.errorCode = 0 Then
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
    End If

End Sub

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

    Const cstrFunctionName As String = "GetErrorNumber"
    
    Dim strErr As String
    
    Select Case genumDbProvider
        Case omiga4DBPROVIDEROracle
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
    
    If genumDbProvider = omiga4DBPROVIDEROracle Then
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
#If GENERIC_SQL Then

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

    Const cstrFunctionName As String = "adoGetValidRecordset"
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

#End If
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
            Case omiga4DBPROVIDEROracle
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
