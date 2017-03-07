Attribute VB_Name = "adoAssistCRUD"
'-----------------------------------------------------------------------------
'Prog   Date        Description
'SR     05/03/2001  New Private function 'getOmigaNumberForDatabaseError'
'AW     20/03/2001  Added 'PARAMMODE' check to adoConvertRecordSetToXML
'AS     31/05/2001  Added adoGetValidRecordset for SQL Server Port
'AS     11/06/2001  Fixed compile error with GENERIC_SQL=1
'LD     12/06/2001  SYS2368 Modify code to get the connection string for Oracle and SQLServer
'LD     11/06/2001  SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     15/06/2001  SYS2368 SQL Server Port - Guids to be generated as type adBinary in function adoCRUDCreateSQLParam
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
'AS             13/11/03        CORE1 Removed GENERIC_SQL.
'SDS    20/01/2004  LIVE00009653 - NOLOCK is used only for SQLServer sql statements
'SDS    22/01/2004  LIVE00009659  Added pieces of code to support both Microsoft and Oracle OLEDB Drivers
'PSC    29/07/2004  BBG1141 Added SQL Server Integrated Security
'------------------------------------------------------------------------
' CORE Change History:
'
'Prog   Date        Defect / Description
'IK     27/07/2005  CORE158
'                   CORE158/1 restrict use of anonymous schema nodes
'                       only allowed when not in schema hierarchy
'                   CORE158/2 adoCRUDReadPopulateChildKeys enhancements
'                       if xpath GETDEFAULT returns null value do not return node
'                       if no KEYTYPE="FOREIGN" values retrieved do not return node
'                   CORE158/3 allow nested ENTITY_REF values
'IK     25/07/2005  CORE158/4 ATTRIBUTE ENTITY_REF values must be resolved in base schema
'IK     03/08/2005  CORE179 preserve & restore REQUEST values when overriden by GETDEFAULT values
'IK     04/08/2005  CORE179/1 remove extra RESPONSE TYPE for embedded FOR XML reads
'IK     04/08/2005  CORE172 tracing on when flag = 1 or 2 (as per printing)
'IK     23/09/2005  CORE190 allow comments in REQUEST & schema(s) - only process ELEMENT child nodes
'IK     23/09/2005  CORE192 remove ResponseValue prefix in adoCRUDReadViaAutoSP
'IK     23/09/2005  CORE194 allow SPs to return out parameters only - fix in adoCRUDExecuteGetRecordSet
'IK     12/10/2005  CORE199 delete processing using READ parameters
'IK     18/10/2005  CORE203 allow xpath: references in REQUEST attributes
'IK     21/10/2005  CORE206 round CURRENCY values on output
'IK     27/10/2005  CORE211 fix for CORE203 when xpath reference not found
'IK     11/11/2005  CORE214 do not retain PUTDEFAULT values in request
'IK     11/11/2005  CORE215 search parent hierarchy for primary key values if not input
'IK     16/11/2005  CORE215/2 also required for CREATE
'IK     30/11/2005  CORE220 - reset read request when no records found
'IK     05/12/2005  CORE223 - add _LOCKHINT to READ operation
'IK     03/02/2006  CORE238 - add XMLSTREAM data type STRING or TEXT field containing XML
'IK     13/02/2006  CORE244 - error in adoCRUDGetSchemaChildNode when column name same as table name
'                           - fix to getDereferencedSchema to allow first node in hierarchy
'                             to be ENTITY_REF node
'IK     01/03/2006  CORE249 - boolean flags via named schema, encoding via schema
'IK     15/03/2006  CORE250 - fix trace file registry look-up
'IK     26/04/2006  CORE261 - (limited?) support for BINARY data (read-only)
'IK     08/08/2006  CORE289 - switch off validation of  XMLSTREAM data
'IK     09/08/2006  CORE290 - add NOREADVALUE to READ_REF, if exists do not add item to read request
'IK     29/09/2006  CORE301 - only add 'group node' when child nodes returned (READ processing)
'IK     08/11/2006  CORE308 - all 'base:' prefix to ENTITY_REF
'IK     08/11/2006  CORE308 - CORE301 backed out (could impact existing code)
'------------------------------------------------------------------------------------------------------
Option Explicit

Private gxmldocBaseSchema As DOMDocument40
Private gcolNamedSchemas As Collection
Private gxmlCRUDCombos As DOMDocument40

Private gxmldocConnections As DOMDocument40
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
Private Enum DBPROVIDER
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

Private Enum COMBOVALIDATIONLOOKUPTYPE
    e_cvluNone = 0
    e_cvluPipeDelim = 1
    e_cvluExtended = 2
End Enum


Private geDateStyle As DATESTYLE
Private gstrBooleanTrue As String
Private gstrBooleanFalse As String
Private Const gstrSchemaNameId As String = "_SCHEMA_"
Private Const gstrExtractTypeId As String = "_EXTRACTTYPE_"
Private Const gstrOutNodeId As String = "_OUTNODENAME_"
Private Const gstrOrderById As String = "_ORDERBY_"
Private Const gstrWhereId As String = "_WHERE_"
Private Const gstrDataTypeId As String = "DATATYPE"
Private Const gstrDataTypeDerivedId As String = "DERIVED"
Private Const gstrDataSourceId As String = "DATASRCE"
Private Const gstrDbRefId As String = "DB_REF"
Private Const gstrComponentRefId As String = "COMPONENT_REF"
'APS - Implemented IK's CRUD changes 20/08/01
Private Const gstrSchemaRefId = "SCHEMA_REF"
Private Const gstrMODULEPREFIX As String = "adoAssist."
'APS - Implemented IK's CRUD changes 20/08/01

Private Const glngENTRYNOTFOUND As Long = -2147024894

Private Const cint_MIN = 10
Private Const cint_MAX = 40

'CORE238 - trace file naming enhancement
Private gstrTraceDateTime As String
Private gintTraceDateTimeQ As Integer

Public Function adoCRUDGetOperationNode( _
    ByVal vstrOperation As String) _
    As IXMLDOMNode
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & "adoGetOperationNode")
    #End If
    If gxmldocBaseSchema Is Nothing Then
        errThrowError _
            "adoGetSchema", _
            oeSchemaNotLoaded
    End If
    If gxmldocBaseSchema.parseError.errorCode <> 0 Then
        errThrowError _
            "adoGetSchema", _
            oeSchemaParseError, _
            gxmldocBaseSchema.parseError.reason
    End If
    Set adoCRUDGetOperationNode = _
               gxmldocBaseSchema.selectSingleNode("OM_SCHEMA/OPERATION[@NAME='" & _
                   vstrOperation & "']")
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
End Function

Public Sub adoCRUD( _
    ByVal vxmlOperationNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlNamedSchemaDoc As DOMDocument40
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlSchemaStartNode As IXMLDOMNode
    Dim xmlRequestDataNode As IXMLDOMNode
    
    Dim strNodeName As String, _
        strPattern As String
    
    Dim intNodeIndex As Long
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & "adoCRUD")
    #End If
        
    ' check OPERATION details
    ' do we have a valid CRUD_OP
    ' ik_20040301
    ' new CRUD_OP i_UPDATE is "intelligent" update
    Select Case UCase(xmlGetAttributeText(vxmlOperationNode, "CRUD_OP"))
        Case "CREATE", "READ", "UPDATE", "DELETE", "I_UPDATE", "IUPDATE"
        Case Else
            Err.Raise _
                oeXMLInvalidAttributeValue, _
                "adoCRUD", _
                "OM_SCHEMA OPERATION invalid CRUD_OP: " & vxmlOperationNode.xml
    End Select
    
' ik_20050610_wip
    
    If xmlAttributeValueExists(vxmlOperationNode, "SCHEMA_NAME") Then
        Set xmlNamedSchemaDoc = adoCRUDGetNamedSchema(xmlGetAttributeText(vxmlOperationNode, "SCHEMA_NAME"))
        If xmlNamedSchemaDoc Is Nothing Then
            Err.Raise _
                oeXMLInvalidRequestNode, _
                "adoCRUD", _
                "error opening schema file: " & _
                    xmlGetAttributeText(vxmlOperationNode, "SCHEMA_NAME")
        End If
    End If
    
    If xmlAttributeValueExists(vxmlOperationNode, "ENTITY_REF") Then
    
        ' ik_20050627_CORE158/3
        getDereferencedSchema _
            xmlNamedSchemaDoc, _
            xmlGetAttributeText(vxmlOperationNode, "ENTITY_REF"), _
            xmlSchemaNode
        ' ik_20050627_CORE158_ends
        
    Else
    
        strPattern = "OM_SCHEMA"
        
        If Not xmlNamedSchemaDoc Is Nothing Then
            Set xmlSchemaNode = xmlNamedSchemaDoc.selectSingleNode(strPattern)
        End If

        If xmlSchemaNode Is Nothing Then
            Set xmlSchemaNode = gxmldocBaseSchema.selectSingleNode(strPattern)
        End If
    
    End If
    
'    If xmlAttributeValueExists(vxmlOperationNode, "ENTITY_REF") Then
'
'        strPattern = _
'            "OM_SCHEMA/ENTITY[@NAME='" & _
'            xmlGetAttributeText(vxmlOperationNode, "ENTITY_REF") & _
'            "']"
'
'    Else
'
'        strPattern = "OM_SCHEMA"
'
'    End If
'
'    ' get schema node for ENTITY_REF
'    If xmlAttributeValueExists(vxmlOperationNode, "SCHEMA_NAME") Then
'        Set xmlNamedSchemaDoc = adoCRUDGetNamedSchema(xmlGetAttributeText(vxmlOperationNode, "SCHEMA_NAME"))
'        If xmlNamedSchemaDoc Is Nothing Then
'            Err.Raise _
'                oeXMLInvalidRequestNode, _
'                "adoCRUD", _
'                "error opening schema file: " & _
'                    xmlGetAttributeText(vxmlOperationNode, "SCHEMA_NAME")
'        End If
'        Set xmlSchemaNode = xmlNamedSchemaDoc.selectSingleNode(strPattern)
'    End If
'
'    If xmlSchemaNode Is Nothing Then
'        Set xmlSchemaNode = gxmldocBaseSchema.selectSingleNode(strPattern)
'    End If
' ik_20050610_wip_ends
    
    If xmlSchemaNode Is Nothing Then
        
        If xmlAttributeValueExists(vxmlOperationNode, "ENTITY_REF") Then
            Err.Raise _
                oeXMLMissingElement, _
                "adoCRUD", _
                "invalid schema reference: " & _
                    xmlGetAttributeText(vxmlOperationNode, "ENTITY_REF")
        Else
            Err.Raise _
                oeXMLMissingElement, _
                "adoCRUD", _
                "invalid schema reference: " & strPattern
        End If
    
    End If
    
    Select Case UCase(xmlGetAttributeText(vxmlOperationNode, "CRUD_OP"))
        
        Case "I_UPDATE", "IUPDATE"
        
            If xmlAttributeValueExists(vxmlOperationNode, "CRUDPREPROC") Then
                adoCRUDPreProcessRequest vxmlRequestNode, xmlSchemaNode
            End If
        
    End Select
        
    ' process each child node on the request
    For Each xmlRequestDataNode In vxmlRequestNode.childNodes
        ' ik_CORE190_250923
        If xmlRequestDataNode.nodeType = NODE_ELEMENT Then
            ' find start node within ENTITY for ENTITY_REF
            Set xmlSchemaStartNode = _
                adoCRUDGetSchemaStartNode(xmlRequestDataNode.nodeName, xmlSchemaNode)
            
            adoCRUDProcessRequestChildNode _
                xmlRequestDataNode, _
                vxmlOperationNode, _
                xmlSchemaStartNode, _
                vxmlResponseNode
            
            Set xmlSchemaStartNode = Nothing
        ' ik_CORE190_250923
        End If
        
    Next
    
    Set xmlSchemaNode = Nothing
    Set xmlSchemaStartNode = Nothing
    Set xmlRequestDataNode = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
End Sub

' ik_20040427_wip
Private Sub adoCRUDPreProcessRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    
    Dim xmlPreProcReadRequest As DOMDocument40
    Dim xmlPreProcReadResponse As DOMDocument40
    Dim xmlPreProcReadRequestNode As IXMLDOMNode
    Dim xmlPreProcReadResponseNode As IXMLDOMNode
    
    Dim xmlUpdateDataNodeList As IXMLDOMNodeList
    Dim xmlExistingDataNodeList As IXMLDOMNodeList
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strNodeName As String
    
    Set xmlPreProcReadRequest = New DOMDocument40
    xmlPreProcReadRequest.async = False

    adoCRUDPreProcGetReadRequest xmlPreProcReadRequest, vxmlRequestNode, vxmlSchemaNode

' ik_debug
    Debug.Print "Pre-Process Request"
    Debug.Print xmlPreProcReadRequest.xml
' ik_debug_ends
    
    Set xmlPreProcReadResponse = New DOMDocument40
    xmlPreProcReadResponse.async = False
    Set xmlElem = xmlPreProcReadResponse.createElement("RESPONSE")
    Set xmlPreProcReadResponseNode = xmlPreProcReadResponse.appendChild(xmlElem)
    
    Set xmlPreProcReadRequestNode = xmlPreProcReadRequest.selectSingleNode("REQUEST")
    
    adoCRUD xmlPreProcReadRequestNode, xmlPreProcReadRequestNode, xmlPreProcReadResponseNode

' ik_debug
    Debug.Print "Pre-Process Response"
    Debug.Print xmlPreProcReadResponseNode.xml
' ik_debug_ends

    strNodeName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
    If Len(strNodeName) = 0 Then
        strNodeName = xmlGetAttributeText(vxmlSchemaNode, "ENTITY_REF")
    End If

    Set xmlUpdateDataNodeList = vxmlRequestNode.selectNodes(strNodeName)
    Set xmlExistingDataNodeList = xmlPreProcReadResponseNode.selectNodes(strNodeName)
    
    adoCRUDPreProcCompareNodeLists _
        xmlUpdateDataNodeList, xmlExistingDataNodeList, vxmlSchemaNode

' ik_debug
    Debug.Print "Pre-Process updated request"
    Debug.Print vxmlRequestNode.xml
' ik_debug_ends

End Sub

Private Sub adoCRUDPreProcGetReadRequest( _
    ByVal vxmlReadRequestDoc As DOMDocument40, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    
    Dim xmlSchemaStartNode As IXMLDOMNode
    Dim xmlSchemaRefNode As IXMLDOMNode
    Dim xmlSchemaKeyNode As IXMLDOMNode
    Dim xmlRequestChild As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlElem = vxmlReadRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", xmlGetAttributeText(vxmlSchemaNode, "NAME")
    Set xmlNode = vxmlReadRequestDoc.appendChild(xmlElem)
    
    For Each xmlRequestChild In vxmlRequestNode.childNodes
        
        ' ik_CORE190_250923
        If xmlRequestChild.nodeType = NODE_ELEMENT Then
        
            Set xmlElem = vxmlReadRequestDoc.createElement(xmlRequestChild.nodeName)
            xmlNode.appendChild xmlElem
            
            Set xmlSchemaStartNode = _
                adoCRUDGetSchemaStartNode(xmlRequestChild.nodeName, vxmlSchemaNode)
                
            If xmlAttributeValueExists(xmlSchemaStartNode, "ENTITY_REF") Then
                Set xmlSchemaRefNode = adoCRUDGetSchemaRefNode(xmlSchemaStartNode)
            Else
                Set xmlSchemaRefNode = xmlSchemaStartNode
            End If
            
            ' schema node has no attributes - logical grouping
            If xmlSchemaRefNode.selectNodes("ATTRIBUTE").length = 0 Then
            
                For Each xmlSchemaStartNode In xmlSchemaStartNode.childNodes
                    
                    ' ik_CORE190_250923
                    If xmlSchemaStartNode.nodeType = NODE_ELEMENT Then
                    
                        If xmlAttributeValueExists(xmlSchemaStartNode, "ENTITY_REF") Then
                            Set xmlSchemaRefNode = adoCRUDGetSchemaRefNode(xmlSchemaStartNode)
                        Else
                            Set xmlSchemaRefNode = xmlSchemaStartNode
                        End If
                        
                        Set xmlRequestChild = xmlRequestChild.selectSingleNode(xmlGetAttributeText(xmlSchemaRefNode, "NAME"))
                        
                        For Each xmlSchemaKeyNode In xmlSchemaRefNode.selectNodes("ATTRIBUTE[@KEYTYPE and not(@KEYTYPE='NULL')]")
                        
                            If xmlAttributeValueExists(xmlRequestChild, xmlGetAttributeText(xmlSchemaKeyNode, "NAME")) Then
                            
                                xmlCopyAttribute _
                                    xmlRequestChild, _
                                    xmlElem, _
                                    xmlGetAttributeText(xmlSchemaKeyNode, "NAME")
                            
                            End If
                        
                        Next
                    
                    ' ik_CORE190_250923
                    End If
            
                Next
                
            Else
                    
                For Each xmlSchemaKeyNode In xmlSchemaRefNode.selectNodes("ATTRIBUTE[@KEYTYPE and not(@KEYTYPE='NULL')]")
                
                    If xmlAttributeValueExists(xmlRequestChild, xmlGetAttributeText(xmlSchemaKeyNode, "NAME")) Then
                    
                        xmlCopyAttribute _
                            xmlRequestChild, _
                            xmlElem, _
                            xmlGetAttributeText(xmlSchemaKeyNode, "NAME")
                    
                    End If
                
                Next
            
            End If
        
        ' ik_CORE190_250923
        End If
    
    Next

End Sub

Private Sub adoCRUDPreProcCompareNodeLists( _
    ByVal vxmlUpdateDataNodeList As IXMLDOMNodeList, _
    ByVal vxmlExistingDataNodeList As IXMLDOMNodeList, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode)
    
    Dim xmlSchemaRefNode As IXMLDOMNode, _
        xmlUpdateDataNode As IXMLDOMNode, _
        xmlExistingDataNode As IXMLDOMNode, _
        xmlSchemaChildNode As IXMLDOMNode
    
    Dim xmlUpdateDataChildNodeList As IXMLDOMNodeList
    Dim xmlExistingDataChildNodeList As IXMLDOMNodeList
    Dim xmlSchemaMasterChildNodeList As IXMLDOMNodeList
    
    Dim strNodeName As String
        
    Dim intIndex As Integer
    
    Set xmlSchemaRefNode = vxmlSchemaMasterNode
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "ENTITY_REF") Then
        Set xmlSchemaRefNode = adoCRUDGetSchemaRefNode(vxmlSchemaMasterNode)
    End If
        
    For intIndex = 0 To (vxmlUpdateDataNodeList.length - 1)
        
        Set xmlUpdateDataNode = vxmlUpdateDataNodeList.Item(intIndex)
        
        If Not xmlAttributeValueExists(xmlUpdateDataNode, "CRUD_OP") Then
            
            Set xmlExistingDataNode = vxmlExistingDataNodeList.Item(intIndex)
            
            If xmlExistingDataNode Is Nothing Then
            
                xmlSetAttributeValue xmlUpdateDataNode, "CRUD_OP", "CREATE"
            
            Else
            
                If xmlSchemaRefNode.selectNodes("ATTRIBUTE").length > 0 Then
                    adoCRUDPreProcCompareNodes _
                        xmlUpdateDataNode, xmlExistingDataNode, xmlSchemaRefNode
                End If
                
                Set xmlSchemaMasterChildNodeList = _
                    vxmlSchemaMasterNode.selectNodes("ENTITY")
                    
                For Each xmlSchemaChildNode In xmlSchemaMasterChildNodeList
    
                    strNodeName = xmlGetAttributeText(xmlSchemaChildNode, "NAME")
                    If Len(strNodeName) = 0 Then
                        strNodeName = xmlGetAttributeText(xmlSchemaChildNode, "ENTITY_REF")
                    End If
                    
                    Set xmlUpdateDataChildNodeList = _
                        xmlUpdateDataNode.selectNodes(strNodeName)
                        
                    If xmlUpdateDataChildNodeList.length > 0 Then
                    
                        Set xmlExistingDataChildNodeList = _
                            xmlExistingDataNode.selectNodes(strNodeName)
    
                        adoCRUDPreProcCompareNodeLists _
                            xmlUpdateDataChildNodeList, _
                            xmlExistingDataChildNodeList, _
                            xmlSchemaChildNode
                    
                    End If
                
                Next
            
            End If
            
        End If
        
    Next

End Sub


Private Sub adoCRUDPreProcCompareNodes( _
    ByVal vxmlUpdateDataNode As IXMLDOMNode, _
    ByVal vxmlExistingDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlNode As IXMLDOMNode
    Dim strAttribName As String
    
    Dim blnUpdate As Boolean
        
    For Each xmlNode In vxmlSchemaRefNode.selectNodes("ATTRIBUTE[@KEYTYPE and not(@KEYTYPE='NULL')]")
        strAttribName = xmlGetAttributeText(xmlNode, "NAME")
        If xmlAttributeValueExists(vxmlExistingDataNode, strAttribName) Then
            If Not xmlAttributeValueExists(vxmlUpdateDataNode, strAttribName) Then
                xmlCopyAttribute _
                    vxmlExistingDataNode, vxmlUpdateDataNode, strAttribName
            End If
        End If
    Next
    
    For Each xmlAttrib In vxmlUpdateDataNode.Attributes
        If xmlAttrib.Text <> xmlGetAttributeText(vxmlExistingDataNode, xmlAttrib.nodeName) Then
            blnUpdate = True
            Exit For
        End If
    Next
    
    If blnUpdate Then
        xmlSetAttributeValue vxmlUpdateDataNode, "CRUD_OP", "UPDATE"
    End If
    
End Sub


' ik_20040427_wip_ends

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
                vxmlResponseNode, _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode, False
        
        Case "READ"
            
            adoCRUDRead _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode, _
                vxmlResponseNode
        
        Case "UPDATE"
            adoCRUDUpdate _
                vxmlResponseNode, _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode, False
    
        ' ik_20040301
        ' new CRUD_OP i_UPDATE is "intelligent" update
        Case "I_UPDATE", "IUPDATE"
            adoCRUDUpdate _
                vxmlResponseNode, _
                vxmlRequestDataNode, _
                vxmlSchemaNode, _
                xmlSchemaRefNode, _
                True
        
        Case "DELETE"
            adoCRUDDelete _
                vxmlResponseNode, _
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

    If xmlGetAttributeText(vxmlSchemaNode, "NAME") = vstrRequestDataNodeName Then
        Set adoCRUDGetSchemaStartNode = vxmlSchemaNode
    Else
        If xmlGetAttributeText(vxmlSchemaNode, "DATA_REF") = vstrRequestDataNodeName Then
            Set adoCRUDGetSchemaStartNode = vxmlSchemaNode
        Else
            If xmlGetAttributeText(vxmlSchemaNode, "ENTITY_REF") = vstrRequestDataNodeName Then
                Set adoCRUDGetSchemaStartNode = vxmlSchemaNode
            Else
                If adoCRUDGetSchemaStartNode Is Nothing Then
                    '   ik_20060213_CORE244
                    strPattern = _
                        ".//ENTITY[@NAME='" & vstrRequestDataNodeName & "' or " & _
                        "@DATA_REF='" & vstrRequestDataNodeName & "' or " & _
                        "@ENTITY_REF='" & vstrRequestDataNodeName & "']"
                    Set adoCRUDGetSchemaStartNode = _
                        vxmlSchemaNode.selectSingleNode(strPattern)
                End If
            End If
        End If
    End If
        
    If adoCRUDGetSchemaStartNode Is Nothing Then
        Err.Raise _
            oeXMLMissingElement, _
            "adoCRUDGetSchemaStartNode", _
            "missing schema node: (" & vstrRequestDataNodeName & ")" & _
            " in schema node: " & _
            vxmlSchemaNode.cloneNode(False).xml
    End If
End Function
Private Function adoCRUDGetSchemaChildNode( _
    ByVal vxmRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmResponseNode As IXMLDOMNode) _
    As IXMLDOMNode
    
    Dim xmlThisSchemaNode As IXMLDOMNode
    Dim strPattern As String, _
        strRequestDataNodeName
        
    strRequestDataNodeName = vxmRequestNode.nodeName
    
    '   ik_20060213_CORE244
    strPattern = _
        "ENTITY[@NAME='" & strRequestDataNodeName & "' or " & _
        "@DATA_REF='" & strRequestDataNodeName & "' or " & _
        "@ENTITY_REF='" & strRequestDataNodeName & "']"
    
    Set xmlThisSchemaNode = _
        vxmlSchemaNode.selectSingleNode(strPattern)
    
    If xmlThisSchemaNode Is Nothing Then
    
        ' ik_20050627_CORE158/1
        If vxmlSchemaNode.parentNode.nodeName = "OM_SCHEMA" _
        And vxmlSchemaNode.selectNodes("ENTITY").length = 0 _
        Then
        ' ik_20050627_CORE158_ends
            
            strPattern = _
                "OM_SCHEMA/ENTITY[" & _
                "@NAME='" & strRequestDataNodeName & "' or " & _
                "@DATA_REF='" & strRequestDataNodeName & "']"
        
            Set xmlThisSchemaNode = _
                vxmlSchemaNode.ownerDocument.selectSingleNode(strPattern)
        
            If xmlThisSchemaNode Is Nothing Then
            
                If vxmlSchemaNode.ownerDocument.selectSingleNode("OM_SCHEMA[@BASESCHEMA]") Is Nothing Then
            
                    Set xmlThisSchemaNode = _
                        gxmldocBaseSchema.documentElement.selectSingleNode(strPattern)
                        
                End If
        
            End If
        
        ' ik_20050627_CORE158/1
        End If
        ' ik_20050627_CORE158_ends
        
    End If
    
    If xmlThisSchemaNode Is Nothing Then
    
        ' ik_20050627_CORE158/1
        adoCRUDSchemaError _
            "adoCRUDGetSchemaChildNode", _
            "error locating schema node", _
            vxmRequestNode, vxmlSchemaNode, Nothing, vxmResponseNode
        ' ik_20050627_CORE158_ends
    
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
    
    Dim xmlMasterNode As IXMLDOMNode, _
        xmlRefChildNode As IXMLDOMNode
    
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim strPattern As String
    strPattern = _
               "ENTITY[@NAME='" & _
               xmlGetAttributeText(vxmlSchemaNode, "ENTITY_REF") & _
               "']"
               
    Set xmlMasterNode = vxmlSchemaNode.ownerDocument.documentElement.selectSingleNode(strPattern)
    
    If Not xmlMasterNode Is Nothing Then
        Set adoCRUDGetSchemaRefNode = xmlMasterNode.cloneNode(True)
    End If
    
    If adoCRUDGetSchemaRefNode Is Nothing Then
        If Not xmlGetAttributeAsBoolean(vxmlSchemaNode.ownerDocument.documentElement, "BASESCHEMA") Then
            If Not gxmldocBaseSchema.documentElement.selectSingleNode(strPattern) Is Nothing Then
                Set adoCRUDGetSchemaRefNode = gxmldocBaseSchema.documentElement.selectSingleNode(strPattern).cloneNode(True)
            End If
        End If
    End If
    
    If adoCRUDGetSchemaRefNode Is Nothing Then
        Err.Raise _
            oeXMLMissingElement, _
            "adoCRUDGetSchemaRefNode", _
            "invalid schema ENTITY_REF on node: " & _
                vxmlSchemaNode.cloneNode(False).xml
    End If
        
    ' does this entity ref. contain a list of attributes refs?
    If adoCRUDGetSchemaRefNode.selectNodes("ATTRIBUTE[@ENTITY_REF]").length <> 0 _
    Then
        adoCRUDPrepareAttributeEntityRefs adoCRUDGetSchemaRefNode
    End If
    
'   copy any attribute attributes from schema master onto schema ref clone
    For Each xmlMasterNode In vxmlSchemaNode.selectNodes("ATTRIBUTE")
        
        strPattern = _
            "ATTRIBUTE[@NAME='" & _
            xmlGetAttributeText(xmlMasterNode, "NAME") & _
            "']"
        
        Set xmlRefChildNode = adoCRUDGetSchemaRefNode.selectSingleNode(strPattern)
        
        If xmlRefChildNode Is Nothing Then
            adoCRUDGetSchemaRefNode.appendChild xmlMasterNode.cloneNode(False)
        Else
            For Each xmlAttrib In xmlMasterNode.Attributes
                xmlRefChildNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
        End If
        
    Next

'   ik_debug
'    Debug.Print "adoCRUDGetSchemaRefNode: "
'    Debug.Print adoCRUDGetSchemaRefNode.xml
'   ik_debug_ends
    
End Function
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDPrepareAttributeEntityRefs( _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    Dim xmlAttribList As IXMLDOMNodeList
    Dim xmlSchemaSrceNode As IXMLDOMNode, _
        xmlAttribNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim strPattern As String
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & "adoCRUDPrepareAttributeEntityRefs")
    #End If
        
    Set xmlAttribList = vxmlSchemaNode.selectNodes("ATTRIBUTE[@ENTITY_REF]")
    For Each xmlAttribNode In xmlAttribList
        strPattern = _
            "OM_SCHEMA/ENTITY[@NAME='" & _
            xmlGetAttributeText(xmlAttribNode, "ENTITY_REF") & _
            "']/ATTRIBUTE[@NAME='" & _
            xmlGetAttributeText(xmlAttribNode, "NAME") & _
            "']"
        
        Set xmlSchemaSrceNode = gxmldocBaseSchema.selectSingleNode(strPattern)
        
        If xmlSchemaSrceNode Is Nothing Then
            Err.Raise _
                oeXMLMissingElement, _
                "adoCRUDPrepareAttributeEntityRefs", _
                "invalid schema ENTITY_REF on node: " & _
                    xmlAttribNode.cloneNode(False).xml
        End If
        
        For Each xmlAttrib In xmlSchemaSrceNode.Attributes
            If xmlAttribNode.Attributes.getNamedItem(xmlAttrib.baseName) Is Nothing Then
                xmlAttribNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            End If
        Next
        xmlAttribNode.Attributes.removeNamedItem ("ENTITY_REF")
    Next
    Set xmlSchemaSrceNode = Nothing
    Set xmlAttribNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlAttribList = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreate( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Const strFunctionName As String = "adoCRUDCreate"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    
    Dim xmlChild As IXMLDOMNode
    
    Dim strCrudOp As String
   
    ' fix to prevent duplicate create of 'parent' records,
    ' e.g. CUSTOMER when new CUSTOMERVERSION created
    
    strCrudOp = UCase(xmlGetAttributeText(vxmlRequestNode, "CRUD_OP"))
    
    If strCrudOp = "NULL" Then
        
        adoCRUDCreateRequestOffspring _
            vxmlResponseNode, _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vblnIsIntelligentUpdate
    
    ElseIf strCrudOp = "UPDATE" Then
        
        If Not vblnIsIntelligentUpdate Then
            For Each xmlChild In vxmlRequestNode.childNodes
                ' ik_CORE190_250923
                If xmlChild.nodeType = NODE_ELEMENT Then
                    adoCRUDCreateResetCreateOp xmlChild
                ' ik_CORE190_250923
                End If
            Next
        End If
        
        adoCRUDUpdate _
            vxmlResponseNode, _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, True
    
    ElseIf strCrudOp = "IUPDATE" Or strCrudOp = "I_UPDATE" Then
        
        adoCRUDUpdate _
            vxmlResponseNode, _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, True
    
    Else
        
        adoCRUDCreateOp _
            vxmlResponseNode, _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vblnIsIntelligentUpdate
    End If

End Sub

Private Sub adoCRUDCreateOp( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Const strFunctionName As String = "adoCRUDCreateOp"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
    
    Dim strAttribName As String, _
        strAttribType As String

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

            If Not xmlSchemaFieldNode.Attributes.getNamedItem("PUTDEFAULT") Is Nothing Then
                
                Set xmlAttrib = _
                    vxmlRequestNode.ownerDocument.createAttribute(strAttribName)
                xmlAttrib.Text = _
                    xmlSchemaFieldNode.Attributes.getNamedItem("PUTDEFAULT").Text
                
                If Left(xmlAttrib.Text, 6) = "xpath:" Then
                    adoCRUDGetDefaultValueViaXPath _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        xmlSchemaFieldNode, _
                        vblnIsIntelligentUpdate, _
                        True
                Else
                    
                    strAttribType = xmlSchemaFieldNode.Attributes.getNamedItem("DATATYPE").Text
                    
                    If (strAttribType = "DATE" Or strAttribType = "DATETIME") And xmlAttrib.Text = "_NOW" Then
                        xmlAttrib.Text = FormatNow
                    End If
                        
                    vxmlRequestNode.Attributes.setNamedItem xmlAttrib
                    
                End If
            
            End If
            
            If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
                If (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "FOREIGN") Or _
                    (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "PRIMARY") _
                Then
                    If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYSOURCE") Is Nothing Then
                        If xmlSchemaFieldNode.Attributes.getNamedItem("KEYSOURCE").Text = "GENERATED" Then
                            adoCRUDGetGeneratedKey vxmlRequestNode, xmlSchemaFieldNode
                        End If
                    End If
                End If
            End If
            
            'IK 16/11/2005 CORE215
            'search for key values in parent hierarchy
            If vxmlRequestNode.Attributes.getNamedItem(strAttribName) Is Nothing Then
                If Not xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
                    If (xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "FOREIGN") Then
                        Set xmlAttrib = getKeyValueFromParentHierarchy(vxmlRequestNode, strAttribName)
                        If Not xmlAttrib Is Nothing Then
                            vxmlRequestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
                        End If
                    End If
                End If
            End If
            
        Else
            Dim strXPathToKey As String, _
                strValueGainedFromXPath As String
            
            '   revised xpath usage notation (attributeValue)="xpath:...."
'IK 27/10/2005 CORE211 xpath: reference now processed in adoCRUDCreateSQLParam
'            If UCase(Left(xmlGetAttributeText(vxmlRequestNode, strAttribName), 6)) = "XPATH:" Then
'                strXPathToKey = Mid(xmlGetAttributeText(vxmlRequestNode, strAttribName), 7)
'                If Not vxmlRequestNode.ownerDocument.selectSingleNode(strXPathToKey) Is Nothing Then
'                    strValueGainedFromXPath = vxmlRequestNode.ownerDocument.selectSingleNode(strXPathToKey).Text
'                    xmlSetAttributeValue vxmlRequestNode, strAttribName, strValueGainedFromXPath
'                End If
'            End If
        
        End If
    Next
        
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "CREATE_PROC") Then
        adoCRUDCreateViaSP _
            vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
    ElseIf xmlAttributeValueExists(xmlSchemaEntityNode, "CREATE_REF") Then
        adoCRUDCreateViaSQL _
            vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
    Else
        If vxmlSchemaRefNode Is Nothing Then
            If vxmlSchemaMasterNode.selectNodes("ENTITY").length <> 0 Then
                ' ik_20040227
                ' implies logical grouping
                adoCRUDCreateRequestOffspring _
                    vxmlResponseNode, _
                    vxmlRequestNode, _
                    vxmlSchemaMasterNode, _
                    vxmlSchemaRefNode, _
                    vblnIsIntelligentUpdate
            Else
                
                Err.Raise _
                    oeXMLMissingAttribute, _
                    strFunctionName, _
                    "schema missing CREATE operation: " & _
                        vxmlSchemaMasterNode.cloneNode(False).xml
            
            End If
        Else
            If xmlAttributeValueExists(vxmlSchemaRefNode, "CREATE_PROC") Then
                adoCRUDCreateViaSP _
                    vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "CREATE_REF") Then
                adoCRUDCreateViaSQL _
                    vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
            Else
                If vxmlSchemaRefNode.selectNodes("ENTITY").length <> 0 Then
                    ' ik_20040227
                    ' implies logical grouping
                    adoCRUDCreateRequestOffspring _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        vblnIsIntelligentUpdate
                Else
                    Err.Raise _
                        oeXMLMissingAttribute, _
                        strFunctionName, _
                        "schema missing CREATE operation: " & _
                            vxmlSchemaMasterNode.cloneNode(False).xml
                End If
            End If
        End If
    End If
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing

End Sub

Private Sub adoCRUDCreateResetCreateOp(ByVal vxmlRequestNode As IXMLDOMNode)
    Dim xmlChild As IXMLDOMNode
    If xmlAttributeValueExists(vxmlRequestNode, "CRUD_OP") Then
        For Each xmlChild In vxmlRequestNode.childNodes
            ' ik_CORE190_250923
            If xmlChild.nodeType = NODE_ELEMENT Then
                adoCRUDCreateResetCreateOp xmlChild
            ' ik_CORE190_250923
            End If
        Next
    Else
        xmlSetAttributeValue vxmlRequestNode, "CRUD_OP", "CREATE"
    End If
End Sub


'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDGetGeneratedKey( _
    ByVal vxmlDataNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    Dim xmlAttrib As IXMLDOMAttribute
    Dim strDataType As String, _
               strGuid As String, _
               strAttribName As String
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & "adoCRUDGetGeneratedKey")
    #End If
        
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
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
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
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & "adoCRUDGetNextKeySequence")
    #End If
    ' parent node must be ENTITY with CREATE_REF
    strDataSource = xmlGetMandatoryAttributeText(vxmlSchemaNode.parentNode, "CREATE_REF")
    strSequenceFieldName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
       
    Set cmd = New ADODB.Command
    
    For Each xmlNode In vxmlSchemaNode.parentNode.selectNodes("ATTRIBUTE[@KEYTYPE='FOREIGN'][not(@KEYSOURCE) or @KEYSOURCE='NULL']")
                    
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
                       adoCRUDCreateSQLParam( _
                           xmlNode, _
                           vxmlDataNode.Attributes.getNamedItem(strKeyAttribName))
            cmd.Parameters.Append param
        End If
    
    Next
    
    strSQL = _
        "SELECT MAX(" & strSequenceFieldName & _
        ") FROM " & strDataSource
            
    ' SDS LIVE00009653 / 21/01/2004 _ STARTS
    If genumDbProvider = omiga4DBPROVIDERSQLServer Then
        strSQL = strSQL & " WITH (NOLOCK) "
    End If
    ' SDS LIVE00009653 / 21/01/2004_ ENDS
    strSQL = strSQL & " WHERE (" & strSQLWhere & ")"
    'MO     12/09/2002  BMIDS00435 - End
    
'   ik_debug
    Debug.Print "adoAssist.adoCRUDGetNextKeySequence strSQL: " & strSQL
'   ik_debug_ends

    If Len(strSQLWhere) = 0 Then
        ' ik_20050923_CORE194/1
        Err.Raise _
            oeCommandFailed, _
            "adoCRUD", _
            "no primary keys spacified for generated key, command: " & strSQL
        ' ik_20050923_CORE194_ends
    End If
    
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = adoCRUDExecuteGetRecordSet(cmd)
    If Not rst Is Nothing Then
        rst.MoveFirst
        If IsNull(rst.Fields.Item(0)) = False Then
            intThisSequence = rst.Fields.Item(0).Value
        End If
        rst.Close
    End If
    adoCRUDGetNextKeySequence = intThisSequence + 1
    Err.Clear
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
adoCRUDGetNextKeySequenceExit:
    Set rst = Nothing
    Set cmd = Nothing
    Set xmlNode = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    If Err.Number <> 0 Then
        If Err.Source = App.EXEName Then
            Err.Source = "adoCRUDGetNextKeySequence"
        Else
            Err.Source = "adoCRUDGetNextKeySequence." & Err.Source
        End If
        
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreateViaSP( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    On Error GoTo adoCRUDCreateViaSPExit
    
    Const strFunctionName As String = "adoCRUDCreateViaSP"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
    
    Dim cmd As ADODB.Command
    
    Dim param As ADODB.Parameter
    
    Dim strParamName As String, _
        strProcRef As String
    
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[not(@CREATEPARAM) or not(@CREATEPARAM='NO')]")
            
        Set xmlDataAttrib = _
            vxmlRequestNode.Attributes.getNamedItem( _
               xmlGetAttributeText(xmlSchemaFieldNode, "NAME"))
        
        If xmlGetAttributeText(xmlSchemaFieldNode, "CREATEPARAM") = "OUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib, adParamOutput)
        ElseIf xmlGetAttributeText(xmlSchemaFieldNode, "CREATEPARAM") = "INOUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib, adParamInputOutput)
        Else
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib)
        End If
        
        cmd.Parameters.Append param
        
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "CREATE_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "CREATE_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "CREATE_PROC")
    End If

    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = strProcRef
    
    lngRecordsAffected = adoCRUDExecuteSP(cmd)
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@CREATEPARAM='INOUT' or @CREATEPARAM='OUT']")
            
        strParamName = xmlGetAttributeText(xmlSchemaFieldNode, "NAME")
    
        For Each param In cmd.Parameters
        
            If param.Name = strParamName Then
            
                adoCRUDParamToXML param, vxmlResponseNode, xmlSchemaFieldNode
                adoCRUDParamToXML param, vxmlRequestNode, xmlSchemaFieldNode
                
                Exit For
            
            End If
            
        Next
    
    Next
    
    If lngRecordsAffected = 0 Then
        
        If Not xmlGetAttributeAsBoolean(vxmlRequestNode, "_IGNORENOWROWS") Then
            Err.Raise oeNoRowsAffected, strFunctionName, "no records created"
        End If
        
    End If
    
    Set cmd = Nothing

'   ik_debug
   Debug.Print "adoCRUDCreateViaSP records created: " & lngRecordsAffected
'   ik_debug_ends
    
    adoCRUDCreateRequestOffspring _
        vxmlResponseNode, _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode, _
        vblnIsIntelligentUpdate

adoCRUDCreateViaSPExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreateViaSQL( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)

On Error GoTo adoCRUDCreateViaSQLExit
    
    Const strFunctionName As String = "adoCRUDCreateViaSQL"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
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
        strPattern = "ATTRIBUTE[@NAME='" & xmlDataAttrib.Name & "'][not(@DATAMAP) or @DATAMAP='IN' or @DATAMAP='INOUT']"
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
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib)
            cmd.Parameters.Append param
      End If
    Next
    strSQL = "INSERT INTO " & strDataSource & " (" & strColumns & ") VALUES (" & strValues & ")"
    
'   ik_debug
   Debug.Print "adoCreate strSQL: " & strSQL
'   ik_debug_ends
     
   Dim p As Parameter
   Dim str As String
   For Each p In cmd.Parameters
        str = str & p.Value & ", "
   Next
   If Len(str) > 0 Then Debug.Print Left(str, Len(str) - 2)
     
     
    If Len(strValues) > 0 Then
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
            
        lngRecordsAffected = adoCRUDExecuteSQLCommand(cmd)
        If lngRecordsAffected = 0 Then
            If Not xmlGetAttributeAsBoolean(vxmlRequestNode, "_IGNORENOWROWS") Then
                Err.Raise oeNoRowsAffected, strFunctionName, "no records created"
            End If
        End If
    Else
        Err.Raise oeNoDataForCreate, strFunctionName, "no data values for create"
    End If
    Set cmd = Nothing
    
'   ik_debug
   Debug.Print "adoCRUDCreateViaSQL records created: " & lngRecordsAffected
'   ik_debug_ends
    
    adoCRUDCreateRequestOffspring _
        vxmlResponseNode, _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode, _
        vblnIsIntelligentUpdate

adoCRUDCreateViaSQLExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlAttrib = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDCreateRequestOffspring( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisSchemaMasterNode As IXMLDOMNode
    Dim xmlThisSchemaRefNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    Dim strPattern As String
    Dim intNodeIndex As Long
    
    If vxmlRequestNode.hasChildNodes Then
        
        For Each xmlThisRequestNode In vxmlRequestNode.childNodes
            
            ' ik_CORE190_250923
            If xmlThisRequestNode.nodeType = NODE_ELEMENT Then
            
                If xmlGetAttributeText(xmlThisRequestNode, "CRUD_OP") <> "DONE" Then
                
                    Set xmlThisSchemaMasterNode = _
                        adoCRUDGetSchemaChildNode( _
                            xmlThisRequestNode, vxmlSchemaMasterNode, vxmlResponseNode)
                    
                    If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                        
                        Set xmlThisSchemaRefNode = _
                            adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                        adoCRUDPopulateChildKeys _
                            xmlThisRequestNode, _
                            xmlThisSchemaRefNode, _
                            vxmlRequestNode
                    Else
                        Set xmlThisSchemaRefNode = Nothing
                        adoCRUDPopulateChildKeys _
                            xmlThisRequestNode, _
                            xmlThisSchemaMasterNode, _
                            vxmlRequestNode
                    End If
                    
                    adoCRUDCreate _
                        vxmlResponseNode, _
                        xmlThisRequestNode, _
                        xmlThisSchemaMasterNode, _
                        xmlThisSchemaRefNode, _
                        vblnIsIntelligentUpdate
                
                End If
            
            ' ik_CORE190_250923
            End If
        
        Next
    
    End If
    
    Set xmlThisRequestNode = Nothing
    Set xmlThisSchemaMasterNode = Nothing
    Set xmlThisSchemaRefNode = Nothing
    Set xmlSchemaNodeList = Nothing

End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDRead( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    Dim xmlElem As IXMLDOMElement
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim strNodeName As String
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_PROC") Then
            
        ' ik_20050627_CORE158/2
        ' expect SP to take care of this
        vxmlRequestNode.Attributes.removeNamedItem "_nullDefault"
        vxmlRequestNode.Attributes.removeNamedItem "_nullForeignKeys"
        ' ik_20050627_CORE158_ends
    
        adoCRUDReadViaSP _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vxmlResponseNode
    
    ElseIf xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_REF") Then
            
        ' ik_20050627_CORE158/2
        If xmlAttributeValueExists(vxmlRequestNode, "_nullDefault") _
        Or xmlAttributeValueExists(vxmlRequestNode, "_nullForeignKeys") _
        Then
            vxmlRequestNode.Attributes.removeNamedItem "_nullDefault"
            vxmlRequestNode.Attributes.removeNamedItem "_nullForeignKeys"
        Else
        ' ik_20050627_CORE158_ends
            adoCRUDReadViaSQL _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                vxmlResponseNode
        ' ik_20050627_CORE158/2
        End If
        ' ik_20050627_CORE158_ends
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
            
            'IK_CORE301_29/09/2006 only add 'group node' when child nodes returned
' CORE308
'            If xmlThisResponseNode.hasChildNodes = False Then
'                vxmlResponseNode.removeChild xmlThisResponseNode
'            End If
            'IK_CORE301_29/09/2006 ends
        Else
            If xmlAttributeValueExists(vxmlSchemaRefNode, "READ_PROC") Then
                adoCRUDReadViaSP _
                    vxmlRequestNode, _
                    vxmlSchemaMasterNode, _
                    vxmlSchemaRefNode, _
                    vxmlResponseNode
            Else
                If xmlAttributeValueExists(vxmlSchemaRefNode, "READ_REF") Then
                    adoCRUDReadViaSQL _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        vxmlResponseNode
                Else
                    ' IK20040227
                    ' allow schema node with no READ_REF or READ_PROC
                    ' implies logical grouping
                    If Not vxmlSchemaRefNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
                        strNodeName = _
                            vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME").Text
                    Else
                        strNodeName = _
                            vxmlSchemaRefNode.Attributes.getNamedItem("NAME").Text
                    End If
                    Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
            
                    'IK_CORE301_29/09/2006 only add 'group node' when child nodes returned
                    adoCRUDReadGetOffspring _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        xmlElem
                    
                    If xmlElem.hasChildNodes Then
                        vxmlResponseNode.appendChild xmlElem
                    End If
                    
                    Set xmlElem = Nothing
                    'IK_CORE301_29/09/2006
                
                End If
            End If
        End If
    End If
    Set xmlElem = Nothing
    Set xmlThisResponseNode = Nothing
End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDReadViaSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    Const strFunctionName As String = "adoCRUDReadViaSP"
    
    On Error GoTo adoCRUDReadViaSPExit
    
    Dim strProcType As String, _
        strProcRef As String
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "READ_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "READ_PROC")
    End If
    
    If Left(UCase(strProcRef), 4) = "XML:" Then
        strProcType = "XML"
    End If
    
    If strProcType = "XML" Then
        adoCRUDReadViaAutoSP vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vxmlResponseNode
    Else
        adoCRUDReadViaNativeSP vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vxmlResponseNode
    End If

adoCRUDReadViaSPExit:
    errCheckError strFunctionName
End Sub

Private Sub adoCRUDReadViaNativeSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDReadViaNativeSPExit
    Const strFunctionName As String = "adoCRUDReadViaNativeSP"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaAttribNode As IXMLDOMNode
    
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    Dim rst As ADODB.Recordset
    
    Dim strAttribName As String, _
        strProcRef As String
        
    Dim xmlDataAttrib As IXMLDOMAttribute
    Dim blnResponsePrefix As Boolean
        
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
    
    For Each xmlSchemaAttribNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@READPARAM][not(@READPARAM='NO')]")
        
        strAttribName = xmlGetMandatoryAttributeText(xmlSchemaAttribNode, "NAME")
        
        blnResponsePrefix = False
        Set xmlDataAttrib = vxmlRequestNode.Attributes.getNamedItem(strAttribName)
        If Not xmlDataAttrib Is Nothing Then
            If Left(xmlDataAttrib.Text, 15) = "responseValue::" Then
                xmlDataAttrib.Text = Mid(xmlDataAttrib.Text, 16)
                blnResponsePrefix = True
            End If
        End If
        
        If xmlGetAttributeText(xmlSchemaAttribNode, "READPARAM") = "OUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib, adParamOutput)
        ElseIf xmlGetAttributeText(xmlSchemaAttribNode, "READPARAM") = "INOUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib, adParamInputOutput)
        Else
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib)
        End If
        
        cmd.Parameters.Append param
            
        If blnResponsePrefix Then
            xmlDataAttrib.Text = "responseValue::" & xmlDataAttrib.Text
        End If
            
    Next
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = strProcRef
    
'   ik_debug
'   Debug.Print strFunctionName & " strSQL: " & cmd.CommandText
'   ik_debug_ends

    '   COMPONENT_REF only implemented for CRUD_OP="READ"
    If xmlAttributeValueExists(xmlSchemaEntityNode, gstrComponentRefId) Then
        Set rst = _
            adoCRUDExecuteGetRecordSet( _
                cmd, _
                UCase(xmlGetAttributeText(xmlSchemaEntityNode, gstrComponentRefId)))
    Else
        Set rst = adoCRUDExecuteGetRecordSet(cmd)
    End If
    
    For Each xmlSchemaAttribNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@READPARAM='INOUT' or @READPARAM='OUT']")
            
        strAttribName = xmlGetAttributeText(xmlSchemaAttribNode, "NAME")
    
        For Each param In cmd.Parameters
        
            If param.Name = strAttribName Then
            
                adoCRUDParamToXML param, vxmlResponseNode, xmlSchemaAttribNode
                
                Exit For
            
            End If
            
        Next
    
    Next
    
    If Not rst Is Nothing Then
        
        adoCRUDConvertRecordSetToXML _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vxmlResponseNode, _
            rst
            
    'IK_CORE220_20051130
    Else
        ReinstateRequest vxmlRequestNode
    'IK_CORE220_20051130_ends
    
    End If
    
    Set rst = Nothing
    Set cmd = Nothing

adoCRUDReadViaNativeSPExit:
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaAttribNode = Nothing
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub

Private Sub adoCRUDReadViaAutoSP( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDReadViaAutoSPExit
    Const strFunctionName As String = "adoCRUDReadViaAutoSP"

    Dim xmlDoc As DOMDocument40
    Dim xmlSchemaKeyNodeList As IXMLDOMNodeList
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaAttribNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute
    
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim adoStr As ADODB.Stream
    
    Dim strAttribName As String, _
        strProcRef As String
        
    ' ik_CORE192_20050923
    Dim blnResponsePrefix As Boolean
        
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "READ_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "READ_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "READ_PROC")
    End If
    
    If Left(UCase(strProcRef), 4) = "XML:" Then
        strProcRef = Right(strProcRef, Len(strProcRef) - 4)
    End If

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = strProcRef

    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    For Each xmlSchemaAttribNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@READPARAM][not(@READPARAM='NO')]")
        
        strAttribName = xmlGetMandatoryAttributeText(xmlSchemaAttribNode, "NAME")
        Set xmlDataAttrib = vxmlRequestNode.Attributes.getNamedItem(strAttribName)
        
        ' ik_CORE192_20050923
        blnResponsePrefix = False
        If Not xmlDataAttrib Is Nothing Then
            If Left(xmlDataAttrib.Text, 15) = "responseValue::" Then
                xmlDataAttrib.Text = Mid(xmlDataAttrib.Text, 16)
                blnResponsePrefix = True
            End If
        End If
        ' ik_CORE192_20050923_ends
        
        If xmlGetAttributeText(xmlSchemaAttribNode, "READPARAM") = "OUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib, adParamOutput)
        ElseIf xmlGetAttributeText(xmlSchemaAttribNode, "READPARAM") = "INOUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib, adParamInputOutput)
        Else
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib)
        End If
        
        cmd.Parameters.Append param
            
        ' ik_CORE192_20050923
        If blnResponsePrefix Then
            xmlDataAttrib.Text = "responseValue::" & xmlDataAttrib.Text
        End If
        ' ik_CORE192_20050923_ends
    
    Next
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        getDbConnectString(UCase(xmlGetAttributeText(xmlSchemaEntityNode, gstrComponentRefId)))
    conn.Open
    cmd.ActiveConnection = conn
    
    Set adoStr = New ADODB.Stream
    adoStr.Open
    cmd.Properties("Output Stream") = adoStr
    cmd.Properties("XML Root") = "RESPONSE"
    
    cmd.Execute , , adExecuteStream
    
    If vxmlResponseNode.parentNode.nodeType <> NODE_DOCUMENT Then
    
        Set xmlElem = vxmlResponseNode
        
        Set xmlDoc = New DOMDocument40
        xmlDoc.async = False
        
        xmlDoc.loadXML adoStr.ReadText
        
        For Each xmlNode In xmlDoc.selectSingleNode("RESPONSE").childNodes
        
            ' ik_CORE190_250923
            If xmlNode.nodeType = NODE_ELEMENT Then
                vxmlResponseNode.appendChild xmlNode.cloneNode(True)
            ' ik_CORE190_250923
            End If
        
        Next
    
    Else
    
        vxmlResponseNode.ownerDocument.loadXML adoStr.ReadText
    
        Set xmlElem = vxmlResponseNode.ownerDocument.selectSingleNode("RESPONSE")

        ' ik_20050803_CORE179/1
        xmlElem.setAttribute "TYPE", "SUCCESS"
        ' ik_20050803_CORE179_ends
        
    End If
    
' ik_20050803_CORE179/1
'    xmlElem.setAttribute "TYPE", "SUCCESS"
' ik_20050803_CORE179_ends
    
    For Each xmlSchemaAttribNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@READPARAM='INOUT' or @READPARAM='OUT']")
            
        strAttribName = xmlGetAttributeText(xmlSchemaAttribNode, "NAME")
    
        For Each param In cmd.Parameters
        
            If param.Name = strAttribName Then
            
                adoCRUDParamToXML param, xmlElem, xmlSchemaAttribNode
                
                Exit For
            
            End If
            
        Next
    
    Next
    
    adoStr.Close
    conn.Close
    Set adoStr = Nothing
    Set conn = Nothing
    Set cmd = Nothing
    
adoCRUDReadViaAutoSPExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaAttribNode = Nothing
    Set xmlSchemaKeyNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlDataAttrib = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub
' ik_20050114_wip_ends

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDReadViaSQL( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDReadViaSQLExit
    Const strFunctionName As String = "adoCRUDReadViaSQL"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
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
        
    Dim blnHasOrderBy As Boolean
    Dim blnResponsePrefix As Boolean
    'IK_05/12/2005_CORE223
    Dim blnHasLockHint As Boolean
        
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
    
        ' IK CORE290 09/08/2006
        strPattern = _
                   "ATTRIBUTE[@NAME='" & xmlDataAttrib.Name & "'][@DATATYPE][not(@NOREADVALUE)]"
        
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
            
            blnResponsePrefix = False
            If Left(xmlDataAttrib.Text, 15) = "responseValue::" Then
                xmlDataAttrib.Text = Mid(xmlDataAttrib.Text, 16)
                blnResponsePrefix = True
            End If
            
            ' APS 03/09/01 - Explicit handling of null columns
            If Len(xmlDataAttrib.Text) = 0 Then
                strSQLWhere = strSQLWhere & strDbColumnName & " is NULL"
            Else
                strSQLWhere = strSQLWhere & strDbColumnName & "=?"
                Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib)
                cmd.Parameters.Append param
            End If
            
            If blnResponsePrefix Then
                xmlDataAttrib.Text = "responseValue::" & xmlDataAttrib.Text
            End If
        
        End If
    
    Next
        
    'IK_05/12/2005_CORE223
    If cmd.Parameters.Count <> 0 Then
        strSQL = "SELECT * FROM " & strDataSource
    Else
        ' no WHERE clause only available at top of hierarchy
        If vxmlSchemaMasterNode.parentNode.nodeName = "OM_SCHEMA" Then
            strSQL = "SELECT * FROM " & strDataSource
        End If
    End If
        
    If Len(strSQL) > 0 Then
        
        blnHasLockHint = False
        If xmlAttributeValueExists(vxmlRequestNode, "_LOCKHINT") Then
            strPattern = xmlGetAttributeText(vxmlRequestNode, "_LOCKHINT")
            blnHasLockHint = True
        ElseIf xmlAttributeValueExists(vxmlRequestNode.ownerDocument.documentElement, "_LOCKHINT") Then
            strPattern = xmlGetAttributeText(vxmlRequestNode.ownerDocument.documentElement, "_LOCKHINT")
            blnHasLockHint = True
        ElseIf xmlAttributeValueExists(vxmlSchemaMasterNode, "_LOCKHINT") Then
            strPattern = xmlGetAttributeText(vxmlSchemaMasterNode, "_LOCKHINT")
            blnHasLockHint = True
        End If
    
        If blnHasLockHint Then
            strSQL = strSQL & " WITH(" & strPattern & ")"
        End If
    
        If cmd.Parameters.Count <> 0 Then
            strSQL = strSQL & " WHERE (" & strSQLWhere & ")"
        End If
    
        blnHasOrderBy = False
        If xmlAttributeValueExists(vxmlRequestNode, "_ORDERBY") Then
            strPattern = xmlGetAttributeText(vxmlRequestNode, "_ORDERBY")
            blnHasOrderBy = True
        ElseIf xmlAttributeValueExists(vxmlSchemaMasterNode, "_ORDERBY") Then
            strPattern = xmlGetAttributeText(vxmlSchemaMasterNode, "_ORDERBY")
            blnHasOrderBy = True
        End If
        
        If blnHasOrderBy Then
            If Left(strPattern, 1) = "(" Then
                strPattern = Mid(strPattern, 2, Len(strPattern) - 2)
            End If
            strSQL = strSQL & " ORDER BY " & strPattern
        End If
        
    End If
    'IK_05/12/2005_CORE223_ends
    
'   ik_debug
   Debug.Print strFunctionName & " strSQL: " & strSQL
'   ik_debug_ends
    
    If Len(strSQL) > 0 Then
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
'       COMPONENT_REF only implemented for CRUD_OP="READ"
        If xmlAttributeValueExists(xmlSchemaEntityNode, gstrComponentRefId) Then
            Set rst = _
                adoCRUDExecuteGetRecordSet( _
                    cmd, _
                    UCase(xmlGetAttributeText( _
                        xmlSchemaEntityNode, gstrComponentRefId)))
        Else
            Set rst = adoCRUDExecuteGetRecordSet(cmd)
        End If
        If Not rst Is Nothing Then
            adoCRUDConvertRecordSetToXML _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                vxmlResponseNode, _
                rst
            'IK_CORE220_20051130
            Else
                ReinstateRequest vxmlRequestNode
            'IK_CORE220_20051130_ends
        End If
    End If
    
    Set rst = Nothing
    Set cmd = Nothing

adoCRUDReadViaSQLExit:
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaAttribNode = Nothing
    Set xmlDataAttrib = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDConvertRecordSetToXML( _
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
    
    Dim enumComboTypeLookUp As COMBOVALIDATIONLOOKUPTYPE
    
    Set xmlSchemaEntityNode = vxmlSchemaMasterNode
'   if not ENTITY_REF assume ehnancement to schema (see DEFAULTs)
    If vxmlSchemaMasterNode.selectNodes("ATTRIBUTE/@ENTITY_REF").length = 0 Then
        If Not vxmlSchemaRefNode Is Nothing Then
            Set xmlSchemaEntityNode = vxmlSchemaRefNode
        End If
    End If
    strNodeName = "UNDEFINED"
        
    If Not vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
        strNodeName = _
            vxmlSchemaMasterNode.Attributes.getNamedItem("OUTNAME").Text
    Else
        If Not vxmlSchemaMasterNode.Attributes.getNamedItem("NAME") Is Nothing Then
            strNodeName = _
                vxmlSchemaMasterNode.Attributes.getNamedItem("NAME").Text
        Else
            If Not vxmlSchemaRefNode Is Nothing Then
                If Not vxmlSchemaRefNode.Attributes.getNamedItem("OUTNAME") Is Nothing Then
                    strNodeName = _
                        vxmlSchemaRefNode.Attributes.getNamedItem("OUTNAME").Text
                Else
                    strNodeName = _
                        vxmlSchemaRefNode.Attributes.getNamedItem("NAME").Text
                End If
            End If
        End If
    End If
    
    blnDoComboLookUp = IsComboLookUpRequired(vxmlSchemaMasterNode, vxmlRequestNode)
    enumComboTypeLookUp = IsComboTypeLookUpRequired(vxmlSchemaMasterNode, vxmlRequestNode)
    Do While Not vrst.EOF
        If vxmlResponseNode.nodeName = strNodeName Then
            Set xmlElem = vxmlResponseNode
        Else
            Set xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName)
        End If
        
        Set xmlSchemaAttributeList = xmlSchemaEntityNode.selectNodes("ATTRIBUTE[not(@DATAMAP) or @DATAMAP='INOUT' or @DATAMAP='OUT']")
        For Each xmlSchemaAttribNode In xmlSchemaAttributeList

            ' IK 05/11/2001 DB_REF support
            If xmlAttributeValueExists(xmlSchemaAttribNode, gstrDbRefId) Then
                strAttribName = _
                    xmlGetAttributeText(xmlSchemaAttribNode, gstrDbRefId)
            Else
                strAttribName = xmlGetAttributeText(xmlSchemaAttribNode, "NAME")
            End If
            
            Set fld = vrst.Fields.Item(strAttribName)
            If Not fld Is Nothing Then
                If Not IsNull(fld) Then
                    If Not IsNull(fld.Value) Then
                        adoCRUDFieldToXml fld, xmlElem, xmlSchemaAttribNode, blnDoComboLookUp, enumComboTypeLookUp
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
Private Sub adoCRUDReadGetOffspring( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
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
    
    'IK_CORE220_20051130
    ReinstateRequest vxmlRequestNode
    
    If vxmlRequestNode.hasChildNodes Then
        
        For Each xmlThisRequestNode In vxmlRequestNode.childNodes
        
            ' ik_CORE190_250923
            If xmlThisRequestNode.nodeType = NODE_ELEMENT Then
                
                Set xmlThisSchemaMasterNode = _
                    adoCRUDGetSchemaChildNode( _
                        xmlThisRequestNode, vxmlSchemaMasterNode, vxmlResponseNode)
                If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                    
                    Set xmlThisSchemaRefNode = _
                        adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                    adoCRUDReadPopulateChildKeys _
                        xmlThisRequestNode, _
                        xmlThisSchemaRefNode, _
                        vxmlResponseNode
                Else
                    Set xmlThisSchemaRefNode = Nothing
                    adoCRUDReadPopulateChildKeys _
                        xmlThisRequestNode, _
                        xmlThisSchemaMasterNode, _
                        vxmlResponseNode
                End If
                adoCRUDRead _
                    xmlThisRequestNode, _
                    xmlThisSchemaMasterNode, _
                    xmlThisSchemaRefNode, _
                    vxmlResponseNode
            
            ' ik_CORE190_250923
            End If
        
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
                    adoCRUDReadPopulateChildKeys _
                        vxmlRequestNode, _
                        xmlThisSchemaRefNode, _
                        vxmlResponseNode
                Else
                    Set xmlThisSchemaRefNode = Nothing
                    adoCRUDReadPopulateChildKeys _
                        vxmlRequestNode, _
                        xmlThisSchemaMasterNode, _
                        vxmlResponseNode
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
                            adoCRUDReadPopulateChildKeys _
                                vxmlRequestNode, _
                                xmlThisSchemaRefNode, _
                                vxmlResponseNode
                        Else
                            Set xmlThisSchemaRefNode = Nothing
                            adoCRUDReadPopulateChildKeys _
                                vxmlRequestNode, _
                                xmlThisSchemaMasterNode, _
                                vxmlResponseNode
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

'IK_CORE220_20051130
Private Sub ReinstateRequest(ByVal vxmlRequestNode As IXMLDOMElement)
    
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' remove any values derived from previous response nodes
    For Each xmlAttrib In vxmlRequestNode.Attributes
        If Left(xmlAttrib.Text, 15) = "responseValue::" Then
            vxmlRequestNode.Attributes.removeNamedItem xmlAttrib.nodeName
        End If
    Next
    
    ' restore previous request values
    For Each xmlAttrib In vxmlRequestNode.Attributes
        If Left(xmlAttrib.Name, 9) = "_request_" Then
            vxmlRequestNode.setAttribute Mid(xmlAttrib.Name, 10), xmlAttrib.Text
            vxmlRequestNode.Attributes.removeNamedItem xmlAttrib.nodeName
        End If
    Next

End Sub


'APS - Implemented IK's CRUD changes 20/08/01
'   ik_20040301 i_UPDATE support
Private Sub adoCRUDUpdate( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Const strFunctionName As String = "adoCRUDUpdate"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    
    Dim xmlAttrib As IXMLDOMAttribute

    adoCRUDGetUpdateDefaults _
        vxmlResponseNode, _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode, _
        vblnIsIntelligentUpdate

    '   ik_20040301 i_UPDATE support
    ' test CRUD_OP on REQUEST node
    
    If Not vblnIsIntelligentUpdate Then
        
        adoCRUDUpdateOp _
            vxmlResponseNode, _
            vxmlRequestNode, _
            vxmlSchemaMasterNode, _
            vxmlSchemaRefNode, _
            vblnIsIntelligentUpdate
        
    Else
        
        Set xmlAttrib = vxmlRequestNode.Attributes.getNamedItem("CRUD_OP")
        
        If xmlAttrib Is Nothing Then
            
            adoCRUDUpdateRequestOffspring _
                vxmlResponseNode, _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                vblnIsIntelligentUpdate
        
        ElseIf xmlAttrib.Value = "NULL" Then
            
            adoCRUDUpdateRequestOffspring _
                vxmlResponseNode, _
                vxmlRequestNode, _
                vxmlSchemaMasterNode, _
                vxmlSchemaRefNode, _
                vblnIsIntelligentUpdate
        
        Else
            
            Select Case UCase(xmlAttrib.Text)
                
                Case "UPDATE"
                    adoCRUDUpdateOp _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        vblnIsIntelligentUpdate
                        
                Case "CREATE"
                    adoCRUDCreate _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        True
                
                Case "DELETE"
                    adoCRUDDelete _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode
                
                Case Else
                    adoCRUDUpdateRequestOffspring _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        vblnIsIntelligentUpdate
            
            End Select
        
        End If
        
    End If

End Sub

Private Sub adoCRUDUpdateOp( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Const strFunctionName As String = "adoCRUDUpdateNode"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    
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
            vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
    
    ElseIf xmlAttributeValueExists(xmlSchemaEntityNode, "UPDATE_REF") Then
    
        adoCRUDUpdateViaSQL _
            vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
            
    Else
    
        If vxmlSchemaRefNode Is Nothing Then
        
            If vxmlSchemaMasterNode.selectNodes("ENTITY").length <> 0 Then
            
                adoCRUDUpdateRequestOffspring _
                    vxmlResponseNode, _
                    vxmlRequestNode, _
                    vxmlSchemaMasterNode, _
                    vxmlSchemaRefNode, _
                    vblnIsIntelligentUpdate
            
            Else
            
                Err.Raise _
                    oeXMLMissingAttribute, _
                    strFunctionName, _
                    "schema missing UPDATE operation: " & _
                    vxmlSchemaMasterNode.cloneNode(False).xml
            
            End If
        
        Else
    
            If xmlAttributeValueExists(vxmlSchemaRefNode, "UPDATE_PROC") Then
            
                adoCRUDUpdateViaSP _
                    vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
            
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "UPDATE_REF") Then
            
                adoCRUDUpdateViaSQL _
                    vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode, vblnIsIntelligentUpdate
                    
            Else
            
                If vxmlSchemaRefNode.selectNodes("ENTITY").length <> 0 Then
                
                    adoCRUDUpdateRequestOffspring _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        vblnIsIntelligentUpdate
                
                Else
                    
                    Err.Raise _
                        oeXMLMissingAttribute, _
                        strFunctionName, _
                        "schema missing UPDATE operation: " & _
                        vxmlSchemaMasterNode.cloneNode(False).xml
                        
                End If
                    
            End If
            
        End If
    
    End If
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDUpdateViaSP( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    On Error GoTo adoCRUDUpdateViaSPExit
    Const strFunctionName As String = "adoCRUDUpdateViaSP"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    
    Dim xmlDataAttrib As IXMLDOMAttribute
    
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strParamName As String, _
        strProcRef As String
    
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[not(@UPDATEPARAM) or not(@UPDATEPARAM='NO')]")
    
        Set xmlDataAttrib = _
            vxmlRequestNode.Attributes.getNamedItem( _
               xmlGetAttributeText(xmlSchemaFieldNode, "NAME"))
                   
        'IK 11/11/2005 CORE215
        'search for parameter values in parent hierarchy
        If xmlDataAttrib Is Nothing Then
            If xmlGetAttributeText(xmlSchemaFieldNode, "UPDATEPARAM") = "IN" _
            Or xmlGetAttributeText(xmlSchemaFieldNode, "UPDATEPARAM") = "INOUT" Then
                Set xmlDataAttrib = getKeyValueFromParentHierarchy(vxmlRequestNode, xmlGetAttributeText(xmlSchemaFieldNode, "NAME"))
            End If
        End If
        
        If xmlGetAttributeText(xmlSchemaFieldNode, "UPDATEPARAM") = "OUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib, adParamOutput)
        ElseIf xmlGetAttributeText(xmlSchemaFieldNode, "UPDATEPARAM") = "INOUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib, adParamInputOutput)
        Else
            Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib)
        End If
        
        cmd.Parameters.Append param
    
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "UPDATE_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "UPDATE_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "UPDATE_PROC")
    End If
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = strProcRef
    
    lngRecordsAffected = adoCRUDExecuteSP(cmd)
    
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@UPDATEPARAM='INOUT' or @UPDATEPARAM='OUT']")
            
        strParamName = xmlGetAttributeText(xmlSchemaFieldNode, "NAME")
    
        For Each param In cmd.Parameters
        
            If param.Name = strParamName Then
            
                adoCRUDParamToXML param, vxmlResponseNode, xmlSchemaFieldNode
                adoCRUDParamToXML param, vxmlRequestNode, xmlSchemaFieldNode
                
                Exit For
            
            End If
            
        Next
    
    Next
    
    If lngRecordsAffected = 0 Then
        If Not xmlGetAttributeAsBoolean(vxmlRequestNode, "_IGNORENOWROWS") Then
            Err.Raise oeNoRowsAffected, strFunctionName, "no records updated"
        End If
    End If
    
    Set cmd = Nothing
    
'   ik_debug
   Debug.Print "adoCRUDUpdateViaSP records updated: " & lngRecordsAffected
'   ik_debug_ends
    
    adoCRUDUpdateRequestOffspring _
        vxmlResponseNode, _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode, _
        vblnIsIntelligentUpdate

adoCRUDUpdateViaSPExit:
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlDataAttrib = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If
    
End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDUpdateViaSQL( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)

On Error GoTo adoCRUDUpdateViaSQLExit
    
    Const strFunctionName As String = "adoCRUDUpdateViaSQL"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
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
    Set xmlSchemaNodeList = xmlSchemaEntityNode.selectNodes("ATTRIBUTE[not(@DATAMAP) or @DATAMAP='IN' or @DATAMAP='INOUT']")
    For Each xmlSchemaFieldNode In xmlSchemaNodeList
        blnDoSet = True
        Set xmlSchemaAttrib = xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE")
        If Not xmlSchemaAttrib Is Nothing Then
            
            If (xmlSchemaAttrib.Value = "FOREIGN") Or _
                (xmlSchemaAttrib.Value = "PRIMARY") _
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
                Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib)
                cmd.Parameters.Append param
                intSetCount = intSetCount + 1
            End If
        End If
    Next
    Set xmlSchemaNodeList = xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@KEYTYPE and not(@KEYTYPE='NULL')]")
    intKeys = xmlSchemaNodeList.length
    For Each xmlSchemaFieldNode In xmlSchemaNodeList
        strAttribName = xmlGetAttributeText(xmlSchemaFieldNode, "NAME")
        Set xmlDataAttrib = _
                   vxmlRequestNode.Attributes.getNamedItem(strAttribName)
                   
        'IK 11/11/2005 CORE215
        'search for primary keys in parent hierarchy
        If xmlDataAttrib Is Nothing Then
            If xmlSchemaFieldNode.Attributes.getNamedItem("KEYTYPE").Text = "FOREIGN" Then
                Set xmlDataAttrib = getKeyValueFromParentHierarchy(vxmlRequestNode, strAttribName)
            End If
        End If
        
        If Not xmlDataAttrib Is Nothing Then
            If xmlDataAttrib.Value <> "" Then
                If Len(strSQLWhere) <> 0 Then
                    strSQLWhere = strSQLWhere & " AND "
                End If
                strSQLWhere = strSQLWhere & strAttribName & "=?"
                Set param = adoCRUDCreateSQLParam(xmlSchemaFieldNode, xmlDataAttrib)
                cmd.Parameters.Append param
                intKeyCount = intKeyCount + 1
            End If
        End If
    Next
    strSQL = strSQL & " SET " & strSQLSet
    strSQL = strSQL & " WHERE " & strSQLWhere
    
'   ik_debug
   Debug.Print "adoUpdate strSQL: " & strSQL
'   ik_debug_ends
    
    ' TODO
    ' modify schema to allow non-unique updates
    ' (schema specific?)
    If (blnNonUniqueInstanceAllowed = False) And (intKeys <> intKeyCount) Then
        Err.Raise oeXMLMissingAttribute, strFunctionName, "not all key values specified"
    End If
    
    If intSetCount > 0 Then
        
        cmd.CommandType = adCmdText
        cmd.CommandText = strSQL
            
        lngRecordsAffected = adoCRUDExecuteSQLCommand(cmd)
    
        If lngRecordsAffected = 0 Then
            If Not xmlGetAttributeAsBoolean(vxmlRequestNode, "_IGNORENOWROWS") Then
                Err.Raise oeNoRowsAffected, strFunctionName, "no records updated"
            End If
        End If
    Else
        Err.Raise oeNoRowsAffected, strFunctionName, "no data for update"
    End If
    
'   ik_debug
   Debug.Print "adoUpdate records updated: " & lngRecordsAffected
'   ik_debug_ends
    
    adoCRUDUpdateRequestOffspring _
        vxmlResponseNode, _
        vxmlRequestNode, _
        vxmlSchemaMasterNode, _
        vxmlSchemaRefNode, _
        vblnIsIntelligentUpdate
    
adoCRUDUpdateViaSQLExit:
    
    Set xmlSchemaNodeList = Nothing
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaFieldNode = Nothing
    Set xmlSchemaAttrib = Nothing
    Set xmlDataAttrib = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDUpdateRequestOffspring( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Const strFunctionName As String = "adoCRUDUpdateRequestOffspring"
    
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisSchemaMasterNode As IXMLDOMNode
    Dim xmlThisSchemaRefNode As IXMLDOMNode
    Dim xmlSchemaNodeList As IXMLDOMNodeList
    Dim strPattern As String
    Dim intNodeIndex As Long
    
    For Each xmlThisRequestNode In vxmlRequestNode.childNodes
            
        ' ik_CORE190_250923
        If xmlThisRequestNode.nodeType = NODE_ELEMENT Then
        
            If xmlGetAttributeText(xmlThisRequestNode, "CRUD_OP") <> "DONE" Then
                
                Set xmlThisSchemaMasterNode = _
                    adoCRUDGetSchemaChildNode( _
                        xmlThisRequestNode, vxmlSchemaMasterNode, vxmlResponseNode)
                
                If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
                    
                    Set xmlThisSchemaRefNode = _
                        adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
                    
                    adoCRUDPopulateChildKeys _
                        xmlThisRequestNode, _
                        xmlThisSchemaRefNode, _
                        vxmlRequestNode
                Else
                    
                    adoCRUDPopulateChildKeys _
                        xmlThisRequestNode, _
                        xmlThisSchemaMasterNode, _
                        vxmlRequestNode
                End If
                
                adoCRUDUpdate _
                    vxmlResponseNode, _
                    xmlThisRequestNode, _
                    xmlThisSchemaMasterNode, _
                    xmlThisSchemaRefNode, _
                    vblnIsIntelligentUpdate
            
            End If
            
        ' ik_CORE190_250923
        End If
    
    Next
    
    errCheckError strFunctionName

End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDDelete( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    Const strFunctionName As String = "adoCRUDDelete"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    Dim xmlSchemaEntityNode As IXMLDOMNode
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "DELETE_PROC") Then
        adoCRUDDeleteViaSP _
            vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
    ElseIf xmlAttributeValueExists(xmlSchemaEntityNode, "DELETE_REF") Then
        Err.Raise oeNotImplemented, strFunctionName, "DELETE_REF not supported"
            
    Else
        If vxmlSchemaRefNode Is Nothing Then
            If vxmlSchemaMasterNode.selectNodes("ENTITY").length = 0 Then
                Err.Raise oeXMLMissingAttribute, strFunctionName, _
                    "schema missing DELETE_PROC operation: " & _
                    vxmlSchemaMasterNode.cloneNode(False).xml
            End If
        Else
            If xmlAttributeValueExists(vxmlSchemaRefNode, "DELETE_PROC") Then
                adoCRUDDeleteViaSP _
                    vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
            ElseIf xmlAttributeValueExists(vxmlSchemaRefNode, "DELETE_REF") Then
                Err.Raise oeNotImplemented, strFunctionName, "DELETE_REF not supported"
            Else
                Err.Raise oeXMLMissingAttribute, strFunctionName, _
                    "schema missing DELETE_PROC operation: " & _
                    vxmlSchemaRefNode.cloneNode(False).xml
            End If
        End If
    End If
End Sub
'APS - Implemented IK's CRUD changes 20/08/01
Private Sub adoCRUDDeleteViaSP( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDDeleteViaSPExit
    Const strFunctionName As String = "adoCRUDDeleteViaSP"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If

    Dim xmlSchemaAttribNodeList As IXMLDOMNodeList
    Dim xmlSchemaEntityNode As IXMLDOMNode
    Dim xmlSchemaAttribNode As IXMLDOMNode
    
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
    
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Dim strSQLParamClause As String, _
        strProcRef As String, _
        strAttribName
    
    Dim lngRecordsAffected As Long
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    Set cmd = New ADODB.Command
    
    Set xmlSchemaAttribNodeList = _
        xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@DELETEPARAM][not(DELETEPARAM='NO')]")
    
    For Each xmlSchemaAttribNode In xmlSchemaAttribNodeList
        
        strAttribName = xmlGetMandatoryAttributeText(xmlSchemaAttribNode, "NAME")
        Set xmlDataAttrib = vxmlRequestNode.Attributes.getNamedItem(strAttribName)
        
        'IK 11/11/2005 CORE215
        'search for parameter values in parent hierarchy
        If xmlDataAttrib Is Nothing Then
            Set xmlDataAttrib = getKeyValueFromParentHierarchy(vxmlRequestNode, strAttribName)
        End If
        
        'IK 12/10/2005 CORE199
        If xmlGetAttributeText(xmlSchemaAttribNode, "DELETEPARAM") = "OUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib, adParamOutput)
        ElseIf xmlGetAttributeText(xmlSchemaAttribNode, "DELETEPARAM") = "INOUT" Then
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib, adParamInputOutput)
        Else
            Set param = adoCRUDCreateSQLParam(xmlSchemaAttribNode, xmlDataAttrib)
        End If
        'IK 12/10/2005 CORE199 ends
        
        cmd.Parameters.Append param
    
    Next
    
    If xmlAttributeValueExists(vxmlSchemaMasterNode, "DELETE_PROC") Then
        strProcRef = xmlGetAttributeText(vxmlSchemaMasterNode, "DELETE_PROC")
    Else
        strProcRef = xmlGetAttributeText(vxmlSchemaRefNode, "DELETE_PROC")
    End If
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = strProcRef
    
    lngRecordsAffected = adoCRUDExecuteSP(cmd)
    
    For Each xmlSchemaAttribNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE[@DELETEPARAM='INOUT' or @DELETEPARAM='OUT']")
            
        strAttribName = xmlGetAttributeText(xmlSchemaAttribNode, "NAME")
    
        For Each param In cmd.Parameters
        
            If param.Name = strAttribName Then
            
                adoCRUDParamToXML param, vxmlResponseNode, xmlSchemaAttribNode
                
                Exit For
            
            End If
            
        Next
    
    Next
    
    If lngRecordsAffected = 0 Then
        If Not xmlGetAttributeAsBoolean(vxmlRequestNode, "_IGNORENOWROWS") Then
            Err.Raise oeNoRowsAffected, strFunctionName, "no records deleted"
        End If
    End If
    
    Set cmd = Nothing
    
'   ik_debug
   Debug.Print "adoCRUDDeleteViaSP records deleted: " & lngRecordsAffected
'   ik_debug_ends

adoCRUDDeleteViaSPExit:
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    Set xmlSchemaEntityNode = Nothing
    Set xmlSchemaAttribNode = Nothing
    Set xmlSchemaAttribNodeList = Nothing
    
    If Err.Number <> 0 Then
        If Left(Err.Source, 7) <> "adoCRUD" Then
            adoCRUDSQLError strFunctionName, cmd, vxmlResponseNode, vxmlRequestNode, vxmlSchemaMasterNode, vxmlSchemaRefNode
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

End Sub

'   ik_20040225
'   bmids changes interfere with read process
Private Sub adoCRUDPopulateChildKeys( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlKeyNodes As IXMLDOMNodeList
    Dim xmlKeyNode As IXMLDOMNode
    Dim xmlResponseParentNode As IXMLDOMNode
    Dim xmlRequestParentNode As IXMLDOMNode
    
    Dim strAttribName As String, _
        strAttribValue As String
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & "adoCRUDPopulateChildKeys")
    #End If
    
    Set xmlResponseParentNode = vxmlResponseNode
    Do While xmlResponseParentNode.Attributes.length = 0
        
        Set xmlResponseParentNode = xmlResponseParentNode.parentNode
        
        If xmlResponseParentNode.nodeType = NODE_DOCUMENT Then
            Set xmlResponseParentNode = Nothing
            Exit Do
        End If
        
    Loop
    
    Set xmlRequestParentNode = vxmlRequestNode.parentNode
    If xmlRequestParentNode.nodeName = "REQUEST" Then
        Set xmlRequestParentNode = Nothing
    End If
    
    Set xmlKeyNodes = vxmlSchemaNode.selectNodes("ATTRIBUTE[@KEYTYPE='FOREIGN']")
    
    For Each xmlKeyNode In xmlKeyNodes
        
        strAttribName = _
            xmlGetMandatoryAttributeText(xmlKeyNode, "NAME")
            
        If Not xmlRequestParentNode Is Nothing Then
            If xmlAttributeValueExists(xmlRequestParentNode, strAttribName) Then
                xmlCopyAttribute _
                    xmlRequestParentNode, vxmlRequestNode, strAttribName
            End If
        End If
                   
        If Not xmlResponseParentNode Is Nothing Then
            If xmlAttributeValueExists(xmlResponseParentNode, strAttribName) Then
                
                xmlCopyAttribute _
                    xmlResponseParentNode, vxmlRequestNode, strAttribName
            
            End If
        End If
    
    Next
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If

End Sub

Private Sub adoCRUDReadPopulateChildKeys( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim xmlKeyNodes As IXMLDOMNodeList
    Dim xmlKeyNode As IXMLDOMNode
    Dim xmlResponseParentNode As IXMLDOMNode
    Dim xmlRequestParentNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    ' ik_20050803_CORE179
    Dim xmlAttrib As IXMLDOMAttribute
    ' ik_20050803_CORE179_ends
    
    Dim strAttribName As String, _
        strAttribValue As String

    ' ik_20050627_CORE158/2
    Dim blnFKeyHit As Boolean
    ' ik_20050627_CORE158_ends
    
    Set xmlResponseParentNode = vxmlResponseNode
    Do While xmlResponseParentNode.Attributes.length = 0
        
        Set xmlResponseParentNode = xmlResponseParentNode.parentNode
        
        If xmlResponseParentNode.nodeType = NODE_DOCUMENT Then
            Set xmlResponseParentNode = Nothing
            Exit Do
        End If
        
    Loop
    
    Set xmlRequestParentNode = vxmlRequestNode.parentNode
    If xmlRequestParentNode.nodeName = "REQUEST" Then
        Set xmlRequestParentNode = Nothing
    End If
    
    Set xmlKeyNodes = vxmlSchemaNode.selectNodes("ATTRIBUTE[@GETDEFAULT]")
    
    For Each xmlKeyNode In xmlKeyNodes
        
        strAttribName = _
            xmlGetMandatoryAttributeText(xmlKeyNode, "NAME")
            
        strAttribValue = xmlGetAttributeText(xmlKeyNode, "GETDEFAULT")
        
        If Left(strAttribValue, 6) = "xpath:" Then
        
            strAttribValue = Mid(strAttribValue, 7)
            
            Set xmlElem = vxmlResponseNode.ownerDocument.createElement("_bogus_")
            Set xmlNode = vxmlResponseNode.appendChild(xmlElem)
            
            If Not xmlNode.selectSingleNode(strAttribValue) Is Nothing Then
            
                ' ik_20050803_CORE179
                If Not vxmlRequestNode.Attributes.getNamedItem(strAttribName) Is Nothing Then
                    Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("_request_" & strAttribName)
                    xmlAttrib.Text = vxmlRequestNode.Attributes.getNamedItem(strAttribName).Text
                    vxmlRequestNode.Attributes.setNamedItem xmlAttrib
                End If
                ' ik_20050803_CORE179_ends
    
                xmlSetAttributeValue _
                    vxmlRequestNode, _
                    strAttribName, _
                    "responseValue::" & xmlNode.selectSingleNode(strAttribValue).Text
                    
            Else
            
                ' ik_20050627_CORE158/2
                xmlSetAttributeValue vxmlRequestNode, "_nullDefault", "true"
                ' ik_20050627_CORE158_ends
            
            End If
            
            vxmlResponseNode.removeChild xmlNode
    
        Else
    
            xmlSetAttributeValue _
                vxmlRequestNode, _
                strAttribName, _
                "responseValue::" & strAttribValue ''IK 11/11/2005 CORE214

        
        End If
    
    Next
    
    Set xmlKeyNodes = vxmlSchemaNode.selectNodes("ATTRIBUTE[@KEYTYPE='FOREIGN']")
    
    For Each xmlKeyNode In xmlKeyNodes
        
        strAttribName = _
            xmlGetMandatoryAttributeText(xmlKeyNode, "NAME")

        ' ik_20050627_CORE158/2
        If xmlAttributeValueExists(vxmlRequestNode, strAttribName) Then
            blnFKeyHit = True
        End If
        ' ik_20050627_CORE158_ends
        
        If Not xmlAttributeValueExists(vxmlRequestNode, strAttribName) Then
            If Not xmlRequestParentNode Is Nothing Then
                
                If xmlAttributeValueExists(xmlRequestParentNode, strAttribName) Then
                    xmlCopyAttribute _
                        xmlRequestParentNode, vxmlRequestNode, strAttribName
                ' ik_20050627_CORE158/2
                blnFKeyHit = True
                ' ik_20050627_CORE158_ends
                
                End If
            End If
        End If
                   
        If Not xmlAttributeValueExists(vxmlRequestNode, strAttribName) Then
            If xmlAttributeValueExists(xmlResponseParentNode, strAttribName) Then
                
                xmlSetAttributeValue _
                    vxmlRequestNode, _
                    strAttribName, _
                    "responseValue::" & xmlGetAttributeText(xmlResponseParentNode, strAttribName)

                ' ik_20050627_CORE158/2
                blnFKeyHit = True
                ' ik_20050627_CORE158_ends
            
            End If
        
        End If
    
    Next
            
    ' ik_20050627_CORE158/2
    If xmlKeyNodes.length <> 0 And blnFKeyHit = False Then
        xmlSetAttributeValue vxmlRequestNode, "_nullForeignKeys", "true"
    End If
    ' ik_20050627_CORE158_ends
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
End Sub

'APS - Implemented IK's CRUD changes 20/08/01
Private Function adoCRUDExecuteSP( _
    ByVal cmd As ADODB.Command, _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As Long
On Error GoTo adoCRUDExecuteSPExit
    
    Const strFunctionName As String = "adoCRUDExecuteSP"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
        
    Dim lngRecordsAffected As Long
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = getDbConnectString(vstrComponentRef)
    conn.Open
    cmd.ActiveConnection = conn
    cmd.Execute lngRecordsAffected
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    adoCRUDExecuteSP = lngRecordsAffected

adoCRUDExecuteSPExit:
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Err.Source = App.EXEName Then
            Err.Source = strFunctionName
        Else
            Err.Source = strFunctionName & "." & Err.Source
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Private Sub adoCRUDGetUpdateDefaults( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean)
    
    Const strFunctionName As String = "adoCRUDGetUpdateDefaults"
    
    Dim xmlSchemaEntityNode As IXMLDOMNode, _
        xmlSchemaFieldNode As IXMLDOMNode
    Dim xmlDataAttrib As IXMLDOMAttribute, _
        xmlAttrib As IXMLDOMAttribute
    
    Dim strAttribName As String, _
        strAttribType As String
    
    If vxmlSchemaRefNode Is Nothing Then
        Set xmlSchemaEntityNode = vxmlSchemaMasterNode
    Else
        Set xmlSchemaEntityNode = vxmlSchemaRefNode
    End If
    
    ' get any xpath default values
    For Each xmlSchemaFieldNode In xmlSchemaEntityNode.selectNodes("ATTRIBUTE")
        strAttribName = xmlGetMandatoryAttributeText(xmlSchemaFieldNode, "NAME")
        If vxmlRequestNode.Attributes.getNamedItem(strAttribName) Is Nothing Then
            If Not xmlSchemaFieldNode.Attributes.getNamedItem("PUTDEFAULT") Is Nothing Then
                Set xmlAttrib = _
                    vxmlRequestNode.ownerDocument.createAttribute(strAttribName)
                xmlAttrib.Text = _
                    xmlSchemaFieldNode.Attributes.getNamedItem("PUTDEFAULT").Text
                If Left(xmlAttrib.Text, 6) = "xpath:" Then
                    adoCRUDGetDefaultValueViaXPath _
                        vxmlResponseNode, _
                        vxmlRequestNode, _
                        vxmlSchemaMasterNode, _
                        vxmlSchemaRefNode, _
                        xmlSchemaFieldNode, _
                        vblnIsIntelligentUpdate, _
                        False
                Else
                    
                    strAttribType = xmlSchemaFieldNode.Attributes.getNamedItem("DATATYPE").Text
                    
                    If (strAttribType = "DATE" Or strAttribType = "DATETIME") And xmlAttrib.Text = "_NOW" Then
                        xmlAttrib.Text = FormatNow
                    End If
                        
                    vxmlRequestNode.Attributes.setNamedItem xmlAttrib
                
                End If
            End If
        End If
    Next
    
End Sub

Private Sub adoCRUDGetDefaultValueViaXPath( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vxmlSchemaAttribNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean, _
    ByVal vblnIsCreate As Boolean)
    
    Const strFunctionName As String = "adoCRUDGetDefaultValueViaXPath"
    
    Dim xmlChildNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim strAttribName As String, _
        strPath As String, _
        strChildNodePath As String
        
    Dim blnWasSpawned As Boolean, _
        blnDoUpdate As Boolean
        
    strPath = xmlGetAttributeText(vxmlSchemaAttribNode, "PUTDEFAULT")
    strPath = Mid(strPath, 7)
    
    If vxmlRequestNode.selectSingleNode(strPath) Is Nothing Then
        
        If Left(strPath, 7) = "child::" Then
        
            strChildNodePath = Mid(strPath, 1, InStr(strPath, "/") - 1)
            
            Set xmlChildNode = vxmlRequestNode.selectSingleNode(strChildNodePath)
            
            If Not xmlChildNode Is Nothing Then
            
                adoCRUDSpawnChildNode _
                    vxmlResponseNode, _
                    xmlChildNode, _
                    vxmlRequestNode, _
                    vxmlSchemaMasterNode, _
                    vxmlSchemaRefNode, _
                    vblnIsIntelligentUpdate, _
                    vblnIsCreate
                blnWasSpawned = True
                
            End If
        End If
    End If
        
    If Not vxmlRequestNode.selectSingleNode(strPath) Is Nothing Then
        If vblnIsCreate Then
            blnDoUpdate = True
        ElseIf blnWasSpawned Then
            blnDoUpdate = True
        End If
    End If
        
    If blnDoUpdate Then
        
        strAttribName = xmlGetAttributeText(vxmlSchemaAttribNode, "NAME")
        
        Set xmlAttrib = _
            vxmlRequestNode.ownerDocument.createAttribute(strAttribName)
        
        xmlAttrib.Text = _
            vxmlRequestNode.selectSingleNode(strPath).Text
        
        vxmlRequestNode.Attributes.setNamedItem xmlAttrib
        
        If Not xmlAttributeValueExists(vxmlRequestNode, "CRUD_OP") Then
            If Not vblnIsCreate Then
                xmlSetAttributeValue vxmlRequestNode, "CRUD_OP", "UPDATE"
            End If
        End If
        
    End If
    
    If Err.Number <> 0 Then
        If Err.Source = App.EXEName Then
            Err.Source = strFunctionName
        Else
            Err.Source = strFunctionName & "." & Err.Source
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Sub

Private Sub adoCRUDSpawnChildNode( _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByVal vblnIsIntelligentUpdate As Boolean, _
    ByVal vblnIsCreate As Boolean)
    
    Const strFunctionName As String = "adoCRUDSpawnChildNodes"
    
    Dim xmlThisSchemaMasterNode As IXMLDOMNode
    Dim xmlThisSchemaRefNode As IXMLDOMNode
    
    Dim strOp As String
        
    If vblnIsCreate Then
        strOp = "CREATE"
    Else
        strOp = xmlGetAttributeText(vxmlRequestNode, "CRUD_OP", "iUPDATE")
    End If

    Set xmlThisSchemaMasterNode = _
        adoCRUDGetSchemaChildNode(vxmlRequestNode, vxmlSchemaMasterNode, vxmlResponseNode)
    
    If xmlAttributeValueExists(xmlThisSchemaMasterNode, "ENTITY_REF") Then
        
        Set xmlThisSchemaRefNode = _
            adoCRUDGetSchemaRefNode(xmlThisSchemaMasterNode)
        
        adoCRUDPopulateChildKeys _
            vxmlRequestNode, _
            xmlThisSchemaRefNode, _
            vxmlRequestParentNode
    Else
        adoCRUDPopulateChildKeys _
            vxmlRequestNode, _
            xmlThisSchemaMasterNode, _
            vxmlRequestParentNode
    End If
    
    Select Case strOp
    
        Case "CREATE"
            adoCRUDCreate vxmlResponseNode, vxmlRequestNode, xmlThisSchemaMasterNode, xmlThisSchemaRefNode, vblnIsIntelligentUpdate
    
        Case "UPDATE", "iUPDATE"
            adoCRUDUpdate vxmlResponseNode, vxmlRequestNode, xmlThisSchemaMasterNode, xmlThisSchemaRefNode, vblnIsIntelligentUpdate
    
    End Select
        
    xmlSetAttributeValue vxmlRequestNode, "CRUD_OP", "DONE"
    
    Set xmlThisSchemaMasterNode = Nothing
    Set xmlThisSchemaRefNode = Nothing
    
    If Err.Number <> 0 Then
        If Err.Source = App.EXEName Then
            Err.Source = strFunctionName
        Else
            Err.Source = strFunctionName & "." & Err.Source
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Sub

Private Function adoCRUDExecuteGetRecordSet( _
    ByVal cmd As ADODB.Command, _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As ADODB.Recordset

On Error GoTo executeGetRecordSetExit
    
    Const strFunctionName As String = "adoCRUDExecuteGetRecordSet"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
    ' </VSA>
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    conn.ConnectionString = getDbConnectString(vstrComponentRef)
    conn.Open
    cmd.ActiveConnection = conn
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.Open cmd
    
    ' ik_CORE194_20050923
    If rst.State = adStateOpen Then
        ' disconnect RecordSet
        Set rst.ActiveConnection = Nothing
        conn.Close
        If Not rst.EOF Then
            rst.MoveFirst
            Set adoCRUDExecuteGetRecordSet = rst
        End If
    ' ik_CORE194_20050923
    End If

executeGetRecordSetExit:
    
    Set rst = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Err.Source = App.EXEName Then
            Err.Source = strFunctionName
        Else
            Err.Source = strFunctionName & "." & Err.Source
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Private Function adoCRUDExecuteSQLCommand( _
    ByVal cmd As ADODB.Command, _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As Long

On Error GoTo executeSQLCommandExit
    
    Const strFunctionName As String = "adoCRUDExecuteSQLCommand"
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (gstrMODULEPREFIX & strFunctionName)
    #End If
        
    Dim lngRecordsAffected As Long
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = getDbConnectString(vstrComponentRef)
    conn.Open
    cmd.ActiveConnection = conn
    cmd.Execute lngRecordsAffected, , adExecuteNoRecords
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    adoCRUDExecuteSQLCommand = lngRecordsAffected

executeSQLCommandExit:
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    If Err.Number <> 0 Then
        If Err.Source = App.EXEName Then
            Err.Source = strFunctionName
        Else
            Err.Source = strFunctionName & "." & Err.Source
        End If
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

' APS 29/08/01 Removed Optional Parameter
Private Function adoCRUDCreateSQLParam( _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vxmlAttrib As IXMLDOMAttribute, _
    Optional ByVal paramDirection As Integer = adParamInput) _
    As ADODB.Parameter
    
    On Error GoTo adoCRUDCreateSQLParamExit
    
    Dim param As ADODB.Parameter
    Dim strDataType As String, _
        strDataValue As String
    Dim LngStringLen As Long
    Set param = New ADODB.Parameter
        
    strDataType = vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
    If Not vxmlAttrib Is Nothing Then
        
        strDataValue = Trim(vxmlAttrib.Text)
    
        'IK 18/10/2005 CORE203
        If Left(strDataValue, 6) = "xpath:" Then
            
            Dim xmlNode As IXMLDOMNode
            Dim strXPath As String
            
            strXPath = Mid(strDataValue, 7)
            
            Set xmlNode = vxmlAttrib.selectSingleNode("../.")
            Set xmlNode = xmlNode.selectSingleNode(strXPath)
            If Not xmlNode Is Nothing Then
                ' CORE238
                If strDataType = "XMLSTREAM" Then
                    strDataValue = xmlNode.xml
                    xmlNode.parentNode.removeChild xmlNode
                Else
                    strDataValue = Trim(xmlNode.Text)
                End If
            Else
                strDataValue = ""
            End If
        End If
        'IK 18/10/2005 CORE203 ends
    
    End If
    
    Select Case strDataType
        
        ' CORE238
        ' Case "STRING", "TEXT", "ESTRING"
        Case "STRING", "TEXT", "ESTRING", "XMLSTREAM"
            
            param.Type = adBSTR
            LngStringLen = xmlGetAttributeAsLong(vxmlSchemaNode, "LENGTH")
            
            If Len(strDataValue) > 0 And strDataType = "ESTRING" Then
                strDataValue = Encrypt(strDataValue)
            End If
            
            If LngStringLen > 0 And Len(strDataValue) > LngStringLen Then
                strDataValue = Left(strDataValue, LngStringLen)
            End If
                        
            If paramDirection = adParamOutput Or paramDirection = adParamInputOutput Then
                param.Size = LngStringLen
            Else
                If Not vxmlAttrib Is Nothing Then
                    param.Size = Len(strDataValue)
                Else
                    param.Size = LngStringLen
                End If
            End If
        
        Case "GUID"
            'IK 27/10/2005 CORE211
            param.Type = adBinary
            param.Size = 16
            If Len(strDataValue) > 0 Then
                strDataValue = GuidStringToByteArray(strDataValue)
            End If
        
        Case "SHORT", "COMBO", "LONG"
            param.Type = adInteger
        
        Case "BOOLEAN"
            strDataValue = UCase(strDataValue)
            ' CORE249
            If strDataValue = getBooleanTrue(vxmlSchemaNode) Or _
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
    
    param.Direction = paramDirection
    
    
'    If paramDirection = adParamOutput Or paramDirection = adParamInputOutput Then
'        param.Name = xmlGetAttributeText(vxmlSchemaNode, "NAME")
'    End If
    
    param.Name = xmlGetAttributeText(vxmlSchemaNode, "NAME")
    param.Value = Null
    
    If Not vxmlAttrib Is Nothing Then
        If Len(strDataValue) > 0 Then
            param.Value = strDataValue
        End If
    End If
    
adoCRUDCreateSQLParamExit:
    
    Set adoCRUDCreateSQLParam = param
    Set param = Nothing
    
    If Err.Number <> 0 Then
    
        Err.Description = Err.Description & ", ATTRIBUTE: " & vxmlAttrib.xml
        If Err.Source = App.EXEName Then
            Err.Source = adoCRUDCreateSQLParam
        Else
            Err.Source = adoCRUDCreateSQLParam & "." & Err.Source
        End If
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

Private Function adoBinToString(bytArray() As Byte, binLen As Long) As String
    Dim i As Integer
    Dim strGuid As String
    strGuid = ""
    For i = 0 To (binLen - 1)
        strGuid = strGuid & Right$(Hex$(bytArray(i) + &H100), 2)
    Next i
    adoBinToString = "0x" & strGuid
End Function

' CORE249
Private Function adoDateToString(ByVal vvarDate As Variant, ByVal vgeDateStyle As DATESTYLE) As String
    If vgeDateStyle = e_dateISO8601 Then
        adoDateToString = _
             Right(str(DatePart("yyyy", vvarDate)), 4) & "-" & Right(str(DatePart("m", vvarDate) + 100), 2) & "-" & Right(str(DatePart("d", vvarDate) + 100), 2)
    Else
        adoDateToString = _
            Right(str(DatePart("d", vvarDate) + 100), 2) & "/" & Right(str(DatePart("m", vvarDate) + 100), 2) & "/" & Right(str(DatePart("yyyy", vvarDate)), 4)
    End If
End Function

' CORE249
Private Function adoDateTimeToString(ByVal vvarDate As Variant, ByVal vgeDateStyle As DATESTYLE) As String
    
    adoDateTimeToString = adoDateToString(vvarDate, vgeDateStyle) & " " & _
               Right(str(DatePart("h", vvarDate) + 100), 2) & ":" & _
               Right(str(DatePart("n", vvarDate) + 100), 2) & ":" & _
               Right(str(DatePart("s", vvarDate) + 100), 2)
End Function

Private Sub adoCRUDFieldToXml( _
    ByVal vfld As ADODB.Field, _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean, _
    ByVal venumComboTypeLookUp As COMBOVALIDATIONLOOKUPTYPE)
    
    On Error GoTo adoCRUDFieldToXmlExit
    
    Dim strAttribName As String
    
    If vxmlSchemaNode.nodeName = "ATTRIBUTE" Then
        strAttribName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
    Else
        strAttribName = vxmlSchemaNode.nodeName
    End If
    
    Select Case vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
        
        Case "GUID"
            vxmlOutElem.setAttribute strAttribName, adoGuidToString(vfld.Value)
            
        Case Else
            adoCRUDVarToXML strAttribName, vfld.Value, vxmlOutElem, vxmlSchemaNode, vblnDoComboLookUp, venumComboTypeLookUp
    
    End Select

adoCRUDFieldToXmlExit:
    
    If Err.Number <> 0 Then
    
        If Left(Err.Source, 7) <> "adoCRUD" Then
            Err.Description = Err.Description & ", ATTRIBUTE: " & strAttribName & "=" & vfld.Value
        End If
        
        If Err.Source = App.EXEName Then
            Err.Source = "adoCRUDFieldToXml"
        Else
            Err.Source = "adoCRUDFieldToXml" & "." & Err.Source
        End If
        
        Err.Raise Err.Number, Err.Source, Err.Description
    
    End If

End Sub

Private Sub adoCRUDParamToXML( _
    ByVal vParam As ADODB.Parameter, _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode)
    
    On Error GoTo adoCRUDParamToXMLExit
    
    Dim strAttribName As String
    
    If vxmlSchemaNode.nodeName = "ATTRIBUTE" Then
        strAttribName = xmlGetAttributeText(vxmlSchemaNode, "NAME")
    Else
        strAttribName = vxmlSchemaNode.nodeName
    End If
    
    Select Case vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
        
        Case "GUID"
            vxmlOutElem.setAttribute strAttribName, adoGuidToString(vParam.Value)
            
        Case Else
            adoCRUDVarToXML strAttribName, vParam.Value, vxmlOutElem, vxmlSchemaNode, False, e_cvluNone
    
    End Select
    
adoCRUDParamToXMLExit:
    
    If Err.Number <> 0 Then
    
        If Left(Err.Source, 7) <> "adoCRUD" Then
            Err.Description = Err.Description & ", ATTRIBUTE: " & strAttribName & "=" & vParam.Value
        End If
        
        If Err.Source = App.EXEName Then
            Err.Source = "adoCRUDParamToXML"
        Else
            Err.Source = "adoCRUDParamToXML" & "." & Err.Source
        End If
        
        Err.Raise Err.Number, Err.Source, Err.Description
    
    End If

End Sub

Private Sub adoCRUDVarToXML( _
    ByVal vstrAttribName As String, _
    ByVal vValue As Variant, _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vblnDoComboLookUp As Boolean, _
    ByVal venumComboTypeLookUp As COMBOVALIDATIONLOOKUPTYPE)
    
    On Error GoTo adoCRUDVarToXMLExit
    
    Dim strElementValue As String, _
        strComboGroup As String, _
        strComboValue As String, _
        strFormatMask As String, _
        strFormatted As String
        
    ' CORE261
    Dim lngIndex As Long
    Dim strBuffer As String
        
    If IsNull(vValue) Then
        
        vxmlOutElem.setAttribute vstrAttribName, "null"
        
    Else
    
        Select Case vxmlSchemaNode.Attributes.getNamedItem("DATATYPE").Text
            
            Case "CURRENCY"
                
                ' IK 21/10/2005 CORE206 - round currency values
                strFormatted = CStr(Format(Round(vValue, 2), "0.00"))
                vxmlOutElem.setAttribute vstrAttribName, strFormatted
            
            Case "DOUBLE"
                
                strFormatMask = Empty
                If Not vxmlSchemaNode.Attributes.getNamedItem("FORMATMASK") Is Nothing Then
                    strFormatMask = vxmlSchemaNode.Attributes.getNamedItem("FORMATMASK").Text
                End If
                If Len(strFormatMask) <> 0 Then
                    
                    vxmlOutElem.setAttribute _
                               vstrAttribName & "_RAW", _
                               vValue
                    vxmlOutElem.setAttribute _
                               vstrAttribName, _
                               Format(vValue, strFormatMask)
                Else
                    
                    vxmlOutElem.setAttribute vstrAttribName, vValue
                    
                End If
            
            Case "COMBO"
                
                vxmlOutElem.setAttribute vstrAttribName, vValue
                    
                If vblnDoComboLookUp = True Then
                    
                    If Not vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP") Is Nothing Then
                        strComboGroup = vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP").Text
                        strComboValue = adoCRUDGetComboText(strComboGroup, vValue)
                        If Len(strComboValue) <> 0 Then
                            vxmlOutElem.setAttribute _
                                       vstrAttribName & "_TEXT", strComboValue
                        End If
                    End If
                
                End If
                
                If venumComboTypeLookUp <> e_cvluNone Then
                    
                    If Not vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP") Is Nothing Then
                        
                        strComboGroup = vxmlSchemaNode.Attributes.getNamedItem("COMBOGROUP").Text
                        
                        If venumComboTypeLookUp = e_cvluPipeDelim Then
                        
                            strComboValue = adoCRUDGetComboTypes(strComboGroup, vValue)
                            If Len(strComboValue) <> 0 Then
                                vxmlOutElem.setAttribute _
                                    vstrAttribName & "_TYPES", strComboValue
                            End If
                        ElseIf venumComboTypeLookUp = e_cvluExtended Then
                        
                            adoCRUDGetComboTypesEx vxmlOutElem, vstrAttribName, strComboGroup, vValue
                        
                        End If
                        
                    End If
                
                End If
            
            Case "DATE"
                ' CORE249
                vxmlOutElem.setAttribute _
                           vstrAttribName, _
                           adoDateToString(vValue, getDateStyle(vxmlSchemaNode))
            
            Case "DATETIME"
                ' CORE249
                vxmlOutElem.setAttribute _
                           vstrAttribName, _
                           adoDateTimeToString(vValue, getDateStyle(vxmlSchemaNode))
            
            Case "BOOLEAN"
                
                If vValue = 1 Then
                    ' CORE249
                    vxmlOutElem.setAttribute vstrAttribName, getBooleanTrue(vxmlSchemaNode)
                Else
                    ' CORE249
                    vxmlOutElem.setAttribute vstrAttribName, getBooleanFalse(vxmlSchemaNode)
                End If
            
            Case "ESTRING"
                strFormatted = Decrypt(vValue)
                vxmlOutElem.setAttribute vstrAttribName, strFormatted
                
            ' CORE238
            Case "XMLSTREAM"
                DeStream vstrAttribName, vxmlOutElem, vValue
                
            ' CORE261
            Case "BINARY"
                If VarType(vValue) = (vbArray + vbByte) Then
                    If Len(lngIndex) > 0 Then
                        strFormatted = "0x"
                        For lngIndex = 1 To Len(vValue) * 2
                            strBuffer = "0" + Hex(vValue(lngIndex - 1))
                            strFormatted = strFormatted & Right(strBuffer, 2)
                        Next
                        vxmlOutElem.setAttribute vstrAttribName, strFormatted
                    End If
                End If
            
            Case Else
                If VarType(vValue) = vbString Then
                    vxmlOutElem.setAttribute vstrAttribName, RTrim(vValue)
                Else
                    vxmlOutElem.setAttribute vstrAttribName, vValue
                End If
        
        End Select
        
    End If
    
adoCRUDVarToXMLExit:
    
    If Err.Number <> 0 Then
    
        Err.Description = Err.Description & ", ATTRIBUTE: " & vstrAttribName & "=" & vValue
        Err.Raise Err.Number, "adoCRUDVarToXML", Err.Description
    
    End If

End Sub

Private Function adoCRUDGetComboText( _
    ByVal vstrGroupName As String, _
    ByVal vIntValueID As Integer) _
    As String
    Dim strPattern As String
        
    adoCRUDLoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']" & _
        "/COMBOVALUE[@VALUEID='" & CStr(vIntValueID) & "']"
    
    If Not gxmlCRUDCombos.selectSingleNode(strPattern) Is Nothing Then
        adoCRUDGetComboText = _
            xmlGetAttributeText(gxmlCRUDCombos.selectSingleNode(strPattern), "VALUENAME")
    End If

End Function

Private Function adoCRUDGetComboTypes( _
    ByVal vstrGroupName As String, _
    ByVal vIntValueID As Integer) _
    As String
    
    Dim xmlNode As IXMLDOMNode
    
    Dim strPattern As String, _
        strTypes As String
        
    adoCRUDLoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']" & _
        "/COMBOVALUE[@VALUEID='" & CStr(vIntValueID) & "']" & _
        "/COMBOVALIDATION"
        
    For Each xmlNode In gxmlCRUDCombos.selectNodes(strPattern)
        strTypes = strTypes & xmlGetAttributeText(xmlNode, "VALIDATIONTYPE") & "|"
    Next
    
    If Len(strTypes) > 0 Then
        adoCRUDGetComboTypes = Left(strTypes, Len(strTypes) - 1)
    End If

End Function

Private Function adoCRUDGetComboTypesEx( _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vstrAttribName As String, _
    ByVal vstrGroupName As String, _
    ByVal vIntValueID As Integer) _
    As String
    
    Dim xmlNode As IXMLDOMNode
    
    Dim strPattern As String, _
        strType As String
        
    adoCRUDLoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']" & _
        "/COMBOVALUE[@VALUEID='" & CStr(vIntValueID) & "']" & _
        "/COMBOVALIDATION"
        
    For Each xmlNode In gxmlCRUDCombos.selectNodes(strPattern)
        strType = xmlGetAttributeText(xmlNode, "VALIDATIONTYPE")
        If Len(strType) <> 0 Then
            If ContainsSpecialChars(strType) = False Then
                vxmlOutElem.setAttribute vstrAttribName & "_TYPE_" & strType, "true"
            End If
        End If
    Next

End Function

Private Sub adoCRUDLoadComboGroup(ByVal vstrGroupName As String)
    
    Const strRequestPrefix As String = "<REQUEST CRUD_OP='READ' ENTITY_REF='COMBOGROUP'>"
    Const strNodePrefix As String = "<COMBOGROUP GROUPNAME='"
    Const strRequestTerm As String = "'/></REQUEST>"
    
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlComboListNode As IXMLDOMNode
    Dim xmlComboNode As IXMLDOMNode
    
    Dim strPattern As String, _
        strRequest As String
    
    If gxmlCRUDCombos Is Nothing Then
'       ik_debug
'       App.LogEvent "load combo group, create: gxmlCRUDCombos" & vstrGroupName, vbLogEventTypeInformation
        
        Set gxmlCRUDCombos = New DOMDocument40
        Set xmlElem = gxmlCRUDCombos.createElement("COMBOLIST")
        Set xmlComboListNode = gxmlCRUDCombos.appendChild(xmlElem)
    Else
        Set xmlComboListNode = gxmlCRUDCombos.selectSingleNode("COMBOLIST")
    End If
    
    strPattern = "COMBOLIST/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']"
    
    If gxmlCRUDCombos.selectSingleNode(strPattern) Is Nothing Then
    
'       ik_debug
'       App.LogEvent "load combo group: " & vstrGroupName, vbLogEventTypeInformation
        
        Set xmlRequestDoc = New DOMDocument40
        xmlRequestDoc.async = False
        xmlRequestDoc.loadXML strRequestPrefix & strNodePrefix & vstrGroupName & strRequestTerm
        
        Set xmlResponseDoc = New DOMDocument40
        xmlResponseDoc.async = False
        xmlResponseDoc.loadXML "<RESPONSE/>"
        
        Set xmlRequestNode = xmlRequestDoc.selectSingleNode("REQUEST")
        Set xmlResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
        
        adoCRUD xmlRequestNode, xmlRequestNode, xmlResponseNode
        Set xmlResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
        Set xmlComboNode = xmlResponseNode.selectSingleNode("COMBOGROUP")
        
        If Not xmlComboNode Is Nothing Then
            xmlComboListNode.appendChild xmlComboNode.cloneNode(True)
        End If
        
    End If
    
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlComboListNode = Nothing
    Set xmlComboNode = Nothing
    
    Set xmlElem = Nothing

    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
End Sub

Public Function adoCRUDGetMessageText( _
    ByVal vlngMessageNo As Long, _
    Optional ByVal vintMessageField As Integer = omMESSAGE_TEXT) _
    As String
    
    Const strRequestPrefix As String = "<REQUEST CRUD_OP='READ' ENTITY_REF='MESSAGE'>"
    Const strNodePrefix As String = "<MESSAGE MESSAGENUMBER='"
    Const strRequestTerm As String = "'/></REQUEST>"
    
    Dim xmlRequestDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlMessageNode As IXMLDOMNode
    
    Set xmlRequestDoc = New DOMDocument40
    xmlRequestDoc.async = False
    xmlRequestDoc.loadXML strRequestPrefix & strNodePrefix & vlngMessageNo & strRequestTerm
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.loadXML "<RESPONSE/>"
    
    Set xmlRequestNode = xmlRequestDoc.selectSingleNode("REQUEST")
    Set xmlResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
    
    adoCRUD xmlRequestNode, xmlRequestNode, xmlResponseNode
    
    Set xmlMessageNode = xmlResponseNode.selectSingleNode("MESSAGE")
    
    If Not xmlMessageNode Is Nothing Then
    
        If vintMessageField = omMESSAGE_TEXT Then
            adoCRUDGetMessageText = xmlGetAttributeText(xmlMessageNode, "MESSAGETEXT")
        Else
            adoCRUDGetMessageText = xmlGetAttributeText(xmlMessageNode, "MESSAGETYPE")
        End If
    
    End If
    
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlMessageNode = Nothing

    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing

End Function

' ik_20050627_CORE158/1
Private Sub adoCRUDSchemaError( _
    ByVal vstrFunctionName As String, _
    ByVal vstrDescription As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode, _
    ByRef rxmlResponseNode As IXMLDOMElement)

    Dim xmlElem As IXMLDOMElement
    Dim xmlErrNode As IXMLDOMNode
    
    rxmlResponseNode.ownerDocument.loadXML "<RESPONSE/>"
    Set rxmlResponseNode = rxmlResponseNode.ownerDocument.selectSingleNode("RESPONSE")
    rxmlResponseNode.setAttribute "TYPE", "APPERR"
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("ERROR")
    Set xmlErrNode = rxmlResponseNode.appendChild(xmlElem)
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("NUMBER")
    xmlElem.Text = oeSchemaParseError
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SOURCE")
    xmlElem.Text = vstrFunctionName & "." & Err.Source
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("VERSION")
    If Len(App.Comments) > 0 Then
        xmlElem.Text = App.Comments
    Else
        xmlElem.Text = App.Major & "." & App.Major & "." & App.Revision
    End If
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("DESCRIPTION")
    xmlElem.Text = vstrDescription
    xmlErrNode.appendChild xmlElem

    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("REQUESTDOC")
    xmlElem.Text = vxmlRequestNode.ownerDocument.xml
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("REQUESTNODE")
    xmlElem.Text = vxmlRequestNode.xml
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SCHEMAMASTERNODE")
    xmlElem.Text = vxmlSchemaMasterNode.xml
    xmlErrNode.appendChild xmlElem
    
    If Not vxmlSchemaRefNode Is Nothing Then
        Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SCHEMAREFNODE")
        xmlElem.Text = vxmlSchemaRefNode.xml
        xmlErrNode.appendChild xmlElem
    End If

    Err.Raise oeSchemaParseError, "adoCRUDSchemaError", vstrDescription

End Sub
' ik_20050627_CORE158_ends

Private Sub adoCRUDSQLError( _
    ByVal vstrFunctionName As String, _
    ByVal vcmd As ADODB.Command, _
    ByRef rxmlResponseNode As IXMLDOMElement, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlSchemaRefNode As IXMLDOMNode)
    
    Dim param As ADODB.Parameter
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlErrNode As IXMLDOMNode
    Dim xmlSQLNode As IXMLDOMNode
    
    rxmlResponseNode.ownerDocument.loadXML "<RESPONSE/>"
    Set rxmlResponseNode = rxmlResponseNode.ownerDocument.selectSingleNode("RESPONSE")
    rxmlResponseNode.setAttribute "TYPE", "APPERR"
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("ERROR")
    Set xmlErrNode = rxmlResponseNode.appendChild(xmlElem)
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("NUMBER")
    xmlElem.Text = Err.Number
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SOURCE")
    xmlElem.Text = vstrFunctionName & "." & Err.Source
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("VERSION")
    If Len(App.Comments) > 0 Then
        xmlElem.Text = App.Comments
    Else
        xmlElem.Text = App.Major & "." & App.Major & "." & App.Revision
    End If
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("DESCRIPTION")
    xmlElem.Text = Err.Description
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("REQUESTDOC")
    xmlElem.Text = vxmlRequestNode.ownerDocument.xml
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("REQUESTNODE")
    xmlElem.Text = vxmlRequestNode.xml
    xmlErrNode.appendChild xmlElem
    
    Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SCHEMAMASTERNODE")
    xmlElem.Text = vxmlSchemaMasterNode.xml
    xmlErrNode.appendChild xmlElem
    
    If Not vxmlSchemaRefNode Is Nothing Then
        Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SCHEMAREFNODE")
        xmlElem.Text = vxmlSchemaRefNode.xml
        xmlErrNode.appendChild xmlElem
    End If
    
    If Not IsNull(vcmd) Then
    
        Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SQLCOMMAND")
        If vcmd.CommandType = adCmdText Then
            xmlElem.setAttribute "TYPE", "adCmdText"
        ElseIf vcmd.CommandType = adCmdStoredProc Then
            xmlElem.setAttribute "TYPE", "adCmdStoredProc"
        End If
        xmlElem.setAttribute "COMMANDTEXT", vcmd.CommandText
        Set xmlSQLNode = xmlErrNode.appendChild(xmlElem)
    
        For Each param In vcmd.Parameters
            Set xmlElem = rxmlResponseNode.ownerDocument.createElement("SQLPARAM")
            
            xmlElem.setAttribute "NAME", param.Name
            
            Select Case param.Type
            
                Case adBSTR
                    xmlElem.setAttribute "TYPE", "adBSTR"
                    
                Case adBinary
                    xmlElem.setAttribute "TYPE", "adBinary"
                    
                Case adInteger
                    xmlElem.setAttribute "TYPE", "adInteger"
                    
                Case adDouble
                    xmlElem.setAttribute "TYPE", "adDouble"
                    
                Case adDBTimeStamp
                    xmlElem.setAttribute "TYPE", "adDBTimeStamp"
            
            End Select
            
            Select Case param.Direction
                
                Case adParamInput
                    xmlElem.setAttribute "DIRECTION", "adParamInput"
                
                Case adParamInputOutput
                    xmlElem.setAttribute "DIRECTION", "adParamInputOutput"
                
                Case adParamOutput
                    xmlElem.setAttribute "DIRECTION", "adParamOutput"
                
                Case adParamReturnValue
                    xmlElem.setAttribute "DIRECTION", "adParamReturnValue"
                
                Case adParamUnknown
                    xmlElem.setAttribute "DIRECTION", "adParamUnknown"
            
            End Select
            
            If Not IsNull(param.Value) Then
                If param.Type = adBinary Then
                    xmlElem.setAttribute "VALUE", adoBinToString(param.Value, param.Size)
                Else
                    xmlElem.setAttribute "VALUE", param.Value
                End If
            Else
                xmlElem.setAttribute "VALUE", "null"
            End If
            
            xmlSQLNode.appendChild xmlElem
            
        Next

    End If

    Err.Raise Err.Number, "adoCRUDSQLError", Err.Description

End Sub


Public Sub adoCRUDBuildDbConnectionString( _
    Optional ByVal vstrComponentRef As String = "DEFAULT")
On Error GoTo BuildDbConnectionStringVbErr
    Dim objWshShell As Object
    Dim strThisComponentRef As String, _
        strConnection As String, _
        strProvider As String, _
        strRegSection As String, _
        strRetries As String
        
    ' PSC 29/07/2004 BBG1141 - Start
    Dim strUserID As String
    Dim strPassword As String
    Dim strSource As String
    Dim strDesc As String
    Dim lngErrNo As Long
    ' PSC 29/07/2004 BBG1141 - End
        
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
        
        ' PSC 29/07/2004 BBG1141 - Start
        strUserID = ""
        strPassword = ""
        
        On Error Resume Next
        
        strUserID = objWshShell.RegRead(strRegSection & gstrUID_KEY)
        lngErrNo = Err.Number
        strSource = Err.Source
        strDesc = Err.Description
        
        On Error GoTo BuildDbConnectionStringVbErr
        
        If lngErrNo <> 0 And lngErrNo <> glngENTRYNOTFOUND Then
            Err.Raise lngErrNo, strSource, strDesc
        End If
        
        On Error Resume Next
        
        strPassword = objWshShell.RegRead(strRegSection & gstrPASSWORD_KEY)
        lngErrNo = Err.Number
        strSource = Err.Source
        strDesc = Err.Description
        
        On Error GoTo BuildDbConnectionStringVbErr
        
        If lngErrNo <> 0 And lngErrNo <> glngENTRYNOTFOUND Then
            Err.Raise lngErrNo, strSource, strDesc
        End If
              
        strConnection = _
                   "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & gstrSERVER_KEY) & ";" & _
                   "database=" & objWshShell.RegRead(strRegSection & gstrDATABASE_KEY) & ";"
                   
        ' If User Id is present use SQL Server Authentication else
        ' use integrated security
        If Len(Trim$(strUserID)) > 0 Then
            strConnection = strConnection & "UID=" & strUserID & ";pwd=" & strPassword & ";"
        Else
            strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
        End If
        ' PSC 29/07/2004 BBG1141 - End
    End If
           
    gstrDbConnectionString = strConnection
    strRetries = objWshShell.RegRead(strRegSection & gstrRETRIES_KEY)
    If Len(strRetries) > 0 Then
        gintDbRetries = CInt(strRetries)
    End If
    Set objWshShell = Nothing
    putConnectString vstrComponentRef, strConnection
    
'   ik_debug
'   Debug.Print strConnection
'   ik_debug_ends
    
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
        Set gxmldocConnections = New DOMDocument40
        gxmldocConnections.async = False
        Set xmlElem = gxmldocConnections.createElement("CONN_ROOT")
        gxmldocConnections.appendChild xmlElem
    End If
    Set xmlElem = gxmldocConnections.createElement("CONN")
    xmlElem.setAttribute vstrRef, vstrConn
    gxmldocConnections.firstChild.appendChild xmlElem
    
'   ik_debug
'   Debug.Print gxmldocConnections.xml
'   ik_debug_ends

End Sub

Private Function getDbConnectString( _
    Optional ByVal vstrComponentRef As String = "DEFAULT") _
    As String
    If Len(vstrComponentRef) = 0 Then vstrComponentRef = "DEFAULT"
    If gxmldocConnections.selectSingleNode("CONN_ROOT/CONN/@" & vstrComponentRef) Is Nothing _
    Then
        adoCRUDBuildDbConnectionString vstrComponentRef
    End If
    getDbConnectString = _
        gxmldocConnections.selectSingleNode("CONN_ROOT/CONN/@" & vstrComponentRef).Text
    
'   ik_debug
'   Debug.Print getDbConnectString
'   ik_debug_ends

End Function

Private Function getDbRetries() As Integer
    getDbRetries = gintDbRetries
End Function

Private Function getDbProvider() As DBPROVIDER
    getDbProvider = genumDbProvider
End Function

Public Sub adoCRUDInitComboFile()
    
'   ik_debug
'   App.LogEvent "gxmlCRUDCombos - created", vbLogEventTypeInformation
    
    If gxmlCRUDCombos Is Nothing Then
        Set gxmlCRUDCombos = New DOMDocument40
        gxmlCRUDCombos.async = False
        gxmlCRUDCombos.appendChild gxmlCRUDCombos.createElement("COMBOLIST")
    End If
    
End Sub

Public Sub adoCRUDLoadSchema()
    Dim xmlNode As IXMLDOMNode
    Dim strFileSpec As String
    
'   ik_debug
'   App.LogEvent "adoCRUDLoadSchema", vbLogEventTypeInformation
    
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\" & App.Title & ".xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    Set gxmldocBaseSchema = New DOMDocument40
    gxmldocBaseSchema.async = False
    gxmldocBaseSchema.setProperty "NewParser", True
    gxmldocBaseSchema.validateOnParse = False
    gxmldocBaseSchema.Load strFileSpec
    geDateStyle = e_dateUK
    gstrBooleanTrue = "1"
    gstrBooleanFalse = "0"
    If gxmldocBaseSchema.parseError.errorCode = 0 Then
        Set xmlNode = gxmldocBaseSchema.selectSingleNode("OM_SCHEMA")
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

Private Function adoCRUDGetNamedSchema(ByVal vstrSchemaName As String) As DOMDocument40
    
    Dim xmlDoc As DOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim strFileSpec As String
    
' CORE249
'    geDateStyle = e_dateUK
'    gstrBooleanTrue = "1"
'    gstrBooleanFalse = "0"
' CORE249 ends
        
    If gcolNamedSchemas Is Nothing Then
        Set gcolNamedSchemas = New Collection
    End If
    
    On Error Resume Next
    Set xmlDoc = gcolNamedSchemas.Item(vstrSchemaName)
    On Error GoTo 0
    
    If xmlDoc Is Nothing Then
        
        ' pick up XML map from "...\Omiga 4\XML" directory
        ' Only do the subsitution once to change DLL -> XML
        strFileSpec = App.Path & "\" & vstrSchemaName & ".xml"
        strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
        Set xmlDoc = New DOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
        xmlDoc.validateOnParse = False
        xmlDoc.Load strFileSpec
        
        If xmlDoc.parseError.errorCode = 0 Then
        
            gcolNamedSchemas.Add xmlDoc, vstrSchemaName
'   ik_debug
'          App.LogEvent "adoLoadNamedSchema, gcolNamedSchemas.Count: " & gcolNamedSchemas.Count, vbLogEventTypeInformation
'   ik_debug_ends
        End If
    
    End If
    
' CORE249
'    If xmlDoc.parseError.errorCode = 0 Then
'        Set xmlNode = xmlDoc.selectSingleNode("OM_SCHEMA")
'        If Not xmlNode Is Nothing Then
'            If xmlGetAttributeText(xmlNode, "DATEFORMAT") = gstrISO8601 Then
'                geDateStyle = e_dateISO8601
'            End If
'            If xmlAttributeValueExists(xmlNode, "BOOLEANTRUE") Then
'                gstrBooleanTrue = UCase(xmlGetAttributeText(xmlNode, "BOOLEANTRUE"))
'            End If
'            If xmlAttributeValueExists(xmlNode, "BOOLEANFALSE") Then
'                gstrBooleanFalse = UCase(xmlGetAttributeText(xmlNode, "BOOLEANFALSE"))
'            End If
'            Set xmlNode = Nothing
'        End If
'    End If
' CORE249 ends
    
    Set adoCRUDGetNamedSchema = xmlDoc

End Function

Private Function adoGetSchema(ByVal vstrSchemaName As String) As IXMLDOMNode
    
    Dim strPattern As String
    If gxmldocBaseSchema Is Nothing Then
        errThrowError _
            "adoGetSchema", _
            oeSchemaNotLoaded
    End If
    If gxmldocBaseSchema.parseError.errorCode <> 0 Then
        errThrowError _
            "adoGetSchema", _
            oeSchemaParseError, _
            gxmldocBaseSchema.parseError.reason
    End If
    strPattern = "//" & vstrSchemaName & "[@DATASRCE]"
    Set adoGetSchema = gxmldocBaseSchema.selectSingleNode(strPattern)
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

Private Function getOmigaNumberForDatabaseError(ByVal strErrDesc As String) As Long
'-----------------------------------------------------------------------------------
'Description : Find the omiga equivalent number for a database error. This is used
'              to trap specific errors. Add to the list below, if you want trap a
'              a new error
'Pass        : strErrDesc : Description of the Error Message (from database )
'------------------------------------------------------------------------------------
On Error GoTo GetOmigaNumberForDatabaseErrorVbErr
    Const cstrFunctioName = "getOmigaNumberForDatabaseError"
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
    getOmigaNumberForDatabaseError = lngOmigaErrorNo
    Exit Function
GetOmigaNumberForDatabaseErrorVbErr:
    
    Err.Raise Err.Number, cstrFunctioName, Err.Description
End Function

' AS 31/05/01 CC012 End


' private methods ======================================================================
Private Function IsComboLookUpRequired( _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode) _
    As Boolean

    Dim xmlNode As IXMLDOMNode
    
    IsComboLookUpRequired = xmlGetAttributeAsBoolean(vxmlSchemaMasterNode, "_COMBOLOOKUP_")
    If IsComboLookUpRequired = False Then
        IsComboLookUpRequired = xmlGetAttributeAsBoolean(vxmlSchemaMasterNode, "COMBOLOOKUP")
    End If
    
    If IsComboLookUpRequired = False Then
        
        IsComboLookUpRequired = xmlGetAttributeAsBoolean(vxmlRequestNode, "_COMBOLOOKUP_")
        If IsComboLookUpRequired = False Then
            IsComboLookUpRequired = xmlGetAttributeAsBoolean(vxmlRequestNode, "COMBOLOOKUP")
        End If
        
        If IsComboLookUpRequired = False Then
            Set xmlNode = vxmlRequestNode.parentNode
            Do While Not xmlNode Is Nothing
                If xmlNode.nodeType <> NODE_ELEMENT Then
                    Exit Do
                Else
                    If xmlNode.nodeName = "OPERATION" Or xmlNode.nodeName = "REQUEST" Then
                        IsComboLookUpRequired = xmlGetAttributeAsBoolean(xmlNode, "_COMBOLOOKUP_")
                        If IsComboLookUpRequired = False Then
                            IsComboLookUpRequired = xmlGetAttributeAsBoolean(xmlNode, "COMBOLOOKUP")
                        End If
                    End If
                End If
                If IsComboLookUpRequired = True Then
                    Exit Do
                Else
                    Set xmlNode = xmlNode.parentNode
                End If
            Loop
        End If
    
    End If

End Function

Private Function IsComboTypeLookUpRequired( _
    ByVal vxmlSchemaMasterNode As IXMLDOMNode, _
    ByVal vxmlRequestNode As IXMLDOMNode) _
    As COMBOVALIDATIONLOOKUPTYPE

    Dim xmlNode As IXMLDOMNode
    
    IsComboTypeLookUpRequired = e_cvluNone
    
    If xmlGetAttributeText(vxmlSchemaMasterNode, "COMBOTYPELOOKUP") = "EX" Then
        IsComboTypeLookUpRequired = e_cvluExtended
    ElseIf xmlGetAttributeAsBoolean(vxmlSchemaMasterNode, "COMBOTYPELOOKUP") = True Then
        IsComboTypeLookUpRequired = e_cvluPipeDelim
    End If
    
    If IsComboTypeLookUpRequired = e_cvluNone Then
        
        If xmlGetAttributeText(vxmlRequestNode, "COMBOTYPELOOKUP") = "EX" Then
            IsComboTypeLookUpRequired = e_cvluExtended
        ElseIf xmlGetAttributeAsBoolean(vxmlRequestNode, "COMBOTYPELOOKUP") = True Then
            IsComboTypeLookUpRequired = e_cvluPipeDelim
        End If
        
        If IsComboTypeLookUpRequired = e_cvluNone Then
            Set xmlNode = vxmlRequestNode.parentNode
            Do While Not xmlNode Is Nothing
                If xmlNode.nodeType <> NODE_ELEMENT Then
                    Exit Do
                Else
                    If xmlNode.nodeName = "OPERATION" Or xmlNode.nodeName = "REQUEST" Then
                        If xmlGetAttributeText(xmlNode, "COMBOTYPELOOKUP") = "EX" Then
                            IsComboTypeLookUpRequired = e_cvluExtended
                        ElseIf xmlGetAttributeAsBoolean(xmlNode, "COMBOTYPELOOKUP") = True Then
                            IsComboTypeLookUpRequired = e_cvluPipeDelim
                        End If
                    End If
                End If
                If IsComboTypeLookUpRequired <> e_cvluNone Then
                    Exit Do
                Else
                    Set xmlNode = xmlNode.parentNode
                End If
            Loop
        End If
    
    End If

End Function

Private Function FormatNow() As String

    Dim dtNow As Date
    dtNow = Now()
    
    FormatNow = CStr(Day(dtNow)) & "/" & CStr(Month(dtNow)) & "/" & CStr(Year(dtNow)) & _
                " " & CStr(Hour(dtNow)) & ":" & CStr(Minute(dtNow)) & ":" & CStr(Second(dtNow))

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
            Select Case getDbProvider()
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

Private Function GuidStringToByteArray(ByVal strGuid As String) As Byte()
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

Public Sub adoCRUDSaveRequest(ByVal vxmlDoc As DOMDocument40, ByVal vxmlRequestNode As IXMLDOMNode)

    adoCRUDSaveDoc vxmlDoc, vxmlRequestNode, "_Request"

End Sub

Public Sub adoCRUDSaveResponse(ByVal vxmlDoc As DOMDocument40, ByVal vxmlRequestNode As IXMLDOMNode)

    adoCRUDSaveDoc vxmlDoc, vxmlRequestNode, "_Response"

End Sub

Private Sub adoCRUDSaveDoc(ByVal vxmlDoc As DOMDocument40, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vstrSuffix As String)

    Dim strRegValue As String, _
        strTempName As String, _
        strRef As String, _
        strTraceDateTime As String

    Dim dtStartTime As Date
    
    Dim fso As Scripting.FileSystemObject
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    dtStartTime = Now
        
    'CORE238 - trace file enhancements
    strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\CRUDTRACE")
    
    If Len(strRegValue) = 0 Then
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\")
    End If
    
    ' ik_20050804_CORE172
    If strRegValue = "1" Or strRegValue = "2" Then
    ' ik_20050804_CORE172_ends
    
        ' ik_20060315_CORE250
        strRegValue = ""
        
        'CORE238 - trace file enhancements
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\CRUDTRACEFOLDER")
        If Len(strRegValue) = 0 Then
            strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\FOLDER")
        End If
        If Len(strRegValue) = 0 Then
            App.LogEvent App.Title & "tracing - no trace folder specified", vbLogEventTypeError
        Else
            Set fso = New Scripting.FileSystemObject
            If Not fso.FolderExists(strRegValue) Then
                fso.CreateFolder strRegValue
            End If
            
            If Not fso.FolderExists(strRegValue) Then
                App.LogEvent App.Title & "tracing - cannot create trace folder: " & strRegValue, vbLogEventTypeError
            Else
            
                If xmlAttributeValueExists(vxmlRequestNode, "TRACEREF") Then
                    strRef = "_" & xmlGetAttributeText(vxmlRequestNode, "TRACEREF") & "_"
                ElseIf xmlAttributeValueExists(vxmlRequestNode, "USERID") Then
                    strRef = "_" & xmlGetAttributeText(vxmlRequestNode, "USERID") & "_"
                Else
                    strRef = "_"
                End If
        
                'CORE238 - trace file enhancements
                strTraceDateTime = Format(Now(), "yyyymmdd_hhmmss")
                
                If strTraceDateTime = gstrTraceDateTime Then
                    gintTraceDateTimeQ = gintTraceDateTimeQ + 1
                Else
                    gstrTraceDateTime = strTraceDateTime
                    gintTraceDateTimeQ = 1
                End If
            
                strTempName = _
                    strRegValue & "\" & "omCRUDTrace" & strRef & _
                    strTraceDateTime & "_" & Right(CStr(gintTraceDateTimeQ + 1000), 3) & vstrSuffix & ".xml"
                'CORE238 - trace file enhancement ends
                
                vxmlDoc.Save strTempName
                    
            End If
        End If
    End If
    On Error GoTo 0
    Set fso = Nothing
    Set objWshShell = Nothing

End Sub


Private Function Encrypt(ByVal vstrIn As String) As String

    Dim lngInputLen     As Long
    Dim lngChar         As Long
    Dim intIncrement    As Integer
    Dim intChar1        As Integer
    Dim intChar2        As Integer
    Dim strOut          As String
    
    lngInputLen = Len(vstrIn)
    
    ' initialise the random-number generator
    Rnd -1
    Randomize 1
    
    For lngChar = 1 To lngInputLen
    
        intIncrement = Int(Rnd() * cint_MAX) + cint_MIN
        intChar1 = Asc(Mid(vstrIn, lngChar, 1))
        
        intChar2 = intChar1 + intIncrement
        If intChar2 >= Asc("~") Then
            ' new character exceeds ascii 127 ('~'), thus deduct increment
            ' precede new character with "~" in output string so that Decrypt
            ' knows that the increment should be added and not deducted
            strOut = strOut & "~"
            intChar2 = intChar1 - intIncrement
        End If
        strOut = strOut & Chr(intChar2)
        
    Next

    Encrypt = strOut
        
End Function

Private Function Decrypt(ByVal vstrIn As String) As String
    Dim lngInputLen     As Long
    Dim lngChar         As Long
    Dim intIncrement    As Integer
    Dim intChar1        As Integer
    Dim intChar2        As Integer
    Dim strOut          As String
    Dim blnAdd          As Boolean
    
    lngInputLen = Len(vstrIn)
    
    ' initialise the random-number generator
    Rnd -1
    Randomize 1
    
    lngChar = 1
    
    Do While lngChar <= lngInputLen
        
        intChar1 = Asc(Mid(vstrIn, lngChar, 1))
        If intChar1 = Asc("~") Then
            lngChar = lngChar + 1
            blnAdd = False
        Else
            blnAdd = True
        End If
                
        intIncrement = Int(Rnd() * cint_MAX) + cint_MIN
        intChar1 = Asc(Mid(vstrIn, lngChar, 1))
        
        If blnAdd Then
            intChar2 = intChar1 - intIncrement
        Else
            intChar2 = intChar1 + intIncrement
        End If
            
        strOut = strOut & Chr(intChar2)
        
        lngChar = lngChar + 1
    Loop

    Decrypt = strOut
End Function

' CORE238
Private Sub DeStream( _
    ByVal vstrAttribName As String, _
    ByVal vxmlOutElem As IXMLDOMElement, _
    ByVal vstrIn As String)

    Dim xd As MSXML2.DOMDocument40
    Set xd = New DOMDocument40
    xd.async = False
    xd.setProperty "NewParser", True
    'IK CORE289 08/08/2006
    xd.validateOnParse = False
    
    xd.loadXML vstrIn
    
    If xd.parseError.errorCode = 0 Then
        vxmlOutElem.appendChild xd.documentElement.cloneNode(True)
    Else
        vxmlOutElem.setAttribute vstrAttribName, vstrIn
    End If
    
    Set xd = Nothing

End Sub


' ik_20050627_CORE158/3
Private Sub getDereferencedSchema( _
    ByVal vxmlNamedSchemaDoc As DOMDocument40, _
    ByVal vstrEntityRef As String, _
    ByRef rxmlTargetNode As IXMLDOMNode)
    
    Dim strXPath As String
    
    ' CORE308
    ' if ENTITY_REF value has base: prefix, de-reference against base schema only
    Dim strEntityRef As String
    
    If Left(vstrEntityRef, 5) = "base:" Then
        strEntityRef = Right(vstrEntityRef, Len(vstrEntityRef) - 5)
    Else
        strEntityRef = vstrEntityRef
    End If
    
    ' ik_20060213_CORE244
    strXPath = "OM_SCHEMA/ENTITY" & _
                "[@NAME='" & strEntityRef & "'" & _
                " or @ENTITY_REF='" & strEntityRef & "']"
    
    If Not vxmlNamedSchemaDoc Is Nothing Then
        Set rxmlTargetNode = vxmlNamedSchemaDoc.selectSingleNode(strXPath)
    End If
        
    If Left(vstrEntityRef, 5) <> "base:" Then
        If rxmlTargetNode Is Nothing Then
            Set rxmlTargetNode = gxmldocBaseSchema.selectSingleNode(strXPath)
        End If
    End If
    
    If rxmlTargetNode Is Nothing Then
        ' error handled in calling function
        Exit Sub
    End If
    
    deRefSchemaEntity vxmlNamedSchemaDoc, rxmlTargetNode

End Sub

Private Sub deRefSchemaEntity(ByVal vxmlNamedSchemaDoc As DOMDocument40, ByVal vxmlNewSchemaNode As IXMLDOMNode)
    
    Dim xmlNewSchemaChildNode As IXMLDOMNode
    Dim xmlSchemaEntityRefNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeLiat As IXMLDOMNodeList
    
    Dim strXPath As String
    
    ' CORE308
    Dim strRawEntityRef As String
    Dim strEntityRef As String
    
    Set xmlNodeLiat = vxmlNewSchemaNode.selectNodes("ATTRIBUTE[@NAME][@ENTITY_REF]")
    
    Do While xmlNodeLiat.length <> 0
    
        deRefSchemaAttributeNodes vxmlNamedSchemaDoc, xmlNodeLiat
        
        Set xmlNodeLiat = vxmlNewSchemaNode.selectNodes("ATTRIBUTE[@NAME][@ENTITY_REF]")
        
    Loop

    If xmlAttributeValueExists(vxmlNewSchemaNode, "ENTITY_REF") Then
    
        ' CORE308
        ' if ENTITY_REF value has base: prefix, de-reference against base schema only
        strRawEntityRef = xmlGetAttributeText(vxmlNewSchemaNode, "ENTITY_REF")
        If Left(strRawEntityRef, 5) = "base:" Then
            strEntityRef = Right(strRawEntityRef, Len(strRawEntityRef) - 5)
        Else
            strEntityRef = strRawEntityRef
        End If
    
        strXPath = "OM_SCHEMA/ENTITY[@NAME='" & strEntityRef & "']"
'ik_debug_20050914
'       Debug.Print strXPath
'ik_debug_20050914_ends
        
        ' CORE308
        If Left(strRawEntityRef, 5) <> "base:" Then
            If Not vxmlNamedSchemaDoc Is Nothing Then
                Set xmlSchemaEntityRefNode = vxmlNamedSchemaDoc.selectSingleNode(strXPath)
            End If
        End If
        
        If xmlSchemaEntityRefNode Is Nothing Then
            Set xmlSchemaEntityRefNode = gxmldocBaseSchema.selectSingleNode(strXPath)
        End If
        
        If xmlSchemaEntityRefNode Is Nothing Then
            Err.Raise _
                oeXMLMissingElement, _
                "adoCRUD", _
                "invalid schema reference: " & vxmlNewSchemaNode.xml
        End If
    
        vxmlNewSchemaNode.Attributes.removeNamedItem ("ENTITY_REF")
        
        For Each xmlAttrib In xmlSchemaEntityRefNode.Attributes
        
            If vxmlNewSchemaNode.Attributes.getNamedItem(xmlAttrib.Name) Is Nothing Then
                vxmlNewSchemaNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            End If
            
        Next
            
        For Each xmlNode In xmlSchemaEntityRefNode.childNodes
        
            ' ik_CORE190_250923
            If xmlNode.nodeType = NODE_ELEMENT Then
            
                Set xmlNewSchemaChildNode = Nothing
            
                If xmlAttributeValueExists(xmlNode, "NAME") Then
            
                    strXPath = xmlNode.nodeName & "[@NAME='" & xmlGetAttributeText(xmlNode, "NAME") & "']"
                    'ik_debug_20050914
                    'Debug.Print strXPath
                    'ik_debug_20050914_ends
                    
                    Set xmlNewSchemaChildNode = vxmlNewSchemaNode.selectSingleNode(strXPath)
                    
                End If
            
                If Not xmlNewSchemaChildNode Is Nothing Then
            
                    For Each xmlAttrib In xmlNode.Attributes
                    
                        If xmlNewSchemaChildNode.Attributes.getNamedItem(xmlAttrib.Name) Is Nothing Then
                            xmlNewSchemaChildNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
                        End If
                        
                    Next
                    
                Else
                
                    vxmlNewSchemaNode.appendChild xmlNode.cloneNode(True)
                
                End If
            
            ' ik_CORE190_250923
            End If
        
        Next
        
        deRefSchemaEntity vxmlNamedSchemaDoc, vxmlNewSchemaNode
        
    Else
    
        For Each xmlNode In vxmlNewSchemaNode.selectNodes("ENTITY")
            deRefSchemaEntity vxmlNamedSchemaDoc, xmlNode
        Next
    
    End If
    
    Set xmlNodeLiat = Nothing

End Sub

Private Sub deRefSchemaAttributeNodes(ByVal vxmlNamedSchemaDoc As DOMDocument40, ByVal vxmlNewSchemaAttNodes As IXMLDOMNodeList)
    
    Dim xmlSchemaEntityRefNode As IXMLDOMNode
    Dim xmlSchemaAttRefNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlNode As IXMLDOMNode
    
    Dim strNodeName As String, _
        strXPath As String
    
    For Each xmlNode In vxmlNewSchemaAttNodes
    
        If xmlGetAttributeText(xmlNode, "ENTITY_REF") <> strNodeName Then
        
            strNodeName = xmlGetAttributeText(xmlNode, "ENTITY_REF")
    
            strXPath = "OM_SCHEMA/ENTITY[@NAME='" & strNodeName & "']"
    
' ik_20050627_CORE158/4
'            If Not vxmlNamedSchemaDoc Is Nothing Then
'                Set xmlSchemaEntityRefNode = vxmlNamedSchemaDoc.selectSingleNode(strXPath)
'            End If
'
'            If xmlSchemaEntityRefNode Is Nothing Then
'                Set xmlSchemaEntityRefNode = gxmldocBaseSchema.selectSingleNode(strXPath)
'            End If
            Set xmlSchemaEntityRefNode = gxmldocBaseSchema.selectSingleNode(strXPath)
' ik_20050627_CORE158_ends
        
            If xmlSchemaEntityRefNode Is Nothing Then
                Err.Raise _
                    oeXMLMissingElement, _
                    "adoCRUD", _
                    "invalid schema reference: " & xmlNode.xml
            End If
            
        End If
            
        strXPath = "ATTRIBUTE[@NAME='" & xmlGetAttributeText(xmlNode, "NAME") & "']"
        
        Set xmlSchemaAttRefNode = xmlSchemaEntityRefNode.selectSingleNode(strXPath)
    
        If xmlSchemaAttRefNode Is Nothing Then
            Err.Raise _
                oeXMLMissingElement, _
                "adoCRUD", _
                "invalid schema reference: " & xmlNode.xml
        End If
        
        xmlNode.Attributes.removeNamedItem ("ENTITY_REF")
                
        For Each xmlAttrib In xmlSchemaAttRefNode.Attributes
        
            If xmlNode.Attributes.getNamedItem(xmlAttrib.Name) Is Nothing Then
                xmlNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            End If
            
        Next
    
    Next

End Sub
' ik_20050627_CORE158_ends
        
''IK 11/11/2005 CORE215
'search for primary keys in parent hierarchy
Private Function getKeyValueFromParentHierarchy( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As IXMLDOMAttribute
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Set xmlNode = vxmlRequestNode.parentNode
    
    Do While xmlAttrib Is Nothing
        If xmlNode.nodeName = "REQUEST" Then
            Exit Do
        End If
        If xmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            Set xmlNode = xmlNode.parentNode
        Else
            Set xmlAttrib = xmlNode.Attributes.getNamedItem(vstrAttribName)
        End If
    Loop
    
    Set getKeyValueFromParentHierarchy = xmlAttrib

End Function

' CORE249
Private Function getBooleanTrue(ByVal vxmlSchemaNode As IXMLDOMNode) As String
    If Not vxmlSchemaNode.ownerDocument.documentElement.Attributes.getNamedItem("BOOLEANTRUE") Is Nothing Then
        getBooleanTrue = vxmlSchemaNode.ownerDocument.documentElement.getAttribute("BOOLEANTRUE")
    Else
        getBooleanTrue = gstrBooleanTrue
    End If
End Function

Private Function getBooleanFalse(ByVal vxmlSchemaNode As IXMLDOMNode) As String
    If Not vxmlSchemaNode.ownerDocument.documentElement.Attributes.getNamedItem("BOOLEANFALSE") Is Nothing Then
        getBooleanFalse = vxmlSchemaNode.ownerDocument.documentElement.getAttribute("BOOLEANFALSE")
    Else
        getBooleanFalse = gstrBooleanFalse
    End If

End Function

Private Function getDateStyle(ByVal vxmlSchemaNode As IXMLDOMNode) As DATESTYLE
    If Not vxmlSchemaNode.ownerDocument.documentElement.Attributes.getNamedItem("DATEFORMAT") Is Nothing Then
        If vxmlSchemaNode.ownerDocument.documentElement.getAttribute("DATEFORMAT") = gstrISO8601 Then
            getDateStyle = e_dateISO8601
        Else
            getDateStyle = e_dateUK
        End If
    Else
        getDateStyle = geDateStyle
    End If
End Function
' CORE249 ends

