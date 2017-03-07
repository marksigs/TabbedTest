VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "DBXMLWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Workfile:
' Copyright:    Copyright © 1999 Marlborough Stirling
'
'Description:
'   Uses an ADO Connection, an input XML stream of initial
'   data and an XML processing map to produce an output XML
'   stream. The output stream can be used for whatever.


' Dependencies:
'   ADO, XML, MTS

' Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'MH     23/02/00    Created
'LD     07/11/00    Explicity close recordsets
'LD     07/11/00    Explicity destroy command objects
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'------------------------------------------------------------------------------------------

Implements ObjectControl
Implements IDBXMLWizard
    Dim msngSQLTimer As Single
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As New ErrAssist
Private m_objRulesAssist As MapRulesAssist

Private m_conOmiga4 As Connection
    
Public Function CreateXML(ByVal vstrXMLIn As String, ByVal vstrXMLMap As String, _
                    ByVal vstrADOConnectionString As String)
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".CreateXML")
    Const strFunctionName As String = "CreateXML"
    On Error GoTo CreateXMLVBErr
' description:
'    Given a map and an output document it will populate the output
'    document based on the rules defined in the map and on the data
'    available. The map may also include rules for retrieving additional
'    data.

' pass:
'   vstrXMLIn           The input parameters
'   vstrxmlMap          The map used to define the output
'   vcolDataCache       The base data of recordset used to populate the map
'   vstrADOConnectionString  An ADO connection string
    
    Dim xmlMap As FreeThreadedDOMDocument40
    Dim xmlIn As FreeThreadedDOMDocument40
    
    Set m_conOmiga4 = New Connection
    m_conOmiga4.open vstrADOConnectionString
    
    CreateXML = IDBXMLWizard_CreateXML(xmlIn, xmlMap, m_conOmiga4).XML
    Exit Function

CreateXMLVBErr:
          
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Sub ProcessMap(ByVal vxmlOut As IXMLDOMNode, _
            ByVal vxmlMap As IXMLDOMNode, _
            ByVal vcolDataCache As DataCache)
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".ProcessMap")
' description:
' description:
'    Given a map and an output document it will populate the output
'    document based on the rules defined in the map and on the data
'    available. The map may also include rules for retrieving additional
'    data.

' The code inside the loop is identical to the ProcessMap code
'
' pass:
'   vxmlOut             The Node that will receive the data
'   vxmlMap             The map used to define the output
'   vcolDataCache       The base data of recordset used to populate the map
'
    Const strFunctionName As String = "ProcessMap"
    On Error GoTo ProcessMapVbErr
    
    Dim xmlTYPEattrib As IXMLDOMAttribute
    Dim xmlCurrentOutNode As IXMLDOMNode
    Dim xmlChildMapNode As IXMLDOMNode
          
    '---------------------------------------------------------------------
    'There are several node types in the map as follows:
    '  No attributes - required in the output, could be anythinng
    '  TYPE="Table" - required in the output, has specific data needs
    '  TYPE="Transparent" - contents in the output, but not this node
    '     (used  for logical JOINS)
    '  TYPE="Cache" - just load the data
    '  No TYPE but has other attributes like SOURCE and RULE - data
    '   Type="#Constant" - like a Datanode, but stores parameters
    '---------------------------------------------------------------------
    
    For Each xmlChildMapNode In vxmlMap.childNodes
    
        If TypeOf xmlChildMapNode Is IXMLDOMElement Then
            Debug.Assert xmlChildMapNode.nodeType <> NODE_COMMENT
            Set xmlTYPEattrib = xmlChildMapNode.Attributes.getNamedItem("TYPE")
                
            '  2. - 5.  Has attributes, may need special processing
            If xmlTYPEattrib Is Nothing Then
                '4. No TYPE, but may have other attributes. May be a data field.
                ProcessDataNode vxmlOut, xmlChildMapNode, vcolDataCache, True
            Else
                'Nodes are listed in most likely occurrence order
                Select Case UCase$(xmlTYPEattrib.Value)
                Case "TABLE"
                '  A data source required in the output, has specific data needs
                    ProcessTableNode vxmlOut, xmlChildMapNode, vcolDataCache, False
                Case "TRANSPARENT"
                '  contents in the output, but not this node
                    ProcessTableNode vxmlOut, xmlChildMapNode, vcolDataCache, True
                Case "CACHE"
                '  Not in the output, use to preload data for the current root node
                    ProcessCacheNode xmlChildMapNode, vcolDataCache
                Case "TEXT"
                    ProcessDataNode vxmlOut, xmlChildMapNode, vcolDataCache, False
                Case "#CONSTANT"
                '  Used to define constants
                    ProcessConstantNode xmlChildMapNode, vcolDataCache
                Case "IGNORE"
                '  put in for debugging.
                Case Else
                    'Unrecognised node - error condition
                    Err.Raise oeInvalidParameter, TypeName(Me), "Node """ & xmlChildMapNode.nodeName & _
                        """ has an unrecognised attribute TYPE=""" & xmlTYPEattrib.Text & """"
                End Select
            
            End If
            
        Else
            ' COMMENT nodes are stripped out since they are not data.
            If xmlChildMapNode.nodeType = NODE_COMMENT Then
                'Do Nothing
            Else
                'don't know what it is so let it through without examining it further
                vxmlOut.appendChild xmlChildMapNode.cloneNode(True) 'Recursive clone
            End If
        End If
    Next xmlChildMapNode
    
    Exit Sub

ProcessMapVbErr:
     
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description  '& vbCrLf & NodeSummary(vxmlMap) & vbCrLf

End Sub

Private Sub ProcessDataNode(ByVal vxmlOut As IXMLDOMNode, _
            ByVal vxmlMap As IXMLDOMNode, _
            ByVal vcolDataCache As DataCache, ByVal vblnElementRequired As Boolean)
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".ProcessDataNode")
' description:
'    Data nodes have the following possible attributes:
'      FORMAT - a standard VB format field (optional)
'      LENGTH - the length of the final data (optional)
'      OPTIONAL - the field optionality. Default is FALSE (optional)
'      SOURCE - The cached data is to be used
'      SELECT - Select other than the current row - SQL type SELECT
'      RULE - The name of a rule to process the data
'
' pass:
'   vxmlOut             The Node that will receive the data
'   vxmlMap             The map used to define the output
'   vcolDataCache       The base data of recordset used to populate the map
'   vblnElementRequired Add an element node instead of a text node

    Const strFunctionName As String = "ProcessDataNode"
    On Error GoTo ProcessDataNodeVbErr
    
    Dim xmlFORMATattrib As IXMLDOMAttribute
    Dim xmlSOURCEattrib As IXMLDOMAttribute
    Dim xmlSELECTattrib As IXMLDOMAttribute
    Dim xmlRULEattrib As IXMLDOMAttribute
    Dim xmlLENGTHattrib As IXMLDOMAttribute
    Dim xmlOPTIONALattrib As IXMLDOMAttribute
    Dim xmlCurrentOutNode As IXMLDOMNode
      
    Dim varValue As Variant
    
    With vxmlMap.Attributes
        Set xmlSOURCEattrib = .getNamedItem("SOURCE")
        Set xmlSELECTattrib = .getNamedItem("SELECT")
        Set xmlFORMATattrib = .getNamedItem("FORMAT")
        Set xmlRULEattrib = .getNamedItem("RULE")
        Set xmlLENGTHattrib = .getNamedItem("LENGTH")
        Set xmlOPTIONALattrib = .getNamedItem("OPTIONAL")
    End With
    
    'SOURCE is not always required
    
    If Not xmlSOURCEattrib Is Nothing Then
        
        'If a regular expression has been given, use it to select a row.
        If xmlSELECTattrib Is Nothing Then
            varValue = vcolDataCache.GetValue(xmlSOURCEattrib.Value)
        Else
            varValue = vcolDataCache.GetValueUsingFilter(xmlSOURCEattrib.Value, xmlSELECTattrib.Value)
        End If
        
        'For an optional node do nothing if the data is missing.
        'There is no point checking this for rules which don't reference SOURCE
        'directly since they should always return data.
        'In reality, only Text data nodes should use this, but by the time its' checked
        'varValue would have been potentially modified
        
        If Not xmlOPTIONALattrib Is Nothing Then
            If CBool(xmlOPTIONALattrib.Value) Then
                 If IsNull(varValue) Then Exit Sub
                 If varValue = vbNullString Then Exit Sub
            End If
        End If
        
    Else
        If xmlRULEattrib Is Nothing Then
            'if there is no SOURCE and no RULE then it is not a recognised node and should
            ' be passed through unaltered.
           
            'NB: The code is place here mainly as an optimisation trick. Most of the
            'elements in map will be Data nodes. Only the minority will be hard-coded
            'nodes that are required in the output.
            '
            'These special names are allowed as a compromise - if they have no kids
            '   CRLF - Insert a CRLF
            '   TAB  - Insert a Tab
            '   SP   - Insert a space
            If vxmlMap.hasChildNodes Then
                Set xmlCurrentOutNode = vxmlMap.cloneNode(False) 'Not a Recursive clone
                vxmlOut.appendChild xmlCurrentOutNode
                ProcessMap xmlCurrentOutNode, vxmlMap, vcolDataCache
            Else
                Select Case vxmlMap.nodeName
                Case "CRLF"
                     vxmlOut.appendChild vxmlOut.ownerDocument.createTextNode(vbCrLf)
                Case "TAB" ' A virtual tab
                     vxmlOut.appendChild vxmlOut.ownerDocument.createTextNode("    ")
                Case "SP"
                     vxmlOut.appendChild vxmlOut.ownerDocument.createTextNode(" ")
                Case Else
                    Set xmlCurrentOutNode = vxmlMap.cloneNode(False) 'Not a Recursive clone
                    vxmlOut.appendChild xmlCurrentOutNode
                    ProcessMap xmlCurrentOutNode, vxmlMap, vcolDataCache
                End Select
            End If
            
            Exit Sub

        End If
    End If
    
    'Apply any rule that is provided
    If Not xmlRULEattrib Is Nothing Then
        varValue = m_objRulesAssist.EvaluateRule(xmlRULEattrib.Value, varValue, vcolDataCache)
    End If
    
    'Now add the value to the xml
    If vblnElementRequired Then
        Set xmlCurrentOutNode = vxmlOut.ownerDocument.createElement(vxmlMap.nodeName)
   
        'Apply formatting
        If Not xmlFORMATattrib Is Nothing Then
            varValue = Format$(varValue, xmlFORMATattrib.Value)
        End If
        If xmlLENGTHattrib Is Nothing Then
            xmlCurrentOutNode.Text = varValue
        Else
            xmlCurrentOutNode.Text = Left$(varValue, xmlLENGTHattrib.Value)
        End If
        vxmlOut.appendChild xmlCurrentOutNode
    Else
        'Text nodes are more complicated. They can have an optional Flag and can have
        'child nodes. Child nodes are supported in order that text formatting is easier.
        'When used with child nodes it is assumed that the child nodes will reference anything
        'that is required for output.
        If vxmlMap.hasChildNodes Then
            ProcessMap vxmlOut, vxmlMap, vcolDataCache
        Else
            'Apply formatting
            If Not xmlFORMATattrib Is Nothing Then
                varValue = Format$(varValue, xmlFORMATattrib.Value)
            End If
            If xmlLENGTHattrib Is Nothing Then
                Set xmlCurrentOutNode = vxmlOut.ownerDocument.createTextNode(varValue)
            Else
                Set xmlCurrentOutNode = vxmlOut.ownerDocument.createTextNode(Left$(varValue, xmlLENGTHattrib.Value))
            End If
             vxmlOut.appendChild xmlCurrentOutNode
        End If
    End If
       
    Exit Sub

ProcessDataNodeVbErr:
    'If Not vxmlMap Is Nothing And InStr(1, Err.Description, "NodeName", vbTextCompare) = 0 Then TypeName (Me) & " Line=" & Erl & vbCrLf & Err.Description = Err.Description & vbCrLf & " NodeName=""" & vxmlMap.nodeName & """"
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub



Private Sub ProcessConstantNode(ByVal vxmlMap As IXMLDOMNode, _
            ByVal vcolDataCache As DataCache)
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".ProcessConstantNode")
' description:
'    #Constant nodes have the following possible attributes:
'      FORMAT - a standard VB format field (optional)
'      LENGTH - the length of the final data (optional)
'      OPTIONAL - the field optionality. Default is FALSE (optional)
'      SOURCE - The cached data is to be used
'      SELECT - Select other than the current row - SQL type SELECT
'      RULE - The name of a rule to process the data
'      NAME - The name of the constant
'
' pass:
'   vxmlOut             The Node that will receive the data
'   vxmlMap             The map used to define the output
'   vcolDataCache       The base data of recordset used to populate the map
'
    Const strFunctionName As String = "ProcessConstantNode"
    On Error GoTo ProcessConstantNodeVbErr
    
    Dim xmlFORMATattrib As IXMLDOMAttribute
    Dim xmlSOURCEattrib As IXMLDOMAttribute
    Dim xmlSELECTattrib As IXMLDOMAttribute
    Dim xmlRULEattrib As IXMLDOMAttribute
    Dim xmlLENGTHattrib As IXMLDOMAttribute
    Dim xmlNAMEattrib As IXMLDOMAttribute
    
    Dim varValue As Variant
    Dim strName As String
    
    With vxmlMap.Attributes
        Set xmlSOURCEattrib = .getNamedItem("SOURCE")
        Set xmlSELECTattrib = .getNamedItem("SELECT")
        Set xmlFORMATattrib = .getNamedItem("FORMAT")
        Set xmlRULEattrib = .getNamedItem("RULE")
        Set xmlLENGTHattrib = .getNamedItem("LENGTH")
        Set xmlNAMEattrib = .getNamedItem("NAME")
    End With
    
    'SOURCE is not always required
    
    If Not xmlSOURCEattrib Is Nothing Then
        
        'If a regular expression has been given, use it to select a row.
        If xmlSELECTattrib Is Nothing Then
            varValue = vcolDataCache.GetValue(xmlSOURCEattrib.Value)
        Else
            varValue = vcolDataCache.GetValueUsingFilter(xmlSOURCEattrib.Value, xmlSELECTattrib.Value)
        End If
        
    End If
    
    'Apply any rule that is provided
    If Not xmlRULEattrib Is Nothing Then varValue = m_objRulesAssist.EvaluateRule(xmlRULEattrib.Value, varValue, vcolDataCache)
       
    If xmlNAMEattrib Is Nothing Then
        strName = vxmlMap.nodeName
    Else
        strName = xmlNAMEattrib.Value
    End If
    
    'Apply formatting
    If Not xmlFORMATattrib Is Nothing Then
        varValue = Format$(varValue, xmlFORMATattrib.Value)
    End If
    
    If xmlLENGTHattrib Is Nothing Then
        vcolDataCache.AddorReplaceConstant strName, varValue
    Else
        vcolDataCache.AddorReplaceConstant strName, Left$(varValue, xmlLENGTHattrib.Value)
    End If
    
    Exit Sub

ProcessConstantNodeVbErr:
'    If Not vxmlMap Is Nothing Then TypeName (Me) & " Line=" & Erl & vbCrLf & Err.Description = Err.Description & vbCrLf & " TYPE=""#Constant""  NodeName=""" & vxmlMap.nodeName & """"
          
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ProcessTableNode(ByVal vxmlOut As IXMLDOMNode, _
            ByVal vxmlMap As IXMLDOMNode, _
            ByVal vcolDataCache As DataCache, _
            Optional ByVal vblnTransparent As Boolean = False)
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".ProcessTableNode")
' description:
'    Table Nodes can contain the following attributes:
'      TYPE - Either "Table" or "Transparent" - used to get here
'      ALIAS - The alternative name to use for caching - instead of SOURCE.
'      SOURCE - The name of the table
'      SELECT - SQL Select to retrieve the table if it isn't cached
'      MIN - minimum number of instances required - default is 1
'      MAX - maximum number of instances - default is 1.
'
'    If MIN is specified then MAX is assumed to be unlimited.
'    If MAX is specified then MIN is asusmed to be 0.
'
'    Data will be added to the cache if not already present. The child
'    nodes will then be processed recursively.
'
' pass:
'   vxmlOut             The Node that will receive the data
'   vxmlMap             The map used to define the output
'   vcolDataCache       The base data of recordset used to populate the map
'
    Const strFunctionName As String = "ProcessTableNode"
    On Error GoTo ProcessTableNodeVbErr
    
    Dim xmlCurrentOutNode As IXMLDOMElement
    
    Dim xmlSOURCEattrib As IXMLDOMAttribute
    Dim xmlSELECTattrib As IXMLDOMAttribute
    Dim xmlMINattrib As IXMLDOMAttribute
    Dim xmlMAXattrib As IXMLDOMAttribute
    Dim xmlALIASattrib As IXMLDOMAttribute
    
    Dim objIDataCacheItem As IDataCacheItem
    Dim colNewCache As DataCache
    Dim rstData As Recordset
    
    Dim strSource As String
    Dim strSelect As String
   
    Dim lngMin As Long
    Dim lngMax As Long
    Dim lngAdded As Long
    
    
    'Copy the cache so that it can be altered if necessary.
    If m_objContext Is Nothing Then
        Set colNewCache = New DataCache
    Else
        Set colNewCache = m_objContext.CreateInstance(App.Title & ".DataCache")
    End If
        
    colNewCache.Clone vcolDataCache
    
    ' Examine all the attributes and set the default values if necessary
    With vxmlMap.Attributes
        Set xmlSOURCEattrib = .getNamedItem("SOURCE")
        Set xmlALIASattrib = .getNamedItem("ALIAS")
        Set xmlSELECTattrib = .getNamedItem("SELECT")
        Set xmlMINattrib = .getNamedItem("MIN")
        Set xmlMAXattrib = .getNamedItem("MAX")
    End With
    
    ' Look at MIN and MAX
    If xmlMINattrib Is Nothing And xmlMAXattrib Is Nothing Then
        'No attributes so set min and max to the defaults
        lngMin = 1
        lngMax = 1
    ElseIf Not (xmlMINattrib Is Nothing Or xmlMAXattrib Is Nothing) Then
        ' Both present, so use both
        lngMin = xmlMINattrib.Value
        lngMax = xmlMAXattrib.Value
    ElseIf Not xmlMINattrib Is Nothing Then
        'Just min is present. Use Min and set Max to be huge
        lngMin = xmlMINattrib.Value
        lngMax = &H7FFFFFFF   ' - Largest possible long
    Else
        'Just max is present. Set min to be 0.
        lngMin = 0
        lngMax = xmlMAXattrib.Value
    End If
    
    'look at SOURCE and SELECT
    If xmlSELECTattrib Is Nothing Then
        'If no SELECT criteria is given then the table must exist in cache, otherwise
        ' it is an error. Set the recordset for this table to be the default
        'Also store under the alias name if given
               
        colNewCache.SetDefault xmlSOURCEattrib.Value
        If Not xmlALIASattrib Is Nothing Then
            colNewCache.AddorReplaceRecordset xmlALIASattrib.Value, colNewCache.Default
        End If
        
        
    Else
        'If a SELECT criteria is given then a recordset needs to be retrieved.
        'If SOURCE is absent an error will be raised.
         
         'Replace all the field references with data values from the cache
         'strSelect = colNewCache.ReplaceFieldsWithData(xmlSELECTattrib.Value)
         strSource = xmlSOURCEattrib.Value
         'Set rstData = m_conOmiga4.Execute("SELECT * FROM " & strSource & " WHERE " & strSelect)
         Set rstData = GetRecordset(colNewCache, strSource, xmlSELECTattrib.Value)
         If xmlALIASattrib Is Nothing Then
            colNewCache.AddorReplaceRecordset strSource, rstData
        Else
            colNewCache.AddorReplaceRecordset xmlALIASattrib.Value, rstData
        End If
		rstData.Close
		Set rstData = Nothing
    End If
    
    'Finally, the table node is ready to be processed. Any data that is required
    'has been retrieved from the database and is set as the default recordset
    
    'Now it is time to iterate through the recordset
    
    lngAdded = 0
    
    Set objIDataCacheItem = colNewCache.Default
    
    If objIDataCacheItem.HasData Then
        
        objIDataCacheItem.GotoFirstEntry
        
        Do While lngAdded < lngMax And Not objIDataCacheItem.AtEndOfData
            'Need to creaet a container node for the output
            If Not vblnTransparent Then
                'Place the node in the output
                Set xmlCurrentOutNode = vxmlOut.ownerDocument.createElement(vxmlMap.nodeName)
                vxmlOut.appendChild xmlCurrentOutNode
            Else
                'Use the node supplied as the output node
                Set xmlCurrentOutNode = vxmlOut
            End If
            
            'Now process each node on the map in turn recursively
             ProcessMap xmlCurrentOutNode, vxmlMap, colNewCache
            
            lngAdded = lngAdded + 1
            objIDataCacheItem.GotoNextEntry
            
        Loop
        
        'Other things may need to use this table. Reset the cursor
        objIDataCacheItem.GotoFirstEntry
    
    End If
    
    ' Need to create empty nodes until sufficient have been created
    Do While lngAdded < lngMin
        
        'Need to creaet a container node for the output
        If Not vblnTransparent Then
            'Place the node in the output
            Set xmlCurrentOutNode = vxmlOut.ownerDocument.createElement(vxmlMap.nodeName)
            vxmlOut.appendChild xmlCurrentOutNode
        Else
            'Use the node supplied as the output node
            Set xmlCurrentOutNode = vxmlOut
        End If
        
        'Now process each node on the map in turn recursively
         ProcessMap xmlCurrentOutNode, vxmlMap, colNewCache
        lngAdded = lngAdded + 1
    Loop
    
    Exit Sub

ProcessTableNodeVbErr:

    ' Close the recordset if still open
	If Not rstData Is Nothing Then
		If rstData.State = adStateOpen Then
			rstData.Close
		End If
	End If
	Set rstData = Nothing

'    If Not vxmlMap Is Nothing And InStr(1, Err.Description, "NodeName", vbTextCompare) = 0 Then TypeName (Me) & " Line=" & Erl & vbCrLf & Err.Description = Err.Description & vbCrLf & " TYPE=""Table/Transparent""  NodeName=""" & vxmlMap.nodeName & """"
'    If Not strSelect = vbNullString Then TypeName (Me) & " Line=" & Erl & vbCrLf & Err.Description = Err.Description & vbCrLf & _
'        "SQL=""" & "SELECT * FROM " & strSource & " WHERE " & strSelect & """"

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ProcessCacheNode(ByVal vxmlMap As IXMLDOMNode, _
            ByVal vcolDataCache As DataCache)
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".ProcessCacheNode")
' description:
'    Cache Nodes can contain the following attributes:
'      TYPE - "Cache"
'      ALIAS - The alternative name to use for caching - instead of SOURCE.
'      SOURCE - The name of the table
'      SELECT - SQL Select to retrieve the table if it isn't cached
'
'  Cache nodes will add/replace recordsets in the supplied data cache.
' pass:
'   vxmlMap             The map used to define the output
'   vcolDataCache       The base data of recordset used to populate the map
'
    Const strFunctionName As String = "ProcessCacheNode"
    On Error GoTo ProcessCacheNodeVbErr
    
    Dim xmlSOURCEattrib As IXMLDOMAttribute
    Dim xmlSELECTattrib As IXMLDOMAttribute
    Dim xmlALIASattrib As IXMLDOMAttribute
    
    Dim rstData As Recordset
    Dim strSelect As String
    Dim strSource As String
    
    ' Examine all the attributes and set the default values if necessary
    With vxmlMap.Attributes
        Set xmlSOURCEattrib = .getNamedItem("SOURCE")
        Set xmlALIASattrib = .getNamedItem("ALIAS")
        Set xmlSELECTattrib = .getNamedItem("SELECT")
    End With
    
    'If a SELECT criteria is given then a recordset needs to be retrieved.
    'If SOURCE is absent an error will be raised.
     
     'Replace all the field references with data values from the cache
    strSource = xmlSOURCEattrib.Value
    
    If Not xmlSELECTattrib Is Nothing Then
        'strSelect = vcolDataCache.ReplaceFieldsWithData(xmlSELECTattrib.Value)
        'Set rstData = m_conOmiga4.Execute("SELECT * FROM " & strSource & " WHERE " & strSelect)
         Set rstData = GetRecordset(vcolDataCache, strSource, xmlSELECTattrib.Value)
        If xmlALIASattrib Is Nothing Then
            vcolDataCache.AddorReplaceRecordset strSource, rstData, vblnChangeDefault:=False
        Else
            vcolDataCache.AddorReplaceRecordset xmlALIASattrib.Value, rstData, vblnChangeDefault:=False
        End If
		rstData.Close
		Set rstData = Nothing
    Else
        vcolDataCache.AddAlias strSource, xmlALIASattrib.Value
    End If
    
    Exit Sub

ProcessCacheNodeVbErr:

    ' Close the recordset if still open
	If Not rstData Is Nothing Then
		If rstData.State = adStateOpen Then
			rstData.Close
		End If
	End If
	Set rstData = Nothing

'    If Not vxmlMap Is Nothing Then TypeName (Me) & " Line=" & Erl & vbCrLf & Err.Description = Err.Description & vbCrLf & " TYPE=""Cache""  NodeName=""" & vxmlMap.nodeName & """"
'    If Not strSelect = vbNullString Then TypeName (Me) & " Line=" & Erl & vbCrLf & Err.Description = Err.Description & vbCrLf & _
'        "SQL=""" & "SELECT * FROM " & strSource & " WHERE " & strSelect & """"
          
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Function GetOwnerDocument(ByRef rxmlNode As IXMLDOMNode) As FreeThreadedDOMDocument40
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".GetOwnerDocument")
' description:

'    Returns the node as a document if it is or the owner document if it
'    isn't

' pass:
'   xmlNode            The node
'
' return:
'            The document
    Const strFunctionName As String = "GetOwnerDocument"
    On Error GoTo GetOwnerDocumentVbErr
    
    If Not rxmlNode.ownerDocument Is Nothing Then
        Set GetOwnerDocument = rxmlNode.ownerDocument
    Else
        Set GetOwnerDocument = rxmlNode
    End If
    
    Exit Function
    
GetOwnerDocumentVbErr:
          
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function NodeSummary(ByVal vxmlNode As IXMLDOMNode) As String
    'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".NodeSummary")
' description:
'    Summarises the content of a node. Used for error messages
' pass:
'   xmlNode            The node to summarise'
'
    Const strFunctionName As String = "NodeSummary"
    On Error GoTo NodeSummaryVbErr
    
    Dim strOut As String
    Dim xmlAttrib As IXMLDOMAttribute
    
    If vxmlNode.nodeName <> vbNullString Then
        strOut = vxmlNode.nodeName
    Else
        strOut = "#NodeHasNoName#"
    End If

    strOut = strOut & " NodeType=""" & vxmlNode.nodeTypeString & """"
    
    If Not vxmlNode.Attributes Is Nothing Then
        For Each xmlAttrib In vxmlNode.Attributes
            strOut = strOut & " " & xmlAttrib.Name & "=""" & xmlAttrib.Value & """"
        Next xmlAttrib
    End If
       
    NodeSummary = "Node: " & strOut
    Exit Function
NodeSummaryVbErr:
          
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



Private Function IDBXMLWizard_CreateXML(ByVal vXMLIn As IXMLDOMNode, _
        ByVal vxmlMapIn As IXMLDOMNode, ByVal vconDB As ADODB.Connection) As FreeThreadedDOMDocument40
     'Dim objTimer As TimerHelper: Set objTimer = g_objtimers.GetMethodTimer(TypeName(Me) & ".IDBXMLWizard_CreateXML")
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IDBXMLWizard_CreateXML"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim colDataCache As DataCache
     
    #If WRITETABLESTOFILE Then
        On Error Resume Next
        Kill App.Path & "\TABLES.xml"
        On Error GoTo GetDataVbErr
    #End If
 
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
   
    '---------------------------------------------------------------------
    'Collect the base data essential for navigating the Databasde
    '---------------------------------------------------------------------
    'Open the connection - using the helpers.
    If m_objContext Is Nothing Then
        Set m_objRulesAssist = New MapRulesAssist
        Set colDataCache = New DataCache
    Else
        Set m_objRulesAssist = m_objContext.CreateInstance(App.Title & ".MapRulesAssist")
        Set colDataCache = m_objContext.CreateInstance(App.Title & ".DataCache")
    End If

    colDataCache.AddXML "XMLIN", vXMLIn.XML
    
    Set m_conOmiga4 = vconDB
        
    'Create the rules helper and load up the values for translation
    m_objRulesAssist.Create m_conOmiga4
    
    '---------------------------------------------------------------------
    'Now process the Map document recursively. An output document is
    'constructed based on the Map document. The collection of recordsets
    'retrieved so far is used as a base for retrieving other recordsets -
    'It is the minimum needed to identify an application.
    '---------------------------------------------------------------------
    ProcessMap xmlOut, vxmlMapIn, colDataCache
    
    Set IDBXMLWizard_CreateXML = xmlOut
    
    Set m_conOmiga4 = Nothing
    Set m_objRulesAssist = Nothing
    
    'Set objTimer = Nothing
    'Debug.Print g_objtimers.Report

'    Debug.Print "SQL TIME="; msngSQLTimer
    Debug.Print "NEW, "; msngSQLTimer

    Exit Function
    
GetDataVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    Set m_objRulesAssist = Nothing
    Set m_conOmiga4 = Nothing
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
       
    Err.Raise oeCommandFailed, TypeName(Me), "Mapping failed" & vbCrLf & Err.Description

End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

'vcolCache defines where field translation occurs
'vstrTable is the table or view that is required
'vstrWhere is the Where clause that will be translated
Private Function GetRecordset(ByVal vcolCache As DataCache, vstrTable As String, vstrWhere As String) As Recordset
    On Error GoTo GetRecordsetVbErr
    Const strMethodName = "GetRecordset"
    
    Const strMyStringDelim As String = "'"
    Const strSQLStringDelim As String = "'"
    
    Dim strExpr As String
    Dim strOut As String
    Dim strFieldName As String
    Dim strFieldValue As String
    Dim lngFieldStart As Long
    Dim lngFieldEnd As Long
    Dim cmd As Command
    
    msngSQLTimer = msngSQLTimer - Timer
   
    Set cmd = New Command
    Set cmd.ActiveConnection = m_conOmiga4

'Old way of doing things
'    cmd.CommandText = "SELECT * FROM " & vstrTable & " WHERE " & vcolCache.ReplaceFieldsWithData(vstrWhere)
'    Set GetRecordset = cmd.Execute
'    msngSQLTimer = msngSQLTimer + Timer
'    Exit Function
        
        strOut = vbNullString
    strExpr = vstrWhere
    'All fields are delimited by [] - eg [Table.Field], [Field].
    
    lngFieldStart = InStr(strExpr, "[")
    Do While lngFieldStart > 0
        lngFieldEnd = InStr(lngFieldStart + 1, strExpr, "]")
        'Grab the field name
        strFieldName = Mid$(strExpr, lngFieldStart + 1, lngFieldEnd - lngFieldStart - 1)
        strFieldValue = vcolCache.GetValue(strFieldName)

        If Mid$(strExpr, lngFieldStart - 1, 1) = "'" Then
            'It's a string Mrs Robinson, it's a string
            'Grab the bit before the field
            strOut = strOut & Left$(strExpr, lngFieldStart - 2)
            strOut = strOut & "?"
            'remove the processed field from the expression and continue parsing
            strExpr = Mid$(strExpr, lngFieldEnd + 2)
            cmd.Parameters.Append cmd.CreateParameter(, adBSTR, adParamInput, Len(strFieldValue), strFieldValue)
        Else
            'It's a number
            'Grab the bit before the field
            strOut = strOut & Left$(strExpr, lngFieldStart - 1)
            strOut = strOut & "?"
            'remove the processed field from the expression and continue parsing
            strExpr = Mid$(strExpr, lngFieldEnd + 1)
            If strFieldValue = vbNullString Then
                cmd.Parameters.Append cmd.CreateParameter(, adNumeric, adParamInput, , 0)
            Else
                'Debug.Assert IsNumeric(strValue)
                cmd.Parameters.Append cmd.CreateParameter(, adNumeric, adParamInput, , strFieldValue)
            End If
        End If
        
        lngFieldStart = InStr(strExpr, "[")
    Loop
    
    'Grab whatever is left of the expression
    strOut = strOut & strExpr
    
    cmd.CommandText = "SELECT * FROM " & vstrTable & " WHERE " & strOut
    Set GetRecordset = cmd.Execute
	Set cmd = Nothing

DoTimer:
    msngSQLTimer = msngSQLTimer + Timer


Exit Function

GetRecordsetVbErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
	Set cmd = Nothing

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


