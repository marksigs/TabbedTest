VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Workfile:      ParserAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Helper object for XML parser.
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' RF    01/07/99 Created
' RF    21/12/99 Use ErrAssist.ThrowError() not ErrAssist.RaiseError().
' SR    29/12/99 New parameter for the function GetTagValue
' SR    02/02/00 New Method - GetNodeFromClassDefByNodeValue
' DJP   08/02/00 Added GetNode and CheckNode
' DJP   17/02/00 Added GetRequestNodeEx
' DJP   17/02/00 Updated GetRequestNode and GetRequestNodeEx to create a new DOM document
'                when cloning
' SR    09/03/00 Modified methods GetRequestNodeEx - Check the REQUEST node is the top
'                level node in the Input to the function.
' MH    17/03/00 Changed so that WriteXMLtoFile uses SPM to make the decision
' PSC   11/04/00 AQR SYS0602: Amend GetNodeValue to take in a further optional parameter of
'                blnTextMandatory to check if the tag text is present
' PSC   28/04/00 AQR SYS0687: Amend GetNodeValue to only throw error if the node
'                itself is mandatory as well
' SR    03/07/02 SYS2433 - Modified method GetRequestNodeEx
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        AQR     Description
'TK     22/11/2004  BBG1821 Performance related fixes
'------------------------------------------------------------------------------------------
Option Explicit
' Private data
Private m_objErrAssist As ErrAssist
Public Enum BOOPERATIONTYPE
    booNone = 0
    booCreate = 1
    booDelete = 2
    booUpdate = 3
End Enum
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
' Prog  Date        AQR         Description
' MV    12/08/2002  BMIDS00323  Core AQR: SYS4754 - Performance. Replace all For...Each...Next with For...Next
'                               Modified AddNode(),CopyNode(),GroupNodesIntoList(),RemoveDuplicates(),DetermineOperation(),
'                               DetermineOperationEx(),AreAllNonKeyValuesNull(),AreAllNonKeyValuesNullEx(),
'                               AttachResponseData(),GetNodeFromClassDefByNodeValue()
'------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_objErrAssist = New ErrAssist
End Sub
Private Sub Class_Terminate()
    Set m_objErrAssist = Nothing
End Sub
Public Function load( _
    ByVal vstrXMLData As String, _
    ByVal vstrCallingObject As String, _
    ByVal vstrCallingFunction As String _
    ) As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Load an xml document from a string. Raise any resultant parser error.
' pass:
'   vstrXMLData         xml data stream containing data to be loaded
'   vstrCallingObject    name of calling object
'   vstrCallingFunction  name of calling function
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "load"
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.async = False
    objXmlDoc.setProperty "NewParser", True
    objXmlDoc.validateOnParse = False
    objXmlDoc.loadXML vstrXMLData
    If objXmlDoc.parseError.errorCode <> 0 Then
        RaiseParserError objXmlDoc.parseError, vstrCallingObject, vstrCallingFunction
    End If
    Set load = objXmlDoc
    Set objXmlDoc = Nothing
End Function
Private Sub RaiseParserError( _
    ByVal objParseError As IXMLDOMParseError, _
    ByVal vstrCallingObject As String, _
    ByVal vstrCallingFunction As String)
' header ----------------------------------------------------------------------------------
' description:
'   Raise XML parser error
' pass:
'   objParseError
'   vstrCallingObject    name of original calling object
'   vstrCallingFunction  name of original calling function
'------------------------------------------------------------------------------------------
On Error GoTo RaiseParserErrorVbErr
    Dim strFunctionName As String
    strFunctionName = "RaiseParserError"
    Dim strParserError As String
    strParserError = m_objErrAssist.FormatParserError(objParseError)
    m_objErrAssist.ThrowError vstrCallingObject, vstrCallingFunction, _
        oeXMLParserError, strParserError
RaiseParserErrorVbErr:
    
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        
End Sub
Public Function GetTagValue( _
    ByVal vxmlElementIn As IXMLDOMElement, _
    ByVal vstrTagName, _
    Optional ByRef rblnTagExists As Boolean, _
    Optional ByVal blnDeep As Boolean = True) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the text associated with the first tag of name vstrTagName, if one
'   exists. It is not an error for the tag to be absent.
' pass:
'   vxmlElementIn   An xml element, e.g. one containing xml string "<CustomerNumber>1234</CustomerNumber>"
'   vstrTagName     Tag name, e.g. "CustomerNumber"
'   rblnTagExists   Returned as True if the tag exists in the element,
'                   otherwise false
'   blnDeep         Should the request node be searched recursively. If set to False, the
'                   tag value will be searched for in the first level only.
'
' return:
'   GetTagValue     Tag value, e.g. "1234"
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "GetTagValue"
    Dim objXmlElem As IXMLDOMElement
    If vxmlElementIn.nodeName = vstrTagName Then
        Set objXmlElem = vxmlElementIn
    ElseIf Not blnDeep Then
        Set objXmlElem = vxmlElementIn.selectSingleNode(vstrTagName)
    Else
        Set objXmlElem = vxmlElementIn.selectSingleNode(".//" & vstrTagName)
    End If
    If Not objXmlElem Is Nothing Then
        rblnTagExists = True
        GetTagValue = objXmlElem.Text
    Else
        GetTagValue = ""
        rblnTagExists = False
    End If
    Set objXmlElem = Nothing
    Exit Function
End Function
Public Function GetAttributeValue( _
    ByVal vxmlElementIn As IXMLDOMElement, _
    ByVal vstrTagName, _
    ByVal vstrAttributeName, _
    Optional ByRef rblnAttributeExists As Boolean) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the text associated with the attribute vstrAttributeName, if one
'   exists. It is not an error for the tag to be absent.
' pass:
'   vxmlElementIn         An xml element, e.g. one containing xml string
'                         "<CustomerNumber TYPE="Omiga">1234</CustomerNumber>"
'   vstrTagName           Tag name, e.g. "CustomerNumber"
'   vstrAttributeName     Attribute name, e.g. "TYPE"
'   rblnAttributeExists   Returned as True if the attribute exists in the element,
'                         otherwise false
' return:
'   GetTagValue     Attribute value, e.g. "1234"
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "GetAttributeValue"
    Dim objXmlAttribute As IXMLDOMAttribute
    Dim objXmlElem As IXMLDOMElement
    If vxmlElementIn.nodeName = vstrTagName Then
        Set objXmlElem = vxmlElementIn
    Else
        Set objXmlElem = vxmlElementIn.selectSingleNode(".//" & vstrTagName)
    End If
       
    Set objXmlAttribute = objXmlElem.getAttributeNode(vstrAttributeName)
    If Not objXmlAttribute Is Nothing Then
        rblnAttributeExists = True
        GetAttributeValue = objXmlAttribute.Text
    Else
        GetAttributeValue = ""
        rblnAttributeExists = False
    End If
    Set objXmlAttribute = Nothing
    Set objXmlElem = Nothing
End Function
Public Function CopyNode(ByVal vstrSourceNodeName As String, _
                        ByVal vstrTargetNodeName As String, _
                        ByVal vobjXmlSourceElement As IXMLDOMElement, _
                        ByRef robjXmlTargetElement As IXMLDOMElement, _
                        ByVal vblnCloneChilden As Boolean) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Copies a node from the Source Element to the Target Element and
'   copies all children nodes
'   i.e.
'       [Source Element]                [Target Element]
'       <vstrSourceNodeName>    --->    <vstrTargetNodeName>
'           <childnode1>                    <childnode1>
'               <grandchild1>                   <grandchild1> (only if vblnCloneChildren)
'           <childnode2>                    <childnode2>
'               <grandchild1>                   <grandchild1> (only if vblnCloneChildren)
' pass:
'   vstrSourceNodeName      String containing the source node name
'   vstrTargetNodeName      String containing the target node name
'   vobjXmlSourceElement    IXMLDOMElement containing the source node
'   robjXmlTargetElement    IXMLDOMElement to contain the copied node
'return:
'   A boolean indicating success
' Raise Errors:
'   n/a
'------------------------------------------------------------------------------------------
    On Error GoTo CopyNodeVbErr
    Dim strFunctionName As String
    strFunctionName = "CopyNode"
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNewNode As IXMLDOMNode
    Dim objXmlChildNode As IXMLDOMNode
    'Dim objXmlParentNode As IXMLDOMNode
    Dim objXmlNodeList As IXMLDOMNodeList
    'SA SYS4754
    Dim intNodeCount As Integer
    Dim intNodeCount2 As Integer
    Dim intListLength As Integer
    Dim intListLength2 As Integer
    intListLength = 0
    intListLength2 = 0
    ' SR 19-01-00
    ' if vstrSourceNodeName is at top level of vobjXmlSourceElement, add all the children
    ' to the target
    If vobjXmlSourceElement.nodeName = vstrSourceNodeName Then
        Set objXmlNode = robjXmlTargetElement.ownerDocument.createElement(vstrTargetNodeName)
        Set objXmlNewNode = robjXmlTargetElement.appendChild(objXmlNode)
        'SA SYS4754
        intListLength = vobjXmlSourceElement.childNodes.length - 1
        For intNodeCount = 0 To intListLength
            Set objXmlChildNode = vobjXmlSourceElement.childNodes(intNodeCount).cloneNode(True)
            objXmlNewNode.appendChild objXmlChildNode
        Next
    Else
        ' obtain the list of nodes matching the vstrSourceNodeName
        Set objXmlNodeList = vobjXmlSourceElement.selectNodes(".//" & vstrSourceNodeName)
        ' for every occurrence of a node matching the vstrSourceNodeName create a
        ' new node on the target and copy its children
        'SA - Core AQR : SYS4754
        intListLength = objXmlNodeList.length - 1
        For intNodeCount2 = 0 To intListLength
            Set objXmlNode = robjXmlTargetElement.ownerDocument.createElement(vstrTargetNodeName)
            If Not objXmlNode Is Nothing Then
                Set objXmlNewNode = robjXmlTargetElement.appendChild(objXmlNode)
                intListLength2 = objXmlNodeList(intNodeCount2).childNodes.length - 1
                For intNodeCount = 0 To intListLength2
                    Set objXmlChildNode = objXmlNodeList(intNodeCount2).childNodes(intNodeCount).cloneNode(vblnCloneChilden)    '++ SA SYS4754
                    Set objXmlNode = objXmlNewNode.appendChild(objXmlChildNode)
                Next
            End If
        Next
    End If
    CopyNode = True
    Set objXmlNode = Nothing
    Set objXmlNewNode = Nothing
    Set objXmlChildNode = Nothing
    Set objXmlNodeList = Nothing
    Exit Function
CopyNodeVbErr:
    
    CopyNode = False
    Set objXmlNode = Nothing
    Set objXmlNewNode = Nothing
    Set objXmlChildNode = Nothing
    Set objXmlNodeList = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function AddNode(ByVal vxmlElem As IXMLDOMElement, _
                        ByVal vstrTargetNodeName As String, _
                        ByVal vblnCloneChildren As Boolean, _
                        ByRef rxmlDocumentElement As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Adds the specified node (vxmlElem) to every parent tag in the document element specified
'   specified by vstrTargetNodeName
'   i.e.
'       [rxmlDocumentElement]        [rxmlDocumentElement]
'       <vstrTargetNodeName>    --->    <vstrTargetNodeName>
'           <childnode1>                    <childnode1>
'           <childnode2>                    <childnode2>
'                                           <vxmlElem>
'                                                <childnode1> (if vblnCloneChildean = true)
' pass:
'   vxmlElem
'       IXMLDOMElement which specifies the element to add
'   vstrTargetNodeName
'       String specifying the target node name in rxmlDocumentElement
'   vblnCloneChildren
'       Boolean specifying if we intent to clone the children of the vxmlElem
'   rxmlDocumentElement
'       IXMLDOMElement specifying the target element to add into
'return:
'   A boolean indicating success
' Raise Errors:
'   omiga4invalidparameter
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'RF    01/07/99 Fixed to work when vstrTargetNodeName is at top level of
'               rxmlDocumentElement
'------------------------------------------------------------------------------------------
On Error GoTo AddNodeVbErr
  
    Dim strFunctionName As String
    strFunctionName = "AddNode"
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim intNodeCounter As Integer
    Dim intNodeListLength As Integer
     
    If rxmlDocumentElement Is Nothing Or _
        Len(vstrTargetNodeName) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If
    If rxmlDocumentElement.nodeName = vstrTargetNodeName Then
        Set xmlNewNode = vxmlElem.cloneNode(vblnCloneChildren)
        Set xmlTempNode = rxmlDocumentElement.appendChild(xmlNewNode)
    Else
        
        Set xmlNodeList = rxmlDocumentElement.selectNodes(".//" & vstrTargetNodeName)
        'SA - Core AQR - SYS4754
        intNodeListLength = xmlNodeList.length - 1              '++ SA SYS4754 Review changes
        For intNodeCounter = 0 To intNodeListLength
            Set xmlNewNode = vxmlElem.cloneNode(vblnCloneChildren)
            Set xmlTempNode = xmlNodeList(intNodeCounter).appendChild(xmlNewNode)
        Next
    End If
    AddNode = True
    Set xmlNodeList = Nothing
    Set xmlTempNode = Nothing
    Set xmlNewNode = Nothing
    Exit Function
AddNodeVbErr:
    AddNode = False
    Set xmlNodeList = Nothing
    Set xmlTempNode = Nothing
    Set xmlNewNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    ' re-raise the error back to the calling function
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function GetRequestNodeEx(ByVal vobjXmlIn As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the <REQUEST> node from an xml element, without any children. This is safer than
'   GetRequestNode because it doesn't start from the DOM document. What may happen in
'   code is that the XML passed into GetRequestNode could be passed in as follows:
'   GetRequestNode (vxmlIn.OwnerDocument)
'   This is unsafe because there is no guarantee that the Ownderdocument's xml is the
'   same as the original node. This method will clone from the point the request is found,
'   in the same node.
' pass:
' return:
'   GetRequestFromNode
' Raise Errors:
'   omiga4MissingPrimaryTag
'------------------------------------------------------------------------------------------
On Error GoTo GetRequestNodeExVbErr
    
    Dim strFunctionName As String
    Dim xmlDOMDocument As New FreeThreadedDOMDocument40
    Dim xmlDOMNode As IXMLDOMNode
    strFunctionName = "GetRequestNodeEx"
    Dim objXmlNode As IXMLDOMNode
    CheckNode vobjXmlIn, TypeName(Me), strFunctionName
    ' If the top most tag of the node is 'REQUEST', just return the node, else
    ' search for 'REQUEST'
    If vobjXmlIn.nodeName = "REQUEST" Then
        Set xmlDOMNode = vobjXmlIn.cloneNode(False)
    Else
        'SR 03/07/01:SYS2433 - Search in the context of the node
        'Set objXmlNode = vobjXmlIn.selectSingleNode("//REQUEST")
        Set objXmlNode = vobjXmlIn.selectSingleNode(".//REQUEST")
        If objXmlNode Is Nothing Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, "Expected REQUEST tag"
        End If
        Set xmlDOMNode = objXmlNode.cloneNode(False)
    End If
    ' Attach the cloned node to a new dom document to ensure safety of using selectsinglenode
    ' with something like this: "/REQUEST/SOMETHING"
    Set xmlDOMDocument = New FreeThreadedDOMDocument40
    Set xmlDOMDocument.documentElement = xmlDOMNode
    Set GetRequestNodeEx = xmlDOMNode
    Set xmlDOMDocument = Nothing
    Set xmlDOMNode = Nothing
    Set objXmlNode = Nothing
    Exit Function
GetRequestNodeExVbErr:
    Set xmlDOMDocument = Nothing
    Set xmlDOMNode = Nothing
    Set objXmlNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetRequestNode(ByVal vobjXmlIn As FreeThreadedDOMDocument40) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the <REQUEST> node from an xml document, without any children.
' pass:
' return:
'   GetRequestNode
' Raise Errors:
'   omiga4MissingPrimaryTag
'------------------------------------------------------------------------------------------
On Error GoTo GetRequestNodeVbErr
    
    Dim strFunctionName As String
    Dim xmlDOMDocument As New FreeThreadedDOMDocument40
    strFunctionName = "GetRequestNode"
    Dim objXmlNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Set objXmlNode = vobjXmlIn.selectSingleNode("//REQUEST")
    If objXmlNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "Expected REQUEST tag"
    End If
    Set xmlRequestNode = objXmlNode.cloneNode(False)
    ' Need to have new dom document owner for the clone
    Set xmlDOMDocument.documentElement = xmlRequestNode
    Set GetRequestNode = xmlRequestNode
    Set xmlDOMDocument = Nothing
    Set xmlRequestNode = Nothing
    Set objXmlNode = Nothing
    Exit Function
GetRequestNodeVbErr:
    Set xmlDOMDocument = Nothing
    Set xmlRequestNode = Nothing
    Set objXmlNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub AddNodeWithUCaseText(vxmlNode As IXMLDOMNode, vstrTagName As String)
' header ----------------------------------------------------------------------------------
' description:
'   Append a node with a tag name of '"UCASE" & vstrTagName', which contains the text
'   from tag vstrTagName converted to upper case. This is useful for sorting.
'   E.g.
'   <CUSTOMERVERSION>                   ->  <CUSTOMERVERSION>
'       <SURNAME>Smith</SURNAME>                <SURNAME>Smith</SURNAME>
'   </CUSTOMERVERSION>                          <UCASESURNAME>SMITH</UCASESURNAME>
'                                           </CUSTOMERVERSION>
' pass:
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddNodeWithUCaseTextVbErr
    Dim strFunctionName As String
    strFunctionName = "AddNodeWithUCaseText"
    Dim strTagValue As String
    Dim xmlElem As IXMLDOMElement
    strTagValue = Me.GetTagValue(vxmlNode, vstrTagName)
    Set xmlElem = vxmlNode.ownerDocument.createElement("UCASE" & vstrTagName)
    xmlElem.Text = UCase(strTagValue)
    vxmlNode.appendChild xmlElem
    Set xmlElem = Nothing
    Exit Sub
AddNodeWithUCaseTextVbErr:
    Set xmlElem = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ChangeNodeName(ByRef rxmlNode As IXMLDOMNode, _
                          ByVal vstrOldName As String, _
                          ByVal vstrNewName As String)
' header ----------------------------------------------------------------------------------
' description:
'   Changes all nodes present in rxmlNode that have a node name
'   of vstrOldName to vstrNewName
'
' pass:
'       rxmlNode    Xml Node
'       vstrOldName Name of nodes to be changed
'       vstrNewName Name nodes are to be changed to
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ChangeNodeNameVbErr
    
    Dim strFunctionName As String
    strFunctionName = "ChangeNodeName"
           
    Dim objNewNode As IXMLDOMNode
    Dim intChildIndex As Integer
    Dim intAttribIndex As Integer
                                     
    ' If this node is to be renamed create a new node with the new name
    ' and copy all the attributes across
    If rxmlNode.nodeName = vstrOldName Then
        Set objNewNode = rxmlNode.ownerDocument.createNode(rxmlNode.nodeType, vstrNewName, rxmlNode.namespaceURI)
        ' Copy attributes if this is an element node
        If objNewNode.nodeType = NODE_ELEMENT Then
            
            Dim objElement As IXMLDOMElement
            Set objElement = objNewNode
            For intAttribIndex = 0 To rxmlNode.Attributes.length - 1
                objElement.setAttributeNode rxmlNode.Attributes(intAttribIndex).cloneNode(True)
            Next
        End If
    End If
             
    ' For all children of this node change their name if it is vstrOldName and
    ' append them to the new node
    For intChildIndex = 0 To rxmlNode.childNodes.length - 1
        ChangeNodeName rxmlNode.childNodes(intChildIndex), vstrOldName, vstrNewName
        If Not objNewNode Is Nothing Then
            objNewNode.appendChild rxmlNode.childNodes(intChildIndex).cloneNode(True)
        End If
    Next
    ' Replace the original node with the new node
    If Not objNewNode Is Nothing Then
        If Not (rxmlNode.parentNode Is Nothing) Then
            rxmlNode.parentNode.replaceChild objNewNode, rxmlNode
        End If
        Set rxmlNode = objNewNode
    End If
    Set objNewNode = Nothing
    Set objElement = Nothing
    Exit Sub
ChangeNodeNameVbErr:
    Set objNewNode = Nothing
    Set objElement = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set objNewNode = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function CountNonNullNodes(ByRef robjElement As IXMLDOMElement) As Integer
' Header ----------------------------------------------------------------------------------
' Description:
'   Returns TRUE if all child nodes existing for the passed element represent 'null' values
'
' Pass:
'       robjElement (in) IXMLDOMElement
'            The element to be interrogated.
' Returns:
'       Boolean
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim i As Integer
Dim intReturn As Integer
    intReturn = 0
    For i = 0 To robjElement.childNodes.length - 1
        If Len(Trim$(robjElement.childNodes.Item(i).Text)) > 0 Then
            ' Found a non-null node value
            intReturn = intReturn + 1
        End If
    Next i
    CountNonNullNodes = intReturn
End Function
Public Sub GroupNodesIntoList(ByRef robjElement As IXMLDOMElement, _
                              ByVal vstrTagName As String)
' Header ----------------------------------------------------------------------------------
' Description:
'
' Pass:
'       robjElement     (in) IXMLDOMElement
'            The element to be interrogated.
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim i As Integer
'    Dim xmlNode     As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlListNode As IXMLDOMNode
    Dim intListLength As Integer
    Const strFunctionName = "GroupNodesIntoList"
    On Error GoTo GroupNodesIntoListVbErr
    Set xmlNodeList = robjElement.selectNodes(vstrTagName)
    If xmlNodeList.length > 0 Then
        Set xmlListNode = robjElement.appendChild(robjElement.ownerDocument.createElement(vstrTagName & "LIST"))
        'SA Core AQR : SYS4754
        ' Move each node in the node list under the new LIST tag
        intListLength = xmlNodeList.length - 1
        For i = 0 To intListLength
            xmlListNode.appendChild xmlNodeList(i)
        Next
    End If
'    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlListNode = Nothing
    Exit Sub
GroupNodesIntoListVbErr:
'    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlListNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Sub RemoveDuplicates(ByRef robjElement As IXMLDOMElement, _
                            Optional ByVal vstrCompareTag As String = "")
' Header ----------------------------------------------------------------------------------
' Description:
'       Uses the passed in tag name to compare each child node of robjElement. If the
'       value for this tag is the same for two child nodes, the second child node is removed.
'       Then, if the second node has children of its own its data is saved by moving the data across
'       to the first node.
'
'       If no tag is specified then the nodes are compared in their entirety for a match, including
'       all child nodes.
'
' Pass:
'       robjElement     (in) IXMLDOMElement
'            The element to be interrogated.
'       vstrCompareTag  (in) String [Optional]
'           The tag to use for comparison of nodes.
' Returns:
'       Boolean
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim intNodeCount As Integer
    Dim intListLength As Integer
    Dim xmlThisNode As IXMLDOMNode
    Dim xmlTestNode As IXMLDOMNode
    Const strFunctionName = "RemoveDuplicates"
    On Error GoTo RemoveDuplicatesVbErr
    i = 0
    Do While i <= robjElement.childNodes.length - 1
        Set xmlThisNode = robjElement.childNodes.Item(i)
        j = i + 1
        Do While j < robjElement.childNodes.length
            Set xmlTestNode = robjElement.childNodes.Item(j)
            If NodesMatch(xmlThisNode, xmlTestNode, vstrCompareTag) Then
                ' The IDs of the node and the test node match
                If vstrCompareTag <> "" Then
                    'SA - Core AQR : SYS4754
                    ' Copy the contents of the test node to the current node.
                    intListLength = xmlTestNode.childNodes.length - 1
                    For intNodeCount = 0 To intListLength
                        If xmlTestNode.childNodes(intNodeCount).childNodes.length > 1 Then
                            xmlThisNode.appendChild xmlTestNode.childNodes(intNodeCount).cloneNode(True)
                        End If
                    Next
                End If
                ' Contents of the test node have now been copied, so destroy it
                robjElement.removeChild xmlTestNode
            Else
                j = j + 1
            End If
        Loop
        i = i + 1
    Loop
    Set xmlThisNode = Nothing
    Set xmlTestNode = Nothing
    Exit Sub
RemoveDuplicatesVbErr:
    Set xmlThisNode = Nothing
    Set xmlTestNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Function NodesMatch(ByRef robjNode1 As IXMLDOMNode, _
                           ByRef robjNode2 As IXMLDOMNode, _
                           Optional ByVal vstrCompareTag As String = "") As Boolean
' Header ----------------------------------------------------------------------------------
' Description:
'   Determines if the two passed nodes 'match' each other.
'   The comparison is executed in one of two ways:
'       1) By comparing the value of a specific tag on each node.
'       2) By comparing the nodes in their entirety. This includes comparing all the child
'          nodes of the nodes.
'
'   Method 1 is used if the vstrCompareTag is non-blank.
'
' Pass:
'       robjNode1       (in) IXMLDOMNode
'       robjNode2       (in) IXMLDOMNode
'       vstrCompareTag  (in) String [Optional]
'           The tag to use for comparison of nodes.
' Returns:
'       Boolean
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim objNode As IXMLDOMNode
Dim i As Integer, _
    j As Integer
Dim blnReturn As Boolean
Dim xmlValue1 As IXMLDOMNode, _
    xmlValue2 As IXMLDOMNode
Const strFunctionName = "NodesMatch"
    On Error GoTo NodesMatchVbErr
    blnReturn = True
    If vstrCompareTag = "" Then
        ' Compare the entire nodes
        If robjNode1.hasChildNodes And (robjNode1.childNodes.length = robjNode2.childNodes.length) And _
            (robjNode1.nodeName = robjNode2.nodeName) Then
            ' Both nodes have the same number of children and the same tag names
            For i = 0 To robjNode1.childNodes.length - 1
                blnReturn = blnReturn And NodesMatch(robjNode1.childNodes.Item(i), robjNode2.childNodes.Item(i))
            Next i
        ElseIf Not (robjNode1.hasChildNodes Or robjNode2.hasChildNodes) And (robjNode1.nodeName = robjNode2.nodeName) Then
            ' Nodes do not have children (i.e. are 'leaf' nodes), so compare the text of both nodes instead
            blnReturn = (robjNode1.Text = robjNode2.Text)
        Else
            blnReturn = False
        End If
    Else
        ' Compare the specified tag in the nodes
        Set xmlValue1 = robjNode1.selectSingleNode(vstrCompareTag)
        Set xmlValue2 = robjNode2.selectSingleNode(vstrCompareTag)
        If (xmlValue1 Is Nothing) Or (xmlValue2 Is Nothing) Then
            blnReturn = False
        Else
            blnReturn = (xmlValue1.Text = xmlValue2.Text)
        End If
    End If
    Set objNode = Nothing
    Set xmlValue1 = Nothing
    Set xmlValue2 = Nothing
    NodesMatch = blnReturn
    Exit Function
NodesMatchVbErr:
    Set objNode = Nothing
    Set xmlValue1 = Nothing
    Set xmlValue2 = Nothing
    blnReturn = False
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Sub GroupNodes(ByRef robjElement As IXMLDOMElement, _
                      ByVal vstrGroupName As String, _
                      ByVal vstrFields As String)
' Header ----------------------------------------------------------------------------------
' Description:
'   Takes each child node of the passed element and attempts to group all specified tags under
'   a new group heading.
'
'   In practice, this means, for each child node of the element, taking each child-child node
'   and determining whether its tag appears in the vstrFields parameter. If it does, that tag
'   and its value are moved to become children of a new child-child node representing a new 'group'.
'   The name of the group is specified in the vstrGroupName parameter.
'
'   For example, if the following element was passed in:
'
'           <TABLE>
'               <FIELD1>text1</FIELD1>
'               <FIELD2>text2</FIELD2>
'               <FIELD3>text3</FIELD3>
'           </TABLE>
'
'       and vstrGroupName = "MYGROUP" and vstrFields = "FIELD2,FIELD3"; the resulting element would be:
'
'           <TABLE>
'               <FIELD1>text1</FIELD1>
'               <MYGROUP>
'                   <FIELD2>text2</FIELD2>
'                   <FIELD3>text3</FIELD3>
'               </MYGROUP>
'           </TABLE>
'
' Pass:
'       robjElement     (in) IXMLDOMElement
'       vstrGroupName   (in) String
'       vstrFields      (in) String
' Returns:
'       N/A
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim objXmlDoc As New FreeThreadedDOMDocument40
Dim objXMLGroupElement As IXMLDOMElement
Dim objThisNode As IXMLDOMNode
Dim blnCreatedGroupTag As Boolean
Dim strFields As String
Dim i As Integer
Const strFunctionName = "NodesMatch"
    On Error GoTo GroupNodesVbErr
    strFields = "," & vstrFields & ","
    blnCreatedGroupTag = False
    i = 0
    Do While i <= robjElement.childNodes.length - 1
        Set objThisNode = robjElement.childNodes.Item(i)
        If InStr(1, strFields, "," & objThisNode.nodeName & ",") > 0 Then
            ' The field represented by the tag is a member of the fields string so move the tag to
            ' the specified grouping element
            If Not blnCreatedGroupTag Then
                Set objXMLGroupElement = robjElement.appendChild(objXmlDoc.createElement(vstrGroupName))
                blnCreatedGroupTag = True
            End If
            objXMLGroupElement.appendChild objThisNode.cloneNode(True)
            robjElement.removeChild objThisNode
        Else
            i = i + 1
        End If
    Loop
    Set objXmlDoc = Nothing
    Set objXMLGroupElement = Nothing
    Set objThisNode = Nothing
    Exit Sub
GroupNodesVbErr:
    Set objXmlDoc = Nothing
    Set objXMLGroupElement = Nothing
    Set objThisNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Sub GroupNodesEx(ByRef rxmlElement As IXMLDOMElement, _
                        ByVal vstrGroupName As String, _
                        ByVal vstrFields As String)
' Header ----------------------------------------------------------------------------------
' Description:
'   Takes each child node of the passed element and attempts to group all specified tags under
'   a new group heading.
'
'   In practice, this means, for each child node of the element, taking each child-child node
'   and determining whether its tag appears in the vstrFields parameter. If it does, that tag
'   and its value are moved to become children of a new child-child node representing a new 'group'.
'   The name of the group is specified in the vstrGroupName parameter.
'
'   For example, if the following element was passed in:
'
'           <TABLE>
'               <FIELD1>text1</FIELD1>
'               <FIELD2>text2</FIELD2>
'               <FIELD3>text3</FIELD3>
'           </TABLE>
'
'       and vstrGroupName = "MYGROUP" and vstrFields = "FIELD2,FIELD3"; the resulting element would be:
'
'           <TABLE>
'               <FIELD1>text1</FIELD1>
'               <MYGROUP>
'                   <FIELD2>text2</FIELD2>
'                   <FIELD3>text3</FIELD3>
'               </MYGROUP>
'           </TABLE>
'
' Pass:
'       rxmlElement     (in) IXMLDOMElement
'       vstrGroupName   (in) String
'       vstrFields      (in) String
' Returns:
'       N/A
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim xmlGroupElement As IXMLDOMElement
Dim xmlThisNode As IXMLDOMNode
Dim blnCreatedGroupTag As Boolean
Dim strFields As String
Dim i As Integer
Const strFunctionName = "NodesMatch"
    On Error GoTo GroupNodesExVbErr
    strFields = "," & vstrFields & ","
    blnCreatedGroupTag = False
    i = 0
    Do While i <= rxmlElement.childNodes.length - 1
        Set xmlThisNode = rxmlElement.childNodes.Item(i)
        If InStr(1, strFields, "," & xmlThisNode.nodeName & ",") > 0 Then
            ' The field represented by the tag is a member of the fields string so move the tag to
            ' the specified grouping element
            If Not blnCreatedGroupTag Then
                Set xmlGroupElement = rxmlElement.appendChild(rxmlElement.ownerDocument.createElement(vstrGroupName))
                blnCreatedGroupTag = True
            End If
            xmlGroupElement.appendChild xmlThisNode.cloneNode(True)
            rxmlElement.removeChild xmlThisNode
        Else
            i = i + 1
        End If
    Loop
    Set xmlGroupElement = Nothing
    Set xmlThisNode = Nothing
    Exit Sub
GroupNodesExVbErr:
    Set xmlGroupElement = Nothing
    Set xmlThisNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Friend Sub WriteXMLToFile(ByVal vstrXML As String)
' Header ----------------------------------------------------------------------------------
' Description:
'   Writes the passed XML to a file on the C drive.
'
'   Note the conditional compilation constants around the code. These are present so that the
'   file is only created when specifically required.
'
' Pass:
'   vstrXML
'       The XML to write to the file.
' Returns:
'       N/A
' Raise Errors:
'------------------------------------------------------------------------------------------
End Sub
Public Function BuildComboValueList(ByVal vstrGroupName As String, _
                                    ByVal vstrValue As String) As String
' Header ----------------------------------------------------------------------------------
' Description:
'       Generates XML to pass to routines in the ComboDO such as GetComboValue.
'
'
' Pass:
'       vstrGroupName   (in) String
'       vstrValue       (in) String
' Returns:
'       String
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim objXMLListElement As IXMLDOMElement
Dim objXMLElement As IXMLDOMElement
Dim objXmlDoc As New FreeThreadedDOMDocument40
Const strFunctionName = "BuildComboValueList"
    On Error GoTo BuildComboValueListVbErr
    Set objXMLListElement = objXmlDoc.createElement("LIST")
    objXmlDoc.appendChild objXMLListElement
    Set objXMLElement = objXmlDoc.createElement("GROUPNAME")
    objXMLElement.Text = vstrGroupName
    objXMLListElement.appendChild objXMLElement
    Set objXMLElement = objXmlDoc.createElement("VALUEID")
    objXMLElement.Text = vstrValue
    objXMLListElement.appendChild objXMLElement
    BuildComboValueList = objXmlDoc.xml
    Set objXmlDoc = Nothing
    Set objXMLElement = Nothing
    Set objXMLListElement = Nothing
    Exit Function
BuildComboValueListVbErr:
    Set objXmlDoc = Nothing
    Set objXMLElement = Nothing
    Set objXMLListElement = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function DetermineOperation( _
                    ByVal vstrXMLRequest As String, _
                    ByVal vstrClassDefinition As String) As BOOPERATIONTYPE
' header ----------------------------------------------------------------------------------
' description:  Determines the appropriate operation to conduct on the passed XML given
'               the primary key field values.
'
'               If the value specified for the final key field in the paramarray is blank then
'               booInsert is returned.
'
'               If all values for the key fields in the paramarray are non-blank and ALL VALUES
'               FOR NON-KEY FIELDS ARE BLANK the booDelete is returned.
'
'               If all values for the key fields in the paramarray are non-blank and at least one
'               value for a non-key field is non-blank then booUpdate is returned.
'
' pass:         vstrXMLRequest
'               vstrClassDefinition
' return:       BOOPERATIONTYPE
'------------------------------------------------------------------------------------------
    On Error GoTo DetermineOperationVbErr
    Dim strFunctionName As String
    strFunctionName = "DetermineOperation"
    Dim xmlIn  As New FreeThreadedDOMDocument40
    Dim xmlClassDefinition As New FreeThreadedDOMDocument40
    Dim xmlNodeList      As IXMLDOMNodeList
    Dim xmlValueNodeList As IXMLDOMNodeList
    Dim xmlKeyNodeList   As IXMLDOMNodeList
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlValueNode As IXMLDOMNode
    Dim blnCompleteKey         As Boolean
    Dim blnHasNonPrimaryValues As Boolean
    Dim strTableName As String
    Dim strValue As String
    Dim i As Integer
    Dim intNodeCount As Integer
    Dim intListLength As Integer
    Dim botReturn As BOOPERATIONTYPE
    '
    ' Initialise
    '
    botReturn = booNone
    Set xmlIn = load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlClassDefinition = load(vstrClassDefinition, TypeName(Me), strFunctionName)
    '
    ' Get table name
    '
    Set xmlValueNodeList = xmlClassDefinition.selectNodes("TABLENAME")
    If xmlValueNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTableName = xmlValueNodeList.Item(0).firstChild.Text
    End If
    '
    ' Find table tag in the request XML
    '
    Set xmlNodeList = xmlIn.selectNodes(".//" & strTableName)
    If xmlNodeList.length = 0 Then
        Exit Function
    Else
        Set xmlRootElement = xmlNodeList.Item(0)
    End If
    '
    ' Get the primary key nodes from the class definition
    '
    Set xmlKeyNodeList = xmlClassDefinition.selectNodes(".//PRIMARYKEY")
    If xmlKeyNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "No keys found"
    End If
    '
    ' Ascertain values of primary key fields
    '
    blnCompleteKey = True
    'SA Core AQR : SYS4754
    intListLength = xmlKeyNodeList.length - 1           '++ SA SYS4754 code review
    For intNodeCount = 0 To intListLength
        strValue = ""
        ' Get node in the data XML corresponding to this key
        Set xmlValueNodeList = xmlRootElement.selectNodes(".//" & xmlKeyNodeList(intNodeCount).firstChild.Text)
        If xmlValueNodeList.length <> 0 Then
            Set xmlValueNode = xmlValueNodeList.Item(0)
            strValue = xmlValueNode.Text
            xmlRootElement.removeChild xmlValueNode
        End If
        blnCompleteKey = blnCompleteKey And (strValue <> "")
    Next
    ' Determine operation
    '
    blnHasNonPrimaryValues = Not AllChildNodesAreNull(xmlRootElement)
    If Not blnCompleteKey And blnHasNonPrimaryValues Then
        ' INSERT
        botReturn = booCreate
    ElseIf blnCompleteKey And blnHasNonPrimaryValues Then
        ' UPDATE
        botReturn = booUpdate
    ElseIf blnCompleteKey And Not blnHasNonPrimaryValues Then
        ' DELETE
        botReturn = booDelete
    Else
        ' Could not ascertain process to perform
        botReturn = booNone
    End If
DetermineOperationExit:
    
    Set xmlRootElement = Nothing
    Set xmlIn = Nothing
    Set xmlClassDefinition = Nothing
    Set xmlNodeList = Nothing
    Set xmlValueNodeList = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNode = Nothing
    DetermineOperation = botReturn
    Exit Function
DetermineOperationVbErr:
    Set xmlRootElement = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlValueNodeList = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlClassDefinition = Nothing
    Set xmlValueNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function DetermineOperationEx( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    ByVal vxmlClassDefinition As FreeThreadedDOMDocument40) As BOOPERATIONTYPE
' header ----------------------------------------------------------------------------------
' description:  Determines the appropriate operation to conduct on the passed XML given
'               the primary key field values.
'
'               If the value specified for the final key field in the paramarray is blank then
'               booInsert is returned.
'
'               If all values for the key fields in the paramarray are non-blank and ALL VALUES
'               FOR NON-KEY FIELDS ARE BLANK the booDelete is returned.
'
'               If all values for the key fields in the paramarray are non-blank and at least one
'               value for a non-key field is non-blank then booUpdate is returned.
'
' pass:         vstrXMLRequest
'               vstrClassDefinition
' return:       BOOPERATIONTYPE
'------------------------------------------------------------------------------------------
    On Error GoTo DetermineOperationExVbErr
    Dim strFunctionName As String
    strFunctionName = "DetermineOperationEx"
    Dim xmlKeyNodeList   As IXMLDOMNodeList
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlValueNode As IXMLDOMNode
    Dim blnCompleteKey As Boolean
    Dim blnHasNonKeyValues As Boolean
    Dim strTableName As String
    Dim strValue As String
    Dim i As Integer
    Dim intKeyValueCount As Integer
    Dim intNodeCount As Integer
    Dim intListLength As Integer
    Dim botReturn As BOOPERATIONTYPE
    '
    ' Initialise
    '
    botReturn = booNone
    '
    ' Get table name
    '
    Set xmlValueNode = vxmlClassDefinition.selectSingleNode("TABLENAME")
    If xmlValueNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTableName = xmlValueNode.firstChild.Text
    End If
    '
    ' Find table tag in the request XML
    '
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRootElement = vxmlRequest
    Else
        Set xmlRootElement = vxmlRequest.selectSingleNode(".//" & strTableName)
    End If
    If xmlRootElement Is Nothing Then
        Exit Function
    End If
    '
    ' Get the primary key nodes from the class definition
    '
    Set xmlKeyNodeList = vxmlClassDefinition.documentElement.selectNodes("PRIMARYKEY")
    If xmlKeyNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "No keys found"
    End If
    '
    ' Ascertain values of primary key fields
    '
    blnCompleteKey = True
    intKeyValueCount = 0
    'SA Core AQR : SYS4754
    intListLength = xmlKeyNodeList.length - 1
    For intNodeCount = 0 To intListLength
        strValue = ""
        ' Get node in the data XML corresponding to this key
        Set xmlValueNode = xmlRootElement.selectSingleNode(xmlKeyNodeList(intNodeCount).firstChild.Text)
        If Not (xmlValueNode Is Nothing) Then
            strValue = xmlValueNode.Text
            intKeyValueCount = intKeyValueCount + 1
        End If
        blnCompleteKey = blnCompleteKey And (strValue <> "")
    Next
    '
    ' Determine operation
    '
    blnHasNonKeyValues = Not AreAllNonKeyValuesNullEx(xmlRootElement, vxmlClassDefinition)
    If Not blnCompleteKey Then
        ' INSERT
        botReturn = booCreate
    ElseIf blnCompleteKey And blnHasNonKeyValues Then
        ' UPDATE
        botReturn = booUpdate
    ElseIf blnCompleteKey And Not blnHasNonKeyValues Then
        ' DELETE
        botReturn = booDelete
    Else
        ' Could not ascertain process to perform
        botReturn = booNone
    End If
DetermineOperationExExit:
    
    Set xmlRootElement = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNode = Nothing
    DetermineOperationEx = botReturn
    Exit Function
DetermineOperationExVbErr:
    Set xmlRootElement = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function AreAllNonKeyValuesNull(ByVal vstrXMLRequest As String, _
                                       ByVal strClassDef As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Cheecks whether all non-key values are Null. If so return True elese False
'
' pass:         vstrXMLRequest  - xml Request (to be tested for),
'               strClassDef - corresponding Clas definition
' return:       Boolean
'------------------------------------------------------------------------------------------
    On Error GoTo AreAllNonKeyValuesNullVbErr
    Dim strFunctionName As String
    strFunctionName = "AreAllNonKeyValuesNull"
    Dim objXmlAssist As New XMLAssist
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlClassDef As New FreeThreadedDOMDocument40
    Dim xmlNonPKNodeList As IXMLDOMNodeList
    Dim strNodeName As String
    Dim intNodeCount As Integer
    Dim intListLength As Integer
    'Load Class Definition and the request
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlClassDef = objXmlAssist.load(strClassDef, TypeName(Me), strFunctionName)
    Set xmlNonPKNodeList = xmlClassDef.selectNodes(".//OTHERS")
    AreAllNonKeyValuesNull = True
    'SA - Core AQR :SYS4754
    intListLength = xmlNonPKNodeList.length - 1
    For intNodeCount = 0 To intListLength
        ' Get node in the data XML corresponding to this key
        strNodeName = xmlNonPKNodeList(intNodeCount).firstChild.Text
        If Len(GetTagValue(xmlIn.documentElement, strNodeName)) <> 0 Then
            AreAllNonKeyValuesNull = False
            Exit For
        End If
    Next
    Set xmlNonPKNodeList = Nothing
    Set xmlIn = Nothing
    Set xmlClassDef = Nothing
    Set objXmlAssist = Nothing
    Exit Function
AreAllNonKeyValuesNullVbErr:
    Set xmlNonPKNodeList = Nothing
    Set xmlIn = Nothing
    Set xmlClassDef = Nothing
    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function AreAllNonKeyValuesNullEx(ByVal vXMLRequestElem As IXMLDOMElement, _
                                         ByVal vxmlClassDef As FreeThreadedDOMDocument40) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Cheecks whether all non-key values are Null. If so return True elese False
'
' pass:         vXMLRequestElem  - xml Request ,
'               vxmlClassDef - corresponding Clas definition
' return:       Boolean
'------------------------------------------------------------------------------------------
    On Error GoTo AreAllNonKeyValuesNullExVbErr
    Dim strFunctionName As String
    strFunctionName = "AreAllNonKeyValuesNullEx"
    Dim objXmlAssist As New XMLAssist
    Dim xmlNonPKNodeList As IXMLDOMNodeList
    Dim intNodeCount As Integer
    Dim strNodeName As String
    Dim intListLength As Integer
    Set xmlNonPKNodeList = vxmlClassDef.selectNodes(".//OTHERS")
    AreAllNonKeyValuesNullEx = True
    'SA - Core AQR :SYS4754
    intListLength = xmlNonPKNodeList.length - 1
    For intNodeCount = 0 To intListLength
        ' Get node in the data XML corresponding to this key
        strNodeName = xmlNonPKNodeList(intNodeCount).firstChild.Text
        If Len(GetTagValue(vXMLRequestElem, strNodeName, blnDeep:=False)) <> 0 Then
            AreAllNonKeyValuesNullEx = False
            Exit For
        End If
    Next
    Set xmlNonPKNodeList = Nothing
    Set objXmlAssist = Nothing
    Exit Function
AreAllNonKeyValuesNullExVbErr:
    
    
    Set xmlNonPKNodeList = Nothing
    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function AllChildNodesAreNull(ByRef robjElement As IXMLDOMElement) As Boolean
' Header ----------------------------------------------------------------------------------
' Description:
'   Returns TRUE if all child nodes existing for the passed element represent 'null' values
'
' Pass:
'       robjElement (in) IXMLDOMElement
'            The element to be interrogated.
' Returns:
'       Boolean
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim i As Integer
Dim blnReturn As Boolean
    blnReturn = True
    For i = 0 To robjElement.childNodes.length - 1
        If Len(Trim$(robjElement.childNodes.Item(i).Text)) > 0 Then
            ' Found a non-null node value
            blnReturn = False
        End If
    Next i
    AllChildNodesAreNull = blnReturn
End Function
Public Function GetMandatoryAttribute(ByVal vobjNode As IXMLDOMNode, _
                                            ByVal vstrAttribName As String) As String
    Dim strValue As String
    If Not vobjNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
        strValue = vobjNode.Attributes.getNamedItem(vstrAttribName).Text
    End If
    If Len(strValue) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                "GetMandatoryNode", _
                                oeMissingElement, _
                                "Missing mandatory attribute - [" _
                                & vobjNode.nodeName & "." & vstrAttribName & "]"
    End If
    GetMandatoryAttribute = strValue
End Function
Public Function GetElementText(ByVal vobjParent As IXMLDOMNode, _
                               ByVal vstrPattern As String) As String
    If Not vobjParent.selectSingleNode(vstrPattern) Is Nothing Then
        GetElementText = vobjParent.selectSingleNode(vstrPattern).Text
    End If
End Function
Public Function GetElementTextBln(ByVal vobjParent As IXMLDOMNode, _
                                  ByVal vstrPattern As String, _
                                  ByRef rstrElementText) _
                                  As Boolean
    Dim strText As String
    If Not vobjParent.selectSingleNode(vstrPattern) Is Nothing Then
        strText = vobjParent.selectSingleNode(vstrPattern).Text
    End If
    rstrElementText = strText
    GetElementTextBln = (Len(strText) > 0)
End Function
Public Function GetMandatoryElementText(ByVal vobjParent As IXMLDOMNode, _
                                        ByVal vstrPattern As String) As String
    Dim objNodes As IXMLDOMNodeList
    Dim strElementText As String
    Set objNodes = vobjParent.selectNodes(vstrPattern)
    If objNodes.length > 0 Then
        strElementText = objNodes.Item(0).Text
    End If
    Set objNodes = Nothing
    If Len(strElementText) = 0 Then
        m_objErrAssist.ThrowError _
            TypeName(Me), "GetMandatoryElementText", oeMissingElementValue, vstrPattern
    End If
    GetMandatoryElementText = strElementText
    Set objNodes = Nothing
End Function
Public Function GetMandatoryNode(ByVal vobjParent As IXMLDOMNode, _
                                    ByVal vstrMatchPattern As String) As IXMLDOMNode
    Dim objNodes As IXMLDOMNodeList
    Set objNodes = vobjParent.selectNodes(vstrMatchPattern)
    If objNodes.length <> 1 Then
        
        Set objNodes = Nothing
        m_objErrAssist.ThrowError _
            TypeName(Me), "GetMandatoryNode", oeMissingElement, vstrMatchPattern
    End If
    Set GetMandatoryNode = objNodes.Item(0)
    Set objNodes = Nothing
End Function
Public Sub AttachResponseData(ByVal vxmlNodeToAttachTo As IXMLDOMNode, _
                              ByVal vxmlResponse As IXMLDOMElement)
' Header ----------------------------------------------------------------------------------
' Description:
'   Gets the data nodes out of vxmlResponse and appends them to vxmlNodeToAttachTo
' Pass:
'       vxmlNodeToAttachTo  xml node to attach the data to
'       vxmlResponse        xml element containing the response whos data is to be
'                            extracted
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "AttachResponseData"
    Dim xmlChildList As IXMLDOMNodeList
    Dim intNodeCount As Integer
    Dim intListLength As Integer
    If vxmlNodeToAttachTo Is Nothing Or vxmlResponse Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "Node to attach to or Response missing"
     End If
    If vxmlResponse.nodeName <> "RESPONSE" Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "RESPONSE must be top level tag"
    End If
         
     Set xmlChildList = vxmlResponse.childNodes
     'SA Core AQR : SYS4754
     intListLength = xmlChildList.length - 1
     For intNodeCount = 0 To intListLength
        If xmlChildList(intNodeCount).nodeName <> "MESSAGE" Then
            vxmlNodeToAttachTo.appendChild xmlChildList(intNodeCount).cloneNode(True)
        End If
     Next
    Set xmlChildList = Nothing
     
End Sub
Public Sub PromoteNode(ByRef rxmlData As IXMLDOMNode, _
                       ByVal vxmlNodeToPromote As IXMLDOMNode, _
                       Optional vblnReplaceParent As Boolean = True, _
                       Optional vblnInheritSiblings As Boolean = True)
' header ----------------------------------------------------------------------------------
' description:  Within the specified XML (rxmlData), finds the specified node and promotes
'               it up one level in the XML hierarchy.
'
'               If vblnReplaceParent = True then the node will not only be promoted up one level
'               but will also replace its original parent node too.
'
'               If vblnInheritSiblings = True then the node will inherit all the nodes that
'               were its siblings (BEFORE the promotion) as its children.
'
' pass:         rxmlData (in/out)
'               vxmlNodeToPromote (in)
' return:       N/A
'------------------------------------------------------------------------------------------
    On Error GoTo PromoteNodeVbErr
Dim strFunctionName As String
    strFunctionName = "PromoteNode"
Dim xmlSiblingNode   As IXMLDOMNode, _
    xmlParentNode    As IXMLDOMNode, _
    xmlGrandparentNode As IXMLDOMNode
    '
    ' Initialise
    '
    If vxmlNodeToPromote Is Nothing Then
        Exit Sub
    End If
    If (vxmlNodeToPromote.parentNode Is Nothing) Or _
       (vxmlNodeToPromote Is vxmlNodeToPromote.ownerDocument) Then
        Exit Sub
    End If
    ' Find parent node of the candidate node to promote
    Set xmlParentNode = vxmlNodeToPromote.parentNode
    If xmlParentNode Is Nothing Then
        Exit Sub
    End If
    '
    ' Inherit siblings if necessary
    '
    If vblnInheritSiblings Then
        ' Inherit all the siblings
        For Each xmlSiblingNode In xmlParentNode.childNodes
            If Not (xmlSiblingNode Is vxmlNodeToPromote) Then
                vxmlNodeToPromote.appendChild xmlSiblingNode
            End If
        Next xmlSiblingNode
    End If
    '
    ' Promote the node, including the replacement of the parent node if necessary
    '
    Set xmlGrandparentNode = xmlParentNode.parentNode
    If (xmlGrandparentNode Is Nothing) Or _
       (xmlParentNode Is xmlParentNode.ownerDocument) Or _
       (xmlGrandparentNode Is xmlParentNode.ownerDocument) Then
        ' Parent node is the root node of the document
        If vblnReplaceParent Then
            ' Parent node is the root of the document so point the passed XML pointer to the promoted node itself
            Set rxmlData = vxmlNodeToPromote
        Else
            ' Do nothing - since the parent node is the root of the document the node to be promoted cannot be moved
            ' up to the same level in the XML hierarchy as the parent as that would mean that there were two root nodes
        End If
    Else
        If vblnReplaceParent Then
            ' Remove the original parent
            xmlGrandparentNode.removeChild xmlParentNode
        Else
            ' Do nothing - the node to be promoted can simply be moved up to be a sibling of its parent node
        End If
        ' Promote this node
        xmlGrandparentNode.appendChild vxmlNodeToPromote
    End If
PromoteNodeExit:
    
    Set xmlSiblingNode = Nothing
    Set xmlParentNode = Nothing
    Set xmlGrandparentNode = Nothing
    Exit Sub
PromoteNodeVbErr:
    Set xmlSiblingNode = Nothing
    Set xmlParentNode = Nothing
    Set xmlGrandparentNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Sub SelectSchemaFields(ByVal varList As Variant, _
                               ByVal xmlTableNode As IXMLDOMNode)
    Dim xmlNode As IXMLDOMNode
    Dim xmlNextSibling As IXMLDOMNode
    Dim blnHit As Boolean
    Dim int0 As Integer
    Set xmlNode = xmlTableNode.firstChild
    Do While Not xmlNode Is Nothing
        Set xmlNextSibling = xmlNode.nextSibling
            
        If xmlNode.nodeType = NODE_ELEMENT Then
            blnHit = False
            For int0 = 0 To (UBound(varList) - 1)
                If varList(int0) = xmlNode.firstChild.Text Then
                    blnHit = True
                    Exit For
                End If
            Next
            If Not blnHit Then
                xmlTableNode.removeChild xmlNode
            End If
        End If
        Set xmlNode = xmlNextSibling
    Loop
    Set xmlNode = Nothing
    Set xmlNextSibling = Nothing
End Sub
Public Function GetAttributeFromNode(ByVal vobjNode As IXMLDOMNode, _
                                     ByVal vstrAttribName As String) _
                                     As String
    Dim strValue As String
    If Not vobjNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
        strValue = vobjNode.Attributes.getNamedItem(vstrAttribName).Text
    End If
    GetAttributeFromNode = strValue
End Function
Public Sub SetMandatoryChildText(ByVal vobjXmlParentNode As IXMLDOMNode, _
                                 ByVal vstrTagName, _
                                 ByVal vstrText)
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Set objXmlNode = vobjXmlParentNode.selectSingleNode(vstrTagName)
    If objXmlNode Is Nothing Then
        Set objXmlElem = vobjXmlParentNode.ownerDocument.createElement(vstrTagName)
        objXmlElem.Text = vstrText
        vobjXmlParentNode.appendChild objXmlElem
    Else
        objXmlNode.Text = vstrText
    End If
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
End Sub
Public Function GetNodeFromClassDefByNodeValue( _
                            ByVal xmlClassDefDoc As FreeThreadedDOMDocument40, _
                            ByVal strNodeValue As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Within the specified class definition document, checks for a node with
'               specified value, and returns it
'
' pass:         xmlClassDefDoc - Class Definition to query the node from
'               strNodeValue - Value of the node sought
' return:       Queried node
'------------------------------------------------------------------------------------------
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim intNodeCounter As Integer
    Dim intListLength As Integer
     
    ' First check whether the required node is part of Primary Key
    Set xmlNodeList = xmlClassDefDoc.documentElement.selectNodes("PRIMARYKEY")
    'SA - Core AQR: SYS4754
    intListLength = xmlNodeList.length - 1
    For intNodeCounter = 0 To intListLength
        If xmlNodeList(intNodeCounter).firstChild.Text = strNodeValue Then
            Set GetNodeFromClassDefByNodeValue = xmlNodeList(intNodeCounter)
            Exit Function
        End If
    Next
    intNodeCounter = 0
    intListLength = 0
    ' check whether the required node is amongst others
    Set xmlNodeList = xmlClassDefDoc.documentElement.selectNodes("OTHERS")
    'SA - Core AQR: SYS4754
    intListLength = xmlNodeList.length - 1
    For intNodeCounter = 0 To intListLength
        If xmlNodeList(intNodeCounter).firstChild.Text = strNodeValue Then
            Set GetNodeFromClassDefByNodeValue = xmlNodeList(intNodeCounter)
            Exit Function
        End If
    Next
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
End Function
Public Sub CheckNode(ByVal xmlNode As Object, strObjectName As String, strFunction As String, Optional strText As String = "Node Empty")
    Dim strFunctionName As String
    strFunctionName = "CheckNode"
    On Error GoTo CheckNodeErr
    If xmlNode Is Nothing Then
        m_objErrAssist.ThrowError strObjectName, strFunction, oeMissingPrimaryTag, strText
    End If
    Exit Sub
CheckNodeErr:
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetNode(ByVal xmlNodeParent As IXMLDOMNode, _
                        ByVal vstrMatchPattern As String, _
                        Optional blnNodeMandatory As Boolean = True) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Finds the node specified by vstrMatchPattern in the XML node xmlNodeParent
' pass:
'   xmlNodeParent        xml node to be searched
'   vstrMatchPattern     the XSL search pattern
'   blnNodeMandatory     Optional - if true, an error is raised if the pattern is not met.
'                        If false, nothing is returned.
'------------------------------------------------------------------------------------------
    Dim xmlNode As IXMLDOMNode
    Dim strFunctionName As String
    strFunctionName = "GetNode"
    On Error GoTo GetNodeErr
     
    If Not xmlNodeParent Is Nothing Then
        Set xmlNode = xmlNodeParent.selectSingleNode(vstrMatchPattern)
        If blnNodeMandatory = True And xmlNode Is Nothing Then
            m_objErrAssist.ThrowError _
                TypeName(Me), strFunctionName, oeMissingElement, "Match pattern: " & vstrMatchPattern
        End If
    Else
        m_objErrAssist.ThrowError _
            TypeName(Me), strFunctionName, oeMissingElement, "Parent node empty"
    End If
    Set GetNode = xmlNode
    Set xmlNode = Nothing
    Exit Function
GetNodeErr:
    Set xmlNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetNodeValue(ByVal xmlNodeParent As IXMLDOMNode, _
                             ByVal vstrMatchPattern As String, _
                             Optional blnNodeMandatory As Boolean = True, _
                             Optional blnTextMandatory As Boolean = True) As String
    On Error GoTo GetNodeValueErr
    Dim xmlNode As IXMLDOMNode
    Dim strFunctionName As String
    Dim strSearch As String
    strFunctionName = "GetNodeValue"
    Set xmlNode = GetNode(xmlNodeParent, vstrMatchPattern, blnNodeMandatory)
    If Not xmlNode Is Nothing Then
        GetNodeValue = xmlNode.Text
        ' PSC 11/04/00 AQR SYS0602: Check blnTextMandatory instead of blnNodeMandatory
        ' PSC 28/04/00 AQR SYS0687: Only throw error if the node itself is mandatory as well
        If blnNodeMandatory = True And Len(GetNodeValue) = 0 And blnTextMandatory = True Then
            m_objErrAssist.ThrowError _
                TypeName(Me), _
                strFunctionName, _
                oeInvalidParameter, _
                "Tag " & vstrMatchPattern & " exists, but value is empty"
        End If
    End If
    Set xmlNode = Nothing
    Exit Function
GetNodeValueErr:
    Set xmlNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub StartXMLLogging()
    On Error GoTo StartXMLLoggingErr
    Const strMethodName As String = "StartXMLLogging"
    SetXMLLoggingFlag True
    Exit Sub
StartXMLLoggingErr:
    m_objErrAssist.AddToErrSource strMethodName
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub StopXMLLogging()
    On Error GoTo StopXMLLoggingErr
    Const strMethodName As String = "StopXMLLogging"
        
    SetXMLLoggingFlag False
    Exit Sub
StopXMLLoggingErr:
    m_objErrAssist.AddToErrSource strMethodName
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub SetXMLLoggingFlag(vblnFlag As Boolean)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim blnExists As Boolean
        
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("XMLLogOptions", LockSetGet, Process, blnExists)
    Set spmProperty = _
        spmGroup.CreateProperty("Logging", blnExists)
    If vblnFlag = True Then
        spmProperty.Value = "ON"
    Else
        spmProperty.Value = "OFF"
    End If
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing
End Sub
Private Function IsXMLLoggingOn() As Boolean
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim blnGroupExists As Boolean
    Dim blnPropertyExists As Boolean
    IsXMLLoggingOn = False
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
            
    Set spmGroup = spmMgr.CreatePropertyGroup _
                ("XMLLogOptions", LockSetGet, Process, blnGroupExists)
    Set spmProperty = _
                spmGroup.CreateProperty("Logging", blnPropertyExists)
    IsXMLLoggingOn = (spmProperty.Value = "ON")
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing
End Function
