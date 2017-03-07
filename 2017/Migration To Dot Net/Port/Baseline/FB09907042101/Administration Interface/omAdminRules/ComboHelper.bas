Attribute VB_Name = "ComboHelper"
'MARS Specific History
'Prog    Date        AQR     Description
'SD      19/01/2006  1811   Added function GetValidationtype

Option Explicit

Private m_xmldocCombos As FreeThreadedDOMDocument40

Public Function IsValidationType( _
    ByVal vstrGroupName As String, _
    ByVal vIntValueID As Integer, _
    ByVal vstrValidationType As String) _
    As Boolean
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & _
        "' and @VALUEID='" & CStr(vIntValueID) & _
        "' and @VALIDATIONTYPE='" & vstrValidationType & "']"
        
    If Not m_xmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        IsValidationType = True
    End If

End Function
Public Function GetValidationType( _
    ByVal vstrGroupName As String, _
    ByVal vIntValueID As Integer) _
    As String
    
    Dim strPattern As String
    Dim xmlNode As IXMLDOMNode
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & _
        "' and @VALUEID='" & CStr(vIntValueID) & "']"
        
    Set xmlNode = m_xmldocCombos.selectSingleNode(strPattern)
    If Not xmlNode Is Nothing Then
        GetValidationType = GetAttributeText(xmlNode, "VALIDATIONTYPE")
    End If
     
End Function

Private Sub LoadComboGroup( _
    ByVal vstrGroupName As String)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlComboNode As IXMLDOMNode
    Dim xmlComboListNode As IXMLDOMNode
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlSchemaNode As IXMLDOMNode
    
    If m_xmldocCombos Is Nothing Then
        Set m_xmldocCombos = New FreeThreadedDOMDocument40
        m_xmldocCombos.validateOnParse = False
        m_xmldocCombos.setProperty "NewParser", True
        Set xmlElem = m_xmldocCombos.createElement("COMBOLIST")
        Set xmlComboListNode = m_xmldocCombos.appendChild(xmlElem)
    End If
    
    If m_xmldocCombos.selectSingleNode("COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & "']") _
        Is Nothing Then
        
        ' create one DOMDocument to contain schema & request
        Set xmlRequestDoc = New FreeThreadedDOMDocument40
        xmlRequestDoc.validateOnParse = False
        xmlRequestDoc.setProperty "NewParser", True
    
        ' create BOGUS root node
        Set xmlElem = xmlRequestDoc.createElement("BOGUS")
        Set xmlRootNode = xmlRequestDoc.appendChild(xmlElem)
    
        ' create schema for TM_COMBOS view
        Set xmlElem = xmlRequestDoc.createElement("SCHEMA")
        Set xmlSchemaNode = xmlRootNode.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("COMBO")
        xmlElem.setAttribute "ENTITYTYPE", "LOGICAL"
        xmlElem.setAttribute "DATASRCE", "TM_COMBOS"
        Set xmlSchemaNode = xmlSchemaNode.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("GROUPNAME")
        xmlElem.setAttribute "DATATYPE", "STRING"
        xmlElem.setAttribute "LENGTH", "30"
        xmlSchemaNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("VALUEID")
        xmlElem.setAttribute "DATATYPE", "SHORT"
        xmlSchemaNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("VALUENAME")
        xmlElem.setAttribute "DATATYPE", "STRING"
        xmlElem.setAttribute "LENGTH", "50"
        xmlSchemaNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("VALIDATIONTYPE")
        xmlElem.setAttribute "DATATYPE", "STRING"
        xmlElem.setAttribute "LENGTH", "20"
        xmlSchemaNode.appendChild xmlElem
        
        ' create COMBOLIST request
        Set xmlElem = xmlRequestDoc.createElement("COMBOLIST")
        xmlElem.setAttribute "GROUPNAME", vstrGroupName
        Set xmlRequestNode = xmlRootNode.appendChild(xmlElem)
        
        If xmlComboListNode Is Nothing Then
            Set xmlComboListNode = m_xmldocCombos.selectSingleNode("COMBOLIST")
        End If
        
        GetRecordSetAsXML xmlRequestNode, xmlSchemaNode, xmlComboListNode
        
    End If

End Sub

Public Function GetFirstComboTextId(ByVal strComboGroup As String, ByVal strValueName As String) As String

Dim colValueId As Collection

    Set colValueId = New Collection
    GetValueIdsForValueName strComboGroup, strValueName, colValueId
    If colValueId.Count > 0 Then
        GetFirstComboTextId = CStr(colValueId.Item(1))
    Else
        GetFirstComboTextId = ""
    End If

End Function

Public Sub GetValueIdsForValueName( _
    ByVal vstrGroupName As String, _
    ByVal vstrValueName As String, _
    ByVal vcolValueIds As Collection)
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUENAME='" & _
        vstrValueName & "']"
        
    Set xmlNodeList = m_xmldocCombos.selectNodes(strPattern)
    
    For Each xmlNode In xmlNodeList
        vcolValueIds.Add GetAttributeText(xmlNode, "VALUEID")
    Next
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing

End Sub


Public Sub GetValueIdsForValidationType( _
    ByVal vstrGroupName As String, _
    ByVal vstrValidationType As String, _
    ByVal vcolValueIds As Collection)
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALIDATIONTYPE='" & _
        vstrValidationType & "']"
        
    Set xmlNodeList = m_xmldocCombos.selectNodes(strPattern)
    
    For Each xmlNode In xmlNodeList
        vcolValueIds.Add GetAttributeText(xmlNode, "VALUEID")
    Next
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing

End Sub

Public Function GetFirstComboValueId(ByVal strComboGroup As String, ByVal strValidationType As String) As String

Dim colValueId As Collection

    Set colValueId = New Collection
    GetValueIdsForValidationType strComboGroup, strValidationType, colValueId
    If colValueId.Count > 0 Then
        GetFirstComboValueId = CStr(colValueId.Item(1))
    Else
        Err.Raise eRECORDNOTFOUND, _
                  "GetFirstComboValueId", _
                  "Records for GROUP " & _
                  strComboGroup & " and VALIDATIONTYPE " & strValidationType & " not found"
    End If

End Function

Public Function GetComboText( _
    ByVal vstrGroupName As String, _
    ByVal vIntValueID As Integer) _
    As String
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID='" & _
        CStr(vIntValueID) & "']"
        
    If Not m_xmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        GetComboText = _
            GetAttributeText(m_xmldocCombos.selectSingleNode(strPattern), "VALUENAME")
    End If

End Function


