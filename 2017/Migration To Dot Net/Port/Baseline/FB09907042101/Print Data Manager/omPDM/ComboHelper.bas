Attribute VB_Name = "ComboHelper"
'Workfile:      ComboHelper.bas
'Copyright:     Copyright © 2001 Marlborough Stirling
'Description:
'Dependencies:  Add any other dependent components
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    12/03/04    BBG42 Add GetValidationTypesForValueId
'TW     18/10/2005  MAR223
'------------------------------------------------------------------------------------------

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
        "' and  @VALUEID='" & CStr(vIntValueID) & _
        "' and  @VALIDATIONTYPE='" & vstrValidationType & "']"
    If Not m_xmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        IsValidationType = True
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
        Set xmlElem = m_xmldocCombos.createElement("COMBOLIST")
        Set xmlComboListNode = m_xmldocCombos.appendChild(xmlElem)
    End If
    If m_xmldocCombos.selectSingleNode("COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & "']") _
        Is Nothing Then
        ' create one FreeThreadedDOMDocument40 to contain schema & request
        Set xmlRequestDoc = New FreeThreadedDOMDocument40
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
        "'  and  @VALUENAME='" & _
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
        "' and  @VALIDATIONTYPE='" & _
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
        "' and  @VALUEID='" & _
        CStr(vIntValueID) & "']"
    If Not m_xmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        GetComboText = _
            GetAttributeText(m_xmldocCombos.selectSingleNode(strPattern), "VALUENAME")
    End If
End Function
Public Function GetValidationTypeForValueId( _
    ByVal vstrGroupName As String, _
    ByVal vstrValueId As String) As String
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID='" & _
        vstrValueId & "']"
    If Not m_xmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        GetValidationTypeForValueId = _
            GetAttributeText(m_xmldocCombos.selectSingleNode(strPattern), "VALIDATIONTYPE")
    End If
End Function

' PSC 12/03/2004 BBG42 - Start
Public Function GetValidationTypesForValueId( _
    ByVal vstrGroupName As String, _
    ByVal vstrValueId As String) As Collection
    
    Dim strPattern As String
    Dim colValidationTypes As Collection
    Dim xmlValidation As IXMLDOMNode
    Dim xmlValidationList As IXMLDOMNodeList
    
    Set colValidationTypes = New Collection
        
    LoadComboGroup vstrGroupName
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID='" & _
        vstrValueId & "']/@VALIDATIONTYPE"
        
    Set xmlValidationList = m_xmldocCombos.selectNodes(strPattern)
    
    If xmlValidationList.length > 0 Then
        For Each xmlValidation In xmlValidationList
            colValidationTypes.Add (xmlValidation.Text)
        Next
    End If
    
    Set GetValidationTypesForValueId = colValidationTypes
    
End Function
' PSC 12/03/2004 BBG42 - End

