Attribute VB_Name = "comboAssistEx"
Option Explicit

'----------------------------------------------------------------------
'BMIDS Change History
'Prog       Date            Ref
'GD         10/07/2002      BMIDS00165 - New Function IsValidationTypeInValidationList
'BS         13/05/2003      BM0310 Amended test for nothing found in IsValidationTypeInValidationList
'----------------------------------------------------------------------
'MARS Change History
'Prog       Date            Ref
'GHun       13/03/2006      MAR1300 Merge core changes
'PSC        20/04/2006      MAR1649 Add GetValidationTypesForValueId
'----------------------------------------------------------------------

Private gxmldocCombos As FreeThreadedDOMDocument40

Public Function GetComboText( _
    ByVal vstrGroupName As String, _
    ByVal vintValueId As Integer) _
    As String
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID='" & _
        CStr(vintValueId) & "']"
        
    If Not gxmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        GetComboText = _
            xmlGetAttributeText(gxmldocCombos.selectSingleNode(strPattern), "VALUENAME")
    End If

End Function

Public Function IsValidationType( _
    ByVal vstrGroupName As String, _
    ByVal vintValueId As Integer, _
    ByVal vstrValidationType As String) _
    As Boolean
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & _
        "' and @VALUEID='" & CStr(vintValueId) & _
        "' and @VALIDATIONTYPE='" & vstrValidationType & "']"
        
    If Not gxmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        IsValidationType = True
    Else
        IsValidationType = False
    End If

End Function

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
        
    Set xmlNodeList = gxmldocCombos.selectNodes(strPattern)
    
    For Each xmlNode In xmlNodeList
        vcolValueIds.Add xmlGetAttributeText(xmlNode, "VALUEID")
    Next
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing

End Sub

Public Function GetValidationTypeForValueID(ByVal vstrGroupName As String, ByVal vintValueId As Integer) As String
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID= " & vintValueId & "]"
        
    If Not gxmldocCombos.selectSingleNode(strPattern) Is Nothing Then
        GetValidationTypeForValueID = _
            xmlGetAttributeText(gxmldocCombos.selectSingleNode(strPattern), "VALIDATIONTYPE")
    End If
    
End Function

Private Sub LoadComboValidationGroup( _
    ByVal vstrGroupName As String)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlComboListNode As IXMLDOMNode
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlSchemaNode As IXMLDOMNode
    
    If gxmldocCombos Is Nothing Then
        Set gxmldocCombos = New FreeThreadedDOMDocument40
        gxmldocCombos.validateOnParse = False
        gxmldocCombos.setProperty "NewParser", True
        Set xmlElem = gxmldocCombos.createElement("COMBOVALIDATIONLIST")
        Set xmlComboListNode = gxmldocCombos.appendChild(xmlElem)
    End If
    
    If gxmldocCombos.selectSingleNode("COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & "']") _
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
        
        Set xmlElem = xmlRequestDoc.createElement("COMBOVALIDATION")
        xmlElem.setAttribute "ENTITYTYPE", "LOGICAL"
        xmlElem.setAttribute "DATASRCE", "TM_COMBOVALIDATION"
        Set xmlSchemaNode = xmlSchemaNode.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("GROUPNAME")
        xmlElem.setAttribute "DATATYPE", "STRING"
        xmlElem.setAttribute "LENGTH", "30"
        xmlSchemaNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("VALUEID")
        xmlElem.setAttribute "DATATYPE", "SHORT"
        xmlSchemaNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("VALIDATIONTYPE")
        xmlElem.setAttribute "DATATYPE", "STRING"
        xmlElem.setAttribute "LENGTH", "20"
        xmlSchemaNode.appendChild xmlElem
        
        ' create COMBOLIST request
        Set xmlElem = xmlRequestDoc.createElement("COMBOVALIDATIONLIST")
        xmlElem.setAttribute "GROUPNAME", vstrGroupName
        Set xmlRequestNode = xmlRootNode.appendChild(xmlElem)
        
        If xmlComboListNode Is Nothing Then
            Set xmlComboListNode = gxmldocCombos.selectSingleNode("COMBOVALIDATIONLIST")
        End If
        
        adoGetRecordSetAsXML xmlRequestNode, xmlSchemaNode, xmlComboListNode
        
    End If

    Set xmlElem = Nothing
    Set xmlComboListNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRootNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlSchemaNode = Nothing

End Sub

Private Sub LoadComboGroup( _
    ByVal vstrGroupName As String)
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlComboListNode As IXMLDOMNode
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlSchemaNode As IXMLDOMNode
    
    If gxmldocCombos Is Nothing Then
        Set gxmldocCombos = New FreeThreadedDOMDocument40
        gxmldocCombos.validateOnParse = False
        gxmldocCombos.setProperty "NewParser", True
        Set xmlElem = gxmldocCombos.createElement("COMBOLIST")
        Set xmlComboListNode = gxmldocCombos.appendChild(xmlElem)
    End If
    
    If gxmldocCombos.selectSingleNode("COMBOLIST/COMBO[@GROUPNAME='" & vstrGroupName & "']") _
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
            Set xmlComboListNode = gxmldocCombos.selectSingleNode("COMBOLIST")
        End If
        
        adoGetRecordSetAsXML xmlRequestNode, xmlSchemaNode, xmlComboListNode
        
    End If
    
    Set xmlElem = Nothing
    Set xmlComboListNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRootNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlSchemaNode = Nothing

End Sub

Public Function GetFirstComboValueId(ByVal strComboGroup As String, ByVal strValidationType As String) As String

    Dim colValueId As Collection

    Set colValueId = New Collection
    GetValueIdsForValidationType strComboGroup, strValidationType, colValueId
    If colValueId.Count > 0 Then
        GetFirstComboValueId = CStr(colValueId.Item(1))
    Else
        errThrowError "GetFirstComboValueId", oeRecordNotFound
    End If

    Set colValueId = Nothing

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
        
    Set xmlNodeList = gxmldocCombos.selectNodes(strPattern)
    
    For Each xmlNode In xmlNodeList
        vcolValueIds.Add xmlGetAttributeText(xmlNode, "VALUEID")
    Next
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing

End Sub

Public Function GetFirstComboTextId(ByVal strComboGroup As String, ByVal strValueName As String) As String

    Dim colValueId As Collection

    Set colValueId = New Collection
    GetValueIdsForValueName strComboGroup, strValueName, colValueId
    If colValueId.Count > 0 Then
        GetFirstComboTextId = CStr(colValueId.Item(1))
    Else
        GetFirstComboTextId = vbNullString
    End If

    Set colValueId = Nothing
End Function

Public Function IsValidationTypeInValidationList(ByVal vstrGroupName As String, ByVal vstrValidationType As String, ByVal vintValueId As Integer) As Boolean
'GD added 09/07/02
    Dim strPattern As String
    Dim xmlValidationList As IXMLDOMNodeList
    Dim blnResult As Boolean
    
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID= " & vintValueId & " and @VALIDATIONTYPE = '" & vstrValidationType & "']"
    Set xmlValidationList = gxmldocCombos.selectNodes(strPattern)
    
    'BS BM0310 13/05/03
    'If Not xmlValidationList Is Nothing Then
    If xmlValidationList.length > 0 Then
        blnResult = True
    Else
        blnResult = False
    End If
    IsValidationTypeInValidationList = blnResult
    Set xmlValidationList = Nothing
End Function

' PSC 20/04/2006 MAR1649 - Start
Public Sub GetValidationTypesForValueID(ByVal vstrGroupName As String, ByVal vintValueId As String, ByVal vcolValidationTypes As Collection)
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim strPattern As String
        
    LoadComboGroup vstrGroupName
    
    strPattern = _
        "COMBOLIST/COMBO[@GROUPNAME='" & _
        vstrGroupName & _
        "' and @VALUEID='" & _
        CStr(vintValueId) & "']"
        
    Set xmlNodeList = gxmldocCombos.selectNodes(strPattern)
    
    For Each xmlNode In xmlNodeList
        vcolValidationTypes.Add xmlGetAttributeText(xmlNode, "VALIDATIONTYPE")
    Next
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing

End Sub
' PSC 20/04/2006 MAR1649 - End
