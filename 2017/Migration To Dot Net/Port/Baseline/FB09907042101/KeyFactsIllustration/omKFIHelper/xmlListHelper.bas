Attribute VB_Name = "xmlListHelper"
'********************************************************************************
'** Module:         xmlListHelper
'** Created by:     Andy Maggs
'** Date:           20/05/2004
'** Description:    Provides functionality manipulating and extracting items from
'**                 XML node lists.
'********************************************************************************
Option Explicit
Option Compare Text

Private Const mcstrModuleName As String = "xmlListHelper"

'********************************************************************************
'** Function:       HasElementWithIntAttribValue
'** Created by:     Andy Maggs
'** Date:           20/05/2004
'** Description:    Gets whether there is an element in the list having the
'**                 specified attribute value.
'** Parameters:     vxmlList - the list to search.
'**                 vstrAttribName - the name of the attribute whose value is to
'**                 be tested.
'**                 vintValue - the value to test for.
'** Returns:        True if there is an element with the specified value, else
'**                 False.
'** Errors:         None Expected
'********************************************************************************
Public Function HasElementWithIntAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String, ByVal vintValue As Integer) As Boolean
    Const cstrFunctionName As String = "HasElementWithIntAttribValue"
    Dim xmlItem As IXMLDOMNode
    Dim intValue As Integer
    Dim blnResult As Boolean

    On Error GoTo ErrHandler

    For Each xmlItem In vxmlList
        intValue = xmlGetAttributeAsInteger(xmlItem, vstrAttribName)
        If intValue = vintValue Then
            '*-this is the item we are looking for
            blnResult = True
            Exit For
        End If
    Next xmlItem
    
    HasElementWithIntAttribValue = blnResult

    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetElementWithIntAttribValue
'** Created by:     Andy Maggs
'** Date:           20/05/2004
'** Description:    Gets the element in the list having the specified attribute
'**                 value, if there is one.
'** Parameters:     vxmlList - the list to search.
'**                 vstrAttribName - the name of the attribute whose value is to
'**                 be tested.
'**                 vintValue - the value to search for.
'** Returns:        The element if there is one, else Nothing.
'** Errors:         None Expected
'********************************************************************************
Public Function GetElementWithIntAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String, ByVal vintValue As Integer) As IXMLDOMNode
    Const cstrFunctionName As String = "GetElementWithIntAttribValue"
    Dim xmlItem As IXMLDOMNode
    Dim intValue As Integer

    On Error GoTo ErrHandler

    For Each xmlItem In vxmlList
        intValue = xmlGetAttributeAsInteger(xmlItem, vstrAttribName)
        If intValue = vintValue Then
            '*-this is the item we are looking for
            Exit For
        End If
    Next xmlItem
    
    Set GetElementWithIntAttribValue = xmlItem

    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetElementWithMaxLngAttribValue
'** Created by:     Andy Maggs
'** Date:           19/05/2004
'** Description:    Gets the element having the specified attribute with the
'**                 maximum value in the list.
'** Parameters:     vxmlList - the list of items
'**                 vstrAttribName - the name of the attribute to get the value
'**                 for on each item.
'** Returns:        The element with the maximum value of the specified attribute.
'** Errors:         None Expected
'********************************************************************************
Public Function GetElementWithMaxLngAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As IXMLDOMNode
    Const cstrFunctionName As String = "GetElementWithMaxLngAttribValue"
    Dim lngMax As Long
    Dim lngValue As Long
    Dim xmlItem As IXMLDOMNode
    Dim xmlMax As IXMLDOMNode

    On Error GoTo ErrHandler

    For Each xmlItem In vxmlList
        lngValue = xmlGetAttributeAsLong(xmlItem, vstrAttribName)
        If lngValue > lngMax Then
            lngMax = lngValue
            Set xmlMax = xmlItem
        End If
    Next xmlItem
    
    Set GetElementWithMaxLngAttribValue = xmlMax

    Set xmlItem = Nothing
    Set xmlMax = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    Set xmlMax = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetMaxLngAttribValue
'** Created by:     Andy Maggs
'** Date:           19/05/2004
'** Description:    Gets the maximum value from the specified attribute of each
'**                 item in the list.
'** Parameters:     vxmlList - the list of items
'**                 vstrAttribName - the name of the attribute to get the value
'**                 for on each item.
'** Returns:        The maximum value of the specified attribute.
'** Errors:         None Expected
'********************************************************************************
Public Function GetMaxLngAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As Long
    Const cstrFunctionName As String = "GetMaxLngAttribValue"
    Dim lngMax As Long
    Dim lngValue As Long
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler

    For Each xmlItem In vxmlList
        lngValue = xmlGetAttributeAsLong(xmlItem, vstrAttribName)
        If lngValue > lngMax Then
            lngMax = lngValue
        End If
    Next xmlItem
    
    GetMaxLngAttribValue = lngMax
    
    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetMaxDblAttribValue
'** Created by:     Andy Maggs
'** Date:           19/05/2004
'** Description:    Gets the maximum value from the specified attribute of each
'**                 item in the list.
'** Parameters:     vxmlList - the list of items
'**                 vstrAttribName - the name of the attribute to get the value
'**                 for on each item.
'** Returns:        The maximum value of the specified attribute.
'** Errors:         None Expected
'********************************************************************************
Public Function GetMaxDblAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As Double
    Const cstrFunctionName As String = "GetMaxDblAttribValue"
    Dim dblMax As Double
    Dim dblValue As Double
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler

    For Each xmlItem In vxmlList
        dblValue = xmlGetAttributeAsDouble(xmlItem, vstrAttribName)
        If dblValue > dblMax Then
            dblMax = dblValue
        End If
    Next xmlItem
    
    GetMaxDblAttribValue = dblMax
    
    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetElementWithMinIntAttribValue
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Gets the item having the lowest value for the specified
'**                 integer attribute.
'** Parameters:     vxmlList - the list to search.
'**                 vstrNumAttribName - the name of the attribute containing the
'**                 integer value.
'** Returns:        The item with the lowest value.
'** Errors:         None Expected
'********************************************************************************
Public Function GetElementWithMinIntAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As IXMLDOMNode
    Const cstrFunctionName As String = "GetElementWithMinIntAttribValue"
    Dim intMin As Integer
    Dim intValue As Integer
    Dim xmlItem As IXMLDOMNode
    Dim xmlMin As IXMLDOMNode

    On Error GoTo ErrHandler

    intMin = 32767
    For Each xmlItem In vxmlList
        intValue = xmlGetAttributeAsInteger(xmlItem, vstrAttribName)
        If intValue < intMin Then
            intMin = intValue
            Set xmlMin = xmlItem
        End If
    Next xmlItem
    
    Set GetElementWithMinIntAttribValue = xmlMin

    Set xmlItem = Nothing
    Set xmlMin = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    Set xmlMin = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetMinIntAttribValue
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Gets the item having the lowest value for the specified
'**                 integer attribute.
'** Parameters:     vxmlList - the list to search.
'**                 vstrAttribName - the name of the attribute containing the
'**                 integer value.
'** Returns:        The item with the lowest value.
'** Errors:         None Expected
'********************************************************************************
Public Function GetMinIntAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As Integer
    Const cstrFunctionName As String = "GetMinIntAttribValue"
    Dim intMin As Integer
    Dim intValue As Integer
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler

    intMin = 32767
    For Each xmlItem In vxmlList
        intValue = xmlGetAttributeAsInteger(xmlItem, vstrAttribName)
        If (intMin = 0) Or (intValue < intMin) Then
            intMin = intValue
        End If
    Next xmlItem
    
    GetMinIntAttribValue = intMin

    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       GetMinLngAttribValue
'** Created by:     Andy Maggs
'** Date:           19/04/2004
'** Description:    Gets the item having the lowest value for the specified
'**                 long attribute.
'** Parameters:     vxmlList - the list to search.
'**                 vstrAttribName - the name of the attribute containing the
'**                 long value.
'** Returns:        The lowest value.
'** Errors:         None Expected
'********************************************************************************
Public Function GetMinLngAttribValue(ByVal vxmlList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As Integer
    Const cstrFunctionName As String = "GetMinLngAttribValue"
    Dim lngMin As Long
    Dim lngValue As Long
    Dim xmlItem As IXMLDOMNode

    On Error GoTo ErrHandler

    lngMin = 2147483647
    For Each xmlItem In vxmlList
        lngValue = xmlGetAttributeAsLong(xmlItem, vstrAttribName)
        If (lngMin = 0) Or (lngValue < lngMin) Then
            lngMin = lngValue
        End If
    Next xmlItem
    
    GetMinLngAttribValue = lngMin

    Set xmlItem = Nothing
Exit Function
ErrHandler:
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, mcstrModuleName
End Function

'********************************************************************************
'** Function:       TotalDblAttributeValues
'** Created by:     Andy Maggs
'** Date:           29/06/2004
'** Description:    Gets the total of the values of the specified attribute in
'**                 each element in the list.
'** Parameters:     vxmlNodeList - the list to iterate.
'**                 vstrAttribName - the name of the attribute.
'** Returns:        The total as a double.
'** Errors:         None Expected
'********************************************************************************
Public Function TotalDblAttributeValues(ByVal vxmlNodeList As IXMLDOMNodeList, _
        ByVal vstrAttribName As String) As Double
    Const cstrFunctionName As String = "TotalAttributeValues"
    Dim xmlItem As IXMLDOMNode
    Dim dblValue As Double
    
    On Error GoTo ErrHandler
    
    For Each xmlItem In vxmlNodeList
        dblValue = dblValue + xmlGetAttributeAsDouble(xmlItem, vstrAttribName)
    Next xmlItem
    
    TotalDblAttributeValues = dblValue
    
    Set xmlItem = Nothing
    Exit Function
ErrHandler:
    Set xmlItem = Nothing
    errCheckError cstrFunctionName, mcstrModuleName
End Function

