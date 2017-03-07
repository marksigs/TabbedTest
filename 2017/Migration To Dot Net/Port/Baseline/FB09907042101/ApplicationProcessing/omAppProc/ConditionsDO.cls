VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ConditionsDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ConditionsDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:

'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'STB    22/04/02    SYS2175 Do not raise an error in FindConditionsList if no records are found.
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IConditionsDO

Private gobjContext As ObjectContext

'BMIDS History:
'
'Prog   Date        Description
'GD     09/09/02    BMIDS00313 - APWP2 - BM004 Added FindNonProductConditionsList
'------------------------------------------------------------------------------------------
Private Sub IConditionsDO_CreateApplicationConditions( _
                ByVal vxmlRequestNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateApplicationConditionsExit
    
    Const cstrFunctionName As String = "IConditionsDO_CreateApplicationConditions"
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    If vxmlRequestNode.nodeName = "APPLICATIONCONDITIONS" Then
        ' Assign current data & time to column 'DATETIME'
        xmlSetSysDateToNodeAttrib vxmlRequestNode, "DATETIME"
        adoCreateFromNode vxmlRequestNode, "APPLICATIONCONDITIONS"
    Else
        Set xmlNodeList = vxmlRequestNode.selectNodes(".//APPLICATIONCONDITIONS")
        ' Assign current date & time to column 'DATETIME' in all nodes
        xmlSetSysDateToNodeListAttribs xmlNodeList, "DATETIME"
        adoCreateFromNodeList xmlNodeList, "APPLICATIONCONDITIONS"
    End If
    
    gobjContext.SetComplete

CreateApplicationConditionsExit:

    Set xmlNodeList = Nothing
    
    If Err.Number <> 0 Then
        
        Dim lngErrNo As Long, strErrDesc As String, strErrSrc As String
        lngErrNo = Err.Number
        strErrDesc = Err.Description
        strErrSrc = Err.Source
        
        If adoGetOmigaNumberForDatabaseError(Err.Description) = oeDuplicateKey Then
            errThrowError cstrFunctionName, 4811
        Else
            Err.Raise lngErrNo, strErrSrc, strErrDesc
        End If
    Else
        errCheckError cstrFunctionName, TypeName(Me)
    End If

End Sub

Private Sub IConditionsDO_DeleteAllApplicationConditions( _
            ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo DeleteVbExit
    
    Const strFunctionName As String = "IConditionsDO_DeleteAllApplicationConditions"
    Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "APPLICATIONCONDITIONS", False)
    gobjContext.SetComplete
    
DeleteVbExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsDO_DeleteApplicationConditions( _
        ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)


On Error GoTo DeleteVbExit
    
    Const strFunctionName As String = "IConditionsDO_DeleteApplicationConditions"
    Dim lngRecordsAffected As Long

    lngRecordsAffected = adoDeleteFromNode(vxmlRequestNode, "APPLICATIONCONDITIONS")
    gobjContext.SetComplete
    
DeleteVbExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsDO_FindConditionsList( _
            ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo FindConditionsListExit
    
    Const strFunctionName As String = "IConditionsDO_FindConditionsList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CONDITIONS"
    
    'SYS2715 - Do not raise an error if no records are returned.
'    If vxmlResponseNode.hasChildNodes = False Then
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
    'SYS2715 - End.
    
FindConditionsListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)


End Sub

Private Sub IConditionsDO_UpdateApplicationConditions( _
            ByVal vxmlRequestNode As IXMLDOMNode, _
            ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateApplicationConditionsExit
    
    Const cstrFunctionName As String = "IConditionsDO_UpdateApplicationConditions"
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    If vxmlRequestNode.nodeName = "APPLICATIONCONDITIONS" Then
        ' Assign current data & time to column 'DATETIME'
        xmlSetSysDateToNodeAttrib vxmlRequestNode, "DATETIME"
        adoUpdateFromNode vxmlRequestNode, "APPLICATIONCONDITIONS"
    Else
        Set xmlNodeList = vxmlRequestNode.selectNodes(".//APPLICATIONCONDITIONS")
        ' Assign current date & time to column 'DATETIME' in all nodes
        xmlSetSysDateToNodeListAttribs xmlNodeList, "DATETIME"
        adoUpdateFromNodeList xmlNodeList, "APPLICATIONCONDITIONS"
    End If
    
    gobjContext.SetComplete

UpdateApplicationConditionsExit:

    Set xmlNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub IConditionsDO_FindApplicationConditionsList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo FindApplicationConditionsListExit
    
    Const strFunctionName As String = "IConditionsDO_FindApplicationConditionsList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "APPLICATIONCONDITIONS"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindApplicationConditionsListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IConditionsDO_FindNonProductConditionsList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)


    On Error GoTo FindNonProductConditionsListExit
    
    Const strFunctionName As String = "IConditionsDO_FindNonProductConditionsList"
    
    Dim strFilter As String
    Dim colValueIDs As Collection
    Dim strValueID As String
    Set colValueIDs = New Collection
    'Get the value id for ConditionType where the validation type is 'P' - Product Special Condition

    GetValueIdsForValidationType "ConditionType", "P", colValueIDs
    'If no records found in ConditionType with validation type of 'P', raise an error.
    If colValueIDs.Count <> 0 Then
        strValueID = colValueIDs.Item(1)
    Else
        errThrowError strFunctionName, oerecordNotfound, "Can't find any value IDs for ConditionType of 'P'"
    End If
    
    'SELECT * FROM CONDITIONS where .... and CONDITIONTYPE <> 30
    
    strFilter = "CONDITIONTYPE <> " & strValueID
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CONDITIONS", strFilter
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oerecordNotfound
    End If
    
FindNonProductConditionsListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub
