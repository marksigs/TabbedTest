VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omRiskMatrixDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omRiskMatrixDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Risk Matrix Business Data Object is called from RiskMatrixBO.
'               It implements two interfaces, these in turn call the relevant
'               schema's and return records from the database.
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   CreateDate                          Package Description
'CWP    Tuesday, May 22, 2001 1:56:02 PM    04D01.T65

Option Explicit

Implements ObjectControl
Implements IomRiskMatrixDO

Private gobjContext As ObjectContext
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub IomRiskMatrixDO_GetManualDecreaseLimitPercent(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

'------------------------------------------------------------------------------------------
' IomRiskMatrixDO_GetManualDecreaseLimitPercent header
' description:
'   Fetches specified Manual Decrease Limit Percent from Table.
' pass:
'   vxmlRequestNodeList
'       list of XML CASETASK nodes from REQUEST
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo ManualDecreaseLimitPercentExit
    
    Const strFunctionName As String = "IomRiskMatrixDO_GetManualDecreaseLimitPercent"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETMORTGAGEPRODUCTDETAILS"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
ManualDecreaseLimitPercentExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IomRiskMatrixDO_GetUserDecreaseLimitPercent(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

'------------------------------------------------------------------------------------------
' IomRiskMatrixDO_GetManualDecreaseLimitPercent header
' description:
'   Fetches specified ManualDecreaseLimitPercent from Table.
' pass:
'   vxmlRequestNodeList
'       list of XML CASETASK nodes from REQUEST
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo GetUserDecreaseLimitPercentExit
    
    Const strFunctionName As String = "IomRiskMatrixDO_GetUserDecreaseLimitPercent"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETUSERDECREASELIMITPERCENT"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetUserDecreaseLimitPercentExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

