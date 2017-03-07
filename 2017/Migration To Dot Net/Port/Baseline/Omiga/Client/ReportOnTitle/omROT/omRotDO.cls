VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omRotDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IomRotDO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub IomRotDO_CreateAddress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateAddressExit
    
Const strFunctionName As String = "IomRotDO_CreateAddress"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("ADDRESS")
    
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateAddressExit:
    Set xmlSchemaNode = Nothing
            
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IomRotDO_CreateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateReportOnTitleExit
    
Const strFunctionName As String = "IomRotDO_CreateReportOnTitle"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("REPORTONTITLE")
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateReportOnTitleExit:
    Set xmlSchemaNode = Nothing
            
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IomRotDO_CreateROTAddress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateROTAddressExit
    
Const strFunctionName As String = "IomRotDO_CreateROTAddress"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("ROTADDRESSES")
    
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateROTAddressExit:
    Set xmlSchemaNode = Nothing
            
    errCheckError strFunctionName, TypeName(Me)


End Sub

Private Sub IomRotDO_CreateROTSolBankAccount(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateROTSolBankAccountExit
    
Const strFunctionName As String = "IomRotDO_CreateROTSolBankAccount"
Dim xmlSchemaNode As IXMLDOMNode
    
    Set xmlSchemaNode = adoGetSchema("SOLBANKACCT")
    
    adoCreate vxmlRequestNode, xmlSchemaNode
    gobjContext.SetComplete
    
CreateROTSolBankAccountExit:
    Set xmlSchemaNode = Nothing
            
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IomRotDO_GetReportOnTitleData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo GetReportOnTitleDataExit
    
Const strFunctionName As String = "IomRotDO_GetReportOnTitleData"
Dim xmlRotNode As IXMLDOMNode
    
    'Extract REPORTONTITLE node
    If vxmlRequestNode.nodeName = "REPORTONTITLE" Then
        Set xmlRotNode = vxmlRequestNode
    Else
        Set xmlRotNode = vxmlRequestNode.selectSingleNode("REPORTONTITLE")
    End If
    
    adoGetAsXML xmlRotNode, vxmlResponseNode, "REPORTONTITLE"
    gobjContext.SetComplete
    
GetReportOnTitleDataExit:
    Set xmlRotNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
        
End Sub

Private Sub IomRotDO_GetRotAndSolBankAccount(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo GetRotAndSolBankAccountExit
    
Const strFunctionName As String = "IomRotDO_GetRotAndSolBankAccount"
Dim xmlRotNode As IXMLDOMNode
    
    'Extract REPORTONTITLE node
    If vxmlRequestNode.nodeName = "REPORTONTITLE" Then
        Set xmlRotNode = vxmlRequestNode
    Else
        Set xmlRotNode = vxmlRequestNode.selectSingleNode("REPORTONTITLE")
    End If
    
    adoGetAsXML xmlRotNode, vxmlResponseNode, "GETROTSOLBANKACCOUNT"
    gobjContext.SetComplete
    
GetRotAndSolBankAccountExit:
    Set xmlRotNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
        
End Sub

Private Sub IomRotDO_GetRotAddresses(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo GetRotAddressesExit
    
Const strFunctionName As String = "IomRotDO_GetRotAddresses"
Dim xmlRotNode As IXMLDOMNode
    
    'Extract REPORTONTITLE node
    If vxmlRequestNode.nodeName = "REPORTONTITLE" Then
        Set xmlRotNode = vxmlRequestNode
    Else
        Set xmlRotNode = vxmlRequestNode.selectSingleNode("GETROTADDRESSES")
    End If
    
    adoGetAsXML xmlRotNode, vxmlResponseNode, "GETROTADDRESSES"
    gobjContext.SetComplete
    
GetRotAddressesExit:
    Set xmlRotNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
        
End Sub

Private Sub IomRotDO_UpdateAddress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateAddressExit
    
Const strFunctionName As String = "IomRotDO_UpdateAddress"
    
    adoUpdateFromNode vxmlRequestNode, "ADDRESS"
    gobjContext.SetComplete
            
UpdateAddressExit:

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IomRotDO_UpdateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateReportOnTitleExit
    
Const strFunctionName As String = "IomRotDO_UpdateReportOnTitle"
    
    adoUpdateFromNode vxmlRequestNode, "REPORTONTITLE"
    gobjContext.SetComplete
            
UpdateReportOnTitleExit:

    errCheckError strFunctionName, TypeName(Me)
        
End Sub

Private Sub IomRotDO_UpdateROTAddress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateROTAddressExit

Const strFunctionName As String = "IomRotDO_UpdateROTAddress"

    adoUpdateFromNode vxmlRequestNode, "ROTADDRESSES"
    gobjContext.SetComplete

UpdateROTAddressExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IomRotDO_UpdateROTSolBankAccount(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateROTSolBankAccountExit
    
Const strFunctionName As String = "IomRotDO_UpdateROTSolBankAccount"
    
    adoUpdateFromNode vxmlRequestNode, "SOLBANKACCT"
    gobjContext.SetComplete
            
UpdateROTSolBankAccountExit:

    errCheckError strFunctionName, TypeName(Me)

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

