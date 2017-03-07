VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "omRotTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omRotTxBo.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Code template for Report on Title Business Object

'Dependencies:  omRotBO, omRotDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
'
' SG    14/12/01    SYS3482 Amended IomRotTxBO_CreateReportOnTitle to ROTAddresses updated correctly.
'
'ING SPECIFIC
' JD    18/05/2006  MAR1698 UpdateReportOnTitle - the update may include a new address to create.
'-------------------------------------------------------------------------------------------------
Option Explicit

Implements IomRotTxBO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub IomRotTxBO_CreateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateReportOnTitleVbErr
    
    Const strFunctionName As String = "IomRotTxBO_CreateReportOnTitle"
    
    Dim objRotDO As IomRotDO
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlRotSolBankNode As IXMLDOMNode
    Dim xmlRotAddressNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set objRotDO = gobjContext.CreateInstance(App.Title & ".omRotDO")
    
    Set xmlNode = vxmlRequestNode.selectSingleNode("REPORTONTITLE")
    If Not xmlNode Is Nothing Then
        objRotDO.CreateReportOnTitle xmlNode, vxmlResponseNode
    End If
        
    'Extract the ROTSOLICITORSBANKACCOUNT node, create empty element if doesnt exist
    Set xmlRotSolBankNode = vxmlRequestNode.selectSingleNode(".//SOLBANKACCT")
    If xmlRotSolBankNode Is Nothing Then
        Set xmlRotSolBankNode = xmlDoc.createElement("SOLBANKACCT")
    End If
    adoPopulateChildKeys xmlNode, xmlRotSolBankNode
    objRotDO.CreateROTSolBankAccount xmlRotSolBankNode, vxmlResponseNode
    
    'Extract the ADDRESS node, create empty element if doesnt exist
    Set xmlAddressNode = vxmlRequestNode.selectSingleNode(".//ADDRESS")
    If xmlAddressNode Is Nothing Then
        Set xmlAddressNode = xmlDoc.createElement("ADDRESS")
    End If
    adoPopulateChildKeys xmlNode, xmlAddressNode
    objRotDO.CreateAddress xmlAddressNode, vxmlResponseNode
        
    'Create ROTADDRESSES element which is populated with PK values (set out in schema)
    'Requires AddressGuid from Address element & RotGuid from ReportonTitle element
    Set xmlRotAddressNode = xmlDoc.createElement("ROTADDRESSES")
    adoPopulateChildKeys xmlAddressNode, xmlRotAddressNode
    adoPopulateChildKeys xmlNode, xmlRotAddressNode
    
    'SG 14/12/01 SYS3482 - Need to add ADDRESSTYPE into ROTAddresses.
    xmlCopyAttribute xmlAddressNode, xmlRotAddressNode, "ADDRESSTYPE"
    
    objRotDO.CreateROTAddress xmlRotAddressNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
CreateReportOnTitleExit:
    Set objRotDO = Nothing
    Set xmlNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
CreateReportOnTitleVbErr:

    gobjContext.SetAbort
    GoTo CreateReportOnTitleExit

End Sub

Private Sub IomRotTxBO_UpdateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateReportOnTitleVbErr
    
    Const strFunctionName As String = "IomRotTxBO_UpdateReportOnTitle"
    
    Dim objRotDO As IomRotDO
    Dim xmlNode As IXMLDOMNode
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    Dim strRotGuid As String
    Dim bCreateAddress As Boolean
    
    Set objRotDO = gobjContext.CreateInstance(App.Title & ".omRotDO")
        
    Set xmlNode = vxmlRequestNode.selectSingleNode("REPORTONTITLE")
    
    'Get Applicationnumber, Applicationfactfindnumber and RotGuid
    strAppNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    strRotGuid = xmlGetMandatoryAttributeText(xmlNode, "ROTGUID")
        
    'Update Reportontitle
    If Not xmlNode Is Nothing Then
        objRotDO.UpdateReportOnTitle xmlNode, vxmlResponseNode
    End If
    
    'Update RotSolicitorsBankAccount
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//SOLBANKACCT")
    If Not xmlNode Is Nothing Then
        xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strAppNumber
        xmlSetAttributeValue xmlNode, "APPLICATIONFACTFINDNUMBER", strAppFFNumber
        xmlSetAttributeValue xmlNode, "ROTGUID", strRotGuid
        objRotDO.UpdateROTSolBankAccount xmlNode, vxmlResponseNode
    End If
    
    'Update Address and ROTAddresses table
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//ADDRESS")
    If Not xmlNode Is Nothing Then
        'JD MAR1698 if we have an addressguid then update otherwise create
        bCreateAddress = False
        If xmlGetAttributeText(xmlNode, "ADDRESSGUID") <> "" Then
            objRotDO.UpdateAddress xmlNode, vxmlResponseNode
        Else
            objRotDO.CreateAddress xmlNode, vxmlResponseNode
            bCreateAddress = True
        End If
        
        'JR SYS2048
        'ROTGUID is PK of RotAddresses table
        xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strAppNumber
        xmlSetAttributeValue xmlNode, "APPLICATIONFACTFINDNUMBER", strAppFFNumber
        xmlSetAttributeValue xmlNode, "ROTGUID", strRotGuid
        If bCreateAddress Then 'MAR1698
            objRotDO.CreateROTAddress xmlNode, vxmlResponseNode
        Else
            objRotDO.UpdateROTAddress xmlNode, vxmlResponseNode
        End If
    End If
        
    gobjContext.SetComplete
    
UpdateReportOnTitleExit:
    Set objRotDO = Nothing

    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub
    
UpdateReportOnTitleVbErr:

    gobjContext.SetAbort
    GoTo UpdateReportOnTitleExit

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





