VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AddressDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      AddressDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:
'               All other that use address (Ex: CustomerAddress)
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'SR    02/11/99  Created
'LD    19/06/01  SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'------------------------------------------------------------------------------------------

Option Explicit

Implements IAddressDO

Private Sub IAddressDO_CreateAddress(ByVal vxmlTableElement As IXMLDOMElement, _
                                     ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateAddress an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateAddressVbErr

    Const strFunctionName As String = "IAddressDO_CreateAddress"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IAddressClassDef
    
    Dim objIDOAssist As IDOAssist

    Dim xmlGUIDNode As IXMLDOMNode
    Dim strGuid As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Set key for address
    '
    Set xmlGUIDNode = vxmlTableElement.selectSingleNode("ADDRESSGUID")
    If xmlGUIDNode Is Nothing Then
        ' No GUID tag was found so CreateAddress it now
        Set xmlGUIDNode = vxmlTableElement.appendChild( _
                            vxmlTableElement.ownerDocument.createElement("ADDRESSGUID"))
    End If

    ' Generate the GUID
    strGuid = CreateGUID()
    xmlGUIDNode.Text = strGuid

    Set objIThirdPartyClassDef = New AddressClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadAddressData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    ' Append generated GUID to keys XML
    rxmlKeysNode.appendChild xmlGUIDNode
  
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlGUIDNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateAddressVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlGUIDNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IAddressDO_DeleteAddress(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteAddress a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAddressVbErr

    Const strFunctionName As String = "IAddressDO_DeleteAddress"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IAddressClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIThirdPartyClassDef = New AddressClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadAddressData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteAddressVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeChildRecordsFound Then
        ' The THIRDPARTY/NAMEANDADDRESSDIRECTORY record still exists for this address
        Resume Next
    End If

    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IAddressDO_GetAddress(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetAddressVbErr

    Const strFunctionName As String = "IAddressDO_GetAddress"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IAddressClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New AddressClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadAddressData
    Set IAddressDO_GetAddress = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetAddressVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IAddressDO_IsEmptyAddress(ByVal vxmlElement As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Check if a customer address is empty
' pass:         vstrXMLRequest  address to check
'
' return:       IAddressDO_IsEmptyAddress True if address is empty
'                              False if address has data
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsEmptyAddressVbErr
    
    Const strFunctionName As String = "IAddressDO_IsEmptyAddress"
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlAddressNode As IXMLDOMNode
    
    Dim blnIsEmpty As Boolean
    blnIsEmpty = False

    If vxmlElement.nodeName = "ADDRESS" Then
        Set xmlAddressNode = vxmlElement
    Else
        Set xmlAddressNode = vxmlElement.selectSingleNode("ADDRESS")
    End If

    If xmlAddressNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ADDRESS tag is missing"
    End If
    
    If Len(objXmlAssist.GetTagValue(xmlAddressNode, "POSTCODE")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "FLATNUMBER")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "BUILDINGORHOUSENAME")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "BUILDINGORHOUSENUMBER")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "STREET")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "DISTRICT")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "TOWN")) = 0 And _
       Len(objXmlAssist.GetTagValue(xmlAddressNode, "COUNTY")) = 0 Then
       'APS 06/09/99 - UNIT TEST REF 19 - Removed the check for the Country
       'as this maybe populated by default
       'And Len(objXMLAssist.GetTagValue(xmlAddressNode, "COUNTRY")) = 0 Then
        blnIsEmpty = True
    End If
    
    IAddressDO_IsEmptyAddress = blnIsEmpty
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlAddressNode = Nothing
        
    Exit Function
    
IsEmptyAddressVbErr:
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlAddressNode = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Sub IAddressDO_UpdateAddress(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateAddress a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAddressVbErr

    Const strFunctionName As String = "IAddressDO_UpdateAddress"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IAddressClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New AddressClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadAddressData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateAddressVbErr:
        
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub



