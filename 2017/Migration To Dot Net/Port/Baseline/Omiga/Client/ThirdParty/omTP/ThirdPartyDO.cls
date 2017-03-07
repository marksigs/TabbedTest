VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ThirdPartyDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


'Workfile:      ThirdPartyDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for ThirdParty Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function GetXMLFromRecordSet(vrstRecordSet As ADODB.Recordset) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'MCS    17/08/99 Created
'MCS    30/08/99 Update following code review
'MCS    01/09/99 Optional timimg output added
'IW     09/03/00 SYS0343 Put in a check for user specified wildcards on Company Name & Town
'MC     27/04/00 SYS0513 Add searchable sortcode field for lenders
'MH     16/05/00 SY0698 Allow wildcard sorting on Town
'SR     19/05/00 SYS0689 Modified method - UpdateThirdparty
'LD     07/11/00 Explicity close recordsets
'GD     11/04/01 SYS2092 Allow wildcard searches for FindPanelValuerList
'LD     07/06/01 SYS2355 CC012 - omTP - FindMainMortgageLenderList (SQL Server Port)
'LD     19/06/01 SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'DM     09/09/01 SYS2572 Added generic sql processing.
'JR     07/09/01 Omiplus24 Modified IThirdPartyDO_GetThirdParty, _GetDirectorySummary, _FindDirectoryList,
'                _GetDirectory, _FindDirectorySummaryList, _FindThirdPartyList, _FindPanelValuerList and added method AppendContactTelephoneData
'AD     10/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
'JR     26/10/01 SYS2680, comment out all Omiplus24 Telephone number changes until further notice.
'JR     03/12/01 SYS2459 re-apply Omiplus24 telephone changes
'------------------------------------------------------------------------------------------

Option Explicit

Implements IThirdPartyDO
Implements ObjectControl
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
Private m_objIDOAssist As IDOAssist
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================

'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'ASu    08/10/2002  BMIDS00152 - Change to IThirdPartyDO_FindPanelValuerList to include Town
'SA     12/11/2002  BMIDS00693 - DeleteThirdParty - error handling changed
'PSC    25/11/2002  BMIDS00998 - Rewrite FindDirectoryList for performance
'PSC    03/12/2002  BM0105 - Improve performance on FindDirectoryList and FindPanelValuerList.
'                                Add GetDirectoryDetails
'JD     05/10/2004  BMIDS897 added GetDirectoryGUID
'------------------------------------------------------------------------------------------
Private Sub SaveAddress(ByVal vxmlElement As IXMLDOMElement, _
                        ByRef rstrAddressGUID As String, _
                        Optional ByVal vbotOperationType As Variant)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo SaveAddressVbErr

    Const strFunctionName As String = "SaveAddress"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objIAddressBO As IAddressBO

    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlAddressNode As IXMLDOMNode

    Dim botOperationType As Integer
    Dim strReturnXML As String

    '
    ' Initialise
    '
    rstrAddressGUID = ""
    botOperationType = booNone
    If Not IsMissing(vbotOperationType) Then botOperationType = vbotOperationType

    '
    ' Main
    '
    Set xmlAddressNode = vxmlElement.selectSingleNode("ADDRESS")
    If Not (xmlAddressNode Is Nothing) Then
        If objContext Is Nothing Then
            Set objIAddressBO = New AddressBO
        Else
            Set objIAddressBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".AddressBO")
        End If

        strReturnXML = objIAddressBO.SaveAddress(xmlAddressNode, rbotOperationType:=botOperationType).xml
        objErrAssist.CheckXMLResponse strReturnXML, vblnRaiseError:=True
    
        ' Extract the address GUID from the appropriate source
        Select Case botOperationType
            Case booCreate
                Set xmlTempDoc = objXmlAssist.load(strReturnXML, TypeName(Me), strFunctionName)
                rstrAddressGUID = objXmlAssist.GetTagValue(xmlTempDoc.documentElement, "ADDRESSGUID")
            Case booDelete, booNone
                rstrAddressGUID = ""
            Case booUpdate
                rstrAddressGUID = objXmlAssist.GetTagValue(vxmlElement, "ADDRESSGUID")
            Case Else
                ' N/A
        End Select
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIAddressBO = Nothing
    Set xmlAddressNode = Nothing
    Set xmlTempDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

SaveAddressVbErr:
    
    Set objXmlAssist = Nothing
    Set xmlAddressNode = Nothing
    Set objIAddressBO = Nothing
    Set xmlTempDoc = Nothing
    
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

Private Sub SaveContact(ByVal vxmlElement As IXMLDOMElement, _
                        ByRef rstrContactGUID As String, _
                        Optional ByVal vbotOperationType As Variant)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo SaveContactVbErr

    Const strFunctionName As String = "SaveContact"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objIContactBO As IContactBO

    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlContactNode As IXMLDOMNode

    Dim botOperationType As Integer
    Dim strReturnXML As String

    '
    ' Initialise
    '
    rstrContactGUID = ""
    botOperationType = booNone
    If Not IsMissing(vbotOperationType) Then botOperationType = vbotOperationType

    '
    ' Main
    '
    Set xmlContactNode = vxmlElement.selectSingleNode("CONTACTDETAILS")
    If Not (xmlContactNode Is Nothing) Then
        If objContext Is Nothing Then
            Set objIContactBO = New ContactBO
        Else
            Set objIContactBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ContactBO")
        End If

        strReturnXML = objIContactBO.SaveContact(xmlContactNode, rbotOperationType:=botOperationType).xml
        objErrAssist.CheckXMLResponse strReturnXML, vblnRaiseError:=True
    
        ' Extract the Contact GUID from the appropriate source
        Select Case botOperationType
            Case booCreate
                Set xmlTempDoc = objXmlAssist.load(strReturnXML, TypeName(Me), strFunctionName)
                rstrContactGUID = objXmlAssist.GetTagValue(xmlTempDoc.documentElement, "CONTACTDETAILSGUID")
            Case booDelete, booNone
                rstrContactGUID = ""
            Case booUpdate
                rstrContactGUID = objXmlAssist.GetTagValue(vxmlElement, "CONTACTDETAILSGUID")
            Case Else
                ' N/A
        End Select
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIContactBO = Nothing
    Set xmlContactNode = Nothing
    Set xmlTempDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

SaveContactVbErr:
    
    Set objXmlAssist = Nothing
    Set xmlContactNode = Nothing
    Set objIContactBO = Nothing
    Set xmlTempDoc = Nothing
    
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

Private Function AlreadyThere(xmlNewList As IXMLDOMElement, xmlCurrentRowElement As IXMLDOMElement) As Boolean
    Dim iListIndex As Integer
    Dim iNewListIndex As Integer
    Dim blnRetVal As Boolean
    Dim strCurrentDirectoryGUID As String
    Dim iNumInNewList
    Dim xmlCurrentNewRowElement As IXMLDOMElement
    Dim strNewDirectoryGUID As String
    Dim strDirectoryGUID As String
    
    iNumInNewList = xmlNewList.childNodes.length
    If iNumInNewList = 0 Then
        blnRetVal = False
    Else
        blnRetVal = False
        strCurrentDirectoryGUID = m_objXmlAssist.GetElementText(xmlCurrentRowElement, "DIRECTORYGUID")
        iNumInNewList = xmlNewList.childNodes.length
        For iNewListIndex = 0 To (iNumInNewList - 1)
                Set xmlCurrentNewRowElement = xmlNewList.childNodes.Item(iNewListIndex)
                strNewDirectoryGUID = m_objXmlAssist.GetElementText(xmlCurrentNewRowElement, "DIRECTORYGUID")
                strDirectoryGUID = m_objXmlAssist.GetElementText(xmlCurrentRowElement, "DIRECTORYGUID")
                If strNewDirectoryGUID = strDirectoryGUID Then blnRetVal = True
        Next
    End If
    AlreadyThere = blnRetVal
    
    
End Function


Private Function IThirdPartyDO_FindPanelValuerList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IThirdPartyDO_FindPanelValuerList"
    
   
    Dim objComboDO As ComboDO
    Dim objIADOAssist  As IADOAssist
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    
    Dim strData As String
    Dim strTagValue As String
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "COMPANYNAME")
    Set adoParameter = adoCommand.CreateParameter("@pCompanyName", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "PANELID")
    Set adoParameter = adoCommand.CreateParameter("@pPanelId", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "TOWN")
    Set adoParameter = adoCommand.CreateParameter("@pTown", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "VALUATIONTYPE")
    
    Set adoParameter = adoCommand.CreateParameter("@pValuationType", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    If strTagValue <> "" Then
        adoCommand(3) = strTagValue
    Else
        adoCommand(3) = Null
    End If
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "VALUERTYPE")
    
    Set adoParameter = adoCommand.CreateParameter("@pValuerType", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    If strTagValue <> "" Then
        adoCommand(4) = strTagValue
    Else
        adoCommand(4) = Null
    End If
    
    With adoConn
        .ConnectionString = objIADOAssist.GetConnStr()
        .open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.open
    
    With adoCommand
        .CommandText = "USP_PANELVALUERLIST"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "PANELVALUERLISTLIST"
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    Set xmlOut = m_objXmlAssist.load(strData, TypeName(Me), strFunctionName)

    Set IThirdPartyDO_FindPanelValuerList = xmlOut.documentElement
    
    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    
    ' Close the stream if still open
    If Not adoStream Is Nothing Then
        If adoStream.State = adStateOpen Then
                adoStream.Close
        End If
    End If
        
    ' Close the connection if still open
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
                adoConn.Close
        End If
    End If
        
    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
  
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Sub IThirdPartyDO_CreateDirectory(ByVal vxmlTableElement As IXMLDOMElement, _
                                          ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IThirdPartyDO_CreateDirectory"
   
    Dim objErrAssist As New ErrAssist
    
    IThirdPartyDO_CreateThirdParty vxmlTableElement, rxmlKeysNode, blnCreateDirectory:=True
    
    Set objErrAssist = Nothing

    Exit Sub

CreateVbErr:
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IThirdPartyDO_CreateMortgageLenderDirectory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IThirdPartyDO_CreateMortgageLenderDirectory"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadMortgageLenderDirectoryData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
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

Private Sub IThirdPartyDO_CreateThirdParty(ByVal vxmlTableElement As IXMLDOMElement, _
                                           ByRef rxmlKeysNode As IXMLDOMNode, _
                                           Optional ByVal blnCreateDirectory As Boolean = False)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IThirdPartyDO_CreateThirdParty"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objIDOAssist As IDOAssist

    Dim strTable     As String, _
        strGUIDField As String

    Dim strAddressGUID As String, _
        strContactGUID As String

    Dim xmlRootNode As IXMLDOMNode
    Dim xmlGUIDNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIThirdPartyClassDef = New ThirdPartyClassDef

    '
    ' Initialise
    '
    strTable = IIf(blnCreateDirectory, "NAMEANDADDRESSDIRECTORY", "THIRDPARTY")
    strGUIDField = IIf(blnCreateDirectory, "DIRECTORYGUID", "THIRDPARTYGUID")

    If vxmlTableElement.nodeName = strTable Then
        Set xmlRootNode = vxmlTableElement
    Else
        Set xmlRootNode = vxmlTableElement.selectSingleNode(strTable)
    End If

    '
    ' Save Address/Contact details
    '
    SaveAddress vxmlTableElement, strAddressGUID
    SaveContact vxmlTableElement, strContactGUID

    '
    ' Insert address GUID
    '
    Set xmlGUIDNode = xmlRootNode.selectSingleNode("ADDRESSGUID")
    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = xmlRootNode.appendChild(vxmlTableElement.ownerDocument.createElement("ADDRESSGUID"))
    End If
    xmlGUIDNode.Text = strAddressGUID

    '
    ' Insert contact details GUID
    '
    Set xmlGUIDNode = xmlRootNode.selectSingleNode("CONTACTDETAILSGUID")
    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = xmlRootNode.appendChild(vxmlTableElement.ownerDocument.createElement("CONTACTDETAILSGUID"))
    End If
    xmlGUIDNode.Text = strContactGUID

    '
    ' Set key for ThirdParty/Directory
    '
    Set xmlGUIDNode = xmlRootNode.selectSingleNode(strGUIDField)
    If xmlGUIDNode Is Nothing Then
        Set xmlGUIDNode = xmlRootNode.appendChild(vxmlTableElement.ownerDocument.createElement(strGUIDField))
    End If
    ' Generate the GUID
    xmlGUIDNode.Text = CreateGUID()

    '
    ' Insert the ThirdParty/Directory
    '
    Set xmlClassDefDoc = IIf(blnCreateDirectory, objIThirdPartyClassDef.LoadDirectoryData, objIThirdPartyClassDef.LoadThirdPartyData)
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    ' Append generated ThirdParty/Directory GUID to keys XML
    rxmlKeysNode.appendChild xmlGUIDNode
   
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlRootNode = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlGUIDNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlRootNode = Nothing
    Set xmlGUIDNode = Nothing
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

Private Sub IThirdPartyDO_DeleteThirdParty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IThirdPartyDO_DeleteThirdParty"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIADOAssist As IADOAssist
    Dim cmd As ADODB.Command
    Dim objSQLAssist As New SQLAssist

    Dim xmlGUIDNode As IXMLDOMNode
    
    If objContext Is Nothing Then
        Set objIADOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set xmlGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
    If xmlGUIDNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        ' DM 09/09/01 SYS2572
        #If GENERIC_SQL Then
            Dim rguidThirdParty() As Byte
            rguidThirdParty = objSQLAssist.GuidStringToByteArray(xmlGUIDNode.Text)
            .CommandText = "DELETETHIRDPARTY_GS"
            .Parameters.Append .CreateParameter("rawTHIRDPARTY", adVarBinary, adParamInput, UBound(rguidThirdParty) + 1, rguidThirdParty)
        #Else

            .CommandText = "DELETETHIRDPARTY(" & objSQLAssist.FormatGuid(xmlGUIDNode.Text) & ")"
        #End If
        .Execute , , adCmdStoredProc
    End With
    
    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set objErrAssist = Nothing
    Set xmlGUIDNode = Nothing
    Set objSQLAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:

    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set xmlGUIDNode = Nothing
    Set objSQLAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    '-- BMIDS00693 {
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
            objContext.SetComplete
'        End If
'    End If
    '-- BMIDS00693 }
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Function IThirdPartyDO_FindDirectoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDirectoryListVbErr

    Const strFunctionName As String = "IThirdPartyDO_FindDirectoryList"
    
    Dim objComboDO As ComboDO
    Dim objIADOAssist  As IADOAssist
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    
    Dim strData As String
    Dim strTagValue As String
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    Set adoParameter = adoCommand.CreateParameter("@pNameAndAddressType", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "NAMEANDADDRESSTYPE")
    
    If Trim$(strTagValue) <> "" Then
        adoCommand(0) = strTagValue
    Else
        adoCommand(0) = Null
    End If
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "COMPANYNAME")
    Set adoParameter = adoCommand.CreateParameter("@pCompanyName", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "PANELID")
    Set adoParameter = adoCommand.CreateParameter("@pPanelId", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "TOWN")
    Set adoParameter = adoCommand.CreateParameter("@pTown", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "POSTCODE")
    Set adoParameter = adoCommand.CreateParameter("@pPostCode", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "BUILDINGORHOUSENAME")
    Set adoParameter = adoCommand.CreateParameter("@pHouseName", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "BUILDINGORHOUSENUMBER")
    Set adoParameter = adoCommand.CreateParameter("@pHouseNumber", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "FLATNUMBER")
    Set adoParameter = adoCommand.CreateParameter("@pFlatNumber", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
   
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "STREET")
    Set adoParameter = adoCommand.CreateParameter("@pStreet", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    ' Head office indicator
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "FILTER")
    If Trim$(strTagValue) <> "" Then
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        If objComboDO.IsItemInValidation("DirectoryFilter", strTagValue, "H") Then
            ' Must be head office
            strTagValue = "1"
        ElseIf objComboDO.IsItemInValidation("DirectoryFilter", strTagValue, "N") Then
            ' Must be non-head office
            strTagValue = "0"
        Else
            strTagValue = ""
        End If
    End If
    
    Set adoParameter = adoCommand.CreateParameter("@pHeadOfficeInd", adInteger, adParamInput)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    If strTagValue <> "" Then
        adoCommand(9) = strTagValue
    Else
        adoCommand(9) = Null
    End If
    
    strTagValue = m_objXmlAssist.GetTagValue(vxmlTableElement, "SORTCODE")
    Set adoParameter = adoCommand.CreateParameter("@pBankSortCode", adBSTR, adParamInput, Len(strTagValue), strTagValue)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    With adoConn
        .ConnectionString = objIADOAssist.GetConnStr()
        .open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.open
    
    With adoCommand
        .CommandText = "USP_DIRECTORYLIST"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "NAMEANDADDRESSDIRECTORYLIST"
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    Set xmlOut = m_objXmlAssist.load(strData, TypeName(Me), strFunctionName)

    Set IThirdPartyDO_FindDirectoryList = xmlOut.documentElement
    
    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindDirectoryListVbErr:
    
    ' Close the stream if still open
    If Not adoStream Is Nothing Then
        If adoStream.State = adStateOpen Then
                adoStream.Close
        End If
    End If
        
    ' Close the connection if still open
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
                adoConn.Close
        End If
    End If
        
    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
   
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IThirdPartyDO_FindDirectorySummaryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDirectorySummaryListVbErr

    Const strFunctionName As String = "IThirdPartyDO_FindDirectorySummaryList"

    Dim objErrAssist As New ErrAssist
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objIDOAssist  As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objXmlAssist As New XMLAssist, _
        objSQLAssist As New SQLAssist
    Dim objComboDO As ComboDO

    Dim strSQLString As String
    Dim strTagValue As String
    Dim blnDummy As Boolean
    Dim i As Integer
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTopElement As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim cmd As ADODB.Command, _
        rst As ADODB.Recordset

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set objIThirdPartyClassDef = New ThirdPartyClassDef

    '
    ' Get WHERE criteria from XML
    '
    strSQLString = ""
    ' Name and address type
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "NAMEANDADDRESSTYPE")
    If Trim$(strTagValue) <> "" Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(NAMEANDADDRESSTYPE = " & strTagValue & ")"
    End If
   
    ' Company name
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "COMPANYNAME")
        If Trim$(strTagValue) <> "" Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(COMPANYNAME LIKE " & objSQLAssist.FormatWildcardedString(strTagValue & "*", blnDummy) & ")"
    End If

    ' Town
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "TOWN")
        If Trim$(strTagValue) <> "" Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(TOWN LIKE " & objSQLAssist.FormatWildcardedString(strTagValue & "*", blnDummy) & ")"
    End If

    ' Head office indicator
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "FILTER")
        If Trim$(strTagValue) <> "" Then
        Set objComboDO = New ComboDO
        If objComboDO.IsItemInValidation("DirectoryFilter", strTagValue, "H") Then
            ' Must be head office
            strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
                "(HEADOFFICEINDICATOR = 'H')"
        ElseIf objComboDO.IsItemInValidation("DirectoryFilter", strTagValue, "N") Then
            ' Must be non-head office
            strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
                "(HEADOFFICEINDICATOR = 'N')"
        End If
    End If


    '
    ' Build the SQL query and execute
    '
    ' DIRECTORYSUMMARY is a view
    strSQLString = "SELECT * FROM DIRECTORYSUMMARY WHERE " & strSQLString & " ORDER BY DIRECTORYGUID"

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQLString

    Set rst = cmd.Execute(, , adCmdText)

    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    '
    ' Convert the recordset into XML
    '
    Set xmlListNode = xmlOut.appendChild(xmlOut.createElement("DIRECTORYSUMMARYLIST"))
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadDirectorySummaryData

    
    ' AD 10/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
    Do While Not rst.EOF
        objIDOAssist.GetXMLFromRecordset rst, xmlClassDefDoc, xmlListNode
        rst.MoveNext
    Loop
        rst.Close

    'JR - Omiplus24, append ContactTelephoneData to ContactDetailsNode
    Dim xmlCTNode As IXMLDOMNode
    Dim xmlCTNodeList As IXMLDOMNodeList
    Set xmlCTNodeList = xmlListNode.selectNodes(".//CONTACTDETAILS")
    For Each xmlCTNode In xmlCTNodeList
        Call AppendContactTelephoneData(xmlCTNode)
    Next
    'End

    ' Promote the NAMEANDADDRESSDIRECTORY node to replace the DIRECTORYDETAILS node in the list items
    For Each xmlNode In xmlOut.documentElement.selectNodes("DIRECTORYDETAILS/NAMEANDADDRESSDIRECTORY")
        objXmlAssist.PromoteNode xmlOut.documentElement, xmlNode
    Next xmlNode

    ' Move all information pertaining to each NAMEANDADDRESSDIRECTORY guid into the same set of NAMEANDADDRESSDIRECTORY tags
    objXmlAssist.RemoveDuplicates xmlListNode, "DIRECTORYGUID"
    ' Remove any duplicate information in each of the list items (i.e. NAMEANDADDRESSDIRECTORY nodes)
    For Each xmlNode In xmlListNode.childNodes
        objXmlAssist.RemoveDuplicates xmlNode
    Next xmlNode

    Set IThirdPartyDO_FindDirectorySummaryList = xmlOut.documentElement

    Set objComboDO = Nothing
    Set objErrAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindDirectorySummaryListVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing
    
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

Private Function IThirdPartyDO_FindMainMortgageLenderList() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageLenderListVbErr

    Const strFunctionName As String = "IThirdPartyDO_FindMainMortgageLenderList"

    Dim objErrAssist As New ErrAssist
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objIDOAssist  As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objXmlAssist As New XMLAssist
    Dim objComboDO As ComboDO

    Dim strSQLString As String
    Dim i As Integer

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTopElement As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim cmd As ADODB.Command, _
        rst As ADODB.Recordset

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set objIThirdPartyClassDef = New ThirdPartyClassDef

    '
    ' Build the SQL query and execute
    '
#If GENERIC_SQL Then
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
    Dim strSystemDate As String
    strSystemDate = objSQLAssist.GetSystemDate
    Set objSQLAssist = Nothing
    strSQLString = "SELECT * FROM NAMEANDADDRESSDIRECTORY N, MORTGAGELENDERDIRECTORY M WHERE " & _
                   "(M.DIRECTORYGUID = N.DIRECTORYGUID) AND " & _
                   "NOT (M.ORGANISATIONID IS NULL) AND " & _
                   "(M.MAINMORTGAGELENDERIND = 1) AND " & _
                   "(N.NAMEANDADDRESSACTIVEFROM <= " & strSystemDate & ") AND " & _
                   "((N.NAMEANDADDRESSACTIVETO >= " & strSystemDate & ") OR (N.NAMEANDADDRESSACTIVETO IS NULL))"
#Else
    strSQLString = "SELECT * FROM NAMEANDADDRESSDIRECTORY N, MORTGAGELENDERDIRECTORY M WHERE " & _
                   "(M.DIRECTORYGUID = N.DIRECTORYGUID) AND " & _
                   "NOT (M.ORGANISATIONID IS NULL) AND " & _
                   "(M.MAINMORTGAGELENDERIND = 1) AND " & _
                   "(N.NAMEANDADDRESSACTIVEFROM <= SYSDATE) AND " & _
                   "((N.NAMEANDADDRESSACTIVETO >= SYSDATE) OR (N.NAMEANDADDRESSACTIVETO IS NULL))"
#End If

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQLString

    Set rst = cmd.Execute(, , adCmdText)

    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    '
    ' Convert the recordset into XML
    '
    Set xmlListNode = xmlOut.appendChild(xmlOut.createElement("MAINMORTGAGELENDERLIST"))
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadMainMortgageLenderData

    ' AD 10/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
    Do While Not rst.EOF
        objIDOAssist.GetXMLFromRecordset rst, xmlClassDefDoc, xmlListNode
        rst.MoveNext
    Loop
        rst.Close

    ' Promote the NAMEANDADDRESSDIRECTORY node to replace the DIRECTORYDETAILS node in the list items
    For Each xmlNode In xmlOut.documentElement.selectNodes("MAINMORTGAGELENDER/MORTGAGELENDERDIRECTORY")
        objXmlAssist.PromoteNode xmlOut.documentElement, xmlNode
    Next xmlNode

    ' Move all information pertaining to each NAMEANDADDRESSDIRECTORY guid into the same set of NAMEANDADDRESSDIRECTORY tags
    objXmlAssist.RemoveDuplicates xmlListNode, "ORGANISATIONID"
    ' Remove any duplicate information in each of the list items (i.e. NAMEANDADDRESSDIRECTORY nodes)
    For Each xmlNode In xmlListNode.childNodes
        objXmlAssist.RemoveDuplicates xmlNode
        objXmlAssist.GroupNodesIntoList xmlNode, "PANELBANKACCOUNT"
        objXmlAssist.GroupNodesIntoList xmlNode, "PANELVALUERVALUATIONTYPE"
    Next xmlNode

    Set IThirdPartyDO_FindMainMortgageLenderList = xmlOut.documentElement

    Set objComboDO = Nothing
    Set objErrAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindMortgageLenderListVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing

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

Private Function IThirdPartyDO_FindThirdPartyList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindThirdPartyListVbErr

    Const strFunctionName As String = "IThirdPartyDO_FindThirdPartyList"

    Dim objErrAssist As New ErrAssist
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objIDOAssist  As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objXmlAssist As New XMLAssist, _
        objSQLAssist As New SQLAssist
    Dim objComboDO As ComboDO
    Dim xmlNode As IXMLDOMNode

    Dim xmlClassDef As FreeThreadedDOMDocument40

    Dim strSQLString As String
    Dim strTagValue As String
    Dim blnDummy As Boolean
    Dim i As Integer
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTopElement As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode

    Dim cmd As ADODB.Command, _
        rst As ADODB.Recordset

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set objIThirdPartyClassDef = New ThirdPartyClassDef

    '
    ' Get WHERE criteria from XML
    '
    strSQLString = ""
    ' Third party type
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYTYPE")
    If Trim$(strTagValue) <> "" Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(THIRDPARTYTYPE = " & strTagValue & ")"
    End If
    
    ' Company name
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "COMPANYNAME")
    If Trim$(strTagValue) <> "" Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(COMPANYNAME LIKE " & objSQLAssist.FormatWildcardedString(strTagValue, blnDummy) & ")"
    End If

    ' Town
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "TOWN")
    If Trim$(strTagValue) <> "" Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(TOWN LIKE " & objSQLAssist.FormatWildcardedString(strTagValue & "*", blnDummy) & ")"
    End If

    '
    ' Build the SQL query and execute
    '
    ' THIRDPARTYDETAILS is a view
    strSQLString = "SELECT * FROM THIRDPARTYDETAILS WHERE " & strSQLString & " ORDER BY THIRDPARTYGUID"

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQLString

    Set rst = cmd.Execute(, , adCmdText)

    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    '
    ' Convert the recordset into XML
    '
    Set xmlListNode = xmlOut.appendChild(xmlOut.createElement("THIRDPARTYLIST"))
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadThirdPartyDetailsData

    ' AD 10/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
    Do While Not rst.EOF
        objIDOAssist.GetXMLFromRecordset rst, xmlClassDefDoc, xmlListNode
        rst.MoveNext
    Loop
        rst.Close

    'JR - Omiplus24, append ContactTelephoneData to ContactDetailsNode
    Dim xmlCTNode As IXMLDOMNode
    Dim xmlCTNodeList As IXMLDOMNodeList
    Set xmlCTNodeList = xmlListNode.selectNodes(".//CONTACTDETAILS")
    For Each xmlCTNode In xmlCTNodeList
        Call AppendContactTelephoneData(xmlCTNode)
    Next
    'End
    
    ' Promote the THIRDPARTY node to replace the THIRDPARTYDETAILS node in the list items
    For Each xmlNode In xmlOut.documentElement.selectNodes("THIRDPARTYDETAILS/THIRDPARTY")
        objXmlAssist.PromoteNode xmlOut.documentElement, xmlNode
    Next xmlNode

    Set IThirdPartyDO_FindThirdPartyList = xmlOut.documentElement

    Set objComboDO = Nothing
    Set objErrAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindThirdPartyListVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objComboDO = Nothing
    Set objErrAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing
    
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

Private Function IThirdPartyDO_GetDirectory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetDirectory"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadDirectoryDetailsData
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, _
            vstrItemName:="NAMEANDADDRESSDIRECTORY")

    'JR - Omiplus24, append ContactTelephoneData to ContactDetailsNode
    Dim xmlCTNode As IXMLDOMNode
    Set xmlCTNode = xmlReturn.selectSingleNode(".//CONTACTDETAILS")
    If Not xmlCTNode Is Nothing Then
        Call AppendContactTelephoneData(xmlCTNode)
    End If
    'End
    
    ' Promote the NAMEANDADDRESSDIRECTORY tags to replace the DIRECTORYDETAILS tags in the list
    For Each xmlNode In xmlReturn.selectNodes("DIRECTORYDETAILS/NAMEANDADDRESSDIRECTORY")
        objXmlAssist.PromoteNode xmlReturn, xmlNode
    Next xmlNode

    ' Ensure only one record exists in the list for each DIRECTORYGUID (there will be only one DIRECTORYGUID -
    ' the DIRECTORYGUID specified in the vxmlTableElement parameter)
    objXmlAssist.RemoveDuplicates xmlReturn, "DIRECTORYGUID"

    ' Group all similar child nodes into lists
    ' (There should be only one child node by this point, but do the For Each anyway)
    For Each xmlNode In xmlReturn.childNodes
        objXmlAssist.RemoveDuplicates xmlNode
        objXmlAssist.GroupNodesIntoList xmlNode, "PANELBANKACCOUNT"
        objXmlAssist.GroupNodesIntoList xmlNode, "PANELVALUERVALUATIONTYPE"
    Next xmlNode

    ' Note that only the firstChild is returned - the NAMEANDADDRESSDIRECTORYLIST node is not required
    Set IThirdPartyDO_GetDirectory = xmlReturn.firstChild

    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDirectoryVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    
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

Private Function IThirdPartyDO_GetDirectoryMinimal(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryMinimalVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetDirectoryMinimal"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadDirectoryData
    Set IThirdPartyDO_GetDirectoryMinimal = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDirectoryMinimalVbErr:
    
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

Private Function IThirdPartyDO_GetDirectorySummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectorySummaryVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetDirectorySummary"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadDirectorySummaryData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstrItemName:="NAMEANDADDRESSDIRECTORY")

    'JR - Omiplus24
    Dim xmlCTNode As IXMLDOMNode
    Set xmlCTNode = xmlReturn.selectSingleNode(".//CONTACTDETAILS")
    If Not xmlCTNode Is Nothing Then
        Call AppendContactTelephoneData(xmlCTNode)
    End If
    'End

    objXmlAssist.PromoteNode xmlReturn, xmlReturn.selectSingleNode("NAMEANDADDRESSDIRECTORY")

    Set IThirdPartyDO_GetDirectorySummary = xmlReturn

    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlCTNode = Nothing ' JR Omiplus24
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDirectorySummaryVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlCTNode = Nothing ' JR Omiplus24
    
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

Private Function IThirdPartyDO_GetThirdParty(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetThirdParty"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadThirdPartyDetailsData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstrItemName:="THIRDPARTY")

    'JR - Omiplus24
    Dim xmlCTNode As IXMLDOMNode
    Set xmlCTNode = xmlReturn.selectSingleNode(".//CONTACTDETAILS")
    If Not xmlCTNode Is Nothing Then
        Call AppendContactTelephoneData(xmlCTNode)
    End If
    'JR - End
    
    objXmlAssist.PromoteNode xmlReturn, xmlReturn.selectSingleNode("THIRDPARTY")

    Set IThirdPartyDO_GetThirdParty = xmlReturn

    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetThirdPartyVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    
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

Private Function IThirdPartyDO_GetThirdPartyMinimal(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetThirdPartyMinimalVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetThirdPartyMinimal"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadThirdPartyData
    Set IThirdPartyDO_GetThirdPartyMinimal = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetThirdPartyMinimalVbErr:
    
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
Private Function IThirdPartyDO_GetDirectoryGUID(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Returns the DIRECTORYGUID for this companyname. If there is no record found, DIRECTORYGUID is returned empty
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryGUIDVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetDirectoryGUID"
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponse As IXMLDOMNode
    Dim xmlTPNode As IXMLDOMNode
    Dim xmlTmpNode As IXMLDOMNode
    Dim objIADOAssist  As IADOAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim strSQL As String
    Dim strGuid As String
    strGuid = ""
    
    strSQL = "select top 1 directoryguid from nameandaddressdirectory where companyname='"
    strSQL = strSQL & m_objXmlAssist.GetElementText(vxmlTableElement, ".//COMPANYNAME")
    strSQL = strSQL & "' and NameAndAddressType='"
    strSQL = strSQL & m_objXmlAssist.GetElementText(vxmlTableElement, ".//NAMEANDADDRESSTYPE")
    strSQL = strSQL & "'"
    
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    Dim cmd As New ADODB.Command
    
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    
    Dim rstThisRecSet As New ADODB.Recordset

    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    
    While Not rstThisRecSet.EOF
        strGuid = objSQLAssist.GuidToString(rstThisRecSet.Fields.Item("DIRECTORYGUID").Value)
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
    'build the response
    Set xmlTPNode = xmlOut.createElement("NAMEANDADDRESSDIRECTORY")
    Set xmlTmpNode = xmlOut.createElement("DIRECTORYGUID")
    xmlTmpNode.Text = strGuid
    xmlTPNode.appendChild xmlTmpNode
    Set xmlResponse = xmlOut.appendChild(xmlTPNode)
    
    Set IThirdPartyDO_GetDirectoryGUID = xmlResponse
    
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function

GetDirectoryGUIDVbErr:
    
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IThirdPartyDO_UpdateThirdParty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IThirdPartyDO_UpdateThirdParty"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    Dim objIDOAssist As IDOAssist

    Dim xmlGUIDNode As IXMLDOMNode, _
        xmlRootNode As IXMLDOMNode
    Dim strAddressGUID As String, _
        strContactGUID As String

    Dim objIAddressBO As IAddressBO, _
        objIContactBO As IContactBO
    Dim botAddressOperation As BOOPERATIONTYPE, _
        botContactOperation As BOOPERATIONTYPE

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Initialise
    '
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    Set objIAddressBO = New AddressBO
    Set objIContactBO = New ContactBO

    botAddressOperation = objIAddressBO.DetermineAddressOperation(vxmlTableElement)
    botContactOperation = objIContactBO.DetermineContactOperation(vxmlTableElement)

    strAddressGUID = ""
    strContactGUID = ""

    Set xmlRootNode = vxmlTableElement

    '
    ' Create/UpdateThirdParty address/contact
    '
    If (botAddressOperation = booCreate) Or (botAddressOperation = booUpdate) Then
        ' Address should be created/UpdateThirdParty before the third party record is UpdateThirdPartyd
        SaveAddress vxmlTableElement, strAddressGUID, vbotOperationType:=botAddressOperation
    End If

    If (botContactOperation = booCreate) Or (botContactOperation = booUpdate) Then
        ' Contact should be created/UpdateThirdParty before the third party record is UpdateThirdPartyd
        SaveContact vxmlTableElement, strContactGUID, vbotOperationType:=botContactOperation
    End If


    '
    ' Save ThirdParty
    '
    ' Insert Address GUID into third party XML
    'SR 18/05/00 - Check for the name of the root node before finding AddressGUID Node
    If vxmlTableElement.nodeName = "THIRDPARTY" Then
        Set xmlGUIDNode = vxmlTableElement.selectSingleNode("./ADDRESSGUID")
    Else
        Set xmlGUIDNode = vxmlTableElement.selectSingleNode(".//THIRDPARTY/ADDRESSGUID")
    End If
    
    If xmlGUIDNode Is Nothing Then
        ' No GUID tag was found so create it now
        Set xmlGUIDNode = xmlRootNode.appendChild(vxmlTableElement.ownerDocument.createElement("ADDRESSGUID"))
    End If
    
    ' SR 18/05/00 - If Address is to be deleted, assign empty string to AddressGuid
    If botAddressOperation = booDelete Then
        xmlGUIDNode.Text = ""
    Else
        xmlGUIDNode.Text = strAddressGUID
    End If

    ' Insert Contact GUID into third party XML
    'SR 18/05/00 - Check for the name of the root node before finding ContactDetailsGUID Node
    If vxmlTableElement.nodeName = "THIRDPARTY" Then
        Set xmlGUIDNode = vxmlTableElement.selectSingleNode("./CONTACTDETAILSGUID")
    Else
        Set xmlGUIDNode = vxmlTableElement.selectSingleNode(".//THIRDPARTY/CONTACTDETAILSGUID")
    End If

    If xmlGUIDNode Is Nothing Then
        ' No GUID tag was found so create it now
        Set xmlGUIDNode = xmlRootNode.appendChild(vxmlTableElement.ownerDocument.createElement("CONTACTDETAILSGUID"))
    End If
    ' SR 18/05/00 - If contact details are to be deleted, assign empty string to ContactDetailsGuid
    If botContactOperation = booDelete Then
        xmlGUIDNode.Text = ""
    Else
        xmlGUIDNode.Text = strContactGUID
    End If
    
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadThirdPartyData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc


    '
    ' Delete address/contact
    '
    If botAddressOperation = booDelete Then
        ' Address should be deleted only once the thirdparty record has been UpdateThirdPartyd
        SaveAddress vxmlTableElement, strAddressGUID, vbotOperationType:=botAddressOperation
    End If

    If botContactOperation = booDelete Then
        ' Contact should be deleted only once the thirdparty record has been UpdateThirdPartyd
        SaveContact vxmlTableElement, strContactGUID, vbotOperationType:=botContactOperation
    End If
    
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set objErrAssist = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlRootNode = Nothing
    Set objIAddressBO = Nothing
    Set objIContactBO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

UpdateVbErr:
        
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    Set objErrAssist = Nothing
    Set xmlGUIDNode = Nothing
    Set xmlRootNode = Nothing
    Set objIAddressBO = Nothing
    Set objIContactBO = Nothing
    
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

Private Function IThirdPartyDO_GetPaymentDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetPaymentDetailsVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetPaymentDetails"
    
                
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IThirdPartyClassDef
    
    Set objIClassDef = New ThirdPartyClassDef
    Set xmlClassDefDoc = objIClassDef.LoadPaymentDetailsData()
    Set IThirdPartyDO_GetPaymentDetails = m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetPaymentDetailsVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing

End Sub

Private Sub AppendContactTelephoneData(ByVal vxmlContactDetailsNode As IXMLDOMNode)
On Error GoTo AppendContactTelephoneDataVbErr

    Const strFunctionName As String = "AppendContactTelephoneData"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objIContactBO As IContactBO
    Dim xmlCTNode As IXMLDOMNode
    Dim xmlCTResponseNode As IXMLDOMNode
    Dim xmlContactNode As IXMLDOMNode
    Dim xmlCTNodeList As IXMLDOMNodeList
        
    Set objIContactBO = New ContactBO
    
    'Check for ContactDetailsGuid
    If m_objXmlAssist.GetElementText(vxmlContactDetailsNode, "CONTACTDETAILSGUID") <> "" Then
        Set xmlCTResponseNode = objIContactBO.GetContact(vxmlContactDetailsNode)
   
        'Append Contact Telephone data
        If Not xmlCTResponseNode Is Nothing Then
            Set xmlCTNodeList = xmlCTResponseNode.selectNodes(".//CONTACTTELEPHONEDETAILS")
    
            'Now append to ContactDetails Node
            For Each xmlCTNode In xmlCTNodeList
                vxmlContactDetailsNode.appendChild xmlCTNode
            Next
        End If
    End If

    Set objErrAssist = Nothing
    Set objIContactBO = Nothing
    Set xmlCTNode = Nothing
    Set xmlCTResponseNode = Nothing
    Set xmlContactNode = Nothing
    Set xmlCTNodeList = Nothing

    Exit Sub

AppendContactTelephoneDataVbErr:

    Set objIContactBO = Nothing
    Set xmlCTNode = Nothing
    Set xmlCTResponseNode = Nothing
    Set xmlContactNode = Nothing
    Set xmlCTNodeList = Nothing

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

' PSC 03/12/2002 BM0105 - Start
Private Function IThirdPartyDO_GetDirectoryDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryDetailsVbErr

    Const strFunctionName As String = "IThirdPartyDO_GetDirectoryDetails"

    Dim objErrAssist As New ErrAssist
    
    Dim objIADOAssist  As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    
    Dim strData As String
    Dim strTagValue As String
    
    Set objSQLAssist = New SQLAssist
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    strTagValue = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "DIRECTORYGUID")
    Set adoParameter = adoCommand.CreateParameter("@DirectoryGuid", adBinary, adParamInput, 16, objSQLAssist.GuidStringToByteArray(strTagValue))
    
    With adoConn
        .ConnectionString = objIADOAssist.GetConnStr()
        .open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.open
    
    With adoCommand
        .Parameters.Append adoParameter
        .CommandText = "USP_DIRECTORYDETAILS"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
        .Properties("Output Stream") = adoStream
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    Set xmlOut = m_objXmlAssist.load(strData, TypeName(Me), strFunctionName)

    Set IThirdPartyDO_GetDirectoryDetails = xmlOut.documentElement
    
    Set objIADOAssist = Nothing
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetDirectoryDetailsVbErr:
    
    ' Close the stream if still open
    If Not adoStream Is Nothing Then
        If adoStream.State = adStateOpen Then
                adoStream.Close
        End If
    End If
        
    ' Close the connection if still open
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
                adoConn.Close
        End If
    End If
        
    Set objIADOAssist = Nothing
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
' PSC 03/12/2002 BM0105 - End

