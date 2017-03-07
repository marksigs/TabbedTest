VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ContactDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ContactDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for Contact Data Object
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
'LD     19/06/01 SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'MDC    19/06/01 Added ContactTelephoneDetails (OmiPlus 24)
'JR     10/09/01 Omiplus24, Modified GetContact, UpdateContact and added new method CreateContactTelephone
'SG     28/05/02 MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit

Implements IContactDO


 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================




'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MDC    12/12/2002  BM0094 - Legal Rep Contact Details
'MDC    13/01/2003  BM0250
'PSC    30/09/2003  BMIDS612 - Correct CreateContact to only use telephone numbers under the
'                              contact details rather than all the ones in the document
'------------------------------------------------------------------------------------------

Private Sub IContactDO_CreateContact(ByVal vxmlTableElement As IXMLDOMElement, _
                                     ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateContact an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactVbErr

    Const strFunctionName As String = "IContactDO_CreateContact"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIContactClassDef As IContactClassDef
    
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
    ' Set key for Contact
    '
    Set xmlGUIDNode = vxmlTableElement.selectSingleNode("CONTACTDETAILSGUID")
    If xmlGUIDNode Is Nothing Then
        ' No GUID tag was found so CreateContact it now
        Set xmlGUIDNode = vxmlTableElement.appendChild( _
                            vxmlTableElement.ownerDocument.createElement("CONTACTDETAILSGUID"))
    End If

    ' Generate the GUID
    strGuid = CreateGUID
    xmlGUIDNode.Text = strGuid

    Set objIContactClassDef = New ContactClassDef
    Set xmlClassDefDoc = objIContactClassDef.LoadContactData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
        
    '15/06/2001 MDC - OmiPlus 24 - Telephone Numbers
    Dim xmlTelNodeList As IXMLDOMNodeList
    Dim xmlTelNode As IXMLDOMNode
    ' PSC 30/09/2003 BMIDS612
    Set xmlTelNodeList = vxmlTableElement.selectNodes(".//CONTACTTELEPHONEDETAILS")
    Set xmlClassDefDoc = objIContactClassDef.LoadContactTelephoneData()
    For Each xmlTelNode In xmlTelNodeList
        'Append the Key Guid to the ContactTelephoneDetails node...
        xmlTelNode.appendChild xmlGUIDNode
        'xmlTelNode.insertBefore xmlGUIDNode, xmlTelNode.childNodes.Item(0)
        '...and create a new record
        
        'JR - now calls new seperate method
        Call CreateContactTelephone(xmlTelNode)
'        objIDOAssist.GenerateSequenceNumber xmlTelNode, xmlClassDefDoc, "TELEPHONESEQNUM"
'        objIDOAssist.Create xmlTelNode, xmlClassDefDoc
    Next
    
    ' Append generated GUID to keys XML
    rxmlKeysNode.appendChild xmlGUIDNode
    
    Set objIContactClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlGUIDNode = Nothing
    '15/06/2001 MDC - OmiPlus 24 - Telephone Numbers
    Set xmlTelNodeList = Nothing
    Set xmlTelNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateContactVbErr:
    
    Set objIContactClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlGUIDNode = Nothing
    '15/06/2001 MDC - OmiPlus 24 - Telephone Numbers
    Set xmlTelNodeList = Nothing
    Set xmlTelNode = Nothing
    
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

Private Sub IContactDO_DeleteContact(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteContact a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteContactVbErr

    Const strFunctionName As String = "IContactDO_DeleteContact"

    'SG 28/05/02 SYS4767
    Dim lngErrorNo As Long
    
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IContactClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIThirdPartyClassDef = New ContactClassDef
    
    '15/06/2001 MDC: OmiPlus 24 - Telephone Numbers
    'Delete any associated ContactTelephoneDetails records first
    Dim xmlTelNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As XMLAssist
    
    'SG 28/05/02 SYS4767
    'If objContext Is Nothing Then
    '    Set objXmlAssist = New XMLAssist
    'Else
    '    Set objXmlAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".XMLAssist")
    'End If
    'BM0094 MDC 12/12/2002 - Reinstate this code section otherwise objXmlAssist is not instantiated before being used.
    Set objXmlAssist = New XMLAssist
    'BM0094 MDC 12/12/2002 - End
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTelNode = xmlDoc.createElement("CONTACTTELEPHONEDETAILS")
    objXmlAssist.CopyNode "CONTACTDETAILSGUID", "CONTACTDETAILSGUID", vxmlTableElement, xmlTelNode, False
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactTelephoneData()
    
    'SG 28/05/02 SYS4767 START
    'MSMS0044 - If no telephone numbers exist then ignore the error raised.
    On Error Resume Next
    
    'Attempt to delete any telephone numbers for this contact.
    objIDOAssist.DeleteAll xmlTelNode, xmlClassDefDoc
    objErrAssist.SaveErr
    
    On Error GoTo DeleteContactVbErr
    lngErrorNo = objErrAssist.LastErr
    
    'If no telephone numbers were deleted then proceed as normal - don't error.
    If lngErrorNo <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeNoRowsAffectedByDeleteAll Then
            objErrAssist.ReRaise
        End If
    End If
    'MSMS0044 - End.
    'SG 28/05/02 SYS4767 END
    
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactData
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteContactVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeChildRecordsFound Then
        ' The THIRDPARTY/NAMEANDADDRESSDIRECTORY record still exists for this contact
        Resume Next
    End If

    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    '15/06/2001 MDC: OmiPlus 24 - Telephone Numbers
    Set xmlTelNode = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
   
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

Private Function IContactDO_GetContact(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetContactVbErr

    Const strFunctionName As String = "IContactDO_GetContact"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IContactClassDef
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ContactClassDef
    
    '15/06/2001 MDC: OmiPlus 24 - Telephone Numbers
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactData
    
    'JR - Omiplus24, Get ContactDetails
    Dim xmlContactDetailsNode As IXMLDOMNode
    Set xmlContactDetailsNode = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
        
    'JR - Omiplus24, Get/append ContactTelephone(CT) details to ContactDetails
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCTNode As IXMLDOMNode
    Dim xmlTelNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Dim lngErrorNo As Long
    
    Set objXmlAssist = New XMLAssist
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTelNode = xmlDoc.createElement("CONTACTTELEPHONEDETAILS")
    objXmlAssist.CopyNode "CONTACTDETAILSGUID", "CONTACTDETAILSGUID", vxmlTableElement, xmlTelNode, False
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactTelephoneData
    On Error Resume Next
    Set xmlCTNode = objIDOAssist.FindList(xmlTelNode, xmlClassDefDoc)
    objErrAssist.SaveErr
    On Error GoTo GetContactVbErr
    lngErrorNo = objErrAssist.LastErr
    
    If lngErrorNo <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
            objErrAssist.ReRaise
        End If
    Else
        If Not xmlCTNode Is Nothing Then
            xmlContactDetailsNode.appendChild xmlCTNode
        End If
    End If
    'JR - End Omiplus24
    
    Set IContactDO_GetContact = xmlContactDetailsNode
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    'JR - Omiplus24
    Set xmlContactDetailsNode = Nothing
    Set xmlDoc = Nothing
    Set xmlCTNode = Nothing
    Set xmlTelNode = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetContactVbErr:
    
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    'JR - Omiplus24
    Set xmlContactDetailsNode = Nothing
    Set xmlDoc = Nothing
    Set xmlCTNode = Nothing
    Set xmlTelNode = Nothing
    Set objXmlAssist = Nothing
    
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

Private Sub IContactDO_UpdateContact(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateContact a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateContactVbErr

    Const strFunctionName As String = "IContactDO_UpdateContact"
        
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist 'JR - Omiplus24
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyClassDef As IContactClassDef
    Dim objIDOAssist As IDOAssist
    
    'JR
    Dim xmlElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strContactDetailsGuid As String
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    'JR - End
    
    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIThirdPartyClassDef = New ContactClassDef
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    'BM0250 MDC 14/01/2003 - Delete any existing telephone numbers
    Dim xmlTelNode As IXMLDOMNode
    Dim lngErrorNo As Long
    Set xmlTelNode = xmlDoc.createElement("CONTACTTELEPHONEDETAILS")
    objXmlAssist.CopyNode "CONTACTDETAILSGUID", "CONTACTDETAILSGUID", vxmlTableElement, xmlTelNode, False
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactTelephoneData()

    On Error Resume Next

    'Attempt to delete any telephone numbers for this contact.
    objIDOAssist.DeleteAll xmlTelNode, xmlClassDefDoc
    objErrAssist.SaveErr

    On Error GoTo UpdateContactVbErr
    lngErrorNo = objErrAssist.LastErr

    'If no telephone numbers were deleted then proceed as normal - don't error.
    If lngErrorNo <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(lngErrorNo) <> oeNoRowsAffectedByDeleteAll Then
            objErrAssist.ReRaise
        End If
    End If
    'BM0250 MDC 14/01/2003 - End
    
    '15/06/2001 MDC - OmiPlus 24 - Telephone Numbers
    Dim xmlTelNodeList As IXMLDOMNodeList
    Set xmlTelNodeList = vxmlTableElement.selectNodes(".//CONTACTTELEPHONEDETAILS")
    Set xmlClassDefDoc = objIThirdPartyClassDef.LoadContactTelephoneData()
    For Each xmlTelNode In xmlTelNodeList
        
        If objXmlAssist.GetTagValue(xmlTelNode, "CONTACTDETAILSGUID") = "" Then
            strContactDetailsGuid = objXmlAssist.GetTagValue(vxmlTableElement, "CONTACTDETAILSGUID")
            Set xmlElement = xmlDoc.createElement("CONTACTDETAILSGUID")
            xmlElement.Text = strContactDetailsGuid
            xmlTelNode.appendChild xmlElement
        'BM0250 MDC 14/01/2003
        End If
        'BM0250 MDC 14/01/2003 - End
        
        Call CreateContactTelephone(xmlTelNode)
        
        'BM0250 MDC 14/01/2003
'        Else
'            objIDOAssist.Update xmlTelNode, xmlClassDefDoc
'        End If
        'BM0250 MDC 14/01/2003 - End
    Next
   
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    'JR - Omiplus24
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateContactVbErr:
        
    Set objIThirdPartyClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    '15/06/2001 MDC - OmiPlus 24 - Telephone Numbers
    Set xmlTelNodeList = Nothing
    Set xmlTelNode = Nothing
    'JR - Omiplus24
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
   
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

Private Sub CreateContactTelephone(ByVal vxmlContactTelephoneNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateContactTelephone an instance of the persistant data associated with this data object
'   for each set of data in the request to Update/Create Contact
' pass:
'   vxmlContactTelephoneNode - xml node containing the request CONTACTTELEPHONEDETAILS
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateContactTelephoneVbErr

    Const strFunctionName As String = "CreateContactTelephone"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIContactClassDef As IContactClassDef
    Dim objIDOAssist As IDOAssist
    Dim xmlTelNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlClassDefDoc = New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Set objIContactClassDef = New ContactClassDef
    Set xmlTelNode = vxmlContactTelephoneNode.cloneNode(True)
    
    Set xmlClassDefDoc = objIContactClassDef.LoadContactTelephoneData()
    objIDOAssist.GenerateSequenceNumber xmlTelNode, xmlClassDefDoc, "TELEPHONESEQNUM"
    objIDOAssist.Create xmlTelNode, xmlClassDefDoc
    
    Set objIContactClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlTelNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateContactTelephoneVbErr:
    
    Set objIContactClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlTelNode = Nothing
    
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

