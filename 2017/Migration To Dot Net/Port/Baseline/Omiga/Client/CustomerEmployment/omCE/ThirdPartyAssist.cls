VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ThirdPartyAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ThirdPartyAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Helper object for handling errors
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'AD     17/01/00 Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements IThirdPartyAssist

'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        AQR        Description
'SA     12/11/2002  BMIDS00693 Delete Third Party error hadnling changed to not set abort.

Private Sub IThirdPartyAssist_DeleteThirdParty(ByVal vstrThirdPartyGUID As String)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteThirdParty an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyAssist_DeleteThirdParty"
   
    Dim objErrAssist As New ErrAssist

    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim objIThirdPartyDO As IThirdPartyDO

    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyDO")
    End If

    If (vstrThirdPartyGUID <> "") Then
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        xmlTempDoc.appendChild xmlTempDoc.createElement("THIRDPARTY")
        xmlTempDoc.documentElement.appendChild(xmlTempDoc.createElement("THIRDPARTYGUID")).Text = vstrThirdPartyGUID

        objIThirdPartyDO.DeleteThirdParty xmlTempDoc.documentElement
    End If
    
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlTempDoc = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteThirdPartyVbErr:
    
    Set objIThirdPartyDO = Nothing
    Set xmlTempDoc = Nothing
    
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

Private Sub IThirdPartyAssist_GetThirdParty(ByRef vxmlDataNode As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   GetThirdParty an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo GetThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyAssist_GetThirdParty"
   
    Dim objErrAssist As New ErrAssist, _
        objXmlAssist As New XMLAssist

    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode     As IXMLDOMNode
    Dim xmlReturn As IXMLDOMNode
    Dim objIThirdPartyDO As IThirdPartyDO

    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String, _
        strGuid           As String
    Dim strGUIDType As String
    Dim strTable As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyDO")
    End If

    '
    ' Initialise
    '
    strThirdPartyGUID = objXmlAssist.GetTagValue(vxmlDataNode, "THIRDPARTYGUID")
    strDirectoryGUID = objXmlAssist.GetTagValue(vxmlDataNode, "DIRECTORYGUID")

    ' Decide which GUID to use in the get
    If strDirectoryGUID <> "" Then
        strGuid = strDirectoryGUID
        strGUIDType = "DIRECTORYGUID"
        strTable = "NAMEANDADDRESSDIRECTORY"
    ElseIf strThirdPartyGUID <> "" Then
        strGuid = strThirdPartyGUID
        strGUIDType = "THIRDPARTYGUID"
        strTable = "THIRDPARTY"
    End If

    '
    ' Main
    '
    If (strThirdPartyGUID <> "") Or (strDirectoryGUID <> "") Then
        ' Compile the request XML
        Set xmlTempDoc = New FreeThreadedDOMDocument40
        xmlTempDoc.validateOnParse = False
        xmlTempDoc.setProperty "NewParser", True
        xmlTempDoc.appendChild xmlTempDoc.createElement(strTable)
        xmlTempDoc.documentElement.appendChild(xmlTempDoc.createElement(strGUIDType)).Text = strGuid

        ' Execute the get
        If strTable = "THIRDPARTY" Then
            Set xmlReturn = objIThirdPartyDO.GetThirdParty(xmlTempDoc.documentElement)
        Else
            Set xmlReturn = objIThirdPartyDO.GetDirectorySummary(xmlTempDoc.documentElement)
        End If

        ' Attach the returned details in the appropriate place in the primary table XML
        vxmlDataNode.appendChild xmlReturn
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturn = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

GetThirdPartyVbErr:
    
    Set objXmlAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturn = Nothing
    
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

Private Sub IThirdPartyAssist_MaintainThirdParty(ByRef rxmlPrimaryTableNode As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   MaintainThirdParty an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo MaintainThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyAssist_MaintainThirdParty"
   
    Dim objErrAssist As New ErrAssist

    Dim objXmlAssist As New XMLAssist
    Dim xmlThirdPartyNode As IXMLDOMNode
    Dim xmlThirdPartyGUIDNode As IXMLDOMNode, _
        xmlDirectoryGUIDNode  As IXMLDOMNode

    Dim objIThirdPartyBO As IThirdPartyBO
    Dim xmlReturn As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim strOperation     As String, _
        strGeneratedGUID As String
    Dim strDirectoryGUID As String

    Dim xmlNodeList As IXMLDOMNodeList, _
        xmlNode     As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objIThirdPartyBO = New ThirdPartyBO
    Else
        Set objIThirdPartyBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    End If

    '
    ' Initialise
    '
    ' Find the THIRDPARTY details in the XML
    Set xmlThirdPartyNode = rxmlPrimaryTableNode.selectSingleNode("THIRDPARTY")

    Set xmlThirdPartyGUIDNode = rxmlPrimaryTableNode.selectSingleNode("THIRDPARTYGUID")
    Set xmlDirectoryGUIDNode = rxmlPrimaryTableNode.selectSingleNode("DIRECTORYGUID")

    If Not (xmlDirectoryGUIDNode Is Nothing) Then
        strDirectoryGUID = Trim$(xmlDirectoryGUIDNode.Text)
    End If

    '
    ' Main
    '
    If Not (xmlThirdPartyNode Is Nothing) And (strDirectoryGUID = "") Then
        ' THIRDPARTY details have been specified, so save those details
        Set xmlReturn = objIThirdPartyBO.SaveThirdParty(xmlThirdPartyNode)
        objErrAssist.CheckXMLResponse xmlReturn.xml, vblnRaiseError:=True

        ' Check the response received from the save
        strOperation = objXmlAssist.GetAttributeValue(xmlReturn, "RESPONSE", "OPERATION")
        strGeneratedGUID = objXmlAssist.GetTagValue(xmlReturn, "THIRDPARTYGUID")

        ' Find the THIRDPARTYGUID node in the primary table XML
        If xmlThirdPartyGUIDNode Is Nothing Then
            ' Create the THIRDPARTYGUID node
            Set xmlThirdPartyGUIDNode = rxmlPrimaryTableNode.appendChild( _
                rxmlPrimaryTableNode.ownerDocument.createElement("THIRDPARTYGUID"))
        End If

        Select Case strOperation
            Case "CREATE":
                ' Update the primary table XML with the generated THIRDPARTYGUID
                xmlThirdPartyGUIDNode.Text = strGeneratedGUID
            Case "DELETE":
                ' Remove any reference in the primary table XML to a THIRDPARTY
                xmlThirdPartyGUIDNode.Text = ""
            Case Else
                '  N/A
        End Select
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIThirdPartyBO = Nothing
    Set xmlThirdPartyNode = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlReturn = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

MaintainThirdPartyVbErr:
    
    Set objXmlAssist = Nothing
    Set objIThirdPartyBO = Nothing
    Set xmlThirdPartyGUIDNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlThirdPartyNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlReturn = Nothing
    
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

