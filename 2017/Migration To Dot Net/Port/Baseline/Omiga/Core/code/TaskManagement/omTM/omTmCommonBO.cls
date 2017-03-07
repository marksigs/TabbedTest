VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omTmCommonBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OmTmCommonBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Description:   Task Manager Business Object.
'               This class has methods that need to be called within a
'               transaction from omTmBO and within a new transaction from
'               omTmNTTxBO
'Dependencies:  Add any other dependent components
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    22/02/01    SYS3997 Created
'-------------------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IomTmCommonBO
Private gobjContext As ObjectContext
'----------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     14/08/2002  BMIDS00333  Core Ref AQr: SYS4919 ; Modified IomTmCommonBO_GetApplicationData
'----------------------------------------------------------------------------------
Private Sub IomTmCommonBO_GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
                                            ByVal vxmlApplicationDestDoc As FreeThreadedDomDocument40, _
                                            Optional ByVal vstrRBTemplateName As String)
    On Error GoTo IomTmCommonBO_GetApplicationDataExit
    Const cstrFunctionName As String = "IomTmCommonBO_GetApplicationData"
    Dim xmlRequestDoc As FreeThreadedDomDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objRB As Object
    gobjContext.SetComplete
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.async = False
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib
    If Len(vstrRBTemplateName) > 0 Then
        Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
        xmlAttrib.Text = vstrRBTemplateName
        xmlNode.Attributes.setNamedItem xmlAttrib
    End If
    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONNUMBER").Text
    If vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER").Text
    End If
    ' _EXTRACTTYPE_ = "NODE" added for CostModelCheckRequired
    xmlCopyAttribute vxmlApplicationSrceNode, xmlElem, "_EXTRACTTYPE_"
    xmlNode.appendChild xmlElem
    Debug.Print "IomTmCommonBO_GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    vxmlApplicationDestDoc.setProperty "NewParser", True
    vxmlApplicationDestDoc.validateOnParse = False
    vxmlApplicationDestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    Debug.Print "IomTmCommonBO_GetApplicationData response"
    Debug.Print vxmlApplicationDestDoc.xml
    errCheckXMLResponseNode vxmlApplicationDestDoc.selectSingleNode("RESPONSE"), , True
    If vxmlApplicationDestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError _
            "IomTmCommonBO_GetApplicationData", _
            oeXMLMissingElement, _
            "No APPLICATION detail"
    End If
IomTmCommonBO_GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    errCheckError cstrFunctionName
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
