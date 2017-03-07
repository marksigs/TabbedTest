VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LockAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LockAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Locking helper object
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' PSC   24/08/99 Created
' RF    02/12/99 AQR SC017: Enhancement - validate working hours on creating an application lock.
' IW    23/05/00 SYS0774 : DISTRIBUTIONCHANNELID S/B CHANNELID
'------------------------------------------------------------------------------------------
Option Explicit

Public Function CreateLock(ByVal vstrXMLRequest As String, _
                           ByVal vstrLockType As String, _
                           ByVal vstrKeyNumber As String) As String
' header ---------------------------------------------------------------------------------
' description:  Creates a lock based on Lock Type passed in
' pass:
'   vstrXmlRequest
'       Xml Request containing a valid request tag:
'   vstrLockType
'       The type of lock to create
'   vstrKeyNumber
'       The key to apply to the lock record
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLockVbErr

    Const strFunctionName As String = "CreateLock"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objRequestElem As IXMLDOMElement
    Dim objInRequestElem As IXMLDOMElement
    Dim objTypeElem As IXMLDOMElement
    Dim objTableElem As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objApplManagerBO As ApplicationManagerBO
    Dim objCustomerBO As CustomerBO

    Dim strXmlResponse As String
    Dim strTableName As String
    Dim strTypeNumber As String

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objInRequestElem = objXmlIn.getElementsByTagName("REQUEST").Item(0)

    If objInRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "REQUEST tag not found"
    End If

    ' Setup a create request
    strTableName = UCase(vstrLockType) & "LOCK"
    strTypeNumber = UCase(vstrLockType) & "NUMBER"

    Set objRequestElem = objInRequestElem.cloneNode(False)
    objXmlDoc.appendChild objRequestElem
    Set objTypeElem = objXmlDoc.createElement("CREATE")
    objRequestElem.appendChild objTypeElem
    Set objTableElem = objXmlDoc.createElement(strTableName)
    objTypeElem.appendChild objTableElem

    Set objElement = objXmlDoc.createElement(strTypeNumber)
    objElement.Text = vstrKeyNumber
    objTableElem.appendChild objElement
    Set objElement = objXmlDoc.createElement("UNITID")
    objElement.Text = _
        objXmlAssist.GetAttributeValue(objXmlDoc.documentElement, "REQUEST", "UNITID")
    objTableElem.appendChild objElement
'** TO DO **'------------------------------------------------------------------------------
' Check with Ivan, The user Id must be taken from the Tag value rather than attribute value
'------------------------------------------------------------------------------------------
    Set objElement = objXmlDoc.createElement("USERID")
    objElement.Text = _
        objXmlAssist.GetAttributeValue(objXmlDoc.documentElement, "REQUEST", "USERID")
    objTableElem.appendChild objElement
    Set objElement = objXmlDoc.createElement("MACHINEID")
    objElement.Text = _
        objXmlAssist.GetAttributeValue(objXmlDoc.documentElement, "REQUEST", "MACHINEID")
    objTableElem.appendChild objElement
    Set objElement = objXmlDoc.createElement("TYPEOFLOCK")
    objElement.Text = _
        objXmlAssist.GetAttributeValue(objXmlDoc.documentElement, "REQUEST", "TYPEOFLOCK")
    objTableElem.appendChild objElement

    ' Call the appropriate Business Object
    If UCase(vstrLockType) = "CUSTOMER" Then
        If objContext Is Nothing Then
            Set objCustomerBO = New CustomerBO
        Else
            Set objCustomerBO = _
                objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        End If
        strXmlResponse = objCustomerBO.CreateCustomerLock(objXmlDoc.xml)

'         strXmlResponse = "<RESPONSE TYPE=""SUCCESS""></RESPONSE>"
    ElseIf UCase(vstrLockType) = "APPLICATION" Then

        ' RF 02/12/99 AQR SC017: Enhancement -
        ' Validate working hours on creating an application lock.
        Set objElement = objXmlDoc.createElement("CHANNELID")
        objElement.Text = objXmlAssist.GetTagValue(objInRequestElem, "CHANNELID")
        objTableElem.appendChild objElement

         If objContext Is Nothing Then
            Set objApplManagerBO = New ApplicationManagerBO
        Else
            Set objApplManagerBO = _
                objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
        End If

        strXmlResponse = objApplManagerBO.CreateLock(objXmlDoc.xml)
    End If

    CreateLock = strXmlResponse

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlDoc = Nothing
    Set objRequestElem = Nothing
    Set objInRequestElem = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objElement = Nothing
    Set objApplManagerBO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

CreateLockVbErr:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlDoc = Nothing
    Set objRequestElem = Nothing
    Set objInRequestElem = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objElement = Nothing
    Set objApplManagerBO = Nothing

    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function DeleteLock(ByVal vstrXMLRequest As String, _
                           ByVal vstrLockType As String, _
                           ByVal vstrKeyNumber As String) As String
' header ---------------------------------------------------------------------------------
' description:  Deletes a lock based on Lock Type passed in
' pass:         vstrXmlRequest  Xml Request containing a valid request tag
'               vstrLockType    The type of lock to create
'               vstrKeyNumber   The key to apply to the lock record
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLockVbErr

    Const strFunctionName As String = "DeleteLock"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objRequestElem As IXMLDOMElement
    Dim objInRequestElem As IXMLDOMElement
    Dim objTypeElem As IXMLDOMElement
    Dim objTableElem As IXMLDOMElement
    Dim objElement As IXMLDOMElement
    Dim objCustomerBO As CustomerBO
    Dim objApplManagerBO As ApplicationManagerBO

    Dim strXmlResponse As String
    Dim strTableName As String
    Dim strTypeNumber As String

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objInRequestElem = objXmlIn.getElementsByTagName("REQUEST").Item(0)

    If objInRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "REQUEST tag not found"
    End If

    ' Setup a create request
    strTableName = UCase(vstrLockType) & "LOCK"
    strTypeNumber = UCase(vstrLockType) & "NUMBER"

    Set objRequestElem = objInRequestElem.cloneNode(False)
    objXmlDoc.appendChild objRequestElem
    Set objTypeElem = objXmlDoc.createElement("DELETE")
    objRequestElem.appendChild objTypeElem
    Set objTableElem = objXmlDoc.createElement(strTableName)
    objTypeElem.appendChild objTableElem

    Set objElement = objXmlDoc.createElement(strTypeNumber)
    objElement.Text = vstrKeyNumber
    objTableElem.appendChild objElement

    ' Call the appropriate Business Object
    If UCase(vstrLockType) = "CUSTOMER" Then
        If objContext Is Nothing Then
            Set objCustomerBO = New CustomerBO
        Else
            Set objCustomerBO = objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        End If

        strXmlResponse = objCustomerBO.DeleteCustomerLock(objXmlDoc.xml)

    ElseIf UCase(vstrLockType) = "APPLICATION" Then

        If objContext Is Nothing Then
            Set objApplManagerBO = New ApplicationManagerBO
        Else
            Set objApplManagerBO = objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
        End If

        strXmlResponse = objApplManagerBO.DeleteLock(objXmlDoc.xml)
    End If

    DeleteLock = strXmlResponse

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlDoc = Nothing
    Set objRequestElem = Nothing
    Set objInRequestElem = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objElement = Nothing
    Set objCustomerBO = Nothing
    Set objApplManagerBO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

DeleteLockVbErr:

    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlDoc = Nothing
    Set objRequestElem = Nothing
    Set objInRequestElem = Nothing
    Set objTypeElem = Nothing
    Set objTableElem = Nothing
    Set objElement = Nothing
    Set objCustomerBO = Nothing
    Set objApplManagerBO = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function



