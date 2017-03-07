VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PersonDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
''Workfile:      CustomerDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vxmlTableElement As IXMLDOMElement)
'               Public Sub Update(vxmlTableElement As IXMLDOMElement)
'               Public Sub Delete(vxmlTableElement As IXMLDOMElement)
'               Public Function GetData(vxmlTableElement As IXMLDOMElement,
'                                       vxmlInNode As IXMLDOMNode) As IXMLDOMNode
'               Public Function AddDerivedData(vxmlData As IXMLDOMNode)
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'SR    04/11/99  Created
'LD    19/06/01  SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'------------------------------------------------------------------------------------------

Option Explicit
Implements IPersonDO

Private Function IPersonDO_Save(ByVal vxmlTableElement As IXMLDOMElement, _
                            Optional ByRef rbotOperationType As Variant) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Saves the Person records specified in the passed XML. The routine
'               examines each individual transaction within the XML and calls the corresponding
'               Update, Delete or Create method.
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:       If the operation performed is 'Create'. returns the node created by that method
'               and nothing in case of other operation
'------------------------------------------------------------------------------------------
On Error GoTo IPersonDO_SaveVbErr

    Const strFunctionName As String = "IPersonDO_Save"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim objIPersonClassDef As IPersonClassDef
    Dim xmlPersonDefDoc As New FreeThreadedDOMDocument40
    xmlPersonDefDoc.validateOnParse = False
    xmlPersonDefDoc.setProperty "NewParser", True

    Dim botOperationType As BOOPERATIONTYPE

    botOperationType = booNone
    If Not IsMissing(rbotOperationType) Then botOperationType = rbotOperationType
    
    Set objIPersonClassDef = New PersonClassdef
    Set xmlPersonDefDoc = objIPersonClassDef.LoadPersonData
    If botOperationType = booNone Then
        botOperationType = objXmlAssist.DetermineOperationEx(vxmlTableElement, xmlPersonDefDoc)
    End If

    Select Case botOperationType
        Case booCreate: Set IPersonDO_Save = IPersonDO_Create(vxmlTableElement)
        Case booUpdate: IPersonDO_Update vxmlTableElement
        Case booDelete: IPersonDO_Delete vxmlTableElement
        Case Else
            ' N/A
    End Select

    If Not IsMissing(rbotOperationType) Then
        rbotOperationType = botOperationType
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlPersonDefDoc = Nothing

    Exit Function

IPersonDO_SaveVbErr:
    Set objXmlAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set xmlPersonDefDoc = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
        Set objContext = Nothing
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IPersonDO_Create(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with table 'Person'
'   for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Dependencies : CustomerBO, CustomerTxBO
' Raise Errors :
'------------------------------------------------------------------------------------------
On Error GoTo IPersonDO_CreateVbErr

    Const strFunctionName As String = "IPersonDO_Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim objIPersonClassDef As New IPersonClassDef
    Dim objXmlAssist As New XMLAssist

    Dim xmlResponseElem As IXMLDOMElement 'contains the PersonGuid generated
    Dim xmlPersonElem As IXMLDOMElement 'used to add in the XML Out document
    Dim xmlPersonGuidElem As IXMLDOMElement 'Used to refer in the Input
    Dim strGuid As String, strTagName As String

    ' Check whether PersonGuid element exists. If not add one with value of
    ' newly generated Guid. If exists, check whether vlaue is mentioned, and
    ' assign generated Guid if value was not mentioned

    strTagName = "PERSONGUID"
    Set xmlPersonGuidElem = vxmlTableElement.getElementsByTagName(strTagName).Item(0)

    If xmlPersonGuidElem Is Nothing Then
        Set xmlPersonGuidElem = vxmlTableElement.ownerDocument.createElement(strTagName)
        vxmlTableElement.appendChild xmlPersonGuidElem
    End If

    If Len(xmlPersonGuidElem.Text) = 0 Then
        strGuid = CreateGUID()
        xmlPersonGuidElem.Text = strGuid
    Else
        strGuid = xmlPersonGuidElem.Text
    End If

    'Append elements with table name and PrimaryKey to response
    Set xmlResponseElem = vxmlTableElement.ownerDocument.createElement("RESPONSE")
    Set xmlPersonElem = vxmlTableElement.ownerDocument.createElement("PERSON")
    xmlResponseElem.appendChild xmlPersonElem

    strTagName = "PERSONGUID"
    Set xmlPersonGuidElem = vxmlTableElement.ownerDocument.createElement(strTagName)
    xmlPersonGuidElem.Text = strGuid
    xmlPersonElem.appendChild xmlPersonGuidElem

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    'Pass the modified XML request node to create method of DOAssist
    Set objIPersonClassDef = New PersonClassdef
    objIDOAssist.Create vxmlTableElement, objIPersonClassDef.LoadPersonData
    Set IPersonDO_Create = xmlResponseElem

    Set objIDOAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing

    Set xmlPersonGuidElem = Nothing
    Set xmlPersonElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Function

IPersonDO_CreateVbErr:

    Set objIDOAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set objXmlAssist = Nothing

    Set xmlPersonGuidElem = Nothing
    Set xmlPersonElem = Nothing

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '  re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub IPersonDO_Delete(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with table 'Person'
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance to be deleted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IPersonDO_DeleteVbErr

    Const strFunctionName As String = "IPersonDO_Delete"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIPersonClassDef As New IPersonClassDef

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIPersonClassDef = New PersonClassdef
    
    objIDOAssist.Delete vxmlTableElement, objIPersonClassDef.LoadPersonData

    Set objIDOAssist = Nothing
    Set objIPersonClassDef = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

IPersonDO_DeleteVbErr:

    Set objIDOAssist = Nothing
    Set objIPersonClassDef = Nothing
    
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
        Set objContext = Nothing
    End If
    
    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub IPersonDO_Update(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'    Update a single instance of the persistant data associated with table 'Person'
' pass:
'    vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IPersonDO_UpdateVbErr

    Const strFunctionName As String = "IPersonDO_Update"

    Dim objErrAssist As New ErrAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objIDOAssist As IDOAssist
    Dim objIPersonClassDef As New IPersonClassDef

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIPersonClassDef = New PersonClassdef
    objIDOAssist.Update vxmlTableElement, objIPersonClassDef.LoadPersonData

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

    Exit Sub

IPersonDO_UpdateVbErr:

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
        Set objContext = Nothing
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

