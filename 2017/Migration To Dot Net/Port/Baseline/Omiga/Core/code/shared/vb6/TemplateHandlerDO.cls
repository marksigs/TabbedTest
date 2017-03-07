VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TemplateHandlerDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      TemplateHandlerDO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'Description:
'       Handles print template database.
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     02/11/00    Created
'IK     17/02/2003  BM0200 - add TraceAssist support (remove most of OOSS inheritance)
'------------------------------------------------------------------------------------------
Option Explicit
Implements ITemplateHandlerDO
Implements ObjectControl
'=============================================
'Constant Declaration Section
'=============================================
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
Private m_objIDOAssist As IDOAssist

Private Function ITemplateHandlerDO_GetTemplate( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement
'       xml element containing the request
' return:
'   xml node containing retrieved data
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "ITemplateHandlerDO_GetTemplate"

Dim objIClassDef As IomDPSClassDef

Dim xmlClassDefDoc As FreeThreadedDOMDocument40

    ' ik_20030210
    On Error GoTo GetTemplateVbErr
    m_objContext.SetComplete

    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If

    Set objIClassDef = New omDPSClassDef
    Set xmlClassDefDoc = objIClassDef.LoadTemplateData()

    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "DOAssist.GetData call"
    Set ITemplateHandlerDO_GetTemplate = _
        m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "DOAssist.GetData return"
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
GetTemplateVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    m_objErrAssist.AddToErrSource (cstrFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If

    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ITemplateHandlerDO_FindAvailableTemplates( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement
'       xml element containing the request
' return:
'   xml node containing retrieved data
'------------------------------------------------------------------------------------------
Const cstrFunctionName As String = "ITemplateHandlerDO_FindAvailableTemplates"

Dim adoCommand As ADODB.Command
Dim adoConnection As ADODB.Connection
Dim adoParameter As ADODB.Parameter
Dim adoRecordset As ADODB.Recordset

Dim blnOpenedOk As Boolean

Dim intAttempt As Integer
Dim intMaxAttempts As Integer
Dim intRetries As Integer

Dim objIADOAssist As IADOAssist
Dim objIClassDef As IomDPSClassDef
Dim objXmlAssist As XMLAssist

Dim strConnString As String
Dim strLanguage As String
Dim strSQL As String
Dim strSecurityLevel As String
Dim strStageNumber As String

Dim xmlClassDefDoc As FreeThreadedDOMDocument40
Dim xmlListElem As IXMLDOMElement
Dim xmlOut As FreeThreadedDOMDocument40
    
    On Error GoTo FindAvailableTemplatesVbErr
    ' ik_20030210
    m_objContext.SetComplete
    
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeMTSNotFound
    End If
    Set objXmlAssist = New XMLAssist
    ' Check the correct keys have been passed in
    strSecurityLevel = objXmlAssist.GetNodeValue(vxmlTableElement, ".//SECURITYLEVEL", True)
    strStageNumber = objXmlAssist.GetNodeValue(vxmlTableElement, ".//STAGENUMBER", True)
    strLanguage = objXmlAssist.GetNodeValue(vxmlTableElement, ".//LANGUAGE", True)
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    strSQL = "SELECT * FROM TEMPLATE WHERE SECURITYLEVEL <= ? AND " & _
             "STAGENUMBER >=  ? AND LANGUAGE = ?"
    adoCommand.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("SecurityLevel", adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(0) = strSecurityLevel
    Set adoParameter = adoCommand.CreateParameter("StageNumber", adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(1) = strStageNumber
    Set adoParameter = adoCommand.CreateParameter("Language", adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(2) = strLanguage
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    intRetries = objIADOAssist.GetRetries()
    intMaxAttempts = 1 + intRetries
    blnOpenedOk = False
    intAttempt = 1
    strConnString = objIADOAssist.GetConnStr()
    ' Keep trying to open the connection for the specified number of retries
    While (blnOpenedOk = False) And (intAttempt <= intMaxAttempts)
        adoConnection.ERRORS.Clear
        adoConnection.open strConnString
        If adoConnection.ERRORS.Count = 0 Then
            blnOpenedOk = True
        End If
        intAttempt = intAttempt + 1
    Wend
    If blnOpenedOk = False Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeUnableToConnect
    End If
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.1]", "adoRecordset.Open call"
    'Populate recordset
    adoRecordset.open
    
    gobjTrace.TraceMessage TypeName(Me), cstrFunctionName, "[1.2]", "adoRecordset.Open return"
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omDPSClassDef
    Set xmlClassDefDoc = objIClassDef.LoadTemplateData()
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("TEMPLATELIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
    adoRecordset.Close
    Set ITemplateHandlerDO_FindAvailableTemplates = xmlListElem
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
FindAvailableTemplatesVbErr:
    
    ' Close the recordset if still open
    If Not adoRecordset Is Nothing Then
        If adoRecordset.State = adStateOpen Then
            adoRecordset.Close
        End If
    End If
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    m_objErrAssist.AddToErrSource (cstrFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    gobjTrace.TraceMethodError TypeName(Me), cstrFunctionName
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
