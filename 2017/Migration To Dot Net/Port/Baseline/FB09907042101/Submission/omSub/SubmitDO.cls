VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "SubmitDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SubmitDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Submit Data Object
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     29/03/00    Created.
'MH     05/05/00    SYS0571 - use of Now without formatting can be affected by Locale
'MC     30/06/00    Added GetSubmissionStatusForApplication method
'LD     07/11/00    Explicity destroy command objects
'LD     07/11/00    Explicity destroy command objects
'------------------------------------------------------------------------------------------

Option Explicit

Implements ISubmitDO
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


Private Sub ISubmitDO_CreateSubmissionStatus( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an ApplicationSubmissionStatus record.
' pass:
'   vxmlTableElement
'       xml element containing the request
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateSubmissionStatusVbErr

    Const strFunctionName As String = "ISubmitDO_CreateSubmissionStatus"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomSubClassDef
    Dim xmlNewRequestDoc As New FreeThreadedDOMDocument40
    xmlNewRequestDoc.validateOnParse = False
    xmlNewRequestDoc.setProperty "NewParser", True
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    '------------------------------------------------------------------------------------------
    ' set up SUBMISSIONDATETIME
    ' n.b. the request node is copied so that it can be safely changed
    '------------------------------------------------------------------------------------------
    
    xmlNewRequestDoc.appendChild vxmlTableElement.cloneNode(True)
    
    Dim strTableName As String
    strTableName = "APPLICATIONSUBMISSIONSTATUS"
    
    Set xmlTableNode = xmlNewRequestDoc.getElementsByTagName(strTableName).Item(0)
    
    If xmlTableNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "Expected " & strTableName & " tag"
    End If
    
    Set xmlElem = xmlNewRequestDoc.createElement("SUBMISSIONDATETIME")
    xmlElem.Text = Format$(Now(), "DD/MM/YYYY HH:MM:SS")
    xmlTableNode.appendChild xmlElem
    
    Set objIClassDef = New omSubClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationSubmissionStatusData()
    m_objIDOAssist.Create xmlNewRequestDoc.documentElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNewRequestDoc = Nothing
    Set xmlTableNode = Nothing
    Set xmlElem = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateSubmissionStatusVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNewRequestDoc = Nothing
    Set xmlTableNode = Nothing
    Set xmlElem = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function ISubmitDO_GetSubmissionStatusForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the ApplicationSubmissionStatus record for the specified application
' pass:
'   vxmlTableElement
'       xml element containing the request
' return:
'   xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetSubmissionStatusForApplicationVbErr

    Const strFunctionName As String = "ISubmitDO_GetSubmissionStatusForApplication"
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomSubClassDef
    Dim objIADOAssist As IADOAssist      ' for GetConnStr
    Dim objIDOAssist As IDOAssist
    Dim objResponseNode As IXMLDOMNode
    
    Set objResponseNode = vxmlTableElement.ownerDocument.createElement("RESPONSE")
    
    Set objIClassDef = New omSubClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationSubmissionStatusData()
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strSQL As String
    
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, ".//APPLICATIONNUMBER")
    strApplicationFFNumber = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, ".//APPLICATIONFACTFINDNUMBER")
        
    strSQL = "select * from APPLICATIONSUBMISSIONSTATUS where "
    strSQL = strSQL & "APPLICATIONNUMBER = '" & strApplicationNumber & "'"
    strSQL = strSQL & " and APPLICATIONFACTFINDNUMBER = " & strApplicationFFNumber
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    Set rst = cmd.Execute(, , adCmdText)
    
    If rst.BOF And rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
        
    'Dim xmlNode As IXMLDOMNode
    Set ISubmitDO_GetSubmissionStatusForApplication = objIDOAssist.GetXMLFromRecordset(rst, xmlClassDefDoc) ', objResponseNode)

    'Set ISubmitDO_GetSubmissionStatusForApplication = objResponseNode
    
    Set cmd = Nothing
	Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetSubmissionStatusForApplicationVbErr:
    
    Set cmd = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ISubmitDO_UpdateSubmissionStatus( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this data object.
' pass:
'   vxmlTableElement
'       xml element containing the request
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSubmissionStatusVbErr

    Const strFunctionName As String = "ISubmitDO_UpdateSubmissionStatus"
            
    Dim xmlNewRequestDoc As New FreeThreadedDOMDocument40
    xmlNewRequestDoc.validateOnParse = False
    xmlNewRequestDoc.setProperty "NewParser", True
    Dim xmlElem As IXMLDOMElement
    Dim xmlTableNode As IXMLDOMNode
    
    '------------------------------------------------------------------------------------------
    ' set SUBMISSIONDATETIME
    ' n.b. the request node is copied so that it can be safely changed
    '------------------------------------------------------------------------------------------
    
    xmlNewRequestDoc.appendChild vxmlTableElement.cloneNode(True)
    
    Dim strTableName As String
    strTableName = "APPLICATIONSUBMISSIONSTATUS"
    
    Set xmlTableNode = xmlNewRequestDoc.getElementsByTagName(strTableName).Item(0)
    
    If xmlTableNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "Expected " & strTableName & " tag"
    End If
    
    Set xmlElem = xmlNewRequestDoc.createElement("SUBMISSIONDATETIME")
    xmlElem.Text = Format$(Now(), "DD/MM/YYYY HH:MM:SS")
    xmlTableNode.appendChild xmlElem
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomSubClassDef
    
    Set objIClassDef = New omSubClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationSubmissionStatusData()
    m_objIDOAssist.Update xmlNewRequestDoc.documentElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNewRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlTableNode = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateSubmissionStatusVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNewRequestDoc = Nothing
    Set xmlElem = Nothing
    Set xmlTableNode = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    
    If m_objContext Is Nothing Then
        Set m_objIDOAssist = New DOAssist
    Else
        Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

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

Private Function ISubmitDO_GetSubmissionStatusData( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the oldest ApplicationSubmissionStatus record having a status matching an
'   input value.
' pass:
'   vxmlTableElement
'       xml element containing the request
' return:
'   xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetSubmissionStatusDataVbErr

    Const strFunctionName As String = "ISubmitDO_GetSubmissionStatusData"
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomSubClassDef
    Dim objIADOAssist As IADOAssist      ' for GetConnStr
    Dim objIDOAssist As IDOAssist
    
    Set objIClassDef = New omSubClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationSubmissionStatusData()
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim strStatus, strSQL As String
    
    strStatus = m_objXmlAssist.GetMandatoryElementText( _
        vxmlTableElement, ".//SUBMISSIONSTATUS")
        
    strSQL = "select * from APPLICATIONSUBMISSIONSTATUS where "
    strSQL = strSQL & "SUBMISSIONSTATUS = " & strStatus
    strSQL = strSQL & " and SUBMISSIONDATETIME = "
    strSQL = strSQL & "(select min(SUBMISSIONDATETIME) from APPLICATIONSUBMISSIONSTATUS where "
    strSQL = strSQL & "SUBMISSIONSTATUS = " & strStatus & ")"
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    Set rst = cmd.Execute(, , adCmdText)
    
    If rst.BOF And rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
        
    Set ISubmitDO_GetSubmissionStatusData = _
        objIDOAssist.GetXMLFromRecordset(rst, xmlClassDefDoc)

    Set cmd = Nothing
	Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetSubmissionStatusDataVbErr:
    
    Set cmd = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


