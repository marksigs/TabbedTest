VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "OrganisationBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OrganisationBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Code template for omiga4 Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (OrganisationTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)
'Dependencies:  OrganisationTxBO, OrganisationDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' RF    24/09/99 Created.
' RF    03/11/99 ValidateUserLogon returns UserName.
' RF    04/11/99 Improved error reporting in ValidateUserLogon and ValidateCurrentPassword.
' RF    11/11/99 Use ErrAssist.ThrowError rather than RaiseError.
' RF    18/11/99 Change to error numbers used in ValidatePasswordFormat.
' RF    01/12/99 AQR SC014: Problem in ValidateCurrentPassword error handling.
' RF    02/12/99 AQR SC017: Enhancement - make validate working hours public.
' RF    27/01/00 Pick up ScriptInterface object from omBase.
' SR    08/02/00 Change the checking condition in ValidataWorkingHours
' SR    10/03/00 Method:ValidateUserLogOn ConvertAssist.CSafeBoolean instead of CBool
' MC    16/05/00 SYS0210 - Synchronise the password change date/time with corresponding
'                access audit record
' MC    17/05/00 SYS0166 - Update the ChangePasswordIndicator when password changed
' IW    23/05/00 SYS0774 DISTRIBUTIONCHANNELID S/B/ CHANNNELID
' MS    23/06/00 Added timing output
' MV    21/11/00 CORE000011 : Added New FindDistributionChannelList,FindDepartmentList
'                FindUserList,FindUnitList Functions
' CL    27/11/00  CORE00005 Modifications to meet new standards
' DJP   14/12/00 Added FindUserNameList and FindUnitNameList
' DJP   04/01/01 SYS1739 Added ValidateUserAuthority
' CL    27/02/01 SYS1982 Changed ValidateCurrentPassword to Public to enable use by GUI#
' MV    05/03/01 SYS2001 Commented all the AccessAuditGUID processing
' DRC   05/04/01 SYS1925 Added GetActionOwnerDetails
' GD    18/04/01 SYS2092 removed setabort from GetActionOwnerDetails
' DRC   05/10/01 SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
' STB   10/04/02 SYS2339 Re-enabled the AccessAudit functionality removed in SYS2001.
' IK    05/08/04 BBG844  re-write of FindUserList for performance
' SAB   21/09/05 CORE191 Added IsDateWorkingDay()
' JD    02/02/06 CORE237 Added GetUserCompetency
' AS    30/05/06 CORE275 Password expiry only uses date and not time.
' AS    07/06/06 CORE278 MaxLogonAttempts unnecessary in AD mode.
' PSC   23/06/06 CORE278 Take into account global parameter not being present
'------------------------------------------------------------------------------------------
Option Explicit
Implements IOrganisationBO
Implements ObjectControl
'=============================================
'Constant Declaration Section
'=============================================
Private m_objContext As ObjectContext
'=============================================
'Variable Declaration Section
'=============================================
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMids
'   Aw  15/08/02    BM029   Added UpdateHighChequeNum()
'
Private Sub Validate(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   performs validation for this object
'   if hard-coded rules completed OK calls scripting interface to perform
'   user scripted rules as appropriate
' pass:
'   vstrXMLRequest
'       xml Request data stream containing details of action to be
'       performed and data required
'       This is the full request block as received from the client
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr
    Const strFunctionName As String = "Validate"
        
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlBeforeData As IXMLDOMElement
    Dim xmlAfterData As IXMLDOMElement
    Dim xmlAfterNode As IXMLDOMNode
    Dim objScriptInterface As ScriptInterface
    Dim objIOrganisationDO As IOrganisationDO
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
            
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
            
' TODO ************************************************************************
'   add processing here
'   e.g.
'   examine input xml request & extract before image node (if present)
'   & after image node
' *****************************************************************************
    
    If xmlIn.getElementsByTagName("UPDATE").Length > 0 Then
        Dim int0 As Integer
        For int0 = 0 To xmlIn.getElementsByTagName("UPDATE").Length - 1
            Select Case xmlIn.getElementsByTagName("UPDATE").Item(int0).Attributes.getNamedItem("TYPE").Text
                Case "BEFORE"
                    Set xmlBeforeData = xmlIn.getElementsByTagName("UPDATE").Item(int0).firstChild
                Case "AFTER"
                    Set xmlAfterNode = xmlIn.getElementsByTagName("UPDATE").Item(int0)
                    Set xmlAfterData = xmlAfterNode.firstChild
            End Select
        Next
    Else
        If xmlIn.getElementsByTagName("CREATE").Length > 0 Then
            Set xmlAfterNode = xmlIn.getElementsByTagName("CREATE").Item(int0)
            Set xmlAfterData = xmlAfterNode.firstChild
        End If
    End If
    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' call scripting interface
    ' -----------------------------------------------------------------------------
    'Dim objScriptInterface As New omBase.ScriptInterface
    Set objScriptInterface = New ScriptInterface
    objScriptInterface.RunScript App.Title, TypeName(Me), strFunctionName, vstrXMLRequest
    ' fixme - check xml response
    '   clean-up code
    Set objScriptInterface = Nothing
    Set xmlBeforeData = Nothing
    Set xmlAfterData = Nothing
    Set xmlAfterNode = Nothing
    Set xmlIn = Nothing
    Set objIOrganisationDO = Nothing
    Exit Sub
ValidateVbErr:
    '   clean-up code
    Set objScriptInterface = Nothing
    Set xmlBeforeData = Nothing
    Set xmlAfterData = Nothing
    Set xmlAfterNode = Nothing
    Set xmlIn = Nothing
    Set objIOrganisationDO = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function ValidateUserLogon(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Validate the UserId, UserPassword and UnitId entered by the current user.
'   Processing is delegated to DoValidateUserLogon and CreateAccessAudit in order to ensure
'   an AccessAudit record is created even upon error, and that in doing so the error info is
'   not lost.
' pass:
'   vstrXMLRequest
'       Format:
'        <REQUEST USERID= MACHINEID=>
'           <OMIGAUSER>
'               <USERID></USERID>
'               <PASSWORDVALUE></PASSWORDVALUE>
'               <UNIT></UNIT>
'           </OMIGAUSER>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     04/11/99 Improved error handling
'------------------------------------------------------------------------------------------
On Error GoTo ValidateUserLogonVbErr
    
    Const strFunctionName As String = "ValidateUserLogon"
           
    Dim xmlOut As New FreeThreadedDOMDocument40
    'Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objIAuditDO As IAuditDO
    Dim strResponse As String
    Dim lngMsgNo As Long
    Dim blnIsAttemptOK
    Dim objIOrganisationDO As IOrganisationDO
    'Dim strTempResponse As String
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    Dim strLogonAuditRecType As String
    strLogonAuditRecType = objIAuditDO.GetLogonValueId()
    '---------------------------------------------------------------------------
    ' Delegate the validate user logon processing to DoValidateUserLogon
    '---------------------------------------------------------------------------
    strResponse = DoValidateUserLogon(vstrXMLRequest, strLogonAuditRecType)
       
    lngMsgNo = m_objErrAssist.CheckXMLResponse(strResponse, False)
    blnIsAttemptOK = True
    If lngMsgNo <> 0 Then
        If m_objErrAssist.IsWarning(lngMsgNo) = False Then
            blnIsAttemptOK = False
        End If
    End If
    '---------------------------------------------------------------------------
    ' create the access audit record even on non-warning error
    '---------------------------------------------------------------------------
    'STB: SYS2339 - ACCESSAUDIT records added during logon (both sucess and failure).
    CreateAccessAudit vstrXMLRequest, strLogonAuditRecType, blnIsAttemptOK
    'STB: SYS2339 - End.
    'strTempResponse = CreateAccessAudit(vstrXMLRequest, strLogonAuditRecType, blnIsAttemptOK)
    'Set xmlDoc = m_objXmlAssist.load(strTempResponse, TypeName(Me), strFunctionName)
    'Set xmlTempNode = xmlDoc.getElementsByTagName("ACCESSAUDITGUID").Item(0)
    'xmlResponseElem.appendChild
    '---------------------------------------------------------------------------
    ' ensure validate error is raised
    '---------------------------------------------------------------------------
    If lngMsgNo <> 0 Then
        m_objErrAssist.CheckXMLResponse strResponse, True
    End If
    '---------------------------------------------------------------------------
    ' set up return data
    '---------------------------------------------------------------------------
    Set xmlElem = xmlOut.createElement("ACCESSAUDITATTEMPT")
    If blnIsAttemptOK = True Then
        xmlElem.Text = "1"
    Else
        xmlElem.Text = "0"
    End If
    xmlResponseElem.appendChild xmlElem
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    'xmlDoc.firstChild.appendChild xmlTempNode
    Set xmlNode = xmlDoc.getElementsByTagName("RESPONSE").Item(0)
    If xmlNode Is Nothing Then
       m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInternalError, _
           "RESPONSE tag missing"
    End If
    For Each xmlTempNode In xmlNode.childNodes
        xmlResponseElem.appendChild xmlTempNode.cloneNode(True)
    Next xmlTempNode
    'if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    ValidateUserLogon = xmlOut.xml
    m_objContext.SetComplete
ValidateUserLogonExit:
    ' clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseNode = Nothing
    Set xmlNode = Nothing
    Set objIAuditDO = Nothing
    Set xmlResponseNode = Nothing
    Set xmlNode = Nothing
    Set xmlTempNode = Nothing
    Set objIAuditDO = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElem = Nothing
     
    Exit Function
ValidateUserLogonVbErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    ValidateUserLogon = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume ValidateUserLogonExit
End Function
Public Function FindCurrentUnitList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get Unit, UserRole, Department data for the supplied UserId where all are active
' pass:
'   vstrXMLRequest
'       xml Request data stream containing data to which identifies
'       instance of the persistant data to be retrieved
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS" and xml representation of data
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindCurrentUnitListVbErr
    
    Const strFunctionName As String = "FindCurrentUnitList"
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlData As New FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objIOrganisationDO As IOrganisationDO
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlReturnNode As IXMLDOMNode
    Dim strUserId As String
    Dim strXMLData As String
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
            
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    strUserId = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "USERID")
    If Len(strUserId) = 0 Then
      m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 106
    End If
                
    ' Convert input into dom and output into string for compatibility
    Set xmlTempDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Err.Clear
    On Error Resume Next
    Set xmlReturnNode = objIOrganisationDO.FindCurrentUnitList(xmlTempDoc.documentElement)
    m_objErrAssist.SaveErr
    On Error GoTo FindCurrentUnitListVbErr
    If m_objErrAssist.LastErr = 0 Then
        strXMLData = xmlReturnNode.xml
    Else
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
            ' "You have no access assigned to you.  Please contact your System Supervisor."
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 115
        Else
            m_objErrAssist.ReRaise
        End If
    End If
    Set xmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlNode = xmlDataNode.appendChild(xmlData.documentElement)
    ' if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    FindCurrentUnitList = xmlOut.xml
    m_objContext.SetComplete
FindCurrentUnitListExit:
    
    Set xmlDataNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
              
    Exit Function
FindCurrentUnitListVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindCurrentUnitList = m_objErrAssist.CreateErrorResponse
        
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume FindCurrentUnitListExit
End Function
Private Sub ValidateAccessAttempts(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Validate the number of previously failed access attempts by the current user
'   against the parameter maximum.
' pass:
'   vstrXMLRequest
'       Format:
'           <REQUEST>
'               <USERID></USERID>
'               <AUDITRECORDTYPE></AUDITRECORDTYPE>
'           </REQUEST>
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateAccessAttemptsVbErr
    
    Const strFunctionName As String = "ValidateAccessAttempts"
               
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim objAuditBO As AuditBO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim dblMaxPermittedAttempts As Double
    Dim strAuditRecType As String
    Dim lngFailedAttempts As Long
    Dim strResponse As String
    Dim objIAuditDO As IAuditDO
    Dim objIOrganisationDO As IOrganisationDO
    Dim strSecurityType As String
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objAuditBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditBO")
    Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
            
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    strAuditRecType = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "AUDITRECORDTYPE")
    '------------------------------------------------------------------------------------------
    ' use AuditBO to get number of failed attempts
    '------------------------------------------------------------------------------------------
    strResponse = objAuditBO.GetNumberOfFailedAttempts(vstrXMLRequest)
    m_objErrAssist.CheckXMLResponse strResponse, True
    Set xmlTempDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    lngFailedAttempts = CLng(m_objXmlAssist.GetNodeValue(xmlTempDoc.documentElement, "//NUMBEROFATTEMPTS"))
    '------------------------------------------------------------------------------------------
    ' get max permitted failed attempts
    '------------------------------------------------------------------------------------------
         
    ' PSC 23/06/2006 CORE278 - Start
    On Error Resume Next
    strSecurityType = UCase$(Trim$(objGlobalParameterDO.GetCurrentParameterByType("SecurityCredentialsType", "STRING")))
    
    m_objErrAssist.SaveErr
    On Error GoTo ValidateAccessAttemptsVbErr
    If m_objErrAssist.LastErr <> 0 And m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound Then
        m_objErrAssist.ReRaise
    End If
    ' PSC 23/06/2006 CORE278 - End
    
    ' AS 07/06/2006 CORE278 - Do not check whether max login attempts have been exceeded
    ' if using windows authentication, as this will be handled by the O/S.
    If strSecurityType <> "WINDOWSAUTHENTICATION" Then
        If objIAuditDO.IsLogon(strAuditRecType) = True Then
            dblMaxPermittedAttempts = _
                objGlobalParameterDO.GetCurrentParameterByType("MaxLogonAttempts", "AMOUNT")
        Else
            If objIAuditDO.IsPasswordChange(strAuditRecType) = True Then
                dblMaxPermittedAttempts = _
                    objGlobalParameterDO.GetCurrentParameterByType( _
                        "MaxChangePasswordAttempts", "AMOUNT")
            Else
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, "Invalid AUDITRECORDTYPE parameter"
            End If
        End If
        '------------------------------------------------------------------------------------------
        ' do the comparison
        '------------------------------------------------------------------------------------------
        If lngFailedAttempts > dblMaxPermittedAttempts Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 117
        End If
    End If
    
    Set xmlIn = Nothing
    Set objIAuditDO = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlTempDoc = Nothing
    Set objIOrganisationDO = Nothing
    Set objAuditBO = Nothing
          
    Exit Sub
ValidateAccessAttemptsVbErr:
    Set xmlIn = Nothing
    Set objIAuditDO = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlTempDoc = Nothing
    Set objIOrganisationDO = Nothing
    Set objAuditBO = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function ValidateCurrentPassword(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
'   Format:
'       <REQUEST>
'           <PASSWORD>
'               <USERID></USERID>
'               <PASSWORDVALUE></PASSWORDVALUE>
'           </PASSWORD>
'       </REQUEST>
' return:
'   xml Response data stream containing results of operation
'   Format:
'       <RESPONSE>
'           <ACCESSAUDITATTEMPT></ACCESSAUDITATTEMPT>
'           <CHANGEPASSWORDINDICATOR></CHANGEPASSWORDINDICATOR>
'       </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo ValidateCurrentPasswordVbErr
    
    Const strFunctionName As String = "ValidateCurrentPassword"
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim objConvertAssist As New ConvertAssist
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlReturnNode As IXMLDOMNode
    Dim blnIsAttemptOK As Boolean       ' referred to as ACCESSAUDITATTEMPT in spec
    Dim blnChangePasswordIndicator
    Dim strInUserId As String
    Dim strInPassword As String
    Dim strDbPassword As String
    Dim strResponse As String
    Dim dblValidDays As Double
    Dim strCreationDate As String
    Dim dteCreationDate As Date
    Dim strSecurityCredentialsType As String
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
               
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlNode = xmlOut.appendChild(xmlResponseElem)
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        
    strInUserId = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "USERID")
    If Len(strInUserId) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "USERID must be specified"
    End If
                
    strInPassword = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "PASSWORDVALUE")
    If Len(strInPassword) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "PASSWORDVALUE must be specified"
    End If
    blnIsAttemptOK = True
    blnChangePasswordIndicator = False
    
    ' Convert input into dom and output into string for compatibility
    Set xmlTempDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Err.Clear
    
    '------------------------------------------------------------------------------------------
    ' determine the authentification mechanism (defaults to 'Omiga' if the global parameter
    ' has not been setup
    '------------------------------------------------------------------------------------------
    On Error Resume Next
    
    ' PSC 23/06/2006 CORE278 - Start
    strSecurityCredentialsType = UCase$(Trim$(objGlobalParameterDO.GetCurrentParameterByType("SecurityCredentialsType", "STRING")))
    m_objErrAssist.SaveErr
    On Error GoTo ValidateCurrentPasswordVbErr
    If m_objErrAssist.LastErr <> 0 And m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound Then
        m_objErrAssist.ReRaise
    End If
    
    If Len(strSecurityCredentialsType) = 0 Then
        strSecurityCredentialsType = "OMIGA"
    End If
    ' PSC 23/06/2006 CORE278 - End
    
    '------------------------------------------------------------------------------------------
    ' validate according the authentification mechanism
    '------------------------------------------------------------------------------------------
    If strSecurityCredentialsType = "WINDOWSAUTHENTICATION" Then
    
        '------------------------------------------------------------------------------------------
        ' NT Authentification
        '------------------------------------------------------------------------------------------
        Dim strDomain As String
        strDomain = GetComputerNameDnsHostname()
        If AuthenticateUser(strDomain, strInUserId, strInPassword) = False Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 112
        End If
        
    ' PSC 23/06/2006 CORE278
    ElseIf strSecurityCredentialsType = "OMIGA" Then
        
        '------------------------------------------------------------------------------------------
        ' Omiga Authentification
        '------------------------------------------------------------------------------------------
        ' ... check that the password is correct for the UserId input
        '------------------------------------------------------------------------------------------
    
        On Error Resume Next
        Set xmlReturnNode = objIOrganisationDO.GetCurrentPassword(xmlTempDoc.documentElement)
        'strResponse = objOrganisationDO.GetCurrentPassword(vstrXMLRequest)
        m_objErrAssist.SaveErr
        On Error GoTo ValidateCurrentPasswordVbErr
        If m_objErrAssist.LastErr = 0 Then
            strResponse = xmlReturnNode.xml
        Else
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
                blnIsAttemptOK = False
                '"You have no current password registered.  Please contact your System Supervisor."
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 233
            Else
                m_objErrAssist.ReRaise
            End If
        End If
        If blnIsAttemptOK = True Then
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            strDbPassword = m_objXmlAssist.GetTagValue( _
                xmlDoc.documentElement, "PASSWORDVALUE")
            ' do the comparison with the encrypted input value
            If strDbPassword <> Encrypt(strInPassword) Then
                blnIsAttemptOK = False
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 112
            End If
        End If
        
        '------------------------------------------------------------------------------------------
        ' ... check if the password needs changing
        '------------------------------------------------------------------------------------------
        If blnIsAttemptOK = True Then
            
            dblValidDays = objGlobalParameterDO.GetCurrentParameterByType( _
                "PasswordValidDays", "AMOUNT")
            strCreationDate = m_objXmlAssist.GetTagValue( _
                xmlDoc.documentElement, "PASSWORDCREATIONDATE")
            dteCreationDate = objConvertAssist.CSafeDate(strCreationDate)
            dteCreationDate = Format(dteCreationDate, "DD/MM/YYYY")                 'AS 30/05/06 CORE275 Password expiry only uses date and not time.
            If (dteCreationDate + dblValidDays) < Format(Now(), "DD/MM/YYYY") Then  'AS 30/05/06 CORE275 Password expiry only uses date and not time.
                blnChangePasswordIndicator = True
                blnIsAttemptOK = False
                ' n.b. leave blnIsAttemptOK set to True; throw a warning
                ' "Your Password has expired. You must set a new password to continue."
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 118
            End If
        End If
    Else
        ' unrecognized authentification type raise error
        ' ... Global parameter SecurityCredentialsType has an invalid value.....
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 8537
    End If
    
    '------------------------------------------------------------------------------------------
    ' set the response data
    '------------------------------------------------------------------------------------------
    Set xmlElem = xmlOut.createElement("CHANGEPASSWORDINDICATOR")
    If blnChangePasswordIndicator = True Then
        xmlElem.Text = "1"
    Else
        xmlElem.Text = "0"
    End If
    xmlResponseElem.appendChild xmlElem
    ' if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    ValidateCurrentPassword = xmlOut.xml
    m_objContext.SetComplete
ValidateCurrentPasswordExit:
    'clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseNode = Nothing
    Set xmlNode = Nothing
    Set objIOrganisationDO = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlResponseElem = Nothing
    Set objConvertAssist = Nothing
    Set xmlElem = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
     
    Exit Function
ValidateCurrentPasswordVbErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        ' RF 01/12/99 AQR SC014
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    ValidateCurrentPassword = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume ValidateCurrentPasswordExit
End Function
Public Function ChangePassword(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Processing is delegated to DoChangePassword and CreateAccessAudit in order to ensure
'   an AccessAudit record is created even upon error, and that in doing so the error info is
'   not lost.
' pass:
'   vstrXMLRequest
'       Format:
'       <REQUEST>
'           <OMIGAUSER>
'               <USERID></USERID>
'               <CURRENTPASSWORD></CURRENTPASSWORD>
'               <NEWPASSWORD></NEWPASSWORD>
'               <CONFIRMEDNEWPASSWORD></CONFIRMEDNEWPASSWORD>
'           </OMIGAUSER>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     05/11/99 Improved error handling
'------------------------------------------------------------------------------------------
On Error GoTo ChangePasswordVbErr
    
    Const strFunctionName As String = "ChangePassword"
        
       
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim objIAuditDO As IAuditDO
    Dim objIOrganisationDO As IOrganisationDO
    Dim strInUserId As String
    Dim strInCurrPassword As String
    Dim strInNewPassword As String
    Dim strInConfirmedNewPw As String
    Dim strResponse As String
    Dim blnIsAttemptOK As Boolean       ' referred to as ACCESSAUDITATTEMPT in spec
    Dim strChangePwAuditRecType As String
    Dim lngMsgNo As Long
    Dim strCreateDate As String
             
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    blnIsAttemptOK = True
    strChangePwAuditRecType = objIAuditDO.GetChangePasswordValueId()
    '---------------------------------------------------------------------------
    ' Delegate the main processing to DoChangePassword
    '---------------------------------------------------------------------------
    strResponse = DoChangePassword(vstrXMLRequest, strChangePwAuditRecType)
    lngMsgNo = m_objErrAssist.CheckXMLResponse(strResponse, False)
    If lngMsgNo <> 0 Then
        If m_objErrAssist.IsWarning(lngMsgNo) = False Then
            blnIsAttemptOK = False
        End If
    End If
    '---------------------------------------------------------------------------
    ' create the access audit record even on non-warning error
    '---------------------------------------------------------------------------
    'SYS0210 MDC 15/05/2000. Synchronise date/time of newly created password record
    '                        with corresponding AccessAudit record
    If lngMsgNo = 0 Then
        strCreateDate = GetDatePasswordChanged(vstrXMLRequest)
    End If
    CreateAccessAudit vstrXMLRequest, strChangePwAuditRecType, blnIsAttemptOK, strCreateDate
    '---------------------------------------------------------------------------
    ' ensure change password error is raised
    '---------------------------------------------------------------------------
    If lngMsgNo <> 0 Then
        m_objErrAssist.CheckXMLResponse strResponse, True
    End If
        
    '---------------------------------------------------------------------------
    ' set up return data
    '---------------------------------------------------------------------------
    Set xmlElem = xmlOut.createElement("ACCESSAUDITATTEMPT")
    If blnIsAttemptOK = True Then
        xmlElem.Text = "1"
    Else
        xmlElem.Text = "0"
    End If
    xmlResponseElem.appendChild xmlElem
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    Set xmlNode = xmlDoc.getElementsByTagName("RESPONSE").Item(0)
    If xmlNode Is Nothing Then
       m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInternalError, _
           "RESPONSE tag missing"
    End If
    For Each xmlTempNode In xmlNode.childNodes
        xmlResponseElem.appendChild xmlTempNode.cloneNode(True)
    Next xmlTempNode
    ' if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    ChangePassword = xmlOut.xml
    m_objContext.SetComplete
ChangePasswordExit:
    ' clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestNode = Nothing
    Set xmlNode = Nothing
    Set objIAuditDO = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseNode = Nothing
    Set xmlElem = Nothing
    Set xmlTempNode = Nothing
    Set objIAuditDO = Nothing
    Set objIOrganisationDO = Nothing
    Exit Function
ChangePasswordVbErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    ChangePassword = m_objErrAssist.CreateErrorResponse
    m_objContext.SetAbort
        
    ' go to clean-up section
    Resume ChangePasswordExit
End Function
Public Sub ValidatePasswordFormat(ByVal vstrPassword As String)
' header ----------------------------------------------------------------------------------
' description:
'   Validate that a password is of the correct format (correct length, etc)
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo ValidatePasswordFormatVbErr
    
    Const strFunctionName As String = "ValidatePasswordFormat"
            
    Dim objGeneralAssist As New GeneralAssist
    Dim objGlobalParamDO As GlobalParameterDO
    Dim dblMinLen As Double
    Dim dblMaxLen As Double
    Dim blnAlphaAndNumeric As Boolean
    Dim blnDuplicatedChars As Boolean
    Dim blnUpperAlpha As Boolean
    Dim blnLowerAlpha As Boolean
    Dim blnNumeric As Boolean
    Dim blnSpecialChars As Boolean
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    '------------------------------------------------------------------------------------------
    ' check password length
    '------------------------------------------------------------------------------------------
      
    dblMinLen = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordMinimumLength", "AMOUNT")
    dblMaxLen = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordMaximumLength", "AMOUNT")
    'RF 18/11/99 Change to error number used
    If Len(vstrPassword) < dblMinLen Or _
        Len(vstrPassword) > dblMaxLen Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 144, _
            , CStr(dblMinLen), CStr(dblMaxLen)
    End If
        
    '------------------------------------------------------------------------------------------
    ' check alpha and numeric chars
    '------------------------------------------------------------------------------------------
            
    blnAlphaAndNumeric = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordAlphaNumeric", "BOOLEAN")
    If blnAlphaAndNumeric = True Then
        If objGeneralAssist.IsAlpha(vstrPassword) Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not be all alpha characters"
        End If
        If objGeneralAssist.IsDigits(vstrPassword) = True Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not be all digits"
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' check for duplicated chars
    '------------------------------------------------------------------------------------------
       
    blnDuplicatedChars = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordCharacterDuplication", "BOOLEAN")
    If blnDuplicatedChars = True Then
        If objGeneralAssist.HasDuplicatedChars(vstrPassword, False) = True Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not contain duplicated characters"
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' check the char set
    '------------------------------------------------------------------------------------------
           
    blnUpperAlpha = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordUpperAlpha", "BOOLEAN")
    If blnUpperAlpha = False Then
        If objGeneralAssist.ContainsUpperAlpha(vstrPassword) = True Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not contain upper case alpha characters"
        End If
    End If
    blnLowerAlpha = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordLowerAlpha", "BOOLEAN")
    If blnLowerAlpha = False Then
        If objGeneralAssist.ContainsLowerAlpha(vstrPassword) = True Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not contain lower case alpha characters"
        End If
    End If
    blnNumeric = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordNumeric", "BOOLEAN")
    If blnNumeric = False Then
        If objGeneralAssist.ContainsDigits(vstrPassword) = True Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not contain digits"
        End If
    End If
    blnSpecialChars = objGlobalParamDO.GetCurrentParameterByType( _
        "PasswordSpecialCharacters", "BOOLEAN")
    If blnSpecialChars = False Then
        If objGeneralAssist.ContainsSpecialChars(vstrPassword) = True Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
                "Password may not contain special characters"
        End If
    End If
    Set objGlobalParamDO = Nothing
    Set objGeneralAssist = Nothing
    Set objIOrganisationDO = Nothing
      
    Exit Sub
ValidatePasswordFormatVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set objGlobalParamDO = Nothing
    Set objGeneralAssist = Nothing
    Set objIOrganisationDO = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Sub ValidateWorkingHours(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Validate that the user is accessing the system during their assigned working hours.
' pass:
'   Format:
'       <REQUEST>
'           <OMIGAUSER>
'               <USERID></USERID>
'               <CHANNELID></CHANNELID>
'           </OMIGAUSER>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo ValidateWorkingHoursVbErr
    
    Const strFunctionName As String = "ValidateWorkingHours"
           
    Dim objIOrganisationDO As IOrganisationDO
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objConvertAssist As ConvertAssist
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlReturnNode As IXMLDOMNode
    Dim strResponse As String
    Dim intFromHr As Integer
    Dim intToHr As Integer
    Dim intFromMin As Integer
    Dim intToMin As Integer
    Dim intHrNow As Integer
    Dim intMinNow As Integer
    Dim dteNow As Date
    Dim strTagName As String
    Dim strTagValue As String
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set objConvertAssist = New ConvertAssist
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
           
    ' Convert input into dom and output into string for compatibility
    Set xmlTempDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Err.Clear
    On Error Resume Next
    Set xmlReturnNode = objIOrganisationDO.GetWorkingHoursForToday(xmlTempDoc.documentElement)
    'strResponse = objOrganisationDO.GetWorkingHoursForToday(vstrXMLRequest)
        
    m_objErrAssist.SaveErr
    On Error GoTo ValidateWorkingHoursVbErr
    If m_objErrAssist.LastErr = 0 Then
        strResponse = xmlReturnNode.xml
    Else
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
            ' "You are not authorised access outside of your assigned working hours. Please contact your System Supervisor."
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 116
        Else
            m_objErrAssist.ReRaise
        End If
    End If
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    strTagName = "FROMHOUR"
    strTagValue = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, strTagName)
    If Len(strTagValue) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInternalError, _
            "Invalid " & strTagName & " value on table WORKINGHOURS"
    End If
    intFromHr = objConvertAssist.CSafeInt(strTagValue)
    strTagName = "TOHOUR"
    strTagValue = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, strTagName)
    If Len(strTagValue) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInternalError, _
            "Invalid " & strTagName & " value on table WORKINGHOURS"
    End If
    intToHr = objConvertAssist.CSafeInt(strTagValue)
    strTagName = "FROMMINUTE"
    strTagValue = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, strTagName)
    If Len(strTagValue) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInternalError, _
            "Invalid " & strTagName & " value on table WORKINGHOURS"
    End If
    intFromMin = objConvertAssist.CSafeInt(strTagValue)
    strTagName = "TOMINUTE"
    strTagValue = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, strTagName)
    If Len(strTagValue) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInternalError, _
            "Invalid " & strTagName & " value on table WORKINGHOURS"
    End If
    intToMin = objConvertAssist.CSafeInt(strTagValue)
    dteNow = Now()
    intHrNow = Hour(dteNow)
    intMinNow = Minute(dteNow)
    If intFromHr > intHrNow Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 116
    Else
        If intFromHr = intHrNow And intFromMin > intMinNow Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 116
        End If
    End If
        
    If intToHr < intHrNow Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 116
    Else
        If intToHr = intHrNow And intToMin < intMinNow Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 116
        End If
    End If
    Set xmlDoc = Nothing
    Set objConvertAssist = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
    Set objIOrganisationDO = Nothing
      
    Exit Sub
ValidateWorkingHoursVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set objIOrganisationDO = Nothing
    Set xmlDoc = Nothing
    Set objConvertAssist = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
            
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Private Sub ValidatePasswordNoUserId(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Validate that the password does not contain user details.
' pass:
'   vstrXMLRequest
'       Format as per ChangePassword, i.e.:
'       <REQUEST>
'           <USERID></USERID>
'           <CURRENTPASSWORD></CURRENTPASSWORD>
'           <PASSWORDVALUE></PASSWORDVALUE>
'           <CONFIRMEDNEWPASSWORD></CONFIRMEDNEWPASSWORD>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo ValidatePasswordNoUserIdVbErr
    
    Const strFunctionName As String = "ValidatePasswordNoUserId"
        
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim objIOrganisationDO As IOrganisationDO
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlReturnNode As IXMLDOMNode
    Dim strNewPassword As String
    Dim strResponse As String
    Dim strUserId As String
    Dim strSurname As String
    Dim strForename As String
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
              
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
       
    strNewPassword = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "PASSWORDVALUE")
    ' Convert input into dom and output into string for compatibility
    Set xmlTempDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlReturnNode = objIOrganisationDO.GetOmigaUserDetails(xmlTempDoc.documentElement)
    strResponse = xmlReturnNode.xml
    'strResponse = objOrganisationDO.GetOmigaUserDetails(vstrXMLRequest)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
    strUserId = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "USERID")
    strSurname = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "SURNAME")
    strForename = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "FORENAME")
    If (InStr(1, strNewPassword, strUserId, vbTextCompare)) <> 0 Then
        ' "New password is invalid. Please retype your new password ..." (etc)
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
            "New password must not contain user id"
    End If
    If (InStr(1, strNewPassword, strSurname, vbTextCompare)) <> 0 Then
        ' "New password is invalid. Please retype your new password ..." (etc)
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
            "New password must not contain surname"
    End If
    If (InStr(1, strNewPassword, strForename, vbTextCompare)) <> 0 Then
        ' "New password is invalid. Please retype your new password ..." (etc)
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 113, _
            "New password must not contain forename"
    End If
    Set xmlDoc = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlIn = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
    Set objIOrganisationDO = Nothing
     
    Exit Sub
ValidatePasswordNoUserIdVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set xmlDoc = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlIn = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
    Set objIOrganisationDO = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Private Sub CreateAccessAudit(ByVal vstrXMLRequest As String, _
    ByVal vstrLogonAuditRecType As String, ByVal vblnSuccessIndicator As Boolean, Optional ByVal vstrCreationDate As String)
' header ----------------------------------------------------------------------------------
' description:
'   Create AccessAudit record. Used by ValidateUserLogon
' pass:
'   vstrXMLRequest
'       Format as per ValidateUserLogon
'   vstrLogonAuditRecType
'   vblnSuccessIndicator
'   vstrCreationDate (optional)
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccessAuditVbErr
    
    Const strFunctionName As String = "CreateAccessAudit"
          
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objAuditBO As AuditBO
    Dim strResponse As String
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objAuditBO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditBO")
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlDoc = Nothing
    Set xmlRequestNode = m_objXmlAssist.GetRequestNode(xmlIn)
    xmlDoc.appendChild xmlRequestNode
    Set xmlElem = xmlDoc.createElement("ACCESSAUDIT")
    Set xmlNode = xmlRequestNode.appendChild(xmlElem)
    Set xmlElem = xmlDoc.createElement("SUCCESSINDICATOR")
    If vblnSuccessIndicator = True Then
        xmlElem.Text = "1"
    Else
        xmlElem.Text = "0"
    End If
    xmlNode.appendChild xmlElem
    'Set xmlElem = xmlDoc.createElement("UNITID")
    'xmlElem.Text = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "UNITID")
    'xmlNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("AUDITRECORDTYPE")
    xmlElem.Text = vstrLogonAuditRecType
    xmlNode.appendChild xmlElem
    'SYS0210 MDC 15/05/2000.
    If Len(Trim$(vstrCreationDate)) > 0 Then
        Set xmlElem = xmlDoc.createElement("PASSWORDCREATIONDATE")
        xmlElem.Text = vstrCreationDate
        xmlNode.appendChild xmlElem
    End If
    strResponse = objAuditBO.CreateAccessAudit(xmlDoc.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    'CreateAccessAudit = strResponse
    m_objContext.SetComplete
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set objAuditBO = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlElem = Nothing
    Set objIOrganisationDO = Nothing
    Exit Sub
CreateAccessAuditVbErr:
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set objAuditBO = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlRequestNode = Nothing
    Set xmlElem = Nothing
    Set objIOrganisationDO = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objContext.SetAbort
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function DoValidateUserLogon(ByVal vstrXMLRequest As String, _
    ByVal vstrLogonAuditRecType As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Validate the UserId, UserPassword and UnitId entered by the current user.
' pass:
'   vstrXMLRequest
'       Format:
'        <REQUEST USERID= MACHINEID=>
'           <OMIGAUSER>
'               <USERID></USERID>
'               <PASSWORDVALUE></PASSWORDVALUE>
'               <UNIT></UNIT>
'           </OMIGAUSER>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo DoValidateUserLogonVbErr
    
    Const strFunctionName As String = "DoValidateUserLogon"
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objIOrganisationDO As IOrganisationDO
    Dim objIAuditDO As IAuditDO
    Dim objConvertAssist As ConvertAssist
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlReturnNode As IXMLDOMNode
    Dim blnIsAttemptOK As Boolean       ' referred to as ACCESSAUDITATTEMPT in spec
    Dim strInUserId As String
    Dim strInUserPassword As String
    Dim strInUnitId As String
    Dim strResponse
    Dim strOmigaUserDetails As String
    Dim strUnitDetails As String
    Dim intUserRole As Integer
    Dim blnUnitFound As Boolean
    Dim bnlChangePassword As Boolean
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
        
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    blnIsAttemptOK = True
    '------------------------------------------------------------------------------------------
    ' validate inputs
    '------------------------------------------------------------------------------------------
    strInUserId = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "USERID")
    strInUserPassword = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "PASSWORDVALUE")
    strInUnitId = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "UNITID")
    If Len(strInUserId) = 0 Or _
        Len(strInUserPassword) = 0 Or _
        Len(strInUnitId) = 0 Then
        blnIsAttemptOK = False
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 106
    End If
    ' check that the userid in the REQUEST block is the same as the userid in the body
    ' of the request xml
    Dim strRequestBlockUserId As String
    strRequestBlockUserId = m_objXmlAssist.GetAttributeValue( _
        xmlIn.documentElement, "REQUEST", "USERID")
    If strRequestBlockUserId <> strInUserId Then
        blnIsAttemptOK = False
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "Non-matching USERID values"
    End If
    '------------------------------------------------------------------------------------------
    ' validate that the User, UserRole and Unit are active
    '------------------------------------------------------------------------------------------
       
    If blnIsAttemptOK = True Then
        
        ' Convert input into dom and output into string for compatibility
        Set xmlTempDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        Err.Clear
        On Error Resume Next
        Set xmlReturnNode = objIOrganisationDO.GetOmigaUserDetails(xmlTempDoc.documentElement)
        'strOmigaUserDetails = objOrganisationDO.GetOmigaUserDetails(vstrXMLRequest)
        m_objErrAssist.SaveErr
        On Error GoTo DoValidateUserLogonVbErr
        If m_objErrAssist.LastErr = 0 Then
            strOmigaUserDetails = xmlReturnNode.xml
        Else
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
                blnIsAttemptOK = False
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 111
                Err.Clear
            Else
                m_objErrAssist.ReRaise
            End If
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Validate the user's role record for the selected unit id.
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        
        ' Convert input into dom and output into string for compatibility
        Set xmlTempDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        Err.Clear
        On Error Resume Next
        Set xmlReturnNode = objIOrganisationDO.FindCurrentUnitList(xmlTempDoc.documentElement)
        'strResponse = objOrganisationDO.FindCurrentUnitList(vstrXMLRequest)
        m_objErrAssist.SaveErr
        On Error GoTo DoValidateUserLogonVbErr
        If m_objErrAssist.LastErr = 0 Then
            strResponse = xmlReturnNode.xml
        Else
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
                blnIsAttemptOK = False
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 111
                Err.Clear
            Else
                m_objErrAssist.ReRaise
            End If
        End If
    End If
    If blnIsAttemptOK = True Then
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        Set xmlNodeList = xmlDoc.getElementsByTagName("USERROLE")
        blnUnitFound = False
        While blnUnitFound = False And intUserRole < xmlNodeList.Length
            Set xmlNode = xmlNodeList.Item(intUserRole)
            If strInUnitId = m_objXmlAssist.GetTagValue(xmlNode, "UNITID") Then
                blnUnitFound = True
                strUnitDetails = xmlNode.xml
            End If
            intUserRole = intUserRole + 1
        Wend
        If blnUnitFound = False Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 111
        End If
    End If
       
    '------------------------------------------------------------------------------------------
    ' Validate the working hours
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        
        '------------------------------------------------------------------------------------------
        ' set up the CHANNELID element
        '------------------------------------------------------------------------------------------
        Dim strDISTRIBUTIONCHANNELID As String
        Set xmlNode = xmlDoc.getElementsByTagName("DEPARTMENT").Item(0)
        If xmlNode Is Nothing Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInternalError, "Expected DEPARTMENT tag"
        End If
        strDISTRIBUTIONCHANNELID = m_objXmlAssist.GetTagValue(xmlNode, "CHANNELID")
        Set xmlNode = xmlIn.getElementsByTagName("OMIGAUSER").Item(0)
        If xmlNode Is Nothing Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInternalError, "Expected OMIGAUSER tag"
        End If
        Set xmlElem = xmlIn.createElement("CHANNELID")
        xmlElem.Text = strDISTRIBUTIONCHANNELID
        xmlNode.appendChild xmlElem
        '------------------------------------------------------------------------------------------
        ' call ValidateWorkingHours
        '------------------------------------------------------------------------------------------
        Err.Clear
        On Error Resume Next
        ValidateWorkingHours xmlIn.xml
        m_objErrAssist.SaveErr
        On Error GoTo DoValidateUserLogonVbErr
        If m_objErrAssist.LastErr <> 0 Then
            blnIsAttemptOK = False
            m_objErrAssist.ReRaise
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Check the number of access attempts
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        Set xmlNode = xmlIn.getElementsByTagName("OMIGAUSER").Item(0)
        If xmlNode Is Nothing Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInternalError, "Expected OMIGAUSER tag"
        End If
        Set xmlElem = xmlIn.createElement("AUDITRECORDTYPE")
        xmlElem.Text = vstrLogonAuditRecType
        xmlNode.appendChild xmlElem
        Err.Clear
        On Error Resume Next
        ValidateAccessAttempts xmlIn.xml
        m_objErrAssist.SaveErr
        On Error GoTo DoValidateUserLogonVbErr
        If m_objErrAssist.LastErr <> 0 Then
            blnIsAttemptOK = False
            m_objErrAssist.ReRaise
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Check the userpassword is correct for the userid input and that the
    ' password has not expired
    '------------------------------------------------------------------------------------------
        
    Set objConvertAssist = New ConvertAssist
    If blnIsAttemptOK = True Then
        m_objXmlAssist.ChangeNodeName xmlIn.documentElement, "OMIGAUSER", "PASSWORD"
        strResponse = ValidateCurrentPassword(xmlIn.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        bnlChangePassword = objConvertAssist.CSafeBool(m_objXmlAssist.GetTagValue( _
            xmlDoc.documentElement, "CHANGEPASSWORDINDICATOR"))
    End If
    '---------------------------------------------------------------------------
    ' perform hard-coded validation & call scripting interface
    '---------------------------------------------------------------------------
    Validate vstrXMLRequest
    '---------------------------------------------------------------------------
    ' set up return data
    '---------------------------------------------------------------------------
    ' RF 09/11/99 ACCESSAUDITATTEMPT tag is now added in ValidateUserLogon
'    Set xmlElem = xmlOut.createElement("ACCESSAUDITATTEMPT")
'    If blnIsAttemptOK = True Then
'        xmlElem.Text = "1"
'    Else
'        xmlElem.Text = "0"
'    End If
'    xmlResponseElem.appendChild xmlElem
    If blnIsAttemptOK = True Then
        
        '---------------------------------------------------------------------------
        ' set up remaining return data
        '---------------------------------------------------------------------------
        ' load OmigaUserDetails
        Set xmlDoc = m_objXmlAssist.load(strOmigaUserDetails, _
            TypeName(Me), strFunctionName)
        ' RF 03/11/99 return UserName
        Dim strUserName As String
        strUserName = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "FORENAME")
        strUserName = strUserName & " "
        strUserName = strUserName & _
            m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "SURNAME")
        ' RF 09/11/99 CHANGEPASSWORDINDICATOR value depends both on return value
        ' from ValidateCurrentPassword and value in OmigaUser table
        If bnlChangePassword = False Then
            ' check value in OmigaUser
            ' SR - 09/03/00 - Use CSafeBool instead of CBool
            bnlChangePassword = objConvertAssist.CSafeBool(m_objXmlAssist.GetTagValue( _
                xmlDoc.documentElement, "CHANGEPASSWORDINDICATOR"))
        End If
        Set xmlElem = xmlOut.createElement("CHANGEPASSWORDINDICATOR")
        If bnlChangePassword = True Then
            xmlElem.Text = "1"
        Else
            xmlElem.Text = "0"
        End If
        xmlResponseElem.appendChild xmlElem
        ' add OmigaUserDetails
        xmlResponseElem.appendChild xmlDoc.documentElement
        ' add UserName
        Set xmlElem = xmlOut.createElement("USERNAME")
        xmlElem.Text = strUserName
        xmlResponseElem.appendChild xmlElem
        ' UnitDetails
        Set xmlDoc = m_objXmlAssist.load(strUnitDetails, _
            TypeName(Me), strFunctionName)
        xmlResponseElem.appendChild xmlDoc.documentElement
    End If
    '   if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    DoValidateUserLogon = xmlOut.xml
    m_objContext.SetComplete
DoValidateUserLogonExit:
    ' clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseNode = Nothing
    Set xmlNode = Nothing
    Set objIOrganisationDO = Nothing
    Set objIAuditDO = Nothing
    Set objConvertAssist = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElem = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlNodeList = Nothing
    Set objIAuditDO = Nothing
    Exit Function
DoValidateUserLogonVbErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    DoValidateUserLogon = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume DoValidateUserLogonExit
End Function
Private Function DoChangePassword(ByVal vstrXMLRequest As String, _
    ByVal vstrChangePwAuditRecType As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
'       Format:
'       <REQUEST>
'           <OMIGAUSER>
'               <USERID></USERID>
'               <CURRENTPASSWORD></CURRENTPASSWORD>
'               <NEWPASSWORD></NEWPASSWORD>
'               <CONFIRMEDNEWPASSWORD></CONFIRMEDNEWPASSWORD>
'           </OMIGAUSER>
'       </REQUEST>
' return:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     05/11/99 Improved error handling
'------------------------------------------------------------------------------------------
On Error GoTo DoChangePasswordVbErr
    
    Const strFunctionName As String = "DoChangePassword"
      
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim objIOrganisationDO As IOrganisationDO
    Dim objIOrganisationTxBO As IOrganisationTxBO
    Dim objComboDO As ComboDO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim blnIsAttemptOK As Boolean       ' referred to as ACCESSAUDITATTEMPT in spec
    Dim strInUserId As String
    Dim strInCurrPassword As String
    Dim strInNewPassword As String
    Dim strInConfirmedNewPw As String
    Dim strResponse As String
    Dim blnPasswordNoUserId As Boolean
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objNode As IXMLDOMNode
    Dim objRequestNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim xmlResponse As IXMLDOMNode
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set objIOrganisationTxBO = m_objContext.CreateInstance(App.Title & ".OrganisationTxBO")
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
                   
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
       
    '------------------------------------------------------------------------------------------
    ' validate inputs
    '------------------------------------------------------------------------------------------
    blnIsAttemptOK = True
    strInUserId = m_objXmlAssist.GetTagValue(xmlIn.documentElement, "USERID")
    strInCurrPassword = m_objXmlAssist.GetTagValue( _
        xmlIn.documentElement, "CURRENTPASSWORD")
    strInNewPassword = m_objXmlAssist.GetTagValue( _
        xmlIn.documentElement, "NEWPASSWORD")
    strInConfirmedNewPw = m_objXmlAssist.GetTagValue( _
        xmlIn.documentElement, "CONFIRMEDNEWPASSWORD")
    If Len(strInUserId) = 0 Or _
        Len(strInCurrPassword) = 0 Or _
        Len(strInNewPassword) = 0 Or _
        Len(strInConfirmedNewPw) = 0 Then
        blnIsAttemptOK = False
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter
    End If
    If blnIsAttemptOK = True Then
        If strInNewPassword <> strInConfirmedNewPw Then
            blnIsAttemptOK = False
            ' "New password and confirmed password must be the same.  Please retype
            ' the new and confirmed passwords and retry."
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 234
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' validate that the user has not exceeded max attempts
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        Set xmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        Set xmlNode = xmlDoc.getElementsByTagName("REQUEST").Item(0)
        If xmlNode Is Nothing Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
        End If
    End If
        
    If blnIsAttemptOK = True Then
        Set xmlElem = xmlDoc.createElement("AUDITRECORDTYPE")
        xmlElem.Text = vstrChangePwAuditRecType
        xmlNode.appendChild xmlElem
        ValidateAccessAttempts xmlDoc.xml
    End If
    '------------------------------------------------------------------------------------------
    ' validate that the currrent userpassword is correct
    '------------------------------------------------------------------------------------------
        
    If blnIsAttemptOK = True Then
        Set xmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, _
            "CURRENTPASSWORD", "PASSWORDVALUE"
        strResponse = ValidateCurrentPassword(xmlDoc.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True
    End If
    '------------------------------------------------------------------------------------------
    ' validate that the new password is of the correct format
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        ValidatePasswordFormat strInNewPassword
    End If
    '------------------------------------------------------------------------------------------
    ' if required, validate that the password does not contain user details
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
            blnPasswordNoUserId = objGlobalParamDO.GetCurrentParameterByType( _
            "PasswordNoUserId", "BOOLEAN")
        If blnPasswordNoUserId = True Then
            Set xmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, _
                "NEWPASSWORD", "PASSWORDVALUE"
            ValidatePasswordNoUserId xmlDoc.xml
        End If
    End If
    '---------------------------------------------------------------------------
    ' perform hard-coded validation & call scripting interface
    '---------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        Validate vstrXMLRequest
    End If
    '------------------------------------------------------------------------------------------
    ' save the new password
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        Set xmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, _
            "OMIGAUSER", "PASSWORD"
        m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, _
            "NEWPASSWORD", "PASSWORDVALUE"
        Set xmlResponse = objIOrganisationTxBO.CreatePassword(xmlDoc.documentElement)
        m_objErrAssist.CheckResponse xmlResponse, xmlResponseElem, True
    End If
    '---------------------------------------------------------------------------
    ' SYS0166: update the change password indicator
    '---------------------------------------------------------------------------
     
    Set objElement = objXmlDoc.createElement("REQUEST")
    Set objRequestNode = objXmlDoc.appendChild(objElement)
    Set objElement = objXmlDoc.createElement("OMIGAUSER")
    Set objNode = objRequestNode.appendChild(objElement)
    Set objElement = objXmlDoc.createElement("USERID")
    objElement.Text = strInUserId
    objNode.appendChild objElement
    Set objElement = objXmlDoc.createElement("CHANGEPASSWORDINDICATOR")
    objElement.Text = "0"
    objNode.appendChild objElement
    Set xmlResponse = objIOrganisationTxBO.UpdateChangePasswordIndicator(objXmlDoc.documentElement)
    m_objErrAssist.CheckResponse xmlResponse, xmlResponseElem, True
    '---------------------------------------------------------------------------
    ' set up return data
    '---------------------------------------------------------------------------
    ' RF 09/11/99 ACCESSAUDITATTEMPT tag is now added in ChangePassword
'    Set xmlElem = xmlOut.createElement("ACCESSAUDITATTEMPT")
'    If blnIsAttemptOK = True Then
'        xmlElem.Text = "1"
'    Else
'        xmlElem.Text = "0"
'    End If
'    xmlResponseElem.appendChild xmlElem
    
    ' if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    DoChangePassword = xmlOut.xml
    m_objContext.SetComplete
DoChangePasswordExit:
    ' clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set objIOrganisationDO = Nothing
    Set objComboDO = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElem = Nothing
    Set objIOrganisationTxBO = Nothing
    Set objXmlDoc = Nothing
    Set objNode = Nothing
    Set objRequestNode = Nothing
    Set objElement = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlResponse = Nothing
    Exit Function
DoChangePasswordVbErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    DoChangePassword = m_objErrAssist.CreateErrorResponse
    m_objContext.SetAbort
        
    ' go to clean-up section
    Resume DoChangePasswordExit
End Function
Public Function FindUnitList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindUnitListVbErr
    Const strFunctionName As String = "FindUnitList"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_FindUnitList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindUnitList = xmlResponseElem.xml
    m_objContext.SetComplete
FindUnitListExit:
    'clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIOrganisationDO = Nothing
       
    Exit Function
FindUnitListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindUnitList = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume FindUnitListExit
End Function

Public Function FindUserList(ByVal vstrXMLRequest As String) As String

On Error GoTo FindUserListVbErr
    
    Const cstrFunctionName As String = "FindUserList"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xsltDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    
    Dim objIADOAssist As IADOAssist
    
    Dim strUnitID As String, _
        strSqlXml As String, _
        strXsltFileSpec As String
    
    m_objContext.SetComplete
        
    strXsltFileSpec = App.Path & "\FindUserList.xslt"
    strXsltFileSpec = Replace(strXsltFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
        
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.async = True
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.async = True
    xmlIn.loadXML vstrXMLRequest
    SimpleXmlParserCheck xmlIn, TypeName(Me)
            
    strUnitID = xmlIn.selectSingleNode("REQUEST/USERLIST/UNITID").Text
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    
    Set objIADOAssist = New omBase.ADOAssist

    With adoConnection
        .ConnectionString = objIADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set objIADOAssist = Nothing

    With adoCommand
        .CommandText = "usp_FindUserList"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("UNITID", adBSTR, adParamInput, Len(strUnitID), strUnitID)
    End With
    
    Set adoStream = New ADODB.Stream
    adoStream.Open
        
    With adoCommand
        .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
    
    strSqlXml = adoStream.ReadText
    
    adoConnection.Close
    Set adoConnection = Nothing
    
    xmlIn.loadXML strSqlXml
    SimpleXmlParserCheck xmlIn, TypeName(Me)
    
    If xmlIn.selectNodes("RESPONSE/OMIGAUSER").Length = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), cstrFunctionName, omiga4RecordNotFound
    End If
    
    Set xsltDoc = New FreeThreadedDOMDocument40
    xsltDoc.async = True
    xsltDoc.load strXsltFileSpec
    SimpleXmlParserCheck xsltDoc, TypeName(Me)
    
    FindUserList = xmlIn.transformNode(xsltDoc)

FindUserListExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xsltDoc = Nothing
       
    Exit Function

FindUserListVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    FindUserList = m_objErrAssist.CreateErrorResponse
    
    ' go to clean-up section
    Resume FindUserListExit

End Function

Private Sub SimpleXmlParserCheck( _
    ByVal vxmlDoc As FreeThreadedDOMDocument40, _
    ByVal vstrFunctionName As String)
    
    If vxmlDoc.parseError.errorCode <> 0 Then
        m_objErrAssist.ThrowError _
            TypeName(Me), _
            vstrFunctionName, _
            oeXMLParserError, _
            m_objErrAssist.FormatParserError(vxmlDoc.parseError)
    End If

End Sub

Public Function FindDepartmentList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindDepartmentListVbErr
    Const strFunctionName As String = "FindDepartmentList"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
       
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_FindDepartmentList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindDepartmentList = xmlResponseElem.xml
    m_objContext.SetComplete
FindDepartmentListExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIOrganisationDO = Nothing
       
    Exit Function
FindDepartmentListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindDepartmentList = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindDepartmentListExit
End Function
Public Function FindDistributionChannelList() As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindDistributionChannelListVbErr
    Const strFunctionName As String = "FindDistributionChannelList"
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_FindDistributionChannelList
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindDistributionChannelList = xmlResponseElem.xml
    m_objContext.SetComplete
FindDistributionChannelListExit:
    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIOrganisationDO = Nothing
          
    Exit Function
FindDistributionChannelListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindDistributionChannelList = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindDistributionChannelListExit
End Function
Private Function IOrganisationBO_FindDepartmentList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindDepartmentListVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindDepartmentList"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
      
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
             
    If vxmlRequest.nodeName = "DEPARTMENT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("DEPARTMENT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, " DEPARTMENT tag not found"
    End If
        
    Set xmlDataNode = objIOrganisationDO.FindDepartmentList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_FindDepartmentList = xmlResponseElem
FindDepartmentListExit:
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIOrganisationDO = Nothing
       
    Exit Function
FindDepartmentListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindDepartmentList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindDepartmentListExit
End Function
Private Function IOrganisationBO_FindDistributionChannelList() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindDistributionChannelListVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindDistributionChannelList"
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlDataNode = objIOrganisationDO.FindDistributionChannelList
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_FindDistributionChannelList = xmlResponseElem
FindDistributionChannelListExit:
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIOrganisationDO = Nothing
          
    Exit Function
FindDistributionChannelListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindDistributionChannelList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindDistributionChannelListExit
End Function
Private Function IOrganisationBO_FindUnitList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitListVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindUnitList"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If vxmlRequest.nodeName = "UNIT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("UNIT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, " UNIT tag not found"
    End If
        
    Set xmlDataNode = objIOrganisationDO.FindUnitList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_FindUnitList = xmlResponseElem
FindUnitListExit:
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIOrganisationDO = Nothing
          
    Exit Function
FindUnitListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindUnitList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUnitListExit
End Function
Private Function IOrganisationBO_FindUnitSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitSummaryVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindUnitSummary"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
                
    If vxmlRequest.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("REQUEST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "UNITID" & " tag not found"
    End If
        
    Set xmlDataNode = objIOrganisationDO.FindUnitSummary(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IOrganisationBO_FindUnitSummary = xmlResponseElem
    m_objContext.SetComplete
FindUnitSummaryExit:
    
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
FindUnitSummaryVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindUnitSummary = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUnitSummaryExit
End Function
Private Function IOrganisationBO_FindUserList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindUserListVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindUserList"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    If vxmlRequest.nodeName = "OMIGAUSER" Then
        Set xmlRequestNode = vxmlRequest
    Else
        ' DJP Change to match DTD
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("USERLIST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        ' DJP Change to match DTD
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, " USERLIST tag not found"
    End If
        
    Set xmlDataNode = objIOrganisationDO.FindUserList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_FindUserList = xmlResponseElem
FindUserListExit:
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIOrganisationDO = Nothing
        
    Exit Function
FindUserListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindUserList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUserListExit
End Function

Private Function IOrganisationBO_GetUserCompetency(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetUserCompetencyVbErr
    
    Const strFunctionName As String = "IOrganisationBO_GetUserCompetency"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If vxmlRequest.nodeName = "OMIGAUSER" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("OMIGAUSER").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, " OMIGAUSER tag not found"
    End If
        
    Set xmlDataNode = objIOrganisationDO.ValidateUserMandateLevel(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_GetUserCompetency = xmlResponseElem
GetUserCompetencyExit:
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
          
    Exit Function
GetUserCompetencyVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_GetUserCompetency = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetUserCompetencyExit
End Function

Private Function IOrganisationBO_ValidateUserMandateLevel(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateUserMandateLevelVbErr
    
     Const strFunctionName As String = "IOrganisationBO_ValidateUserMandateLevel"
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Dim dblUserLoanAmountMandate As Double
    Dim dblUserLTVMandate As Double
    Dim dblReqLoanAmountMandate  As Double
    Dim dblReqLTVMandate As Double
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlDataNode = objIOrganisationDO.ValidateUserMandateLevel(vxmlRequest)
    dblReqLoanAmountMandate = Val(m_objXmlAssist.GetTagValue(vxmlRequest, "AMOUNTREQUESTED"))
    dblReqLTVMandate = Val(m_objXmlAssist.GetTagValue(vxmlRequest, "LTV"))
    dblUserLoanAmountMandate = Val(m_objXmlAssist.GetTagValue(xmlDataNode, "LOANAMOUNTMANDATE"))
    dblUserLTVMandate = Val(m_objXmlAssist.GetTagValue(xmlDataNode, "LTVMANDATE"))
    If (dblUserLoanAmountMandate >= dblReqLoanAmountMandate) And (dblUserLTVMandate >= dblReqLTVMandate) Then
        xmlResponseElem.appendChild(xmlResponseElem.ownerDocument.createElement("VALIDUSER")).Text = "1"
    Else
        xmlResponseElem.appendChild(xmlResponseElem.ownerDocument.createElement("VALIDUSER")).Text = "0"
    End If
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
        
    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    Set IOrganisationBO_ValidateUserMandateLevel = xmlResponseElem.cloneNode(True)
ValidateUserMandateLevelExit:
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
ValidateUserMandateLevelVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_ValidateUserMandateLevel = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume ValidateUserMandateLevelExit
End Function
Public Function ValidateUserMandateLevel(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo ValidateUserMandateLevelVbErr
    Const strFunctionName As String = "ValidateUserMandateLevel"
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIOrganisationDO As IOrganisationDO
            
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
         
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
            
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_ValidateUserMandateLevel(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    ValidateUserMandateLevel = xmlResponseElem.xml
    m_objContext.SetComplete
ValidateUserMandateLevelExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIOrganisationDO = Nothing
    Exit Function
ValidateUserMandateLevelVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    ValidateUserMandateLevel = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume ValidateUserMandateLevelExit
End Function
Private Function GetDatePasswordChanged(ByVal strRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the date/time of the last password change
' pass:         strRequest  xml Request data stream
' return:                   string holding the date/time of last password change
'------------------------------------------------------------------------------------------
On Error GoTo GetDatePasswordChanged_vberr
'SYS0210 MDC 15/05/2000
     Const strFunctionName As String = "GetDatePasswordChanged"
      
    Dim objIOrganisationDO As IOrganisationDO
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlReturnNode As IXMLDOMNode
    Dim xmlPass As FreeThreadedDOMDocument40
    Dim strResponse As String
    Dim strCreateDate As String
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set xmlTempDoc = m_objXmlAssist.load(strRequest, TypeName(Me), strFunctionName)
    Set xmlReturnNode = objIOrganisationDO.GetCurrentPassword(xmlTempDoc.documentElement)
    strResponse = xmlReturnNode.xml
    'strResponse = objOrganisationDO.GetCurrentPassword(strRequest)
    Set xmlPass = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    strCreateDate = m_objXmlAssist.GetTagValue(xmlPass.documentElement, "PASSWORDCREATIONDATE")
        
    GetDatePasswordChanged = strCreateDate
GetDatePasswordChanged_exit:
    Set xmlPass = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
      
    Exit Function
        
GetDatePasswordChanged_vberr:
        
    Set xmlPass = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlReturnNode = Nothing
      
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
End Function
Private Function IOrganisationBO_GetActionOwnerDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetActionOwnerDetailsVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetActionOwnerDetails"
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIOrganisationDO = New OrganisationDO
    Else
        Set objIOrganisationDO = objContext.CreateInstance(App.Title & ".OrganisationDO")
    End If
                
    If vxmlRequest.nodeName = "ACTIONOWNER" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("ACTIONOWNER").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ACTIONOWNER" & " tag not found"
    End If
        
    Set xmlDataNode = objIOrganisationDO.GetActionOwnerDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
    Set IOrganisationBO_GetActionOwnerDetails = xmlResponseElem.cloneNode(True)
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
GetActionOwnerDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
GetActionOwnerDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    objErrAssist.AddToErrSource strFunctionName
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_GetActionOwnerDetails = objErrAssist.CreateErrorResponseEx()
    'GD SYS2092
    'If Not objContext Is Nothing Then
    '    objContext.setcomplete
    'End If
    ' go to clean-up section
    Resume GetActionOwnerDetailsExit
End Function
Public Function GetActionOwnerDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetActionOwnerDetailsVbErr
    Dim strFunctionName As String
    strFunctionName = "GetActionOwnerDetails"
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_GetActionOwnerDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetActionOwnerDetails = xmlResponseElem.xml
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
GetActionOwnerDetailsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
GetActionOwnerDetailsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    objErrAssist.AddToErrSource strFunctionName
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetActionOwnerDetails = objErrAssist.CreateErrorResponse
    'GD SYS2092
    'If Not objContext Is Nothing Then
    '    objContext.setcomplete
    'End If
    ' go to clean-up section
    Resume GetActionOwnerDetailsExit
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
Private Function IOrganisationBO_FindUserNameList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindUserNameListVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindUserNameList"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
        
    Set xmlRequestNode = m_objXmlAssist.GetNode(vxmlRequest, "/REQUEST/USERLIST")
        
    Set xmlDataNode = objIOrganisationDO.FindUserNameList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_FindUserNameList = xmlResponseElem
FindUserNameListExit:
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIOrganisationDO = Nothing
        
    Exit Function
FindUserNameListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindUserNameList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUserNameListExit
End Function
Public Function FindUserNameList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindUserNameListVbErr
    Const strFunctionName As String = "FindUserNameList"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_FindUserNameList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindUserNameList = xmlResponseElem.xml
    m_objContext.SetComplete
FindUserNameListExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    Exit Function
FindUserNameListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindUserNameList = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUserNameListExit
End Function
Private Function IOrganisationBO_FindUnitNameList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitNameListVbErr
    
    Const strFunctionName As String = "IOrganisationBO_FindUnitNameList"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    Set xmlRequestNode = m_objXmlAssist.GetNode(vxmlRequest, "/REQUEST/UNITLIST")
        
    Set xmlDataNode = objIOrganisationDO.FindUnitNameList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IOrganisationBO_FindUnitNameList = xmlResponseElem
FindUnitNameListExit:
    Set xmlRequestNode = Nothing
    Set objIOrganisationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIOrganisationDO = Nothing
        
    Exit Function
FindUnitNameListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_FindUnitNameList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUnitNameListExit
End Function
Public Function FindUnitNameList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindUnitNameListVbErr
    Const strFunctionName As String = "FindUnitNameList"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_FindUnitNameList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindUnitNameList = xmlResponseElem.xml
    m_objContext.SetComplete
FindUnitNameListExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    Exit Function
FindUnitNameListVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindUnitNameList = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUnitNameListExit
End Function
Public Function ValidateUserAuthority(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Validates the user authority passed in against the user authority global
'               parameter passed in
' pass:         vstrXmlRequest  xml Request data stream containing data to be validated
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo ValidateUserAuthorityVbErr
    Const strFunctionName As String = "ValidateUserAuthority"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_ValidateUserAuthority(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    ValidateUserAuthority = xmlResponseElem.xml
    m_objContext.SetComplete
ValidateUserAuthorityExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    Exit Function
ValidateUserAuthorityVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    ValidateUserAuthority = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume ValidateUserAuthorityExit
End Function
Private Function IOrganisationBO_ValidateUserAuthority(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateUserAuthorityVbErr
    
    Const strFunctionName As String = "IOrganisationBO_ValidateUserAuthority"
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim strUserAuthority As String
    Dim strUserAuthRequirement As String
    Dim strValid As String
    Dim objGlobalParameterBO As GlobalParameterBO
    Dim objConvertAssist As ConvertAssist
    strValid = "0"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objGlobalParameterBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    Set objConvertAssist = New ConvertAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                    
    ' Get User authority
    strUserAuthority = m_objXmlAssist.GetNodeValue(vxmlRequest, "/REQUEST/VALIDATION/USERROLE")
    strUserAuthRequirement = m_objXmlAssist.GetNodeValue(vxmlRequest, "/REQUEST/VALIDATION/AUTHORITYREQUIREMENT")
    ' Get the user authority requirement
    Dim strXmlResponse As String, strAuthority As String
    Dim xmlDoc As FreeThreadedDOMDocument40
        
    strXmlResponse = objGlobalParameterBO.GetCurrentParameter(strUserAuthRequirement)
    Set xmlDoc = m_objXmlAssist.load(strXmlResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    strAuthority = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT")
    ' Is it valid?
    If objConvertAssist.CSafeLng(strUserAuthority) >= objConvertAssist.CSafeLng(strAuthority) Then
        strValid = "1"
    End If
    Set xmlDataNode = xmlResponseElem.appendChild(xmlOut.createElement("VALIDITY"))
    xmlDataNode.Text = strValid
        
    m_objContext.SetComplete
    Set IOrganisationBO_ValidateUserAuthority = xmlResponseElem
ValidateUserAuthorityExit:
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objConvertAssist = Nothing
    Exit Function
ValidateUserAuthorityVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_ValidateUserAuthority = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume ValidateUserAuthorityExit
End Function
Public Function FindUnitSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Find unit summary information from unitid
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'history:
'BG SYS1860 22/01/01 Added Method
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitSummaryVbErr
    Const strFunctionName As String = "FindUnitSummary"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_FindUnitSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindUnitSummary = xmlResponseElem.xml
    m_objContext.SetComplete
FindUnitSummaryExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
FindUnitSummaryVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindUnitSummary = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindUnitSummaryExit
End Function
Public Function UpdateHighChequeNum(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
'       Format:
'       <REQUEST>
'               <UNITID></UNITID>
'               <UNHIGHCHEQUENUMBER></UNHIGHCHEQUENUMBER>
'       </REQUEST>
' return:
'   Aw  15/08/02    BM029
On Error GoTo UpdateHighChequeNumVbErr
    
    Const strFunctionName As String = "UpdateHighChequeNum"
        
       
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    Dim objIOrganisationTxBO As IOrganisationTxBO
    Dim strInUserId As String
    Dim strInChequeNum As String
    Dim strResponse As String
    Dim lngMsgNo As Long
    Dim strCreateDate As String
             
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
            
            
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    '---------------------------------------------------------------------------
    ' Delegate the main processing to DoChangePassword
    '---------------------------------------------------------------------------
    Set objIOrganisationTxBO = m_objContext.CreateInstance(App.Title & ".OrganisationTxBO")
    Set xmlResponse = objIOrganisationTxBO.UpdateUnitChequeNo(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlResponse, xmlResponseElem, True
    ' if we get here, everything has completed OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    UpdateHighChequeNum = xmlOut.xml
    m_objContext.SetComplete
UpdateHighChequeNumExit:
    ' clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestNode = Nothing
    Set xmlNode = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseNode = Nothing
    Set xmlElem = Nothing
    Set xmlTempNode = Nothing
    Set objIOrganisationTxBO = Nothing
    Exit Function
UpdateHighChequeNumVbErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    UpdateHighChequeNum = m_objErrAssist.CreateErrorResponse
    m_objContext.SetAbort
        
    ' go to clean-up section
    Resume UpdateHighChequeNumExit
End Function

'=================================================================================
'MSla BBGR13 - System logon for BBG Retail
'=================================================================================
Public Function ValidateUserPassword(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Validate the user logon credentials against the password entered
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'history:
'BBGR13 - Add for System logon on BBG Retail
'------------------------------------------------------------------------------------------

On Error GoTo ValidateUserPasswordErr

    Const strFunctionName As String = "ValidateUserPassword"
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objIAuditDO As IAuditDO
    Dim objConvertAssist As ConvertAssist
    
    Dim lngMsgNo As Long
    Dim strResponse As String
    Dim strLogonAuditRecType As String

    Dim bnlChangePassword As Boolean
    Dim blnIsAttemptOK As Boolean
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set objIAuditDO = m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO")
    
    blnIsAttemptOK = True
     
    
    '------------------------------------------------------------------------------------------
    ' Check the number of access attempts
    '------------------------------------------------------------------------------------------
    If blnIsAttemptOK = True Then
        strLogonAuditRecType = objIAuditDO.GetLogonValueId()
        Set xmlNode = xmlIn.getElementsByTagName("OMIGAUSER").Item(0)
        If xmlNode Is Nothing Then
            blnIsAttemptOK = False
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInternalError, "Expected OMIGAUSER tag"
        End If
        Set xmlElem = xmlIn.createElement("AUDITRECORDTYPE")
        xmlElem.Text = strLogonAuditRecType
        xmlNode.appendChild xmlElem
        Err.Clear
        On Error Resume Next
        ValidateAccessAttempts xmlIn.xml
        m_objErrAssist.SaveErr
        On Error GoTo ValidateUserPasswordErr
        If m_objErrAssist.LastErr <> 0 Then
            blnIsAttemptOK = False
            m_objErrAssist.ReRaise
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Check the userpassword is correct for the userid input and that the
    ' password has not expired
    '------------------------------------------------------------------------------------------
     
    Set objConvertAssist = New ConvertAssist
    If blnIsAttemptOK = True Then
        m_objXmlAssist.ChangeNodeName xmlIn.documentElement, "OMIGAUSER", "PASSWORD"
        strResponse = ValidateCurrentPassword(xmlIn.xml)
        m_objErrAssist.CheckXMLResponse strResponse, True
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        bnlChangePassword = objConvertAssist.CSafeBool(m_objXmlAssist.GetTagValue( _
            xmlDoc.documentElement, "CHANGEPASSWORDINDICATOR"))
        Set xmlResponseElem = xmlDoc.documentElement

    End If
    
    lngMsgNo = m_objErrAssist.CheckXMLResponse(strResponse, False)
    'blnIsAttemptOK = True
    If lngMsgNo <> 0 Then
        If m_objErrAssist.IsWarning(lngMsgNo) = False Then
            blnIsAttemptOK = False
        End If
    End If
    
    '---------------------------------------------------------------------------
    ' create the access audit record even on non-warning error
    '---------------------------------------------------------------------------
    If blnIsAttemptOK = False Then
        CreateAccessAudit vstrXMLRequest, strLogonAuditRecType, blnIsAttemptOK
    End If
    
    
    '---------------------------------------------------------------------------
    ' set up return data
    '---------------------------------------------------------------------------
    
    Set xmlElem = xmlOut.createElement("ACCESSAUDITATTEMPT")
    If blnIsAttemptOK = True Then
        xmlElem.Text = "1"
    Else
        xmlElem.Text = "0"
    End If
    xmlResponseElem.appendChild xmlElem
    
    ValidateUserPassword = xmlResponseElem.xml

ValidateUserPasswordExit:
    ' clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set objIAuditDO = Nothing
    Set objConvertAssist = Nothing
    
    Exit Function
ValidateUserPasswordErr:
    If m_objErrAssist.IsWarning = True Then
        ' add message element to response block
        Dim xmlErrDoc As New FreeThreadedDOMDocument40
        Set xmlErrDoc = m_objXmlAssist.load( _
            m_objErrAssist.FormatMessageNode, TypeName(Me), strFunctionName)
        Set xmlResponseNode = xmlResponseElem.appendChild(xmlErrDoc.documentElement)
        Set xmlErrDoc = Nothing
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    ValidateUserPassword = m_objErrAssist.CreateErrorResponse
    '---------------------------------------------------------------------------
    ' create the access audit record
    '---------------------------------------------------------------------------
    CreateAccessAudit vstrXMLRequest, strLogonAuditRecType, False

        
    ' go to clean-up section
    Resume ValidateUserPasswordExit
End Function
Public Function GetUserCompetency(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  returns user, competency info
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo GetUserCompetencyVbErr
    Const strFunctionName As String = "GetUserCompetency"
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
            
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
         
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_GetUserCompetency(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetUserCompetency = xmlResponseElem.xml
    m_objContext.SetComplete
GetUserCompetencyExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
GetUserCompetencyVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetUserCompetency = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
        
    ' go to clean-up section
    Resume GetUserCompetencyExit
End Function

'CORE191 - New Method
Public Function IsDateWorkingDay(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Validate that the password does not contain user details.
' pass:
'   vstrXMLRequest - Format:
'                    <REQUEST>
'                       <CHANNELID></CHANNELID>
'                       <DATE></DATE>
'                    </REQUEST>
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS" or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo IsDateWorkingDayVbErr
    Const strFunctionName As String = "IsDateWorkingDay"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IOrganisationBO_IsDateWorkingDay(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    IsDateWorkingDay = xmlResponseElem.xml
    m_objContext.SetComplete
    
    ' Clean-up
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    Exit Function

IsDateWorkingDayVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    IsDateWorkingDay = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    
    ' Clean-up
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

End Function

Private Function IOrganisationBO_IsDateWorkingDay(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Checks if the date passed in falls on a working day
' pass:         vxmlRequest xml Request data stream containing data to be persisted
' return:       xml Response Node format:
'                       <RESPONSE>
'                           <STATUS></STATUS>
'                       </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo IsDateWorkingDayVbErr
    Const strFunctionName As String = "IOrganisationBO_IsDateWorkingDay"
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlStatusElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Dim objConvertAssist As New ConvertAssist
    Dim strTargetDate As String
    Dim strChannelID As String
    Dim dtmTargetDate As Date
    Dim blnReturnStatus As Boolean
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    ' Extract key parameters from the XML request
    strChannelID = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//CHANNELID")
    strTargetDate = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//DATE")
    dtmTargetDate = objConvertAssist.CSafeDate(strTargetDate)
    
    blnReturnStatus = True
    
    If objIOrganisationDO.IsDateBankHoliday(strChannelID, dtmTargetDate) Then
        blnReturnStatus = False
    Else
        If IsValidationType("DaysOfWeek", Weekday(dtmTargetDate), "NBD") Then
            blnReturnStatus = False
        End If
    End If
        
    Set xmlStatusElem = xmlOut.createElement("STATUS")
    xmlStatusElem.Text = IIf(blnReturnStatus, "1", "0")
    xmlResponseElem.appendChild xmlStatusElem
        
    m_objContext.SetComplete
    Set IOrganisationBO_IsDateWorkingDay = xmlResponseElem

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlStatusElem = Nothing
    Set objConvertAssist = Nothing
    Set objIOrganisationDO = Nothing
        
    Exit Function

IsDateWorkingDayVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationBO_IsDateWorkingDay = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlStatusElem = Nothing
    Set objConvertAssist = Nothing
    Set objIOrganisationDO = Nothing

End Function

'CORE191 - End
