VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AuditDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AuditDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Data objects class for omAU
'
'Dependencies:  ADOAssist
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date         Description
'MCS    02/09/99    Created
'RF     29/09/99    Added CreateAccessAudit and GetNumberOfFailedAttempts
'RF     19/11/99    Change to search logic in GetNumberOfFailedAttempts
'MH     05/05/00    SYS0571 Use of Now() without formatting
'MC     16/05/00    SYS0210 - Synchronise the password change date/time with corresponding
'                   access audit record
'CL     18/10/00    Core00004 Modifications made to conform to coding templates
'PSC    12/12/00    CORE00004 Change RaiseError to ThrowError
'APS    27/02/01    SYS1986
'MV     06/03/01    SYS2001: Commenting in CreateAccessaudit , as CreateAccessaudit is a sub there is no return Data
'                    at the end of the proc
'LD    11/06/01     SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD    19/06/01     SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        Description
'TK     30/11/2004  E2EM00002504 - Performance related fixes.
'------------------------------------------------------------------------------------------
Option Explicit
Implements IAuditDO
Implements ObjectControl
'=============================================
'Variable Declaration Section
'=============================================
 
'=============================================
'Constant Declaration Section
'=============================================
Private Const cstrLogon = "L"
Private Const cstrPasswordChange = "C"
Private Const cstrApplicationAccess = "AL"
Private Const cstrApplicationRelease = "AR"
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     12/08/2002  BMIDS00322  Core Ref AQR: SYS2341 ; Modified GetHighestSequenceNumberForAttempt
'DB     10/03/2003  BM0383      Commented out code that increments AttemptNumber.
'------------------------------------------------------------------------------------------
Private Function IAuditDO_GetNumberOfFailedAttempts(ByVal vxmlTableElement As IXMLDOMElement) As Long
' header ----------------------------------------------------------------------------------
' description:
' Get number of failed attempts for a given audit record type and user id.
' Search logic is as follows:
'    DateL = datetime of most recent successful logon
'    DateCP  = datetime of most recent change password
'    if DateL is found
'        if DateCP is found
'            if DateL > DateCP
'                searchdate = DateL
'            else (DateCP > DateL)
'                searchdate = DateCP
'            End If
'        else (DateCP not found)
'            searchdate = DateL
'        End If
'    else (DateL not found)
'        if DateCP is found
'            searchdate = DateCP
'        Else
'            searchdate = Null
'        End If
'    End If
'    if searchdate not null
'        NumFailedAttempts = number of failures since searchdate
'    Else
'        NumFailedAttempts = number of failures ever
'    End If
' pass:
'   vstrXMLRequest
'       Format:
'           <REQUEST>
'               <USERID></USERID>
'               <AUDITRECORDTYPE></AUDITRECORDTYPE>
'           </REQUEST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNumberOfFailedAttemptsVbErr
    
    Const strFunctionName As String = "IAuditDO_GetNumberOfFailedAttempts"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
                
    Dim objSQLAssist As New SQLAssist
    Dim objADOAssist As ADOAssist
'    Dim xmlIn As New FreeThreadedDOMDocument40
        
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    Dim strCriteria, strRecTypeTagValue, _
        strRecTypeFieldName, strDateFieldName, _
        strTableName, strUserIdFieldName, _
        strUserIdTagValue As String
    '------------------------------------------------------------------------------------------
    ' validate parameters
    '------------------------------------------------------------------------------------------
    strRecTypeFieldName = "AUDITRECORDTYPE"
    strUserIdFieldName = "USERID"
    strRecTypeTagValue = m_objXmlAssist.GetNodeValue( _
    vxmlTableElement, ".//" & strRecTypeFieldName)
        
    strUserIdTagValue = m_objXmlAssist.GetNodeValue( _
        vxmlTableElement, ".//" & strUserIdFieldName)
    If IAuditDO_IsLogon(strRecTypeTagValue) = False And _
        IAuditDO_IsPasswordChange(strRecTypeTagValue) = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, _
            "Expected Logon or ChangePassword Audit record type"
    End If
    Dim dteLogon, dteChangePassword, dteSearch As Date
    Dim blnLogonDateFound, blnChangePasswordDateFound, blnSearchDateFound As Boolean
    blnChangePasswordDateFound = False
    blnLogonDateFound = False
    blnSearchDateFound = False
    strTableName = "ACCESSAUDIT"
    strDateFieldName = "ACCESSDATETIME"
    '------------------------------------------------------------------------------------------
    ' Get datetime of most recent successful change password record
    '------------------------------------------------------------------------------------------
    ' Build criteria string of format:
    ' "date = (select max date from accessaudit where
    ' userid = x and rectype = y and successindicator = true)"
    strCriteria = _
        strDateFieldName & " = (SELECT MAX (" & strDateFieldName & ")" & _
        " FROM " & strTableName & _
        " WHERE " & _
        strUserIdFieldName & " = " & objSQLAssist.FormatString(strUserIdTagValue) & _
        " AND " & strRecTypeFieldName & " = " & IAuditDO_GetChangePasswordValueId() & _
        " AND SUCCESSINDICATOR = 1)"
    Err.Clear
    On Error Resume Next
    objADOAssist.GetValueFromTable strTableName, strCriteria, _
        strDateFieldName, dteChangePassword
    m_objErrAssist.SaveErr
    On Error GoTo GetNumberOfFailedAttemptsVbErr
    If m_objErrAssist.LastErr = 0 Then
        blnChangePasswordDateFound = True
    Else
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound Then
            m_objErrAssist.ReRaise
        End If
        Err.Clear
    End If
    If IAuditDO_IsPasswordChange(strRecTypeTagValue) = True Then
        
        '------------------------------------------------------------------------------------------
        ' set the search date
        '------------------------------------------------------------------------------------------
        If blnChangePasswordDateFound = True Then
            dteSearch = dteChangePassword
            blnSearchDateFound = True
        End If
    Else
        
        '------------------------------------------------------------------------------------------
        ' Get datetime of most recent successful logon
        '------------------------------------------------------------------------------------------
        strCriteria = _
            strDateFieldName & " = (SELECT MAX (" & strDateFieldName & ")" & _
            " FROM " & strTableName & _
            " WHERE " & _
            strUserIdFieldName & " = " & objSQLAssist.FormatString(strUserIdTagValue) & _
            " AND " & strRecTypeFieldName & " = " & strRecTypeTagValue & _
            " AND SUCCESSINDICATOR = 1)"
        Err.Clear
        On Error Resume Next
        objADOAssist.GetValueFromTable strTableName, strCriteria, _
            strDateFieldName, dteLogon
        m_objErrAssist.SaveErr
        On Error GoTo GetNumberOfFailedAttemptsVbErr
        If m_objErrAssist.LastErr = 0 Then
            blnLogonDateFound = True
        Else
            If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound Then
                m_objErrAssist.ReRaise
            End If
            Err.Clear
        End If
        '------------------------------------------------------------------------------------------
        ' set the search date
        '------------------------------------------------------------------------------------------
        If blnLogonDateFound = True Then
            If blnChangePasswordDateFound = True Then
                If dteLogon > dteChangePassword Then
                    dteSearch = dteLogon
                    blnSearchDateFound = True
                Else
                    dteSearch = dteChangePassword
                    blnSearchDateFound = True
                End If
            Else
                dteSearch = dteLogon
                blnSearchDateFound = True
            End If
        Else
            If blnChangePasswordDateFound = True Then
                dteSearch = dteChangePassword
                blnSearchDateFound = True
            End If
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' set up the criteria for the count of failed attempts
    '------------------------------------------------------------------------------------------
    If blnSearchDateFound = True Then
        ' set criteria of format:
        ' "userid = x and rectype = y and SUCCESSINDICATOR = 0 and date > searchdate"
        strCriteria = _
            strUserIdFieldName & " = " & objSQLAssist.FormatString(strUserIdTagValue) & _
            " and " & strRecTypeFieldName & " = " & strRecTypeTagValue & _
            " and SUCCESSINDICATOR = 0" & _
            " and " & strDateFieldName & " > " & _
            objSQLAssist.FormatDate(dteSearch, dtfDateTime)
    Else
        ' set criteria of format:
        ' "userid = x and rectype = y and SUCCESSINDICATOR = 0"
            strCriteria = _
            strUserIdFieldName & " = " & objSQLAssist.FormatString(strUserIdTagValue) & _
            " and " & strRecTypeFieldName & " = " & strRecTypeTagValue & _
            " and SUCCESSINDICATOR = 0"
    End If
    '------------------------------------------------------------------------------------------
    ' do the count of failed attempts
    '------------------------------------------------------------------------------------------
    IAuditDO_GetNumberOfFailedAttempts = _
        objADOAssist.GetNumberOfRecords(strTableName, strCriteria)
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetNumberOfFailedAttemptsVbErr:
        
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
               
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetHighestSequenceNumberForAttempt(ByVal vstrUserId As String, _
                                                        ByVal vstrInAuditRecordType As String) As Long
' header ----------------------------------------------------------------------------------
' description:
'   get the currently highest sequence number from table ACCESSAUDIT for this value
'   of vstrUserId and vstrInAuditRecordType
' pass:
'   vstrUserId
'   vstrInAuditRecordType
'------------------------------------------------------------------------------------------
    
    On Error GoTo GetHighestSequenceNumberForAttemptVbErr
    Const strFunctionName As String = "GetHighestSequenceNumberForAttempt"
    Dim objSQLAssist As SQLAssist
    Dim objIADOAssist As IADOAssist
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim adoConnection As ADODB.Connection
    Dim adoParameter As ADODB.Parameter
    Dim strSQL As String
    Dim sSQLNoLock As String
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Set objSQLAssist = New SQLAssist
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoParameter = New ADODB.Parameter
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
           
    If Not Len(vstrUserId) > 0 Then
        ' "Invalid parameter value: empty string received"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Invalid USERID"
    End If
    If Not Len(vstrInAuditRecordType) > 0 Then
        ' "Invalid parameter value: empty string received"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Invalid AUDITRECORDTYPE"
    End If
            
    'SYS2341 - The following SQL places an exclusive locks which can cause deadlocks during heavy
    'periods of logging on. It does mean, however, that if two or more people login with the same
    'userid at the same time, the ATTEMPTNUMBER returned will be the same, rather than increment.
    If objIADOAssist.GetDBEngine = SQLServer Then
        sSQLNoLock = " WITH (NOLOCK)"
    End If
            
    strSQL = "SELECT MAX(ATTEMPTNUMBER)" & _
            " FROM ACCESSAUDIT" & sSQLNoLock & _
            " WHERE USERID = ? AND AUDITRECORDTYPE = ?"
    cmd.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(vstrUserId), vstrUserId)
    cmd.Parameters.Append adoParameter
    Set adoParameter = cmd.CreateParameter("AuditRecordType", adInteger, adParamInput)
    cmd.Parameters.Append adoParameter
    cmd(1) = vstrInAuditRecordType
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
        
    intRetries = objIADOAssist.GetRetries()
    intMaxAttempts = 1 + intRetries
    blnOpenedOk = False
    intAttempt = 1
    While (blnOpenedOk = False) And (intAttempt <= intMaxAttempts)
        adoConnection.ERRORS.Clear
        adoConnection.Open objIADOAssist.GetConnStr
        If adoConnection.ERRORS.Count = 0 Then
            blnOpenedOk = True
        End If
        intAttempt = intAttempt + 1
    Wend
    If blnOpenedOk = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeUnableToConnect
    End If
    cmd.ActiveConnection = adoConnection
    Set rst.Source = cmd
    'Populate recordset
    rst.Open
    'Disconnect the recordset
    Set rst.ActiveConnection = Nothing
    adoConnection.Close
          
    'rst!FieldName does not work where field value is null
    If IsNull(rst.Fields(0).Value) Then
        ' sequence number not yet set on this field
        GetHighestSequenceNumberForAttempt = 0
    Else
        GetHighestSequenceNumberForAttempt = rst.Fields(0).Value
    End If
    rst.Close
         
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Set adoConnection = Nothing
    Set adoParameter = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetHighestSequenceNumberForAttemptVbErr:
    m_objErrAssist.SaveErr
    If objIADOAssist.IsDBConnectionError(Err.Description) Then
        Resume Next
    End If
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    ' SYS1986 APS 27/02/01
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Set adoConnection = Nothing
    Set adoParameter = Nothing
        
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    m_objErrAssist.ReRaise
End Function
Private Function IAuditDO_IsPasswordChange(ByVal vstrAuditRecType As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if access audit type is a password change
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo IsPasswordChangeVbErr
    Const strFunctionName As String = "IAuditDO_IsPasswordChange"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
                
    Dim objComboDO As ComboDO
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
        
    Dim blnIsPasswordChange As Boolean
    If objComboDO.IsItemInValidation( _
        "AccessAuditType", vstrAuditRecType, cstrPasswordChange) Then
        blnIsPasswordChange = True
    Else
        blnIsPasswordChange = False
    End If
    IAuditDO_IsPasswordChange = blnIsPasswordChange
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
IsPasswordChangeVbErr:
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
       
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_IsApplicationAccess(ByVal vstrAuditRecType As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if access audit type is an application access
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo IsApplicationAccessVbErr
    Const strFunctionName As String = "IAuditDO_IsApplicationAccess"
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim objComboDO As ComboDO
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim blnIsApplicationAccess As Boolean
    If objComboDO.IsItemInValidation( _
        "AccessAuditType", vstrAuditRecType, cstrApplicationAccess) Then
        blnIsApplicationAccess = True
    Else
        blnIsApplicationAccess = False
    End If
    IAuditDO_IsApplicationAccess = blnIsApplicationAccess
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
IsApplicationAccessVbErr:
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
       
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_IsApplicationRelease(ByVal vstrAuditRecType As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if access audit type is an application release
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo IsApplicationReleaseVbErr
    Const strFunctionName As String = "IAuditDO_IsApplicationRelease"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim objComboDO As ComboDO
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim blnIsApplicationRelease As Boolean
    If objComboDO.IsItemInValidation( _
        "AccessAuditType", vstrAuditRecType, cstrApplicationRelease) Then
        blnIsApplicationRelease = True
    Else
        blnIsApplicationRelease = False
    End If
    IAuditDO_IsApplicationRelease = blnIsApplicationRelease
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
IsApplicationReleaseVbErr:
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
          
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_IsLogon(ByVal vstrAuditRecType As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if access audit type is a system logon or logoff
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo IsLogonVbErr
    Const strFunctionName As String = "IAuditDO_IsLogon"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim objComboDO As ComboDO
          
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim blnIsLogon As Boolean
    If objComboDO.IsItemInValidation("AccessAuditType", vstrAuditRecType, cstrLogon) Then
        blnIsLogon = True
    Else
        blnIsLogon = False
    End If
    IAuditDO_IsLogon = blnIsLogon
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
IsLogonVbErr:
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_GetApplicationLockValueId() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationLockValueIdVbErr
    Const strFunctionName As String = "IAuditDO_GetApplicationLockValueId"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim objComboDO As ComboDO
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("AccessAuditType", cstrApplicationAccess)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    IAuditDO_GetApplicationLockValueId = _
        m_objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "VALUEID")
    Set xmlResponseDoc = Nothing
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetApplicationLockValueIdVbErr:
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    Set xmlResponseDoc = Nothing
    Set objComboDO = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_GetApplicationReleaseValueId() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationReleaseValueIdVbErr
    Const strFunctionName As String = "IAuditDO_GetApplicationReleaseValueId"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Dim objComboDO As ComboDO
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("AccessAuditType", cstrApplicationRelease)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    IAuditDO_GetApplicationReleaseValueId = _
        m_objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "VALUEID")
    Set objComboDO = Nothing
    Set xmlResponseDoc = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetApplicationReleaseValueIdVbErr:
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    Set objComboDO = Nothing
    Set xmlResponseDoc = Nothing
        
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
      
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_GetChangePasswordValueId() As String
' header ----------------------------------------------------------------------------------
' description:
'   Get combo value id for a change password audit record type
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetChangePasswordValueIdVbErr
    Const strFunctionName As String = "IAuditDO_GetChangePasswordValueId"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim objComboDO As ComboDO
               
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("AccessAuditType", cstrPasswordChange)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    IAuditDO_GetChangePasswordValueId = _
        m_objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "VALUEID")
    Set xmlResponseDoc = Nothing
    Set objComboDO = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetChangePasswordValueIdVbErr:
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    Set xmlResponseDoc = Nothing
    Set objComboDO = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
       
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IAuditDO_GetLogonValueId() As String
' header ----------------------------------------------------------------------------------
' description:
'   Get combo value id for a logon audit record type
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetLogonValueIdVbErr
    Const strFunctionName As String = "IAuditDO_GetLogonValueId"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim objComboDO As ComboDO
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("AccessAuditType", cstrLogon)
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    IAuditDO_GetLogonValueId = _
        m_objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "VALUEID")
    Set xmlResponseDoc = Nothing
    Set objComboDO = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetLogonValueIdVbErr:
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    Set xmlResponseDoc = Nothing
    Set objComboDO = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
           
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub IAuditDO_AddDerivedData(ByVal vxmlData As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vxmlData          base XML node
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr
    Const strFunctionName As String = "IAuditDO_AddDerivedData"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
AddDerivedDataVBErr:
      
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub IAuditDO_CreateAccessAudit(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlTableElement  xml Request data stream containing data to be persisted.
'                   Format:
'                    <ACCESSAUDIT>
'                        <USERID></USERID>
'                        <AUDITRECORDTYPE></AUDITRECORDTYPE>
'                        <MACHINEID></MACHINEID>
'                        <SUCCESSINDICATOR></SUCCESSINDICATOR>
'                        <ONBEHALFOFUSERID></ONBEHALFOFUSERID>
'                        <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                        <PASSWORDCREATIONDATE>optional</PASSWORDCREATIONDATE>
'                    </ACCESSAUDIT>
'------------------------------------------------------------------------------------------
On Error GoTo CreateAccessAuditVbErr
    Const strFunctionName As String = "IAuditDO_CreateAccessAudit"
        
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim strCreationDate As String
    'Dim strAccessAuditGUID As String
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
               
    'initialise some generally used values
    Dim strUserId As String
    Dim strAuditRecType As String
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
    strAuditRecType = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//AUDITRECORDTYPE")
    '------------------------------------------------------------------------------------------
    ' create ACCESSAUDIT record
    '------------------------------------------------------------------------------------------
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.appendChild vxmlTableElement.cloneNode(True)
            
    Set xmlTableNode = m_objXmlAssist.GetNode(xmlDoc, ".//ACCESSAUDIT")
       
    'set up ACCESSAUDITGUID, ACCESSDATETIME and ATTEMPTNUMBER
    Set xmlElem = xmlDoc.createElement("ACCESSAUDITGUID")
    xmlElem.Text = CreateGUID()
    'strAccessAuditGUID = xmlElem.Text
    xmlTableNode.appendChild xmlElem
    'SYS0210 MDC 15/05/2000. If exists (ie for change password access audit type) use password
    '                        creation Date
    strCreationDate = m_objXmlAssist.GetNodeValue(xmlTableNode, _
                                                  ".//PASSWORDCREATIONDATE", _
                                                  False, False)
    If Len(Trim$(strCreationDate)) = 0 Then
        strCreationDate = Format$(Now(), "DD/MM/YYYY HH:MM:SS")
    End If
    Set xmlElem = xmlDoc.createElement("ACCESSDATETIME")
    xmlElem.Text = strCreationDate 'Format$(Now(), "DD/MM/YYYY HH:MM:SS")
    xmlTableNode.appendChild xmlElem
    'DB BM0383 - AttemptNumber no longer needed
    'Set xmlElem = xmlDoc.createElement("ATTEMPTNUMBER")
    'xmlElem.Text = _
    '    CStr(GetHighestSequenceNumberForAttempt(strUserId, strAuditRecType) + 1)
    'xmlTableNode.appendChild xmlElem
    'DB End
    Dim objIomAUClassDef As IomAUClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Set objIomAUClassDef = New omAUClassDef
        
    Set xmlClassDefDoc = objIomAUClassDef.LoadAccessAuditData()
    m_objIDOAssist.Create xmlDoc.documentElement, xmlClassDefDoc
    Set xmlClassDefDoc = Nothing
    '------------------------------------------------------------------------------------------
    ' create CHANGEPASSWORD record if password has been changed by someone other
    ' than the user (e.g. system administrator)
    '------------------------------------------------------------------------------------------
    Dim strONBEHALFOFUSERID As String
    'Find the xmlnode ONBEHALFOFUSERID in the current path
    strONBEHALFOFUSERID = _
        m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//ONBEHALFOFUSERID", False, False)
    'If a value has been found
    If Len(strONBEHALFOFUSERID) > 0 Then
        
        If strONBEHALFOFUSERID <> strUserId Then
            
            If IAuditDO_IsPasswordChange(strAuditRecType) Then
                
                'create CHANGEPASSWORD record
                Dim strTableName As String
                strTableName = "CHANGEPASSWORD"
                m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, _
                                            "ACCESSAUDIT", strTableName
                Set xmlTableNode = m_objXmlAssist.GetNode(xmlDoc, ".//" & strTableName)
                       
                Set xmlClassDefDoc = objIomAUClassDef.LoadChangePasswordData()
                m_objIDOAssist.Create xmlDoc.documentElement, xmlClassDefDoc
                Set xmlClassDefDoc = Nothing
            End If
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' create APPLICATIONACCESS for application access/release
    '------------------------------------------------------------------------------------------
    If IAuditDO_IsApplicationAccess(strAuditRecType) Or _
        IAuditDO_IsApplicationRelease(strAuditRecType) Then
        ' create APPLICATIONACCESS record
        Dim strAppNo As String
        ' SR 23/03/00 - Search in the current node only (not recursively)
        strAppNo = m_objXmlAssist.GetNodeValue(vxmlTableElement, "APPLICATIONNUMBER")
           
        strTableName = "APPLICATIONACCESS"
        m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, _
            "ACCESSAUDIT", strTableName
        Set xmlTableNode = m_objXmlAssist.GetNode(xmlDoc, ".//" & strTableName)
                       
        Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlElem.Text = strAppNo
        xmlTableNode.appendChild xmlElem
                
        Set xmlClassDefDoc = objIomAUClassDef.LoadApplicationAccessData()
        m_objIDOAssist.Create xmlDoc.documentElement, xmlClassDefDoc
        Set xmlClassDefDoc = Nothing
    End If
    Set xmlDoc = Nothing
    Set xmlIn = Nothing
    Set xmlTableNode = Nothing
    Set xmlElem = Nothing
        
    m_objContext.SetComplete
    'Set xmlTableNode = xmlDoc.createElement("GENERATEDKEYS")
    'Set xmlElem = xmlDoc.createElement("ACCESSAUDITGUID")
    'xmlElem.Text = strAccessAuditGUID
    'xmlTableNode.appendChild xmlElem
    'Set IAuditDO_CreateAccessAudit = xmlElem
    Exit Sub
CreateAccessAuditVbErr:
      
    Set xmlDoc = Nothing
    Set xmlIn = Nothing
    Set xmlTableNode = Nothing
    Set xmlElem = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
       
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
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
