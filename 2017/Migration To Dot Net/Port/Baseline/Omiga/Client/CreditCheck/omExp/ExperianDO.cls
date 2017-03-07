VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ExperianDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'Workfile:      ExperianDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Data Object
'               contains (implied) standard interface:
'               Public Sub GetDelphiData(ByVal strApplicationNumber As String)
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: RequiresNewTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     16/03/00    Created
'MC     18/05/00    Amend to use I-interface, move DemoOptions methods to omCC
'MC     01/09/00    SYS1174 UserID and UnitID no longer mandatory
'LD     07/11/00    Explicitly close database connections
'LD     07/11/00    Explicitly close recordsets
'ASt    08/03/01    CC012 Pass guid as binary into sp_Experian_gs.SetDelphi96DataReply stored
'                   procedure.
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     19/06/01    SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'AD     19/10/01    Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
'------------------------------------------------------------------------------------------
Option Explicit
Implements IExperianDO
Implements ObjectControl

Private m_objXmlAssist As XMLAssist


'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    13/08/2002  BMIDS00336 CCWP1 BM062 Credit Checking and Bureau Download
'MDC    24/09/2002  BMIDS00336 Make Experian.Search late bound.
'MDC    27/09/2002  BMIDS00521 Return error messages from Experian
'MDC    08/10/2002  BMIDS00605 Re-integrate Experian search into BMTB
'MDC    19/11/2002  BMIDS00990 BUR1 prefix offset adjustment
'MV     22/03/2003  BM0449 Amended IExperianDO_RunCreditCheck
'MV     10/04/2003  BM0449 Amended IExperianDO_RunCreditCheck
'HMA    25/11/2003  BM0555 Changed LogToFile to read the file path from Global Parameter.
'GHun   28/11/2003  BMIDS658 Minor performance improvements and clean up
'HMA    16/12/2003  BM0555 New Global Parameter for CreditCheck Log file path
'INR    13/02/2004  BMIDS682 Address Targeting Processing
'INR    19/03/2004  BMIDS730 Address Targeting Processing
'INR    25/04/2004  BMIDS744 Third Part Data changes.
'HMA    17/07/2004  BMIDS748 Allow for Enquiry Credit Check and Upgrade to Full Application Credit Check.
'GHun   22/11/2004  BMIDS933 Changed RunCreditCheck and AppendToStrFromSQLSvrRecSet
'JD     28/02/2005  BMIDS980 Add BUR5 block read
'------------------------------------------------------------------------------------------

'INR SYS4862
Private Const nSizeHeaderplus1 As Integer = 134
Private Const nSizeDelphi96 As Integer = 500

Private Enum LOG_TYPE
    ltRequest
    ltResponse
End Enum

'BMIDS00521 MDC 27/09/2002
Private Const oeCreditCheckError = 4812
'BMIDS00521 MDC 27/09/2002 - End

Private m_objContext    As ObjectContext
Private m_objErrAssist  As ErrAssist

Private Function IExperianDO_DelphiData(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

'BMIDS00336 MDC 14/08/2002
'No longer required

' header ----------------------------------------------------------------------------------
' description:
'   Calls a stored procedure to formulate an experian request.  Calls the Experian RSO.
'   Calls another stored procedure to unpack the data.  This transaction must commit
'   otherwise multiple credit search foot-prints could be made.
' pass:
'   strApplicationNumber Application Number identifying the case to process.
' return:
'   FindList        boolean containing a success indicator
' Raise Errors:
'   if ???, raise ???
'------------------------------------------------------------------------------------------
'On Error GoTo DelphiDataVbErr
'
'Dim strUserId As String
'Dim strUnitId As String
'Dim strApplicationNumber As String
''Error raised by Experian stored procedure if mandatory data is missing
'Const MANDATORY_DATA_ERR = "ORA-20003"
'
'Dim objXMLElement As IXMLDOMElement
'Dim objXmlChildNode As IXMLDOMNode
'Dim objXmlOut As New FreeThreadedDOMDocument40
'Dim objResponseNode As IXMLDOMNode
'
'    Dim strFunctionName As String
'    strFunctionName = "IExperianDO_DelphiData"
'
'    #If PROFILING Then
'        Dim objLogAssist As New LogAssist
'        objLogAssist.StartTimerEx TypeName(Me) & "." & strFunctionName, True
'    #End If
'
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'
'    Dim objErrAssist As New ErrAssist
'    Dim objADOAssist As ADOAssist
'#If GENERIC_SQL Then
'    Dim objSQLAssist As SQLAssist
'#End If
'
'    Set objXMLElement = objXmlOut.createElement("RESPONSE")
'    Set objResponseNode = objXmlOut.appendChild(objXMLElement)
'
'    'SYS0755 Use the UnitID and UserID nodes if they exist. Otherwise default to the Request node attribute values
'    strUserId = m_objXmlAssist.GetElementText(vobjRequestNode, "USERID")
'    If Len(Trim$(strUserId)) = 0 Then
'        'Get the UserID attribute from the REQUEST node
'        'SYS1174 UserID and UnitID no longer mandatory
'        strUserId = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "USERID")
'        If Len(Trim$(strUserId)) = 0 Then
'            'Get default user ID from global parameters
'            strUserId = GetDefaultUserId()
'        End If
'    End If
'    strUnitId = m_objXmlAssist.GetElementText(vobjRequestNode, "UNITID")
'    If Len(Trim$(strUnitId)) = 0 Then
'        'Get the UnitID attribute from the REQUEST node
'        'SYS1174 UserID and UnitID no longer mandatory
'        strUnitId = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "UNITID")
'    End If
'
'    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "APPLICATIONNUMBER")
'
'    Dim strSQL As String
'    strSQL = "{call sp_Experian.GetDelphi96DataRequest(?,?,?," _
'        & "{resultset 10," _
'        & "tableszHeader," _
'        & "tableszNAM1," _
'        & "tableszAUK1," _
'        & "tableszCCN1," _
'        & "tableszAP01," _
'        & "tableszAP42," _
'        & "tableszAM01," _
'        & "tableszAM02," _
'        & "tableszAM03," _
'        & "tableszAM04" _
'        & "})}"
'
'    If objContext Is Nothing Then
'        Set objADOAssist = New ADOAssist
'    Else
'        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
'    End If
'
'    Dim adoConnection As New ADODB.Connection
'    adoConnection.ConnectionString = objADOAssist.GetConnStr
'    Set objADOAssist = Nothing
'    adoConnection.CursorLocation = adUseClient
'    adoConnection.Open
'
'    Dim adoCommand As New ADODB.Command
'    Set adoCommand.ActiveConnection = adoConnection
'    adoCommand.CommandType = adCmdText
'    adoCommand.CommandText = strSQL
'
'    Dim adoParameter As ADODB.Parameter
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strUserId), strUserId)
'    adoCommand.Parameters.Append adoParameter
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strUnitId), strUnitId)
'    adoCommand.Parameters.Append adoParameter
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
'    adoCommand.Parameters.Append adoParameter
'
'    Dim adoRecSet As ADODB.Recordset
'    Set adoRecSet = adoCommand.Execute
'
'    Set adoCommand.ActiveConnection = Nothing
'        Set adoCommand = Nothing
'
'    If adoRecSet.EOF Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
'    End If
'
'    ' make the request from the returned record set
'    Dim strRequest As String
'    strRequest = ""
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszHeader"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszNAM1"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAUK1"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszCCN1"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAP01"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAP42"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM01"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM02"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM03"
'    AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM04"
'    adoRecSet.Close
'    Set adoRecSet = Nothing
'
'    ' make the call to experian
'    Dim strErrorMessage As String
'    Dim strHeader As String
'    Dim strDEC1 As String
'    Dim strDelphi96 As String
'    Dim bSuccessIndicator As Boolean
'    Dim ExpInt As New ExperianInterface1
'
'    'MDC 20/04/2000. Automatic Retry
'    Dim bRetry As Boolean
'    Dim nAttempts As Integer
'    Const MAX_ATTEMPTS = 3
'    Const ERR_NODATA = "the interface to experian returned no data"
'    Do
'        nAttempts = nAttempts + 1
'        ExpInt.Run strRequest, strErrorMessage, strHeader, strDEC1, strDelphi96, bSuccessIndicator
'        If bSuccessIndicator = False And InStr(LCase$(strErrorMessage), ERR_NODATA) And nAttempts < MAX_ATTEMPTS Then
'            bRetry = True
'        'ExpInt is returning SuccessIndicator = True when there is an ODBC error at Egg.
'        'Temporary fix until ExpInt component can be amended.
'        ElseIf bSuccessIndicator = True And Len(Trim$(strErrorMessage)) > 0 Then
'            bRetry = False
'            bSuccessIndicator = False
'        Else
'            bRetry = False
'        End If
'    Loop Until bRetry = False
'
'    Set ExpInt = Nothing
'    If bSuccessIndicator = False Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, 282, strErrorMessage
'    End If
'
'    ' Generate the GUID
'    Dim strGuid As String
'    strGuid = CreateGUID()
'
'    ' send the reply to the database
'#If GENERIC_SQL Then
'    strSQL = "{call sp_Experian_gs.SetDelphi96DataReply(?, ?, ?, ?)}"
'#Else
'    strSQL = "{call sp_Experian.SetDelphi96DataReply(?, ?, ?, ?)}"
'#End If
'
'    Set adoCommand = New ADODB.Command
'    Set adoCommand.ActiveConnection = adoConnection
'    adoCommand.CommandType = adCmdText
'    adoCommand.CommandText = strSQL
'
'    ' ASt 08/03/01 CC012 Pass guid as binary
'#If GENERIC_SQL Then
'    Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
'    adoCommand.Parameters.Append adoParameter
'    Set objSQLAssist = New SQLAssist
'    adoCommand(0) = objSQLAssist.GuidStringToByteArray(strGuid)
'    Set objSQLAssist = Nothing
'#Else
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strGuid), strGuid)
'    adoCommand.Parameters.Append adoParameter
'#End If
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHeader), strHeader)
'    adoCommand.Parameters.Append adoParameter
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDEC1), strDEC1)
'    adoCommand.Parameters.Append adoParameter
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDelphi96), strDelphi96)
'    adoCommand.Parameters.Append adoParameter
'
'    adoCommand.Execute , , adExecuteNoRecords
'
'    Set adoCommand.ActiveConnection = Nothing
'        Set adoCommand = Nothing
'        adoConnection.Close
'        Set adoConnection = Nothing
'
'    If bSuccessIndicator = False Then
'        'objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
'         Err.Raise -1, Err.Source, strErrorMessage
'    End If
'
'    Set objErrAssist = Nothing
'
'    'Set objXMLElement = vobjResponseNode.ownerDocument.createElement("SUCCESSINDICATOR")
'    'objXMLElement.Text = CStr(bSuccessIndicator)
'    'Set objXmlChildNode = vobjResponseNode.appendChild(objXMLElement)
'
'    objXMLElement.setAttribute "TYPE", "SUCCESS"
'    Set IExperianDO_DelphiData = objResponseNode
'
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
'
'    #If PROFILING Then
'        objLogAssist.StopTimerEx
'        Set objLogAssist = Nothing
'    #End If
'
'    Exit Function
'
'DelphiDataVbErr:
'
'    #If PROFILING Then
'        objLogAssist.StopTimerEx
'        Set objLogAssist = Nothing
'    #End If
'
'    ' Close the recordset if still open
'        If Not adoRecSet Is Nothing Then
'                If adoRecSet.State = adStateOpen Then
'                        adoRecSet.Close
'                End If
'        End If
'
'    ' Close the connection if still open
'        If Not adoConnection Is Nothing Then
'                If adoConnection.State = adStateOpen Then
'                        adoConnection.Close
'                End If
'        End If
'
'    'Trap error raised by stored procedure and give more useful error message
'    If InStr(Err.Description, MANDATORY_DATA_ERR) Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, 280
'    End If
'
'    Set adoRecSet = Nothing
'    Set adoConnection = Nothing
'    Set objADOAssist = Nothing
'#If GENERIC_SQL Then
'    Set objSQLAssist = Nothing
'#End If
'
'    Set adoCommand = Nothing
'
'    objErrAssist.AddToErrSource (strFunctionName)
'
'    If objErrAssist.IsSystemError = True Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
'            objContext.SetComplete
'        End If
'    End If
'
'    Set objErrAssist = Nothing
'
'    '   re-raise error for calling object to interpret as appropriate
'    Err.Raise Err.Number, Err.Source, Err.Description
       
End Function

Private Function IExperianDO_DelphiDemoData(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo DelphiDemoDataVbErr

Dim cmd As ADODB.Command
Dim prmParam As ADODB.Parameter
Dim conn As ADODB.Connection
Dim strGuid As String
Dim strUserID As String
Dim strApplicationNumber As String
Dim strApplicationFactFindNumber As String

Dim objXMLElement As IXMLDOMElement
Dim objResponseNode As IXMLDOMNode
Dim objXMLOut As FreeThreadedDOMDocument40

Dim objADOAssist As ADOAssist

    Const strFunctionName As String = "IExperianDO_DelphiDemoData"

    Set objXMLOut = New FreeThreadedDOMDocument40
    objXMLOut.validateOnParse = False
    objXMLOut.setProperty "NewParser", True
    Set objXMLElement = objXMLOut.createElement("RESPONSE")
    Set objResponseNode = objXMLOut.appendChild(objXMLElement)
    
    'Get a GUID for the new record(s)
    strGuid = CreateGUID()
    
    'SYS0755 Use the UserID node if it exists. Otherwise default to the Request node attribute value
    strUserID = m_objXmlAssist.GetElementText(vobjRequestNode, "USERID")
    If Len(Trim$(strUserID)) = 0 Then
        'SYS1174 UserID no longer mandatory
        strUserID = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "USERID")
        If Len(Trim$(strUserID)) = 0 Then
            'Get default user ID from global parameters
            strUserID = GetDefaultUserId()
        End If
    End If
    
    'Get Application Number of current application
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "APPLICATIONFACTFINDNUMBER")
    
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CreditCheck"
    cmd.CommandType = adCmdStoredProc
    
    'Set input parameters required by stored procedure
    Set prmParam = cmd.CreateParameter("szNewApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    cmd.Parameters.Append prmParam
    Set prmParam = cmd.CreateParameter("nNewApplicationFactFindNumber", adInteger, adParamInput, , strApplicationFactFindNumber)
    cmd.Parameters.Append prmParam
    Set prmParam = cmd.CreateParameter("rawNewCreditCheckGUID", adGUID, adParamInput, , strGuid)
    cmd.Parameters.Append prmParam
    Set prmParam = cmd.CreateParameter("szUserID", adBSTR, adParamInput, Len(strUserID), strUserID)
    cmd.Parameters.Append prmParam
    Set prmParam = cmd.CreateParameter("dtDateTime", adDate, adParamInput, , Format$(Now, "dd-mmm-yy hh:mm"))
    cmd.Parameters.Append prmParam

    'Open a connection to the database
    Set conn = New ADODB.Connection
    conn.ConnectionString = objADOAssist.GetConnStr
    conn.CursorLocation = adUseClient
    conn.Open

    'Execute the stored procedure
    cmd.ActiveConnection = conn
    cmd.Execute
    
    conn.Close
    
    'Set objXMLElement = vobjResponseNode.ownerDocument.createElement("SUCCESSINDICATOR")
    'objXMLElement.Text = CStr(True)
    'Set objXmlChildNode = vobjResponseNode.appendChild(objXMLElement)
    
    objXMLElement.setAttribute "TYPE", "SUCCESS"
    Set IExperianDO_DelphiDemoData = objResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Set conn = Nothing
    Set objADOAssist = Nothing
    Set cmd = Nothing
    Set prmParam = Nothing
    Set objXMLOut = Nothing
  
    Exit Function

DelphiDemoDataVbErr:
    
    ' Close the connection if still open
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
    End If

    Set conn = Nothing
    Set objADOAssist = Nothing
    Set cmd = Nothing
    Set prmParam = Nothing
    Set objXMLOut = Nothing
  
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

Private Function GetDefaultUserId() As String

On Error GoTo GetDefaultUserIdVbErr
    
    Dim objGlobalParamDO As GlobalParameterDO
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    GetDefaultUserId = objGlobalParamDO.GetCurrentParameterByType("DefaultUserId", "STRING")
    
GetDefaultUserIdVbExit:
    Set objGlobalParamDO = Nothing
    Exit Function
    
GetDefaultUserIdVbErr:
    
    'If error occurs return empty string
    GetDefaultUserId = vbNullString
    GoTo GetDefaultUserIdVbExit
    
End Function

Private Sub AppendToStrFromRecSet(ByRef strInOut As String, ByRef adoRecSet As ADODB.Recordset, ByVal strItem As String)
'    adoRecSet.MoveFirst
    ' AD 19/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.

    Do While True
        If IsNull(adoRecSet.fields.Item(strItem).Value) Then
            Exit Do
        End If
        strInOut = strInOut & adoRecSet.fields.Item(strItem).Value
        adoRecSet.MoveNext
        If adoRecSet.EOF Then
            Exit Do
        End If
    Loop
End Sub

Private Function IExperianDO_FullBureauData(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

'BMIDS00336 MDC 14/08/2002
'No longer needed
'On Error GoTo FullBureauDataVbErr
'
'    Dim strFunctionName As String
'    strFunctionName = "IExperianDO_FullBureauData"
'
'    Dim xmlElement As IXMLDOMElement
'    Dim xmlOut As New FreeThreadedDOMDocument40
'    Dim xmlResponseNode As IXMLDOMNode
'    Dim xmlTempResponse As IXMLDOMNode
'
'    #If PROFILING Then
'        Dim objLogAssist As New LogAssist
'        objLogAssist.StartTimerEx TypeName(Me) & "." & strFunctionName, True
'    #End If
'
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'
'    Dim objErrAssist As New ErrAssist
'    Dim objADOAssist As ADOAssist
'
'    Dim strFBRequest As String
'
'    Set xmlElement = xmlOut.createElement("RESPONSE")
'    Set xmlResponseNode = xmlOut.appendChild(xmlElement)
'
'    strFBRequest = m_objXmlAssist.GetTagValue(vxmlRequest, "FBREQUEST")
'
'    If objContext Is Nothing Then
'        Set objADOAssist = New ADOAssist
'    Else
'        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
'    End If
'
'    If strFBRequest = "" Then
'        ' FIX ME - Run stored proc to build request, and fetch it
'    End If
'
'    Dim ExpInt As New ExperianInterface1
'    Dim bRetry As Boolean
'    Dim nAttempts As Integer
'    Dim strErrorMessage As String
'    Dim strFBReply As String
'    Dim bSuccessIndicator As Boolean
'
'    Const MAX_ATTEMPTS = 3
'    Const ERR_NODATA = "the interface to experian returned no data"
'    Do
'        nAttempts = nAttempts + 1
'
''        strFBRequest = "TASKCPLKNRNRCKLU620101ERROR5678JB      010116105443000T     FBL     NEW                                  31 3200000013  2532      000"
'
'        ExpInt.RunFullBureau strFBRequest, strErrorMessage, strFBReply, bSuccessIndicator
'        If bSuccessIndicator = False And InStr(LCase$(strErrorMessage), ERR_NODATA) And nAttempts < MAX_ATTEMPTS Then
'            bRetry = True
'        'ExpInt is returning SuccessIndicator = True when there is an ODBC error at Egg.
'        'Temporary fix until ExpInt component can be amended.
'        ElseIf bSuccessIndicator = True And Len(Trim$(strErrorMessage)) > 0 Then
'            bRetry = False
'            bSuccessIndicator = False
'        Else
'            bRetry = False
'        End If
'    Loop Until bRetry = False
'
'    Set ExpInt = Nothing
'
'    If bSuccessIndicator = False Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, 282, strErrorMessage
'    End If
''    strFBReply = "EQFXCPLKNRNRCKLU620101ERROR5678JB      990318100802000T     RETURN                  N                    32Y3200001372  2535      005BE01007301           66 CARNBRAE AVENUE##BELFAST#COUNTY ANTRIM#BT8 6NH#@#BE02047702         00#9800###CHRISTOPHER#J#KEENAN#####@       00#9800###MAUREEN##KEENAN#####@       00#9798###SEAMUS#G#MCEVOY#####@       00#9798###JOANNE##MCEVOY#####@       00#9397###PAULINE##WOODS#####@       00#9397###ROGER#F#WOODS#####@       00#8693###BRUCE#C#MCCONNELL#####@       00#8693###KAREN#C#MCCONNELL#####@       00#8486###JAMES##MURPHY#####@       00#8486###MICHELE##MURPHY#####@       00#8384###KIMBERLY##GARRETT#####@       00#8384##MS#ROBERTA##GARRETT#####@#BF01016303               260400         4#LLOYDS UDT LTD##MR#SEAMUS##MCEVOY##########@             121000         4#PRUDENTIAL BANKI##MR#SEAMUS##MCEVOY##########@"
'
'    Dim strGuid As String
'    strGuid = m_objXmlAssist.GetTagValue(vxmlRequest, "CREDITCHECKGUID")
'
'    Dim strSQL As String
'    Dim adoConnection As New ADODB.Connection
'    Dim adoParameter As ADODB.Parameter
'    adoConnection.ConnectionString = objADOAssist.GetConnStr
'    Set objADOAssist = Nothing
'    adoConnection.CursorLocation = adUseClient
'    adoConnection.Open
'
'    Dim adoCommand As ADODB.Command
'
'    ' send the reply to the database
'    strSQL = "{call sp_FullBureau.SetFullBureauReply(?, ?)}"
'
'    Set adoCommand = New ADODB.Command
'    Set adoCommand.ActiveConnection = adoConnection
'    adoCommand.CommandType = adCmdText
'    adoCommand.CommandText = strSQL
'
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strGuid), strGuid)
'    adoCommand.Parameters.Append adoParameter
'
'    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strFBReply), strGuid)
'    adoCommand.Parameters.Append adoParameter
'
'    adoCommand.Execute , , adExecuteNoRecords
'
'    Set adoCommand.ActiveConnection = Nothing
'    Set adoCommand = Nothing
'    adoConnection.Close
'    Set adoConnection = Nothing
'
'    Set objErrAssist = Nothing
'
'    xmlElement.setAttribute "TYPE", "SUCCESS"
'    Set IExperianDO_FullBureauData = xmlResponseNode
'
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
'
'    #If PROFILING Then
'        objLogAssist.StopTimerEx
'        Set objLogAssist = Nothing
'    #End If
'
'    Exit Function
'
'FullBureauDataVbErr:
'
'    #If PROFILING Then
'        objLogAssist.StopTimerEx
'        Set objLogAssist = Nothing
'    #End If
'
'    ' Close the connection if still open
'    If Not adoConnection Is Nothing Then
'        If adoConnection.State = adStateOpen Then
'            adoConnection.Close
'        End If
'    End If
'
'    Set adoConnection = Nothing
'    Set objADOAssist = Nothing
'    Set adoCommand = Nothing
'
'    objErrAssist.AddToErrSource (strFunctionName)
'
'    If objErrAssist.IsSystemError = True Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
'            objContext.SetComplete
'        End If
'    End If
'
'    Set objErrAssist = Nothing
'
'    '   re-raise error for calling object to interpret as appropriate
'    Err.Raise Err.Number, Err.Source, Err.Description
       
End Function

Private Function IExperianDO_FullBureauDemoData(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
End Function

Private Function IExperianDO_RunCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

'BMIDS00336 MDC 14/08/2002

' header ----------------------------------------------------------------------------------
' description:
'   Calls a stored procedure to formulate an experian request.  Calls Experian through
'   the BM Translation Bridge and then calls another stored procedure to unpack the data.
'   This transaction must commit otherwise multiple credit search foot-prints could be made.
' pass:
'   vobjRequestNode: XML Request Node
' return:
'   Response XML - Success = True/False
'------------------------------------------------------------------------------------------
    On Error GoTo RunCreditCheckExit
    
    Const strFunctionName As String = "IExperianDO_RunCreditCheck"
    
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim prmFactFind As ADODB.Parameter
    Dim prmCustNumSame As ADODB.Parameter
    Dim prmCustVerSame As ADODB.Parameter
    Dim prmCustNumOther As ADODB.Parameter
    Dim prmCustVerOther As ADODB.Parameter
    Dim adoRecSet As ADODB.Recordset
    Dim objExperianInterface As Object
    
    Dim objADOAssist As ADOAssist
    Dim objSqlAssist As SQLAssist
    
    Dim strUserID As String
    Dim strUnitId As String
    Dim strApplicationNumber As String
    Dim strSQL As String
    Dim strRequest As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim strDEC1 As String
    Dim strDelphi96 As String
    Dim strGuid As String
    Dim intDBEngineType As DBENGINETYPE
    Dim blnLog As Boolean
    
    'INR SYS4862
    Dim nAppFactFindNumber As Integer
    Dim szCustomerNumberSameP As String
    Dim nCustomerVersionNumberSameP As Integer
    Dim szCustomerNumberOtherN As String
    Dim nCustomerVersionNumberOtherN As Integer
    Dim strExpReturn As String
    Dim intNumParams As Integer
    'BMIDS682
    Dim sRunFromTM030 As String
    Dim blnRunFromTM030 As Boolean
    Dim sAddressTargetReq As String
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim strAddressTargeting As String
    Dim strOrigRequest As String
    Dim bNoFurtherCalls As Boolean
    Dim xmlReqHeaderElem As IXMLDOMElement
    Dim xmlCaseTaskNode     As IXMLDOMNode
    bNoFurtherCalls = False
    
    'BMIDS730
    Dim nReturnedAUKPos As Integer
    Dim nReturnedBUKPos As Integer
    Dim xmlDeclareNode As IXMLDOMNode
    'BMIDS744
    Dim strTPD1 As String
    Dim strBUR5 As String  'JLD BMIDS980
    
    'BMIDS748
    Dim sEnquiryCreditCheck As String
    Dim iEnquiryCreditCheck As Integer
    
    'Constants
    Const cstrFunctionName As String = "IExperianDO_RunCreditCheck"
    Const MANDATORY_DATA_ERR As String = "ORA-20003"  'Error raised by Experian stored procedure if mandatory data is missing

    '===========================================================================================
    'Initialise
    '===========================================================================================

    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set objSqlAssist = New SQLAssist
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlResponseElem
    
    blnLog = CBool(m_objXmlAssist.GetAttributeValue(vobjRequestNode, "REQUEST", "LOG"))
    'BMIDS682
    sRunFromTM030 = m_objXmlAssist.GetAttributeValue(vobjRequestNode, "REQUEST", "CREDITCHECKINVOKEDFROMACTIONBUTTON")
    If (StrComp(sRunFromTM030, "TRUE") = 0) Then
        blnRunFromTM030 = True
    Else
        blnRunFromTM030 = False
    End If
    
    'BMIDS748
    sEnquiryCreditCheck = m_objXmlAssist.GetAttributeValue(vobjRequestNode, "REQUEST", "ENQUIRYCREDITCHECK")
    If (StrComp(sEnquiryCreditCheck, "TRUE") = 0) Then
        iEnquiryCreditCheck = 1
    Else
        iEnquiryCreditCheck = 0
    End If
    
    
    '===========================================================================================
    ' Validate
    '===========================================================================================
    'SYS0755 Use the UnitID and UserID nodes if they exist.
    'Otherwise default to the Request node attribute values
    strUserID = m_objXmlAssist.GetElementText(vobjRequestNode, "USERID")
    If Len(Trim$(strUserID)) = 0 Then
        'Get the UserID attribute from the REQUEST node
        'SYS1174 UserID and UnitID no longer mandatory
        strUserID = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "USERID")
        If Len(Trim$(strUserID)) = 0 Then
            'Get default user ID from global parameters
            strUserID = GetDefaultUserId()
        End If
    End If
    strUnitId = m_objXmlAssist.GetElementText(vobjRequestNode, "UNITID")
    If Len(Trim$(strUnitId)) = 0 Then
        'Get the UnitID attribute from the REQUEST node
        'SYS1174 UserID and UnitID no longer mandatory
        strUnitId = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "UNITID")
    End If
    
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, ".//APPLICATIONNUMBER")



    '===========================================================================================
    ' Execute Stored Procedure to formulate Experian request
    '===========================================================================================
   
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .Open
    End With

    intDBEngineType = objADOAssist.GetDBEngine()
    
    'BMIDS682 Are we Address Targeting?
    sAddressTargetReq = m_objXmlAssist.GetAttributeValue(vobjRequestNode, "REQUEST", "ADDRESSTARGETREQ")
        
    If (StrComp(sAddressTargetReq, "TRUE") = 0) Then
        'If Address Targeting, everything required has been returned in the request
        strRequest = MakeUpAddressTargetRequest(vobjRequestNode)
        
        'Need to get the saved params required for the save stored proc
        Set xmlReqHeaderElem = vobjRequestNode.selectSingleNode("REQHEADER")
        
        szCustomerNumberOtherN = xmlReqHeaderElem.getAttribute("CUSTNUMBEROTHER")
        szCustomerNumberSameP = xmlReqHeaderElem.getAttribute("CUSTNUMBERSAME")
        nCustomerVersionNumberOtherN = xmlReqHeaderElem.getAttribute("CUSTVERSIONOTHER")
        nCustomerVersionNumberSameP = xmlReqHeaderElem.getAttribute("CUSTVERSIONSAME")
        nAppFactFindNumber = xmlReqHeaderElem.getAttribute("FACTFINDNUMBER")
    
    Else
        Set adoCommand = New ADODB.Command
        With adoCommand
       
            'INR SYS4862 added param1 name
            Set adoParameter = adoCommand.CreateParameter("szUserId", adBSTR, adParamInput, Len(strUserID), strUserID)
            adoCommand.Parameters.Append adoParameter
            Set adoParameter = adoCommand.CreateParameter("szUnitId", adBSTR, adParamInput, Len(strUnitId), strUnitId)
            adoCommand.Parameters.Append adoParameter
            Set adoParameter = adoCommand.CreateParameter("szAppNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
            adoCommand.Parameters.Append adoParameter
            'HMA BMIDS748 Add EnquiryCreditCheck parameter
            Set adoParameter = adoCommand.CreateParameter("iEnquiry", adInteger, adParamInput, Len(iEnquiryCreditCheck), iEnquiryCreditCheck)
            adoCommand.Parameters.Append adoParameter
            
            Select Case intDBEngineType
            Case SQLServer
                Set prmFactFind = .CreateParameter("m_nAppFactFindNumber", adInteger, adParamOutput, , nAppFactFindNumber)
                .Parameters.Append prmFactFind
                Set prmCustNumSame = .CreateParameter("m_szCustomerNumberSameP", adBSTR, adParamOutput, 12, szCustomerNumberSameP)
                .Parameters.Append prmCustNumSame
                Set prmCustVerSame = .CreateParameter("m_nCustomerVersionNumberSameP", adInteger, adParamOutput, , nCustomerVersionNumberSameP)
                .Parameters.Append prmCustVerSame
                Set prmCustNumOther = .CreateParameter("m_szCustomerNumberOtherN", adBSTR, adParamOutput, 12, szCustomerNumberOtherN)
                .Parameters.Append prmCustNumOther
                Set prmCustVerOther = .CreateParameter("m_nCustomerVersionNumberOtherN", adInteger, adParamOutput, , nCustomerVersionNumberOtherN)
                .Parameters.Append prmCustVerOther
            
                strSQL = "USP_GetDelphi96DataRequest"
                .CommandType = adCmdStoredProc
        
            Case Oracle
                strSQL = "{call GetDelphi96DataRequest(?,?,?," _
                    & "{resultset 10," _
                    & "tableszHeader," _
                    & "tableszNAM1," _
                    & "tableszAUK1," _
                    & "tableszCCN1," _
                    & "tableszAP01," _
                    & "tableszAP42," _
                    & "tableszAM01," _
                    & "tableszAM02," _
                    & "tableszAM03," _
                    & "tableszAM04" _
                    & "})}"
                .CommandType = adCmdText
            End Select
            
            'Execute the Stored procedure
            .CommandText = strSQL
            Set .ActiveConnection = adoConnection
            Set adoRecSet = adoCommand.Execute
        End With
        
        If Not objADOAssist.GetValidRecordset(adoRecSet) Then
            ' raise application error to be interpreted by calling object
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
        End If
        '===========================================================================================
        ' Process the response from the Stored Procedure
        '===========================================================================================
        'INR SYS4862
        If prmFactFind.Value > 0 Then
            nAppFactFindNumber = prmFactFind.Value
        End If
        If Len(prmCustNumSame.Value) > 0 Then
            szCustomerNumberSameP = Trim$(prmCustNumSame.Value)
        End If
        If prmCustVerSame.Value > 0 Then
            nCustomerVersionNumberSameP = prmCustVerSame.Value
        End If
        If Len(prmCustNumOther.Value) > 0 Then
            szCustomerNumberOtherN = prmCustNumOther.Value
        End If
        If prmCustVerOther.Value > 0 Then
            nCustomerVersionNumberOtherN = prmCustVerOther.Value
        End If
    
        Set adoCommand.ActiveConnection = Nothing
        Set adoCommand = Nothing
        If adoRecSet.EOF Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
        End If
    
        'INR SYS4862
        ' make the request from the returned record set
        strRequest = vbNullString
        Select Case intDBEngineType
        Case SQLServer
            'BMIDS933 GHun
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszHeader"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszNAM1"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAUK1"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszCCN1"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAP01"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAP42"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAM01"
            'BMIDS744
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszTPD1"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAM02"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAM03"
            'AppendToStrFromSQLSvrRecSet strRequest, adoRecSet, "tableszAM04"
            AppendToStrFromSQLSvrRecSet strRequest, adoRecSet
            'BMIDS933 End
        Case Oracle
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszHeader"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszNAM1"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAUK1"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszCCN1"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAP01"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAP42"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM01"
            'BMIDS744
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszTPD1"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM02"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM03"
            AppendToStrFromRecSet strRequest, adoRecSet, "tableszAM04"
        End Select
        
        'Close recordset
        If Not adoRecSet Is Nothing Then
            If adoRecSet.State = adStateOpen Then
                    adoRecSet.Close
            End If
        End If
    End If
    
    '===========================================================================================
    ' Make call to Experian
    '===========================================================================================
    
'    'Call Experian through the BM Translation Bridge
    'BMIDS00605 MDC 08/10/2002
    'Set objExperianInterface = m_objContext.CreateInstance("Experian.Search")
    Set objExperianInterface = m_objContext.CreateInstance("BMTranslationBridge.ExperianService")
    'BMIDS00605 MDC 08/10/2002 - End
    
    'BMIDS682 Need to be able to do auto processing here
    Do While bNoFurtherCalls = False
    
    'BMIDS682 Move logging so we can see what is sent where auto send
    If blnLog Then
        Call LogToFile(strRequest, strApplicationNumber, ltRequest)
    End If
   
        #If DEBUGDATA Then
    
            'DEBUG - Test Data
    
            'Single
    '         strEXPReturn = "cZiICPLKE1E1CLU62 0301423700000KP      020902151219320TL    RETURN                  N                    09Y3200000596  AAAREH7F8W002FAD1007401C059MRS J TUBERFIELD##5,TRELAWNEY CLOSE#WORCESTER#WORCS#WR24HG##BUR10509{00{0{A51F0GA{A51F0G{00{N00{019A00D19A00A{{N{00{{{{CA{C{{{00{0{B27C1I{{B27C1I{00{N00{209E22A09E00A{{N{00{{{{{{{{{A0{0{0{{{05D05F01E06C07{05D07C00A04F06D06B06B03C02B09A00I05D00{06A00G18D112D2D550{0{550{0{550{0{441WR2 4HG       N00{03H{{{01A0{0A0A0A0A{NN{{{{{{{00{0{{ {{{{    1   NTNN Ã E ?V ä      ?  N{{{{NNN{{0{00{{11H0{00{NN97NB1 {00{0{0{{A{{  0{1  ANT00{0{0H0H{                    {{05HBA7    0{  {{{{ {     Ô   6m??ÙÊ5qNNNNNNN000{      F120780{      F12078000G  0{00{00EA{11 51068B00000{00000{#"
    
            'Joint
            strExpReturn = "byRKCPLKE1E1CLU62 0301423700000KP      020902150705610TL    RETURN                  N                    09Y3200000796  AAARCJ3WT2004FAD1010201C087MR J LEWIS#MRS B LEWIS#GREENAWAY HOUSE#ROSE LANE,WHEATHAMPSTEAD#ST ALBANS,HERTS#AL48RA##FAD1008811P073MR J LEWIS##46,KINGS COURT SOUTH#CHELSEA MANOR GARDENS#LONDON SW3#SW35EG##FAD10084213069MR J LEWIS##CAMBRIDGE MANSIONS#23,CAMBRIDGE ROAD#LONDON SW11#SW114RU##BUR10509{00{0{{00{0{IC{00{0{{00{N00{126F10I26F99I{{0F99I{{{{A{{{{{00{0{A01F6B{{A01F6B{00{N00{000A00A00A00A{{N{00{{{{{{{{{{0{0{0{{{04G09H05H06B09C05C05G06C08B03G06F06G07B03G01A07D04C00{08{00{12{112B2B230B0{111D1D550{0{241AL4 8RA       N00{05E{{F99H0{0A0C0A0ABNN{{{{{{{00{0{{ {{{{    10  NTNN$?ð ^ ?K     ? ?      N{{{{NNN{{1C00AE10C0{00{NN97NB1 {00{0{0A{{{{  0{10  NT00{0{0B0B{                    {{05EIA3    0{  {{{{ {  ô?(??ÙCúNNNNNNN110{      C090344A060344D09034400A  0{00{01{AD11 51098E00000{00000{#"
        
        #Else
            strExpReturn = objExperianInterface.Search(strRequest)
        #End If
        
        If blnLog Then
            Call LogToFile(strExpReturn, strApplicationNumber, ltResponse)
        End If
        
        '===========================================================================================
        ' Decode Response
        '===========================================================================================
    '    If Not DecodeEXPDelphiReturn(strEXPReturn, strErrorMessage, strHeader, strDEC1, strDelphi96) Then
    '        'BMIDS00521 MDC 27/09/2002
    '        'm_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 282, strErrorMessage
    '        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeCreditCheckError, strErrorMessage
    '        'BMIDS00521 MDC 27/09/2002 - End
    '    End If
        'BMIDS744 New TPD1 block returned
        'BMIDS682 Attach any AUK1 and/or BUK1 block info that has beeen returned
        Set xmlTempResponseNode = DecodeEXPDelphiReturn(strExpReturn, strErrorMessage, _
                                    strHeader, strDEC1, strDelphi96, strTPD1, strBUR5, blnRunFromTM030)
       
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        strAddressTargeting = m_objXmlAssist.GetNodeValue(xmlTempResponseNode, ".//ADDRESSTARGETING", True, True)
        
        'BMIDS730 Need to see if we can do an auto return on a BUK1
        'regardless of where we are calling from.
        'Keep the original request the first time in.
        If (Len(strOrigRequest) = 0) Then
            strOrigRequest = strRequest
        End If
        
        nReturnedAUKPos = InStr(1, strExpReturn, "AUK1", 1)
        nReturnedBUKPos = InStr(1, strExpReturn, "BUK1", 1)
        
        'If we only have a BUK1 returned and everything is either assumed or
        'enhanced we want to make up an AUK1 and CCN1 & send it straight back again
        If ((nReturnedBUKPos > 0) And (nReturnedAUKPos = 0)) Then
        
            'BMIDS730 May Need CCN1's from vobjRequestNode request, otherwise if we go
            'from AUK1's returned initially to BUK1's returned theCCN1's won't exist
            strRequest = AddressTargetRequestFromString(strOrigRequest, _
                                            strExpReturn, vobjRequestNode)
            
            If Len(strRequest) > 0 Then
                'we can send it straight back
                bNoFurtherCalls = False
            Else
                'Haven't got just BUK1s with assumed or enhanced
                bNoFurtherCalls = True
            End If
        Else
            'BMIDS682 No auto return, so leave the loop
            bNoFurtherCalls = True
        End If
        
        'BMIDS682 If Manual Address Targeting is Required, No results to save at this time
        'Need to save our requestheader and associated CCN1 blocks for the
        'new address targeting request
        If (StrComp(strAddressTargeting, "YES") = 0) Then
            Dim nAUK1Len As Integer
            Dim nAUKPos As Integer
            Dim strAUK1 As String
            
            If (bNoFurtherCalls) Then
     
                'BMIDS682 attach addresstargeting response after we've done the auto processing
                m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
                
                'BMIDS730 make it easier to handle the declared address in GUI
                Set xmlDeclareNode = xmlResponseDoc.createElement("DECLAREADDRESSLIST")
           
                'Need to know what the original addresses were that we were searching on
                'Possibly dealing with multiple AUK1 blocks
                nAUKPos = InStr(1, strOrigRequest, "AUK1", 1)
                Do While nAUKPos > 0
                    nAUKPos = InStr(nAUKPos, strOrigRequest, "AUK1", 1)
                
                    If nAUKPos > 0 Then
                            
                        nAUK1Len = Mid$(strOrigRequest, nAUKPos + 4, 4)
                        strAUK1 = Mid$(strOrigRequest, nAUKPos, nAUK1Len)
                        
                        Set xmlTempResponseNode = MakeAUK1(strAUK1, nAUK1Len, "DECLAREADDRESS")
                        'ADD OUR RETURNED NODE TO THE RESPONSE
                        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                        'BMIDS730 make it easier to handle the declared address in GUI
                        xmlDeclareNode.appendChild xmlTempResponseNode.selectSingleNode("DECLAREADDRESS").cloneNode(True)
                        nAUKPos = nAUKPos + nAUK1Len
                        
                    End If
                Loop
                
                'BMIDS730 make it easier to handle the declared address in GUI
                xmlResponseElem.appendChild xmlDeclareNode.cloneNode(True)
                
                'need to save szCustomerNumberOtherN,
                'szCustomerNumberSameP,nCustomerVersionNumberOtherN,nCustomerVersionNumberSameP
                
                SaveRequiredRequestElements strOrigRequest, _
                                                szCustomerNumberOtherN, _
                                                szCustomerNumberSameP, _
                                                nCustomerVersionNumberOtherN, _
                                                nCustomerVersionNumberSameP, _
                                                nAppFactFindNumber, _
                                                xmlResponseElem, _
                                                strExpReturn
            End If 'bNoFurtherCalls
                                            
        Else
            'BMIDS730 Not address targeting, but we may still be able to auto
            'send back to Experian
            If ((nReturnedBUKPos > 0) And (bNoFurtherCalls = False)) Then
                'Do nothing. Not Address targeting, but we are auto sending back
                'assumed or enhanced BUK1 data

            ElseIf (((nReturnedBUKPos > 0) And (bNoFurtherCalls = True)) _
                                            Or (nReturnedAUKPos > 0)) Then
                'One or more of the Addresses have been Unmatched (BUK1's) and we can't
                'do an auto send back to Experian,or the Address has Multiple Matches
                '(AUK1's)and this was not run from TM030, so we need to raise
                'an error.
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7018

            Else
                'Not Address targeting continue with normal processing
                bNoFurtherCalls = True
                'BMIDS682 attach addresstargeting response after we've done the auto processing
                m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
            End If
            
        End If 'We are Address targeting
         
    Loop
      
    'Not Address Targeting
    If (StrComp(strAddressTargeting, "YES") <> 0) Then
        '===========================================================================================
        ' Save results to database
        '===========================================================================================
       
        ' Generate the GUID
        strGuid = CreateGUID()
        Set adoCommand = New ADODB.Command
        
        'INR SYS4862 Cater for SQL Server NOTE Oracle still a package.
        Select Case intDBEngineType
        Case SQLServer
            ' SQL Server -----------------------------------------------------------
            Set adoParameter = adoCommand.CreateParameter("m_szUserId", adBSTR, adParamInput, Len(strUserID), strUserID)
            adoCommand.Parameters.Append adoParameter
            
            Set adoParameter = adoCommand.CreateParameter("m_szAppNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
            adoCommand.Parameters.Append adoParameter
            
            Set adoParameter = adoCommand.CreateParameter("m_nAppFactFindNumber", adBSTR, adParamInput, Len(nAppFactFindNumber), nAppFactFindNumber)
            adoCommand.Parameters.Append adoParameter
            
            If (Len(szCustomerNumberOtherN) > 0) Then
                Set adoParameter = adoCommand.CreateParameter("m_szCustomerNumberOtherN", adBSTR, adParamInput, Len(szCustomerNumberOtherN), szCustomerNumberOtherN)
            Else
                Set adoParameter = adoCommand.CreateParameter("m_szCustomerNumberOtherN", adBSTR, adParamInput, Len(szCustomerNumberOtherN), Null)
            End If
            adoCommand.Parameters.Append adoParameter
            
            Set adoParameter = adoCommand.CreateParameter("m_szCustomerNumberSameP", adBSTR, adParamInput, Len(szCustomerNumberSameP), szCustomerNumberSameP)
            adoCommand.Parameters.Append adoParameter
            
            If (nCustomerVersionNumberOtherN > 0) Then
                Set adoParameter = adoCommand.CreateParameter("m_nCustomerVersionNumberOtherN", adInteger, adParamInput, , nCustomerVersionNumberOtherN)
            Else
                Set adoParameter = adoCommand.CreateParameter("m_nCustomerVersionNumberOtherN", adInteger, adParamInput, , Null)
            End If
            adoCommand.Parameters.Append adoParameter
            
            Set adoParameter = adoCommand.CreateParameter("m_nCustomerVersionNumberSameP", adInteger, adParamInput, , nCustomerVersionNumberSameP)
            adoCommand.Parameters.Append adoParameter
            
            Set adoParameter = adoCommand.CreateParameter("szCreditCheckGuid", adBinary, adParamInput, 16)
            adoCommand.Parameters.Append adoParameter
            intNumParams = adoCommand.Parameters.Count - 1
            adoCommand(intNumParams).Value = objSqlAssist.GuidStringToByteArray(strGuid)
        
            strSQL = "USP_SetDelphi96DataReply"
            adoCommand.CommandType = adCmdStoredProc
    
        Case Oracle
            ' Oracle ---------------------------------------------------------------
            adoCommand.CommandType = adCmdText
            #If GENERIC_SQL Then
                strSQL = "{call sp_Experian_gs.SetDelphi96DataReply(?, ?, ?, ?)}"
            #Else
                strSQL = "{call sp_Experian.SetDelphi96DataReply(?, ?, ?, ?)}"
            #End If
            
            ' ASt 08/03/01 CC012 Pass guid as binary
            #If GENERIC_SQL Then
                Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamInput, 16)
                adoCommand.Parameters.Append adoParameter
                
                adoCommand(0) = objSqlAssist.GuidStringToByteArray(strGuid)
            #Else
                Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strGuid), strGuid)
                adoCommand.Parameters.Append adoParameter
            #End If
    
        End Select
        
        Set objSqlAssist = Nothing
        adoCommand.CommandText = strSQL
           
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strHeader), strHeader)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDEC1), strDEC1)
        adoCommand.Parameters.Append adoParameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strDelphi96), strDelphi96)
        adoCommand.Parameters.Append adoParameter
        'BMIDS744 New TPD1 parameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strTPD1), strTPD1)
        adoCommand.Parameters.Append adoParameter
        'JLD BMIDS980 New BUR5 parameter
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strBUR5), strBUR5)
        adoCommand.Parameters.Append adoParameter
        
        Set adoCommand.ActiveConnection = adoConnection
    
        'Execute the Stored Procedure
        adoCommand.Execute , , adExecuteNoRecords
        
    End If 'Of No Address Targeting Required

    'Return response
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set IExperianDO_RunCreditCheck = xmlResponseElem.cloneNode(True)

RunCreditCheckExit:
    '===========================================================================================
    ' Clean up
    '===========================================================================================
    
    'Disconnect command
    If Not adoCommand Is Nothing Then
        Set adoCommand.ActiveConnection = Nothing
    End If
    
    ' Close the recordset if still open
    If Not adoRecSet Is Nothing Then
        If adoRecSet.State = adStateOpen Then
            adoRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    
    'Dereference all objects
    Set xmlResponseDoc = Nothing
    Set xmlTempDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTempElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set prmFactFind = Nothing
    Set prmCustNumSame = Nothing
    Set prmCustVerSame = Nothing
    Set prmCustNumOther = Nothing
    Set prmCustVerOther = Nothing
    Set adoRecSet = Nothing
    Set objExperianInterface = Nothing
    Set objADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set xmlTempResponseNode = Nothing
'BMIDS682
    Set xmlReqHeaderElem = Nothing
    Set xmlCaseTaskNode = Nothing
'BMIDS730
    Set xmlDeclareNode = Nothing

    If Err.Number <> 0 Then
        'Process error
        m_objErrAssist.AddToErrSource cstrFunctionName
        'BMIDS00521 MDC 27/09/2002 - Log all errors
        'If m_objErrAssist.IsSystemError = True Then
            m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
        'End If
        'BMIDS00521 MDC 27/09/2002 - End
        
        If Not m_objContext Is Nothing Then
            If m_objErrAssist.IsSystemError Then
                m_objContext.SetAbort
            Else
                m_objContext.SetComplete
            End If
        End If
        
        'Trap error raised by stored procedure and give more useful error message
        If InStr(Err.Description, MANDATORY_DATA_ERR) Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 280
        End If
        
        '   re-raise error for calling object to interpret as appropriate
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    End If

End Function

'BMIDS748
Private Function IExperianDO_RunUpgradeToFullAppCreditCheck(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:
'   Calls a stored procedure to formulate an experian request.
'   Calls Experian through the BM Translation Bridge.
' pass:
'   vobjRequestNode: XML Request Node
' return:
'   Response XML - Success = True/False
'------------------------------------------------------------------------------------------
    On Error GoTo RunUpgradeToFullAppCreditCheckExit
    
    Const strFunctionName As String = "IExperianDO_RunUpgradeToFullAppCreditCheck"
    
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecSet As ADODB.Recordset
    Dim objExperianInterface As Object
    
    Dim objADOAssist As ADOAssist
    Dim objSqlAssist As SQLAssist
    
    Dim strUserID As String
    Dim strUnitId As String
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strSQL As String
    Dim strRequest As String
    Dim blnLog As Boolean
    
    Dim strExpReturn As String
    
   'Constants
    Const cstrFunctionName As String = "IExperianDO_RunUpgradeToFullAppCreditCheck"
    Const MANDATORY_DATA_ERR As String = "ORA-20003"  'Error raised by Experian stored procedure if mandatory data is missing

    '===========================================================================================
    'Initialise
    '===========================================================================================

    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set objSqlAssist = New SQLAssist
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlResponseElem
    
    blnLog = CBool(m_objXmlAssist.GetAttributeValue(vobjRequestNode, "REQUEST", "LOG"))
    
    '===========================================================================================
    ' Validate
    '===========================================================================================
    'Use the UnitID and UserID nodes if they exist.
    'Otherwise default to the Request node attribute values
    strUserID = m_objXmlAssist.GetElementText(vobjRequestNode, "USERID")
    If Len(Trim$(strUserID)) = 0 Then
        'Get the UserID attribute from the REQUEST node
        strUserID = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "USERID")
        If Len(Trim$(strUserID)) = 0 Then
            'Get default user ID from global parameters
            strUserID = GetDefaultUserId()
        End If
    End If
    strUnitId = m_objXmlAssist.GetElementText(vobjRequestNode, "UNITID")
    If Len(Trim$(strUnitId)) = 0 Then
        'Get the UnitID attribute from the REQUEST node
        strUnitId = m_objXmlAssist.GetAttributeValue(vobjRequestNode.ownerDocument.documentElement, "REQUEST", "UNITID")
    End If
    
    Set xmlApplicationNode = m_objXmlAssist.GetMandatoryNode(vobjRequestNode, "APPLICATION")
   
    strApplicationNumber = m_objXmlAssist.GetMandatoryAttribute(xmlApplicationNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetMandatoryAttribute(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")

    '===========================================================================================
    ' Execute Stored Procedure to formulate Experian request
    '===========================================================================================
   
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    
    adoConnection.ConnectionString = objADOAssist.GetConnStr
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open

    Set adoCommand.ActiveConnection = adoConnection
    
    Set adoParameter = adoCommand.CreateParameter("szHeader", adBSTR, adParamOutput, 133)
    adoCommand.Parameters.Append adoParameter
 
    Set adoParameter = adoCommand.CreateParameter("szUserId", adBSTR, adParamInput, Len(strUserID), strUserID)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter("szAppNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter("szUnitId", adBSTR, adParamInput, Len(strUnitId), strUnitId)
    adoCommand.Parameters.Append adoParameter
               
    strSQL = "USP_RequestUpgrade"
    adoCommand.CommandType = adCmdStoredProc
            
    'Execute the Stored procedure
    adoCommand.CommandText = strSQL
    adoCommand.Execute
        
    ' Raise errors if any while executing the stored procedure
    If adoConnection.Errors.Count <> 0 Then
        Err.Raise adoConnection.Errors.Item(0).Number, _
                  adoConnection.Errors.Item(0).Source, _
                  adoConnection.Errors.Item(0).Description
    End If

    ' Request was returned from Stored Procedure
    strRequest = adoCommand(0)
     
    '===========================================================================================
    ' Make call to Experian
    '===========================================================================================
    
    'Call Experian through the BM Translation Bridge
    Set objExperianInterface = m_objContext.CreateInstance("BMTranslationBridge.ExperianService")
    
    If blnLog Then
        Call LogToFile(strRequest, strApplicationNumber, ltRequest)
    End If
   
    strExpReturn = objExperianInterface.Search(strRequest)
  
    If blnLog Then
        Call LogToFile(strExpReturn, strApplicationNumber, ltResponse)
    End If
        
    'Examine the reply from Experian and report any errors returned.
    Set xmlTempResponseNode = DecodeUpgradeReturn(strExpReturn, strApplicationNumber, strApplicationFactFindNumber, strUserID)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    'Return response
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set IExperianDO_RunUpgradeToFullAppCreditCheck = xmlResponseElem.cloneNode(True)

RunUpgradeToFullAppCreditCheckExit:
    '===========================================================================================
    ' Clean up
    '===========================================================================================
    
    'Disconnect command
    If Not adoCommand Is Nothing Then
        Set adoCommand.ActiveConnection = Nothing
    End If
    
    ' Close the recordset if still open
    If Not adoRecSet Is Nothing Then
        If adoRecSet.State = adStateOpen Then
            adoRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    
    'Dereference all objects
    Set xmlResponseDoc = Nothing
    Set xmlTempDoc = Nothing
    Set xmlResponseElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set adoRecSet = Nothing
    Set objExperianInterface = Nothing
    Set objADOAssist = Nothing
    Set objSqlAssist = Nothing

    If Err.Number <> 0 Then
        'Process error
        m_objErrAssist.AddToErrSource cstrFunctionName
            m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
        
        If Not m_objContext Is Nothing Then
            If m_objErrAssist.IsSystemError Then
                m_objContext.SetAbort
            Else
                m_objContext.SetComplete
            End If
        End If
        
        'Trap error raised by stored procedure and give more useful error message
        If InStr(Err.Description, MANDATORY_DATA_ERR) Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 280
        End If
        
        '   re-raise error for calling object to interpret as appropriate
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    End If

End Function


Private Function IExperianDO_RunFullBureau(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

'BMIDS00336 MDC 14/08/2002

On Error GoTo FullBureauExit

Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlTempDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlRequestElem As IXMLDOMElement
Dim xmlTempElem As IXMLDOMElement

Dim objADOAssist As ADOAssist
Dim objSqlAssist As SQLAssist

Dim objExperianInterface As Object

Dim adoCommand As ADODB.Command
Dim adoConnection As ADODB.Connection
Dim adoParameter As ADODB.Parameter

Dim strErrorMessage As String
Dim strFBReply As String
Dim strFBRequest As String
Dim strGuid As String
Dim strSQL As String
Dim strEXPFBLReturn As String
Dim strApplicationNumber As String
Dim blnLog As Boolean

'Constants
Const cstrFunctionName As String = "IExperianDO_RunFullBureau"


    '=======================================================================================
    ' Initialise
    '=======================================================================================
    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    
    'Initialise objects
    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set objSqlAssist = New SQLAssist
    
    blnLog = CBool(m_objXmlAssist.GetAttributeValue(vXMLRequest, "REQUEST", "LOG"))
    
    '=======================================================================================
    ' Validate Request
    '=======================================================================================
    strApplicationNumber = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//APPLICATIONNUMBER", True, True)
    strFBRequest = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//FBREQUEST", True, True)
'   strFBRequest = "TASKCPLKNRNRCKLU620101ERROR5678JB      010116105443000T     FBL     NEW                                  31 3200000013  2532      000"
    strGuid = m_objXmlAssist.GetNodeValue(vXMLRequest, ".//CREDITCHECKGUID", True, True)
   
    If blnLog Then
        Call LogToFile(strFBRequest, strApplicationNumber, ltRequest)
    End If
    
    '===========================================================================================
    ' Make call to Experian
    '===========================================================================================
    
    #If DEBUGDATA Then
        'Debug data
        Dim xmlDoc As New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        xmlDoc.load "C:\FULLBUREAURESPONSE.XML"
        strEXPFBLReturn = xmlDoc.selectSingleNode("XML").Text
        Set xmlDoc = Nothing
    #Else
        'BMIDS00605 MDC 08/10/2002
        'Set objExperianInterface = m_objContext.CreateInstance("Experian.Search")
        Set objExperianInterface = m_objContext.CreateInstance("BMTranslationBridge.ExperianService")
        'BMIDS00605 MDC 08/10/2002 - End
        strEXPFBLReturn = objExperianInterface.Search(strFBRequest)
    #End If
    
    If blnLog Then
        Call LogToFile(strEXPFBLReturn, strApplicationNumber, ltResponse)
    End If
    
    '===========================================================================================
    ' Decode Response
    '===========================================================================================
    If Not DecodeEXPFBLReturn(strEXPFBLReturn, strErrorMessage, strFBReply) Then
        'BMIDS00521 MDC 27/09/2002
        'm_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, 282, strErrorMessage
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeCreditCheckError, strErrorMessage
        'BMIDS00521 MDC 27/09/2002 - End
    End If
    
'    strFBReply = "EQFXCPLKNRNRCKLU620101ERROR5678JB      990318100802000T     RETURN                  N                    32Y3200001372  2535      005BE01007301           66 CARNBRAE AVENUE##BELFAST#COUNTY ANTRIM#BT8 6NH#@#BE02047702         00#9800###CHRISTOPHER#J#KEENAN#####@       00#9800###MAUREEN##KEENAN#####@       00#9798###SEAMUS#G#MCEVOY#####@       00#9798###JOANNE##MCEVOY#####@       00#9397###PAULINE##WOODS#####@       00#9397###ROGER#F#WOODS#####@       00#8693###BRUCE#C#MCCONNELL#####@       00#8693###KAREN#C#MCCONNELL#####@       00#8486###JAMES##MURPHY#####@       00#8486###MICHELE##MURPHY#####@       00#8384###KIMBERLY##GARRETT#####@       00#8384##MS#ROBERTA##GARRETT#####@#BF01016303               260400         4#LLOYDS UDT LTD##MR#SEAMUS##MCEVOY##########@             121000         4#PRUDENTIAL BANKI##MR#SEAMUS##MCEVOY##########@"
    
    
    '===========================================================================================
    ' Save results to database
    '===========================================================================================
    
    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .Open
    End With

    'INR SYS4862 Cater for SQL Server NOTE Oracle still a package.
    Select Case objADOAssist.GetDBEngine
    Case SQLServer
        'SQL Server
        strSQL = "usp_SetFullBureauReply"
        With adoCommand
            .CommandType = adCmdStoredProc
            Set adoParameter = .CreateParameter("@szCreditCheckGuid", adBinary, adParamInput, 16)
            .Parameters.Append adoParameter
            .Parameters(0).Value = objSqlAssist.GuidStringToByteArray(strGuid)
        End With
        
    Case Oracle
        'Oracle
        strSQL = "{call sp_FullBureau.SetFullBureauReply(?, ?)}"
        With adoCommand
            .CommandType = adCmdText
            Set adoParameter = .CreateParameter("@szCreditCheckGuid", adBSTR, adParamInput, Len(strGuid), strGuid)
            .Parameters.Append adoParameter
        End With
    End Select

    strFBReply = Trim$(strFBReply)
    Set adoParameter = adoCommand.CreateParameter("@szFBreply", adChar, adParamInput, Len(strFBReply), strFBReply)
    
    With adoCommand
        .Parameters.Append adoParameter
        .CommandText = strSQL
        Set .ActiveConnection = adoConnection
        'Execute Stored procedure
        .Execute , , adExecuteNoRecords
        Set .ActiveConnection = Nothing
    End With
    
    
    'Return response
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set IExperianDO_RunFullBureau = xmlResponseElem.cloneNode(True)
    
    
FullBureauExit:
    '===========================================================================================
    ' Clean up
    '===========================================================================================

    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If

    'Dereference objects
    Set objExperianInterface = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    Set adoParameter = Nothing
    Set objADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTempDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlTempElem = Nothing
    
    If Err.Number <> 0 Then
        m_objErrAssist.AddToErrSource (cstrFunctionName)
        'BMIDS00521 MDC 27/09/2002 - Log all errors
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
        If m_objErrAssist.IsSystemError = True Then
'            m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
        'BMIDS00521 MDC 27/09/2002 - End
        
        '   re-raise error for calling object to interpret as appropriate
        Err.Raise Err.Number, Err.Source, Err.Description
    
    Else
        m_objContext.SetComplete
    End If

End Function

'INR SYS4862 New function
'INR BMIDS744 Need to extract the TPD string
Private Function DecodeEXPDelphiReturn(ByVal strExpReturn As String, _
                                        ByRef strErrorMessage As String, _
                                        ByRef strHeader As String, _
                                        ByRef strDEC1 As String, _
                                        ByRef strDelphi96 As String, _
                                        ByRef strTPD1 As String, _
                                        ByRef strBUR5 As String, _
                                        ByVal blnRunFromTM030 As Boolean) As IXMLDOMNode

On Error GoTo DecodeEXPDelphiReturnVbErr

Const strFunctionName As String = "DecodeEXPDelphiReturn"

Dim strErrText As String
Dim nOffsetERR1 As Integer
Dim nOffsetDEC1 As Integer
Dim nOffsetBUR1 As Integer
Dim nOffset As Integer
Dim strErr As String
Dim bSuccessIndicator As Boolean
Dim bHeaderPlusBlocksReturned As Boolean
Dim nMsgSize As Integer
Dim nCounter As Integer
Dim strCompare As String
Dim nSizeErrorMessage As Integer
Dim nOffsetErrorMessageStart As Integer
Dim nOffsetErrorMessageEnd As Integer
Dim nSizeDEC1 As Integer
Dim nOffsetDelphi96 As Integer
Dim nBUKPos As Integer
Dim nAUKPos As Integer
Dim strSuccessFlag As String
'BMIDS682
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
Dim blnAddressTargeting As Boolean
'BMIDS744
Dim nTPDPos As Integer
Dim nBUR5Pos As Integer 'JLD BMIDS980

'Initialise response
Set xmlResponseDoc = New FreeThreadedDOMDocument40
xmlResponseDoc.validateOnParse = False
xmlResponseDoc.setProperty "NewParser", True
Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
xmlResponseDoc.appendChild xmlResponseElem
xmlResponseElem.setAttribute "TYPE", "SUCCESS"

'BMIDS00990 MDC 19/11/2002
Const cintBUR1PrefixOffset = 8
'BMIDS00990 MDC 19/11/2002 - End
'BMIDS682
blnAddressTargeting = False

    If (Len(strExpReturn) > 0) Then
        
        'Successful, so copy return the reply
    
        nOffsetERR1 = 0
        nOffsetDEC1 = 0
        nOffsetBUR1 = 0
        nOffset = 0
        bSuccessIndicator = False
        bHeaderPlusBlocksReturned = False
        nMsgSize = Len(strExpReturn)
        
        ' ... search for the blocks (after the header)
        nOffset = nSizeHeaderplus1
                        
        strErr = Mid$(strExpReturn, nSizeHeaderplus1, 4)
        If (StrComp(strErr, "ERR1") = 0) Then
            nOffsetERR1 = nOffset
        Else
                    
            'BMIDS00334 MDC 02/09/2002 - No DEC1 block
'            For nCounter = nOffset To nMsgSize
'                strCompare = Mid$(strEXPReturn, nCounter, 4)
'                If (StrComp(strCompare, "DEC1") = 0) Then
'                    nOffsetDEC1 = nCounter
'                    nCounter = nMsgSize
'                End If
'            Next nCounter
'
'            For nCounter = nOffset To nMsgSize
'                strCompare = Mid$(strEXPReturn, nCounter, 4)
'                If (StrComp(strCompare, "BUR1") = 0) Then
'                    nOffsetBUR1 = nCounter
'                    nCounter = nMsgSize
'                End If
'            Next nCounter
            nOffsetDEC1 = 0
            nOffsetBUR1 = InStr(nOffset, strExpReturn, "BUR1")
            
            'BMIDS00334 MDC 02/09/2002 - End
    
        End If
    
        '  ... now return the blocks or error messages (if not found)
        strHeader = Left$(strExpReturn, nSizeHeaderplus1)
    
        If Len(strHeader) > 0 Then
            If (nOffsetERR1 > 0) Then
                nOffsetErrorMessageStart = nOffsetERR1 + 16
                nOffsetErrorMessageEnd = nOffsetErrorMessageStart
    
                'BMIDS00521 MDC 27/09/2002 - End
                'For nCounter = nOffsetErrorMessageStart To (nMsgSize - nOffsetErrorMessageStart)
                For nCounter = nOffsetErrorMessageStart To nMsgSize
                'BMIDS00521 MDC 27/09/2002
                    strCompare = Mid$(strExpReturn, nCounter, 1)
                    If (StrComp(strCompare, "#") = 0) Then
                        nOffsetErrorMessageEnd = nCounter
                        nCounter = nMsgSize
                    End If
                Next nCounter
                        
                nSizeErrorMessage = nOffsetErrorMessageEnd - nOffsetErrorMessageStart
                strErrorMessage = Mid$(strExpReturn, nOffsetErrorMessageStart, nSizeErrorMessage)
                bHeaderPlusBlocksReturned = True
                
            ElseIf (nOffsetDEC1 > 0 And nOffsetBUR1 > 0) Then
            
                nSizeDEC1 = nOffsetBUR1 - nOffsetDEC1
                strDEC1 = Mid$(strExpReturn, nOffsetDEC1, nSizeDEC1)
                
                nOffsetDelphi96 = nOffsetBUR1 + cintBUR1PrefixOffset    'BMIDS00990 MDC 19/11/2002
                strDelphi96 = Mid$(strExpReturn, nOffsetDelphi96, nSizeDelphi96)
                bHeaderPlusBlocksReturned = True
            'BMIDS00336 MDC 03/09/2002
            ElseIf (nOffsetDEC1 = 0 And nOffsetBUR1 > 0) Then
                strDEC1 = vbNullString
                nOffsetDelphi96 = nOffsetBUR1 + cintBUR1PrefixOffset    'BMIDS00990 MDC 19/11/2002
                strDelphi96 = Mid$(strExpReturn, nOffsetDelphi96, nSizeDelphi96)
                bHeaderPlusBlocksReturned = True
            'BMIDS00336 MDC 03/09/2002 - End
            
            ElseIf ((nOffsetDEC1 = 0) And (nOffsetBUR1 = 0)) Then
                'need to check for AUK1/BUK1 addresstargeting blocks
                'Possibly dealing with multiple BUK1 blocks
                nBUKPos = InStr(1, strExpReturn, "BUK1", 1)
                Do While nBUKPos > 0
                    nBUKPos = InStr(nBUKPos, strExpReturn, "BUK1", 1)
                    If nBUKPos > 0 Then
                        'BMIDS682 If running from TM030 we want to do address targeting
                        If (blnRunFromTM030) Then
                            blnAddressTargeting = True

                            Dim nBUK1Len As Integer
                            Dim strBUK1 As String
                            nBUK1Len = Mid$(strExpReturn, nBUKPos + 4, 4)
                            strBUK1 = Mid$(strExpReturn, nBUKPos, nBUK1Len)
                            
                            If (StrComp(Right$(strBUK1, 1), "#") = 0) Then
                                Set xmlTempResponseNode = MakeBUK1(strBUK1, nBUK1Len)
                                'ADD OUR RETURNED NODE TO THE RESPONSE
                                m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
                                nBUKPos = nBUKPos + nBUK1Len
                            Else
                                strErrText = "Credit Check: BUK1 block indicates Address is Unmatched, but length of BUK1 block does not match stated Block Length"
                                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                            End If
                        Else
                            'BMIDS730 Don't throw an error here, the BUK1 may require
                            'automatic send back to Experian, leave the loop
                            nBUKPos = 0
                            blnAddressTargeting = False
                            strSuccessFlag = Mid$(strExpReturn, 108, 1)
                            If (StrComp(strSuccessFlag, "Y") = 0) Then
                                bSuccessIndicator = True
                            End If
'                            'Not running from TM030 need to raise an error
'                            strErrText = "Credit Check: One or more of the Addresses have been Unmatched"
'                            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                        End If
                    End If
                Loop
                  
                'Possibly dealing with multiple AUK1 blocks
                nAUKPos = InStr(1, strExpReturn, "AUK1", 1)
                Do While nAUKPos > 0
                    nAUKPos = InStr(nAUKPos, strExpReturn, "AUK1", 1)
                
                    If nAUKPos > 0 Then
                        'BMIDS682 If running from TM030 we want to do address targeting
                        If (blnRunFromTM030) Then
                            blnAddressTargeting = True
                            
                            Dim nAUK1Len As Integer
                            Dim strAUK1 As String
                            nAUK1Len = Mid$(strExpReturn, nAUKPos + 4, 4)
                            strAUK1 = Mid$(strExpReturn, nAUKPos, nAUK1Len)
                            If (StrComp(Right$(strAUK1, 1), "#") = 0) Then
                            
                                Set xmlTempResponseNode = MakeAUK1(strAUK1, nAUK1Len, "ADDRESSTARGET")
                                'ADD OUR RETURNED NODE TO THE RESPONSE
                                m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
                                nAUKPos = nAUKPos + nAUK1Len
                            Else
                                strErrText = "Credit Check: AUK1 block indicates Address has Multiple Matches, but length of AUK1 block does not match stated Block Length"
                                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                            End If
                        Else
                            'BMIDS730 Don't raise error here, leave the loop
                            nAUKPos = 0
                            blnAddressTargeting = False
                            strSuccessFlag = Mid$(strExpReturn, 108, 1)
                            If (StrComp(strSuccessFlag, "Y") = 0) Then
                                bSuccessIndicator = True
                            End If
'                            'Not running from TM030 need to raise an error
'                            strErrText = "Credit Check: Address has Multiple Matches"
'                            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                        End If
                    End If
                Loop
                 
                             
            'BMIDS00334 MDC 02/09/2002 - No DEC1 block
'            ElseIf (nOffsetDEC1 = 0) Then
'                strErrText = "Credit Check: DEC1 block was expected and not returned by Experian"
'                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
            'BMIDS00334 MDC 02/09/2002 - End
            ElseIf (nOffsetBUR1 = 0) Then
                strErrText = "Credit Check: BUR1 block was expected and not returned by Experian"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
            End If
            
            strSuccessFlag = Mid$(strExpReturn, 108, 1)
            If ((bHeaderPlusBlocksReturned = True) And (StrComp(strSuccessFlag, "Y") = 0)) Then
                bSuccessIndicator = True
            End If
            
            'BMIDS744 If there is no Error, need to check for a single TPD1 block
            If (nOffsetERR1 = 0) Then
                nTPDPos = InStr(1, strExpReturn, "TPD1", 1)
                If nTPDPos > 0 Then
                    Dim nTPD1Len As Integer
                    nTPD1Len = Mid$(strExpReturn, nTPDPos + 4, 4)
                    strTPD1 = Mid$(strExpReturn, nTPDPos, nTPD1Len)
                Else
                    strErrText = "Credit Check: TPD1 block was expected and not returned by Experian"
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    
                End If
            End If
            
            'BMIDS980 If there is no Error, need to check for a single BUR5 block
            If (nOffsetERR1 = 0) Then
                nBUR5Pos = InStr(1, strExpReturn, "BUR5", 1)
                If nBUR5Pos > 0 Then
                    Dim nBUR5Len As Integer
                    nBUR5Len = Mid$(strExpReturn, nBUR5Pos + 4, 4)
                    strBUR5 = Mid$(strExpReturn, nBUR5Pos, nBUR5Len)
                Else
                    strErrText = "Credit Check: BUR5 block was expected and not returned by Experian"
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                    'strBUR5 = "BUR50033011UU-00000NN1AA-20000IB#"  'TEST DATA
                End If
            End If
                    
        Else
            strErrText = "Experian: Header block could not be returned"
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    
        End If
        
    Else
                        
        strErrText = "Experian: The interface to Experian returned no data"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    
    End If
    
    'BMIDS682 Check for error other than address targeting
    If ((bSuccessIndicator = False) And (blnAddressTargeting = False)) Then
        strErrText = "Error returned from Experian: " & strErrorMessage
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    End If
                             
    Set xmlTempResponseNode = xmlResponseDoc.createElement("ADDRESSTARGETING")
    If (blnAddressTargeting = True) Then
        xmlTempResponseNode.Text = "YES"
    Else
        xmlTempResponseNode.Text = "NO"
    End If
    xmlResponseElem.appendChild xmlTempResponseNode.cloneNode(True)

    Set DecodeEXPDelphiReturn = xmlResponseElem
   
DecodeEXPDelphiReturnExit:

    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function
    
DecodeEXPDelphiReturnVbErr:

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Set DecodeEXPDelphiReturn = m_objErrAssist.CreateErrorResponseEx()
    
    '   go to clean-up section
    Resume DecodeEXPDelphiReturnExit

End Function
'INR SYS4862 New Function
Private Function DecodeEXPFBLReturn(ByVal strEXPFBLReturn As String, _
        ByRef strErrorMessage As String, ByRef strFBReply As String) As Boolean

On Error GoTo DecodeEXPFBLReturnVbErr

Const strFunctionName As String = "DecodeEXPFBLReturn"

Dim strErrText As String
Dim bSuccessIndicator As Boolean
Dim nOffsetERR1 As Integer
Dim nOffset As Integer
Dim strErr As String
Dim nMsgSize As Integer
Dim nCounter As Integer
Dim nSizeErrorMessage As Integer
Dim strCompare As String
Dim nOffsetErrorMessageStart As Integer
Dim nOffsetErrorMessageEnd As Integer
Dim strSuccessFlag As String

    bSuccessIndicator = False
    
    If (Len(strEXPFBLReturn) > 0) Then
        'Successful, so copy return the reply
        nOffsetERR1 = 0
        nOffset = 0
        nMsgSize = Len(strEXPFBLReturn)
        
        ' ... search for the Err block (after the header)
        nOffset = nSizeHeaderplus1
                        
        strErr = Mid$(strEXPFBLReturn, nSizeHeaderplus1, 4)
        If (StrComp(strErr, "ERR1") = 0) Then
            nOffsetERR1 = nOffset
            nOffsetErrorMessageStart = nOffsetERR1 + 16
            nOffsetErrorMessageEnd = nOffsetErrorMessageStart
    
            'BMIDS00521 MDC 27/09/2002 - End
            'For nCounter = nOffsetErrorMessageStart To (nMsgSize - nOffsetErrorMessageStart)
            For nCounter = nOffsetErrorMessageStart To nMsgSize
            'BMIDS00521 MDC 27/09/2002
                strCompare = Mid$(strEXPFBLReturn, nCounter, 1)
                If (StrComp(strCompare, "#") = 0) Then
                    nOffsetErrorMessageEnd = nCounter
                    nCounter = nMsgSize
                End If
            Next nCounter
                    
            nSizeErrorMessage = nOffsetErrorMessageEnd - nOffsetErrorMessageStart
            strErrorMessage = Mid$(strEXPFBLReturn, nOffsetErrorMessageStart, nSizeErrorMessage)
        
        Else
            strFBReply = strEXPFBLReturn
            strSuccessFlag = Mid$(strEXPFBLReturn, 108, 1)
            If (StrComp(strSuccessFlag, "Y") = 0) Then
                bSuccessIndicator = True
            End If
        End If
        
    Else
                        
        strErrText = "Experian: The FBL interface to Experian returned no data"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    
    End If

    DecodeEXPFBLReturn = bSuccessIndicator

    Exit Function
    
DecodeEXPFBLReturnVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
       
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'INR SYS4862 New Function
Private Sub AppendToStrFromSQLSvrRecSet(ByRef strInOut As String, ByRef adoRecSet As ADODB.Recordset) ', ByVal strItem As String)

On Error GoTo AppendToStrFromSQLSvrVbErr
    
    'BMIDS933 GHun
    Do While Not adoRecSet.EOF
        If Not (IsNull(adoRecSet.fields.Item(0).Value)) Then
            strInOut = strInOut & adoRecSet.fields.Item(0).Value
        End If
        adoRecSet.MoveNext
    Loop
    
    'Do While True
    '    If IsNull(adoRecSet.fields.Item(strItem).Value) Then
    '        Exit Do
    '    End If
    '    strInOut = strInOut & adoRecSet.fields.Item(strItem).Value
    '    adoRecSet.MoveNext
    '    If adoRecSet.EOF Then
    '        Exit Do
    '    End If
    'Loop
    
'NextTag:
    'Set adoRecSet = adoRecSet.NextRecordset
    
    Exit Sub

AppendToStrFromSQLSvrVbErr:
    Err.Raise Err.Number, Err.Source, Err.Description
    'could possibly be more recordsets of information
    'If Err.Number = 3021 Then
    '    Resume NextTag
    'End If
    'BMIDS933 End
        
End Sub

'BMIDS748  New function
Private Function DecodeUpgradeReturn(ByVal strExpReturn As String, _
                                     ByVal strApplicationNumber As String, _
                                     ByVal strApplicationFactFindNumber As String, _
                                     ByVal strUserID As String) As IXMLDOMNode

    On Error GoTo DecodeUpgradeReturnVbErr

    Const strFunctionName As String = "DecodeUpgradeReturn"

    Dim strErrText As String
    Dim nOffsetERR1 As Integer
    Dim nOffset As Integer
    Dim strErr As String
    Dim strHeader As String
    Dim strErrorMessage As String
    Dim bSuccessIndicator As Boolean
    Dim bErrorReturned As Boolean
    Dim nMsgSize As Integer
    Dim nCounter As Integer
    Dim strCompare As String
    Dim strResp As String
    Dim xmlRespDoc As FreeThreadedDOMDocument40
    Dim nSizeErrorMessage As Integer
    Dim nOffsetErrorMessageStart As Integer
    Dim nOffsetErrorMessageEnd As Integer
    Dim strSuccessFlag As String
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim strComboResponse As String
    Dim strMemoPadValueID As String
    Dim xmlMemoPadElement As IXMLDOMElement
    Dim xmlComboRespDoc As FreeThreadedDOMDocument40
    
    Dim xmlTempRequestElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlChildElement As IXMLDOMElement
    Dim objAppBO As ApplicationBO
    Dim objComboDO As ComboDO
    
    'Initialise response
    Set xmlComboRespDoc = New FreeThreadedDOMDocument40
    xmlComboRespDoc.validateOnParse = False
    xmlComboRespDoc.setProperty "NewParser", True
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlRespDoc = New FreeThreadedDOMDocument40
    xmlRespDoc.validateOnParse = False
    xmlRespDoc.setProperty "NewParser", True

    If (Len(strExpReturn) > 0) Then
        
        'Successful return from Experian - search for error message.
        nOffsetERR1 = 0
        nOffset = 0
        bSuccessIndicator = False
        bErrorReturned = False
        nMsgSize = Len(strExpReturn)
        
        ' Search for the error block (after the header)
        nOffset = nSizeHeaderplus1
                        
        strErr = Mid$(strExpReturn, nSizeHeaderplus1, 4)
        If (StrComp(strErr, "ERR1") = 0) Then
            nOffsetERR1 = nOffset
        End If
    
        strHeader = Left$(strExpReturn, nSizeHeaderplus1)
    
        If Len(strHeader) > 0 Then
            If (nOffsetERR1 > 0) Then
                nOffsetErrorMessageStart = nOffsetERR1 + 16
                nOffsetErrorMessageEnd = nOffsetErrorMessageStart
    
                For nCounter = nOffsetErrorMessageStart To nMsgSize
                    strCompare = Mid$(strExpReturn, nCounter, 1)
                    If (StrComp(strCompare, "#") = 0) Then
                        nOffsetErrorMessageEnd = nCounter
                        nCounter = nMsgSize
                    End If
                Next nCounter
                        
                nSizeErrorMessage = nOffsetErrorMessageEnd - nOffsetErrorMessageStart
                strErrorMessage = Mid$(strExpReturn, nOffsetErrorMessageStart, nSizeErrorMessage)
                bErrorReturned = True

            End If
            
            strSuccessFlag = Mid$(strExpReturn, 108, 1)
            If ((bErrorReturned = False) And (StrComp(strSuccessFlag, "Y") = 0)) Then
                bSuccessIndicator = True
            End If
 
        Else
            strErrText = "Experian: Header block could not be returned"
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    
        End If
        
    Else
                        
        strErrText = "Experian: The interface to Experian returned no data"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
    
    End If
    
    If (bSuccessIndicator = False) Then
        'An error has been reported by Experian.
        'Write this error to the Memo Pad
                
        Dim objIExperianNoTxBO As IExperianNoTxBO
    
        If m_objContext Is Nothing Then
            Set objIExperianNoTxBO = New ExperianNoTxBO
        Else
            Set objIExperianNoTxBO = m_objContext.CreateInstance(App.Title & ".ExperianNoTxBO")
        End If
        
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
                      
        Set xmlTempRequestElement = xmlDoc.createElement("REQUEST")
        Set xmlElement = xmlDoc.createElement("MEMOPAD")
        xmlTempRequestElement.appendChild xmlElement
                
        Set xmlChildElement = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlChildElement.Text = strApplicationNumber
        xmlElement.appendChild xmlChildElement
                
        Set xmlChildElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlChildElement.Text = strApplicationFactFindNumber
        xmlElement.appendChild xmlChildElement
                
        Set xmlChildElement = xmlDoc.createElement("ACTIONNAME")
        xmlChildElement.Text = "Experian Upgrade"
        xmlElement.appendChild xmlChildElement
                
        Set xmlChildElement = xmlDoc.createElement("MEMOENTRY")
        xmlChildElement.Text = strErrorMessage
        xmlElement.appendChild xmlChildElement
                            
        Set xmlChildElement = xmlDoc.createElement("SCREENDESCRIPTION")
        xmlChildElement.Text = " "
        xmlElement.appendChild xmlChildElement
                
        Set xmlChildElement = xmlDoc.createElement("SCREENREFERENCE")
        xmlChildElement.Text = " "
        xmlElement.appendChild xmlChildElement
                
        Set xmlChildElement = xmlDoc.createElement("USERID")
        xmlChildElement.Text = strUserID
        xmlElement.appendChild xmlChildElement
                
        Set xmlChildElement = xmlDoc.createElement("ENTRYTYPE")
        
        'EntryType = ValueID corresponding to General (Validation type "G")
        strComboResponse = objComboDO.GetComboValueId("MemoPadEntryType", "G")
        xmlComboRespDoc.loadXML (strComboResponse)
        Set xmlMemoPadElement = xmlComboRespDoc.selectSingleNode(".//VALUEID")
        If xmlMemoPadElement Is Nothing Then
            'RaiseError
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoComboTagValue, "ValueID for validation type 'G'  not found"
        Else
            xmlChildElement.Text = xmlMemoPadElement.Text
        End If
        
        xmlElement.appendChild xmlChildElement
        
        strResp = objIExperianNoTxBO.CreateMemoPad(xmlTempRequestElement.xml)
        xmlRespDoc.loadXML strResp
        m_objErrAssist.CheckResponse xmlRespDoc.documentElement, xmlElement, True
    
        'Raise the error
     
        strErrText = "Error returned from Experian: " & strErrorMessage
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
  
    End If
                             
    Set DecodeUpgradeReturn = xmlResponseElem
    
DecodeUpgradeReturnExit:

    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRespDoc = Nothing
    Set xmlTempRequestElement = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    Set xmlChildElement = Nothing
    Set objIExperianNoTxBO = Nothing
    Exit Function

DecodeUpgradeReturnVbErr:

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Set DecodeUpgradeReturn = m_objErrAssist.CreateErrorResponseEx()

    'Go to clean-up section
    Resume DecodeUpgradeReturnExit

End Function

'BMIDS00336 MDC 04/09/2002
Private Sub LogToFile(ByVal strText As String, ByVal strApplicationNumber As String, _
                    Optional ByVal intLogType As LOG_TYPE = ltRequest)

On Error GoTo LogToFileExit

Dim objFSO As Scripting.FileSystemObject

Dim intFileNo As Integer
Dim blnFileOpen As Boolean
Dim strLine As String

'BM0555 Read the Log File Path from the string in CreditCheckLogFilePath Global Parameter.
'Const cstrLogFolder As String = "C:\LOG"

Dim strLogFolder As String

Dim objGlobalParamDO As GlobalParameterDO
    
    If m_objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    strLogFolder = objGlobalParamDO.GetCurrentParameterByType("CreditCheckLogFilePath", "STRING")
    
    Set objFSO = New Scripting.FileSystemObject
    If Not objFSO.FolderExists(strLogFolder) Then
        objFSO.CreateFolder strLogFolder
    End If
    
    intFileNo = FreeFile
    Open strLogFolder & "\" & strApplicationNumber & ".TXT" For Append As #intFileNo
    blnFileOpen = True
    
    Print #intFileNo, "==================================================================="
    
    If intLogType = ltRequest Then
        strLine = "Data sent to"
    Else
        strLine = "Data received from"
    End If
    strLine = strLine & " Experian " & Format$(Now, "dd/mm/yyyy hh:mm:ss")
    Print #intFileNo, strLine
    Print #intFileNo, ""
    
    Print #intFileNo, "Buffer:"
    Print #intFileNo, strText
    Print #intFileNo, ""
    Print #intFileNo, "Buffer Length: " & CStr(Len(strText))
    Print #intFileNo, ""
    Print #intFileNo, ""
    
LogToFileExit:
    If blnFileOpen Then
        Close #intFileNo
    End If
    Set objFSO = Nothing
    Set objGlobalParamDO = Nothing   'BM0555
    
End Sub
'BMIDS00336 MDC 04/09/2002 - End

'INR BMIDS682 New Function
Private Function MakeBUK1(ByVal strBUK1 As String, ByVal nBUK1Len As Integer) As IXMLDOMNode

On Error GoTo MakeBUK1VbErr

Const strFunctionName As String = "MakeBUK1"

Dim strText As String
Dim nOffset As Integer
Dim nCurPos As Integer
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseElem As IXMLDOMElement

    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    Set xmlTempResponseElem = xmlResponseDoc.createElement("ADDRESSTARGET")
    
    xmlTempResponseElem.setAttribute "BLOCKTYPE", "BUK1"
    xmlTempResponseElem.setAttribute "BLOCKLENGTH", nBUK1Len
    
    'First items are all fixed length
    'Block Sequence Number
    strText = Mid$(strBUK1, 9, 2)
    xmlTempResponseElem.setAttribute "BLOCKSEQNUMBER", strText
   
    'Address Indicator
    strText = Mid$(strBUK1, 11, 1)
    xmlTempResponseElem.setAttribute "ADDRESSINDICATOR", strText
   
    'RMC
    strText = Mid$(strBUK1, 12, 7)
    xmlTempResponseElem.setAttribute "RMC", strText
   
    'Region
    strText = Mid$(strBUK1, 19, 1)
    xmlTempResponseElem.setAttribute "REGION", strText
   
    'Experian Address Key
    strText = Mid$(strBUK1, 20, 20)
    xmlTempResponseElem.setAttribute "ADDRESSKEY", strText
   
    'Enhancement Flags
    'Flat
    strText = Mid$(strBUK1, 40, 1)
    xmlTempResponseElem.setAttribute "FLATENHANCED", strText
   
    'House Name
    strText = Mid$(strBUK1, 41, 1)
    xmlTempResponseElem.setAttribute "HOUSENAMEENHANCED", strText
   
    'House Number
    strText = Mid$(strBUK1, 42, 1)
    xmlTempResponseElem.setAttribute "HOUSENUMBERENHANCED", strText
   
    'Street
    strText = Mid$(strBUK1, 43, 1)
    xmlTempResponseElem.setAttribute "STREETENHANCED", strText
   
    'District
    strText = Mid$(strBUK1, 44, 1)
    xmlTempResponseElem.setAttribute "DISTRICTENHANCED", strText
   
    'Town
    strText = Mid$(strBUK1, 45, 1)
    xmlTempResponseElem.setAttribute "TOWNENHANCED", strText
   
    'County
    strText = Mid$(strBUK1, 46, 1)
    xmlTempResponseElem.setAttribute "COUNTYENHANCED", strText
   
    'Postcode
    strText = Mid$(strBUK1, 47, 1)
    xmlTempResponseElem.setAttribute "POSTCODEENHANCED", strText
   
    'Now for the variable length Address Details
    'Flat
    nCurPos = 51
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "FLAT", strText

    End If
    nCurPos = nOffset + 1

    'House Name
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "HOUSENAME", strText
    End If
    nCurPos = nOffset + 1
    
    'House Number
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "HOUSENUMBER", strText
    End If
    nCurPos = nOffset + 1

    'Street
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "STREET", strText
    End If
    nCurPos = nOffset + 1
    
    'District
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "DISTRICT", strText
    End If
    nCurPos = nOffset + 1
    
    'Town
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "TOWN", strText
    End If
    nCurPos = nOffset + 1
    
    'County
    nOffset = InStr(nCurPos, strBUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "COUNTY", strText
    End If
    nCurPos = nOffset + 1
    
    'Postcode has no trailing comma - block end is a #
    nOffset = InStr(nCurPos, strBUK1, "#", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "POSTCODE", strText
    End If
    
    'Used in RA040 to indicate if the address has been resolved
    xmlTempResponseElem.setAttribute "ADDRESSRESOLVED", "N"
    
    xmlResponseElem.appendChild xmlTempResponseElem.cloneNode(True)
    
    Set MakeBUK1 = xmlResponseElem
   
MakeBUK1Exit:
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing

    Exit Function
    
MakeBUK1VbErr:

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Set MakeBUK1 = m_objErrAssist.CreateErrorResponseEx()
    
    '   go to clean-up section
    Resume MakeBUK1Exit

End Function

'INR BMIDS682 New Function
Private Function MakeAUK1(ByVal strAUK1 As String, ByVal nAUK1Len As Integer, _
                        ByVal strMainElem As String) As IXMLDOMNode

On Error GoTo MakeAUK1VbErr

Const strFunctionName As String = "MakeAUK1"

Dim strText As String
Dim nOffset As Integer
Dim nCurPos As Integer
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseElem As IXMLDOMElement
    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    Set xmlTempResponseElem = xmlResponseDoc.createElement(strMainElem)
    
    xmlTempResponseElem.setAttribute "BLOCKTYPE", "AUK1"
    xmlTempResponseElem.setAttribute "BLOCKLENGTH", nAUK1Len

    'First items are all fixed length
    'Block Sequence Number
    strText = Mid$(strAUK1, 9, 2)
    xmlTempResponseElem.setAttribute "BLOCKSEQNUMBER", strText

    'Address Indicator
    strText = Mid$(strAUK1, 11, 1)
    xmlTempResponseElem.setAttribute "ADDRESSINDICATOR", strText

    'Detail Code
    strText = Mid$(strAUK1, 12, 1)
    xmlTempResponseElem.setAttribute "DETAILCODE", strText

    'RMC Code
    strText = Mid$(strAUK1, 13, 7)
    xmlTempResponseElem.setAttribute "RMCCODE", strText

    'Region Code
    strText = Mid$(strAUK1, 20, 1)
    xmlTempResponseElem.setAttribute "REGIONCODE", strText

    'Address Data Length
    strText = Mid$(strAUK1, 21, 3)
    xmlTempResponseElem.setAttribute "ADDRESSLENGTH", strText

    'Now for the variable length Address Details
    'Flat
    nCurPos = 24
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "FLAT", strText
    End If
    nCurPos = nOffset + 1

    'House Name
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "HOUSENAME", strText
    End If
    nCurPos = nOffset + 1
    
    'House Number
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "HOUSENUMBER", strText
    End If
    nCurPos = nOffset + 1

    'Street
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "STREET", strText
    End If
    nCurPos = nOffset + 1
    
    'District
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "DISTRICT", strText
    End If
    nCurPos = nOffset + 1
    
    'Town
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "TOWN", strText
    End If
    nCurPos = nOffset + 1
    
    'County
    nOffset = InStr(nCurPos, strAUK1, ",", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "COUNTY", strText
    End If
    nCurPos = nOffset + 1
    
    'Postcode has no trailing comma - block end is a #
    nOffset = InStr(nCurPos, strAUK1, "#", 1)
    If (nCurPos = nOffset) Then
        'No Text
    Else
        strText = Mid$(strAUK1, nCurPos, nOffset - nCurPos)
        xmlTempResponseElem.setAttribute "POSTCODE", strText
    End If
    
    'Used in RA040 to indicate if the address has been resolved
    xmlTempResponseElem.setAttribute "ADDRESSRESOLVED", "N"
    xmlTempResponseElem.setAttribute "FULLSTRING", strAUK1
    
    xmlResponseElem.appendChild xmlTempResponseElem.cloneNode(True)
 
    Set MakeAUK1 = xmlResponseElem
   
MakeAUK1Exit:
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing

    Exit Function
    
MakeAUK1VbErr:

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    Set MakeAUK1 = m_objErrAssist.CreateErrorResponseEx()
    
    '   go to clean-up section
    Resume MakeAUK1Exit

End Function

'INR BMIDS682 New Sub
Private Sub SaveRequiredRequestElements(ByVal strReturn As String, _
                                            ByVal szCustomerNumberOtherN As String, _
                                            ByVal szCustomerNumberSameP As String, _
                                            ByVal nCustomerVersionNumberOtherN As String, _
                                            ByVal nCustomerVersionNumberSameP As String, _
                                            ByVal nAppFactFindNumber As String, _
                                            ByVal vxmlResponseNode As IXMLDOMElement, _
                                            ByVal strExpReturn As String)

On Error GoTo SaveRequiredRequestElementsExit

Const strFunctionName As String = "SaveRequiredRequestElements"

    Dim strHeader As String
    Dim strCCN1 As String
    Dim strText As String
    Dim sBlockSeqNo As String
    Dim sAddrIndicator As String
    Dim strErrText As String
    Dim nNumBUK1s As Integer
    Dim nNumAUK1s As Integer
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlDisplayResponseElem As IXMLDOMElement
    Dim objBUK1NodeList As IXMLDOMNodeList
    Dim objBUK1Elem As IXMLDOMElement
    Dim objAUK1NodeList As IXMLDOMNodeList
    Dim objAUK1Elem As IXMLDOMElement
    Dim objNAM1Elem As IXMLDOMElement
    Dim objTempNAM1Elem As IXMLDOMElement
    Dim objCCN1Elem As IXMLDOMElement
    Dim objCCN1Name As IXMLDOMElement
    Dim objDisplayElem As IXMLDOMElement
    Dim nNumNodes As Integer
    Dim nCCNPos As Integer
    Dim strBlockSearch As String
    Dim strNameMatchSearch As String
    Dim nNAM1Pos As Integer
    Dim nCurPos As Integer
    Dim nOffset As Integer
    Dim sName As String
    Dim strPattern As String
    Dim nBlockCount As Integer
    nBlockCount = 1
    'BMIDS730
    Dim strNamesCount As String
    Dim nNamesCount As Integer
    
    'Initialise response
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    Set xmlTempResponseElem = xmlResponseDoc.createElement("REQHEADER")
    
    'Need the header returned from Experian regardless of AUK1 or BUK1 returned
'    strHeader = Left$(strReturn, 133)
    strHeader = Left$(strExpReturn, 133)
    xmlTempResponseElem.setAttribute "HEADER", strHeader

    'Need to keep this lot for the eventual save
    xmlTempResponseElem.setAttribute "CUSTNUMBEROTHER", szCustomerNumberOtherN
    xmlTempResponseElem.setAttribute "CUSTNUMBERSAME", szCustomerNumberSameP
    xmlTempResponseElem.setAttribute "CUSTVERSIONOTHER", nCustomerVersionNumberOtherN
    xmlTempResponseElem.setAttribute "CUSTVERSIONSAME", nCustomerVersionNumberSameP
    xmlTempResponseElem.setAttribute "FACTFINDNUMBER", nAppFactFindNumber

    xmlResponseElem.appendChild xmlTempResponseElem.cloneNode(True)
   
    Set objTempNAM1Elem = xmlResponseDoc.createElement("TempNAM1")
    Set xmlTempResponseElem = xmlResponseDoc.createElement("NAM1")
    
    'Looking for the NAM1 info. Possibly dealing with multiple NAM1 blocks,
    'but we only want the info associated with our TARGETDISPLAY nodes.
    nNAM1Pos = InStr(1, strReturn, "NAM1", 1)
    Do While (nNAM1Pos > 0)
    
        'We are looking for all NAM1 blocks
        nNAM1Pos = InStr(nNAM1Pos, strReturn, "NAM1", 1)
        
        If nNAM1Pos > 0 Then
            Dim nNAM1Len As Integer
            Dim strNAM1SeqNo As String
            Dim strName As String
            Dim strNAM1Text As String
            
            nNAM1Len = Mid$(strReturn, nNAM1Pos + 4, 4)
            strNAM1SeqNo = Mid$(strReturn, nNAM1Pos + 8, 2)
            strNAM1Text = Mid$(strReturn, nNAM1Pos, nNAM1Len)
            If (StrComp(Right$(strNAM1Text, 1), "#") = 0) Then
                xmlTempResponseElem.setAttribute "BLOCKSEQNUMBER", strNAM1SeqNo
                
                'Work through the variable length info to get the applicants name
                'Title
                nCurPos = 22
                nOffset = InStr(nCurPos, strNAM1Text, ",", 1)
                nCurPos = nOffset + 1
            
                'ForeName
                nOffset = InStr(nCurPos, strNAM1Text, ",", 1)
                If (nCurPos = nOffset) Then
                    'No Text
                Else
                    strName = Mid$(strNAM1Text, nCurPos, nOffset - nCurPos) & " "
                End If
                nCurPos = nOffset + 1
                
                'Initials
                nOffset = InStr(nCurPos, strNAM1Text, ",", 1)
                nCurPos = nOffset + 1
                
                'Surname
                nOffset = InStr(nCurPos, strNAM1Text, ",", 1)
                If (nCurPos = nOffset) Then
                    'No Text
                Else
                    strName = strName & Mid$(strNAM1Text, nCurPos, nOffset - nCurPos)
                End If
                nCurPos = nOffset + 1

                xmlTempResponseElem.setAttribute "TEMPNAME", strName

                objTempNAM1Elem.appendChild xmlTempResponseElem.cloneNode(True)
            Else
                strErrText = "Address Targeting: length of NAM1 block does not match stated Block Length"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
            End If
        nNAM1Pos = nNAM1Pos + 1
        End If
    Loop

    Set xmlTempResponseElem = Nothing
    
    'BMIDS730 CCN1 info required before AUK's & BUK's to add the
    'correct names to be displayed
    Set objCCN1Elem = xmlResponseDoc.createElement("CCN1LIST")
    
    'Possibly dealing with multiple CCN1 blocks
    nCCNPos = InStr(1, strReturn, "CCN1", 1)
    Do While (nCCNPos > 0)
    
        Set xmlTempResponseElem = xmlResponseDoc.createElement("CCN1")
        'We are looking for all CCN1 blocks
        nCCNPos = InStr(nCCNPos, strReturn, "CCN1", 1)

        If nCCNPos > 0 Then
            Dim nCCN1Len As Integer
            Dim strCCN1Text As String
            Dim strCCN1SeqNo As String

            strCCN1SeqNo = Mid$(strReturn, nCCNPos + 8, 2)
            nCCN1Len = Mid$(strReturn, nCCNPos + 4, 4)
            strCCN1Text = Mid$(strReturn, nCCNPos, nCCN1Len)
            nNamesCount = Mid$(strCCN1Text, 15, 1)
            
            'BMIDS730 Use the names count to get the correct names to be displayed
            'We don't want to display an alias name or if an alias is returned
            'and we don't have an alias in Omiga, it won't match anything
            'and will display a blank. Therefore we only want to match on the 2nd
            'char.  i.e 1 for main applicant or 2 for joint
            
            If nNamesCount = 2 Then
                'Get applicant 1 & 2
                'Get the associated NAM1 info
                strPattern = "NAM1[@BLOCKSEQNUMBER = '01']"
                
                Set objNAM1Elem = objTempNAM1Elem.selectSingleNode(strPattern)
                If Not objNAM1Elem Is Nothing Then
                    sName = objNAM1Elem.getAttribute("TEMPNAME")
                End If

                strPattern = "NAM1[@BLOCKSEQNUMBER = '02']"
                
                Set objNAM1Elem = objTempNAM1Elem.selectSingleNode(strPattern)
                If Not objNAM1Elem Is Nothing Then
                    sName = sName & " & " & objNAM1Elem.getAttribute("TEMPNAME")
                End If

            Else
                'Get the specific applicant
                strNameMatchSearch = "0" & Right(strCCN1SeqNo, 1)
    
                'Get the associated NAM1 info
                strPattern = "NAM1[@BLOCKSEQNUMBER = '" & strNameMatchSearch & "']"
                
                Set objNAM1Elem = objTempNAM1Elem.selectSingleNode(strPattern)
                If Not objNAM1Elem Is Nothing Then
                    sName = objNAM1Elem.getAttribute("TEMPNAME")
                End If
            
            End If
            
            If (StrComp(Right$(strCCN1Text, 1), "#") = 0) Then
                xmlTempResponseElem.setAttribute "BLOCKSEQNUMBER", strCCN1SeqNo
                xmlTempResponseElem.setAttribute "CCNTEXT", strCCN1Text
                xmlTempResponseElem.setAttribute "NAME", sName
            Else
                strErrText = "Address Targeting: length of CCN1 block does not match stated Block Length"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
            End If
        
            nCCNPos = nCCNPos + 1
            objCCN1Elem.appendChild xmlTempResponseElem.cloneNode(True)

        End If
    
    Loop
    
    'get a nodelist of all BUK1 nodes
    Set objBUK1NodeList = vxmlResponseNode.getElementsByTagName("ADDRESSTARGET[@BLOCKTYPE = 'BUK1']")
    
    'Get the No. of clients assoc with the application
    nNumBUK1s = objBUK1NodeList.length
    
    'For each BUK1 Block (with unique blockseqno) save info to
    'be displayed in the address targeting summary screen
    For nNumNodes = 1 To (nNumBUK1s) Step 1
    
        Set objBUK1Elem = objBUK1NodeList.Item(nNumNodes - 1)
        
        sBlockSeqNo = objBUK1Elem.getAttribute("BLOCKSEQNUMBER")
        sAddrIndicator = objBUK1Elem.getAttribute("ADDRESSINDICATOR")
        objBUK1Elem.setAttribute "BLOCKCOUNT", nBlockCount
        
        'Check that we don't already have this Invalid address flagged for display
        strBlockSearch = "TARGETDISPLAY[@SEQNUMBER='" & sBlockSeqNo & "']"
        Set objDisplayElem = xmlResponseElem.selectSingleNode(strBlockSearch)
        
        'if a node to be displayed for this sequence number doesn't already exist
        If objDisplayElem Is Nothing Then
            'collate the info required to display on RA040
            Set xmlDisplayResponseElem = xmlResponseDoc.createElement("TARGETDISPLAY")
            xmlDisplayResponseElem.setAttribute "ADDRESSINDICATOR", sAddrIndicator
            xmlDisplayResponseElem.setAttribute "SEQNUMBER", sBlockSeqNo
            xmlDisplayResponseElem.setAttribute "ADDRESSTYPE", "Invalid"
            xmlDisplayResponseElem.setAttribute "ADDRESSRESOLVED", "N"
            xmlDisplayResponseElem.setAttribute "BLOCKCOUNT", nBlockCount

            xmlResponseElem.appendChild xmlDisplayResponseElem.cloneNode(True)
            nBlockCount = nBlockCount + 1
            
        End If
    
    Next
    
    Set xmlTempResponseElem = Nothing
    
    'get a nodelist of all AUK1 nodes
    Set objAUK1NodeList = vxmlResponseNode.getElementsByTagName("ADDRESSTARGET[@BLOCKTYPE = 'AUK1']")
    
    'Get the No. of clients assoc with the application
    nNumAUK1s = objAUK1NodeList.length
    
    'For each AUK1 Block (with unique blockseqno) save info to
    'be displayed in the address targeting summary screen
    For nNumNodes = 1 To (nNumAUK1s) Step 1
    
        Set objAUK1Elem = objAUK1NodeList.Item(nNumNodes - 1)
        sBlockSeqNo = objAUK1Elem.getAttribute("BLOCKSEQNUMBER")
        sAddrIndicator = objAUK1Elem.getAttribute("ADDRESSINDICATOR")
        objAUK1Elem.setAttribute "BLOCKCOUNT", nBlockCount
        
        'Check that we don't already have this Ambiguous address flagged for display
        strBlockSearch = "TARGETDISPLAY[@SEQNUMBER='" & sBlockSeqNo & "']"
        Set objDisplayElem = xmlResponseElem.selectSingleNode(strBlockSearch)
        
        'if a node to be displayed for this sequence number doesn't already exist
        If objDisplayElem Is Nothing Then
        
            'collate the info required to display on RA040
            Set xmlDisplayResponseElem = xmlResponseDoc.createElement("TARGETDISPLAY")
            xmlDisplayResponseElem.setAttribute "ADDRESSINDICATOR", sAddrIndicator
            xmlDisplayResponseElem.setAttribute "SEQNUMBER", sBlockSeqNo
            xmlDisplayResponseElem.setAttribute "ADDRESSTYPE", "Ambiguous"
            xmlDisplayResponseElem.setAttribute "ADDRESSRESOLVED", "N"
            xmlDisplayResponseElem.setAttribute "BLOCKCOUNT", nBlockCount

            xmlResponseElem.appendChild xmlDisplayResponseElem.cloneNode(True)
            nBlockCount = nBlockCount + 1
            
        End If
    Next

    xmlResponseElem.appendChild objCCN1Elem.cloneNode(True)
        
    m_objXmlAssist.AttachResponseData vxmlResponseNode, xmlResponseElem.cloneNode(True)

SaveRequiredRequestElementsExit:
   
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlDisplayResponseElem = Nothing
    Set objBUK1NodeList = Nothing
    Set objBUK1Elem = Nothing
    Set objAUK1NodeList = Nothing
    Set objAUK1Elem = Nothing
    Set objNAM1Elem = Nothing
    Set objTempNAM1Elem = Nothing
    Set objCCN1Elem = Nothing
    Set objDisplayElem = Nothing
    Set objCCN1Name = Nothing
   
   Exit Sub
SaveRequiredRequestElementsVbErr:

    m_objErrAssist.AddToErrSource (strFunctionName)

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
   
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlDisplayResponseElem = Nothing
    Set objBUK1NodeList = Nothing
    Set objBUK1Elem = Nothing
    Set objAUK1NodeList = Nothing
    Set objAUK1Elem = Nothing
    Set objNAM1Elem = Nothing
    Set objTempNAM1Elem = Nothing
    Set objCCN1Elem = Nothing
    Set objDisplayElem = Nothing
    Set objCCN1Name = Nothing
    
    '   re-raise error for calling object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'INR BMIDS682 New Function
Private Function MakeUpAddressTargetRequest(ByVal vxmlRequestNode As IXMLDOMNode) As String

On Error GoTo MakeUpAddressTargetRequestVbErr

Const strFunctionName As String = "MakeUpAddressTargetRequest"
    
    Dim strRequest As String
    Dim strHeader As String
    Dim strOrigHeader As String
    Dim strBody As String
    Dim strCCN1 As String
    Dim nTotNum As Integer
    Dim nNumNodes As Integer
    Dim nLenAUK1 As Integer
    Dim nLenBody As Integer
    Dim sLenBody As String
    Dim nDetailCode As Integer
    Dim sTempText As String
    Dim sFixedText As String
    Dim sVariableText As String
    Dim sBlockSeqNo As String
    Dim strSearch As String
    Dim lBUKLen As Long
    Dim nlenCCN As Integer
    Dim nNumReturnBlocks As Integer
    Dim xmlHeaderElem As IXMLDOMElement
    Dim xmlResolvedAddressElem As IXMLDOMElement
    Dim xmlCCN1Elem As IXMLDOMElement
    Dim xmlCCN1Elements As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strAUKReturn As String
        
    Set xmlNodeList = vxmlRequestNode.selectNodes("TARGETREQUEST/ADDRESSTARGET")
    Set xmlCCN1Elements = vxmlRequestNode.selectSingleNode("CCN1LIST")

    'Get the No. of Resolved addresses to make the request
    nTotNum = xmlNodeList.length
    nNumReturnBlocks = 0
    
    For nNumNodes = 1 To (nTotNum) Step 1
    
        Set xmlResolvedAddressElem = xmlNodeList.Item(nNumNodes - 1)
        'Only need the CCN1 block  if a BUK1 returned
        sTempText = xmlResolvedAddressElem.getAttribute("BLOCKTYPE")
        sBlockSeqNo = xmlResolvedAddressElem.getAttribute("BLOCKSEQNUMBER")
        
        If (StrComp(sTempText, "AUK1") = 0) Then
            nNumReturnBlocks = nNumReturnBlocks + 1
            nDetailCode = xmlResolvedAddressElem.getAttribute("DETAILCODE") + 1
            strBody = xmlResolvedAddressElem.getAttribute("FULLSTRING")
            nLenAUK1 = Len(strBody)
            sTempText = Mid(strBody, 13, nLenAUK1 - 12)
            sTempText = Left(strBody, 11) & nDetailCode & sTempText
            strAUKReturn = strAUKReturn & sTempText
        Else
            'We have a BUK1 Block
            'For each  BUK1 Block there is an associated CCN1 info
            nNumReturnBlocks = nNumReturnBlocks + 2
      
            sFixedText = xmlResolvedAddressElem.getAttribute("BLOCKSEQNUMBER")
            sFixedText = sFixedText & xmlResolvedAddressElem.getAttribute("ADDRESSINDICATOR")
            'Detail Code always 8
            sFixedText = sFixedText & "8"
            sFixedText = sFixedText & xmlResolvedAddressElem.getAttribute("RMC")
            sFixedText = sFixedText & xmlResolvedAddressElem.getAttribute("REGION")
            sVariableText = xmlResolvedAddressElem.getAttribute("FLAT") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("HOUSENAME") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("HOUSENUMBER") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("STREET") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("DISTRICT") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("TOWN") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("COUNTY") & ","
            sVariableText = sVariableText & xmlResolvedAddressElem.getAttribute("POSTCODE") & "#"
            
            'Add on the associated CCN1 Block
            strSearch = "CCN1[@BLOCKSEQNUMBER = '" & sBlockSeqNo & "']"
            Set xmlCCN1Elem = xmlCCN1Elements.selectSingleNode(strSearch)
            strCCN1 = xmlCCN1Elem.getAttribute("CCNTEXT")
            
            'Length of the Address area, DON'T INCLUDE THE END #
            nLenBody = Len(sVariableText) - 1
            sLenBody = ZeroPad(nLenBody, 3)
            
            strBody = sFixedText & sLenBody & sVariableText
            lBUKLen = Len(strBody) + 8
            strBody = "AUK1" & ZeroPad(lBUKLen, 4) & strBody
            
            'Change search mode to 1 for a returned BUK1
            nlenCCN = Len(strCCN1)
            sTempText = strBody & Left(strCCN1, 13) & "1" & Right(strCCN1, nlenCCN - 14)
            strAUKReturn = strAUKReturn & sTempText
            
        End If 'End We have a BUK1 Block
        
    Next
    
    nLenBody = Len(strAUKReturn) + 13
    sLenBody = ZeroPad(nLenBody, 5)
    Set xmlHeaderElem = vxmlRequestNode.selectSingleNode("REQHEADER")
    strOrigHeader = xmlHeaderElem.getAttribute("HEADER")
    sTempText = Right(strOrigHeader, 15)
    sTempText = Left(sTempText, 12) & ZeroPad(nNumReturnBlocks, 3)
    strHeader = Left(strOrigHeader, 105) & "12 32000" & sLenBody & sTempText
    
    strRequest = strHeader & strAUKReturn

    MakeUpAddressTargetRequest = strRequest

    Set xmlHeaderElem = Nothing
    Set xmlResolvedAddressElem = Nothing
    Set xmlCCN1Elem = Nothing
    Set xmlCCN1Elements = Nothing
    Set xmlNodeList = Nothing

    Exit Function

MakeUpAddressTargetRequestVbErr:

    m_objErrAssist.AddToErrSource (strFunctionName)

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set xmlHeaderElem = Nothing
    Set xmlResolvedAddressElem = Nothing
    Set xmlCCN1Elem = Nothing
    Set xmlCCN1Elements = Nothing
    Set xmlNodeList = Nothing

    '   re-raise error for calling object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


'INR BMIDS682 New Function
Private Function AddressTargetRequestFromString(ByVal strRequest As String, _
                                                ByVal strExpReturn As String, _
                                                ByVal vxmlRequestNode As IXMLDOMNode) As String

On Error GoTo AddressTargetRequestFromStringVbErr

Const strFunctionName As String = "AddressTargetRequestFromString"
    
Dim strText As String
Dim nOffset As Integer
Dim nBUKPos As Integer
Dim nBUK1Len As Integer
Dim strBUK1 As String
Dim nCurPos As Integer
Dim strBody As String
Dim strFixed As String
Dim strAddress As String
Dim strNewRequest As String
Dim strBlockSeq As String
Dim lBlocklen As Long
Dim strBlocklen As String
Dim lAddresslen As Long
Dim strAddresslen As String
Dim strAUKReplace As String
Dim strHeader As String
Dim strOrigHeader As String
Dim strCCN As String
Dim nCCNPos As Integer
Dim strErrText As String
Dim nLenBody As Integer
Dim sLenBody As String
Dim sTempText As String
Dim bAssumedOrEnhanced As Boolean
Dim bCCN1NotFound As Boolean
Dim nNumReturnBlocks As Integer
Dim nlenCCN As Integer
'BMIDS730
Dim xmlCCN1Elem As IXMLDOMElement
Dim xmlCCN1Elements As IXMLDOMElement
Dim strSearch As String
Dim nCCN1Len As Integer
Dim strCCN1Text As String
Dim strCCN1SeqNo As String

bAssumedOrEnhanced = True
bCCN1NotFound = True
    'need to check for BUK1 addresstargeting blocks
    'Possibly dealing with multiple BUK1 blocks
    nBUKPos = InStr(1, strExpReturn, "BUK1", 1)
    Do While nBUKPos > 0
        nBUKPos = InStr(nBUKPos, strExpReturn, "BUK1", 1)
        If nBUKPos > 0 Then
    
            nBUK1Len = Mid$(strExpReturn, nBUKPos + 4, 4)
            strBUK1 = Mid$(strExpReturn, nBUKPos, nBUK1Len)
            
            If (StrComp(Right$(strBUK1, 1), "#") = 0) Then
                nBUKPos = nBUKPos + nBUK1Len
            Else
                strErrText = "Credit Check: BUK1 block indicates Address is Unmatched, but length of BUK1 block does not match stated Block Length"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
            End If
                
        
             'First items are all fixed length
             'Look at the Enhancement Flags to see if we can send this straight back
             'If we find any with Missing/Required "R" or Not Found "C"
             'don't automatically send it back
             'Flat Enhanced
             strText = Mid$(strBUK1, 40, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'House Name Enhanced
             strText = Mid$(strBUK1, 41, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'House Number Enhanced
             strText = Mid$(strBUK1, 42, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'Street Enhanced
             strText = Mid$(strBUK1, 43, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'District Enhanced
             strText = Mid$(strBUK1, 44, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'Town Enhanced
             strText = Mid$(strBUK1, 45, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'County Enhanced
             strText = Mid$(strBUK1, 46, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
            
             'Postcode Enhanced
             strText = Mid$(strBUK1, 47, 1)
             If ((StrComp(strText, "C") = 0) Or (StrComp(strText, "R") = 0)) Then
                bAssumedOrEnhanced = False
             End If
                   
            If (bAssumedOrEnhanced) Then
                
                'Block Sequence Number
                strBlockSeq = Mid$(strBUK1, 9, 2)
               
                'Address Indicator
                strFixed = strBlockSeq & Mid$(strBUK1, 11, 1)
               
                'Detail Code
                strFixed = strFixed & "8"
               
                'RMC Code
                strFixed = strFixed & Mid$(strBUK1, 12, 7)
               
                'Region Code
                strFixed = strFixed & Mid$(strBUK1, 19, 1)
            
                'Now for the variable length Address Details
                'Everything else to the next #
                nCurPos = 51
                nOffset = InStr(nCurPos, strBUK1, "#", 1)
                If (nCurPos = nOffset) Then
                    'No Terminator?, shouldn't be here
                    strErrText = "Automatic BUK1 processing. BUK1 block # terminator not found."
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                Else
                    strText = Mid$(strBUK1, nCurPos, nOffset - nCurPos)
                    strAddress = strText & "#"
                End If
                  
                lAddresslen = Len(strAddress)
                strAddresslen = ZeroPad(lAddresslen, 3)
                strBody = strFixed & strAddresslen & strAddress
    
                lBlocklen = Len(strBody) + 8
                strBlocklen = ZeroPad(lBlocklen, 4)
                            
                strBody = "AUK1" & strBlocklen & strBody
                
                bCCN1NotFound = True
                'Need the associated CCN1 from original request from the
                'stored Proc or could be in a modified request from the
                'Address Targeting screens.
                'Possibly dealing with multiple CCN1 blocks
                nCCNPos = InStr(1, strRequest, "CCN1", 1)
                    
                If nCCNPos > 0 Then
                    Do While ((nCCNPos > 0) Or bCCN1NotFound)
                    
                        'We are looking for all CCN1 blocks
                        nCCNPos = InStr(nCCNPos, strRequest, "CCN1", 1)
                
                        If nCCNPos > 0 Then
                
                            strCCN1SeqNo = Mid$(strRequest, nCCNPos + 8, 2)
                            nCCN1Len = Mid$(strRequest, nCCNPos + 4, 4)
                            strCCN1Text = Mid$(strRequest, nCCNPos, nCCN1Len)
                            
                            If (StrComp(Right$(strCCN1Text, 1), "#") = 0) Then
                                If (StrComp(strCCN1SeqNo, strBlockSeq) = 0) Then
                                    bCCN1NotFound = False
                                    'BMIDS730 Found our CCN1 block, want to leave the loop
                                    'so go to the end less 1 because it is incremented below.
                                    nCCNPos = Len(strRequest) - 1
                                End If
                            Else
                                strErrText = "Address Targeting: length of CCN1 block does not match stated Block Length"
                                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                            End If
                        
                            nCCNPos = nCCNPos + 1
                
                        End If
                    
                    Loop
                Else
                    'Need the associated CCN1 from original request from Address Target screen
                    'Possibly dealing with multiple CCN1 blocks
                    'BMIDS730 Need CCN1's from the RunCreditCheck request, otherwise if
                    'go from AUK1's to BUK1's theCCN1's won't exist
                    Set xmlCCN1Elements = vxmlRequestNode.selectSingleNode("CCN1LIST")
    
                    strSearch = "CCN1[@BLOCKSEQNUMBER = '" & strBlockSeq & "']"
                    Set xmlCCN1Elem = xmlCCN1Elements.selectSingleNode(strSearch)
                    If Not xmlCCN1Elem Is Nothing Then
                        strCCN1Text = xmlCCN1Elem.getAttribute("CCNTEXT")
                    Else
                        'should have found an associated CCN1 by now
                        strErrText = "Auto Return of BUK1: Did not find associated CCN1 block. Seq No: " & strBlockSeq
                        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, strErrText
                    End If
                End If
            
                'Change search mode to 1 for a returned BUK1
                nlenCCN = Len(strCCN1Text)
                sTempText = Left(strCCN1Text, 13) & "1" & Right(strCCN1Text, nlenCCN - 14)
                
                strAUKReplace = strAUKReplace & strBody & sTempText
                nNumReturnBlocks = nNumReturnBlocks + 2
                
            End If 'If (bAssumedOrEnhanced)
        End If 'If nBUKPos > 0
            
    Loop
    
    If (bAssumedOrEnhanced) Then
        'Need the header returned from Experian
        strOrigHeader = Left$(strExpReturn, 133)
    
        nLenBody = Len(strAUKReplace) + 13
        sLenBody = ZeroPad(nLenBody, 5)
    
        sTempText = Right(strOrigHeader, 15)
        sTempText = Left(sTempText, 12) & ZeroPad(nNumReturnBlocks, 3)
        strHeader = Left(strOrigHeader, 105) & "12 32000" & sLenBody & sTempText
        
        strNewRequest = strHeader & strAUKReplace
        AddressTargetRequestFromString = strNewRequest
    Else
         'Something somewhere is Missing/Required "R" or Not Found "C"
         'can't automatically send it back
        AddressTargetRequestFromString = ""
    End If
   
AddressTargetRequestFromStringExit:

    Set xmlCCN1Elem = Nothing
    Set xmlCCN1Elements = Nothing

    Exit Function
    
AddressTargetRequestFromStringVbErr:

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description, vbLogEventTypeError
    End If

    AddressTargetRequestFromString = m_objErrAssist.CreateErrorResponseEx()
    
    '   go to clean-up section
    Resume AddressTargetRequestFromStringExit


End Function

'INR BMIDS682 New Function
Private Function ZeroPad(ByVal nToPad As Integer, ByVal nLenToPad As Integer) As String
    
    Dim strToPad As String
    Dim nLenOfInput As Integer
    Dim nCounter As Integer
    
    strToPad = nToPad
    nLenOfInput = Len(strToPad)
    
    For nCounter = 1 To nLenToPad - nLenOfInput
        strToPad = "0" & strToPad
    Next nCounter
    
    ZeroPad = strToPad

End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objXmlAssist = New XMLAssist
    Set m_objErrAssist = New ErrAssist
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objErrAssist = Nothing
End Sub

