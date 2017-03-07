VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "SystemDatesDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SystemDatesDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Contains methods that perform date manipulation with regard to working and
'               non-working days in relation to nonworking days specified at country and
'               distribution channel level.
'
'Dependencies:
'               ComboDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    17/01/01    Created.
'DJP    23/01/01    Validate date passed into CheckNonWorkingOccurence.
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'CL     07/05/02    SYS4510  Modifications to class initialise
'CL     10/05/02    SYS4510  Remove class initialize & class terminate
'SDS    25/08/04    BBG369   This AQR is not BBG specific. There is in-consistency in evaluating
'                            the weekday in VB and SQL Server. Hence the fix

'------------------------------------------------------------------------------------------
Option Explicit
Implements ISystemDatesDO
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
Private m_objSQLAssist As SQLAssist
Private Function ISystemDatesDO_CheckNonWorkingOccurence(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Returns whether or not the date passed in is a non-working day or not.
' pass:         vxmlRequest  xml Request data stream
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CheckNonWorkingOccurenceVbErr
    Const strFunctionName As String = "ISystemDates_CheckNonWorkingOccurence"
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlListElem As IXMLDOMElement
    Dim xmlValueList As IXMLDOMNodeList
    Dim xmlComboResp As FreeThreadedDOMDocument40
    Dim objComboDO As ComboDO
    Dim objIADOAssist As IADOAssist
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecordset As ADODB.Recordset
    Dim adoConnection As ADODB.Connection
    Dim blnOpenedOk As Boolean
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim strSQL As String
    Dim strValueSQL As String
    Dim strDayOfWeek As String
    Dim strChannelID As String
    Dim strConnString As String
    Dim strDateToCheck As String
    Dim strIsWorkingDay As String
    Dim strComboValueIDList As String
    Const cstrRecurringValidation As String = "R"
    Const cstrNonWorkingDayGroup As String = "NonWorkingDayType"
    ' Check the correct keys have been passed in
    strDateToCheck = m_objXmlAssist.GetNodeValue(vxmlRequest, _
                                                     ".//DATE")
    ' Validate Date
    If Not IsDate(strDateToCheck) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeInvalidParameter, "DATE must be a validate date or date/time"
    End If
    strDayOfWeek = Weekday(strDateToCheck)
    strChannelID = m_objXmlAssist.GetNodeValue(vxmlRequest, _
                                                     ".//CHANNELID")
    ' Create objects
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    Set objComboDO = m_objContext.CreateInstance(App.Title & ".ComboDO")
    strSQL = "SELECT BH.BANKHOLIDAYDATE " & _
    "FROM " & _
        "BANKHOLIDAY BH, DISTRIBUTIONCHANNEL DC " & _
    "WHERE " & _
        "DC.CHANNELID = ? AND " & _
        "BH.COUNTRYNUMBER= DC.COUNTRYNUMBER AND " & _
        "BH.BANKHOLIDAYDATE = ? " & _
        "UNION "
    'SDS  BBG369  25/08/2004___START
    strSQL = strSQL & "SELECT  NWD.NONWORKINGDAYDATE " & _
    "FROM " & _
        "NONWORKINGDAY NWD " & _
    "WHERE " & _
        "NWD.CHANNELID = ? AND " & _
        "NWD.NONWORKINGDAYDATE = ? UNION "
    strSQL = strSQL & "SELECT  NWD.NONWORKINGDAYDATE " & _
    "FROM " & _
        "NONWORKINGDAY NWD " & _
    "WHERE " & _
        "NWD.CHANNELID = ? AND " & _
        "NWD.NONWORKINGDAYDATE <= ?  AND " & _
        m_objSQLAssist.GetColumnDayOfWeek("NWD.NONWORKINGDAYDATE") & " = " & m_objSQLAssist.GetColumnDayOfWeek("'" & strDateToCheck & "'")
        'm_objSQLAssist.GetColumnDayOfWeek("NWD.NONWORKINGDAYDATE") & " = ? "
        'SDS  BBG369  25/08/2004___END
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ChannelID", adBSTR, adParamInput, Len(strChannelID), strChannelID)
    adoCommand.Parameters.Append adoParameter
    strDateToCheck = Format(strDateToCheck, "yyyy/mm/dd")
    Set adoParameter = adoCommand.CreateParameter("DateToCheck", adDBDate, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(adoCommand.Parameters.Count - 1) = strDateToCheck
    Set adoParameter = adoCommand.CreateParameter("ChannelID", adBSTR, adParamInput, Len(strChannelID), strChannelID)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter("DateToCheck", adDBDate, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(adoCommand.Parameters.Count - 1) = strDateToCheck
    Set adoParameter = adoCommand.CreateParameter("ChannelID", adBSTR, adParamInput, Len(strChannelID), strChannelID)
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = adoCommand.CreateParameter("DateToCheck", adDBDate, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(adoCommand.Parameters.Count - 1) = strDateToCheck
    'Set adoParameter = adoCommand.CreateParameter("DayOfWeek", adInteger, adParamInput)
    'adoCommand.Parameters.Append adoParameter
    'adoCommand(adoCommand.Parameters.Count - 1) = strDayOfWeek
    '''''''''''''' Get Combo ValueID's ''''''''''''''''''
    strComboValueIDList = objComboDO.GetComboValueId(cstrNonWorkingDayGroup, cstrRecurringValidation)
    Set xmlComboResp = m_objXmlAssist.load(strComboValueIDList, TypeName(Me), strFunctionName)
    ' Need to add the comboid values returned to the SQL - we only want recurring dates to be checked if the
    ' NonWorkingType field has a valueid who's validation type is "R"
    Set xmlValueList = xmlComboResp.getElementsByTagName("VALUEID")
    Dim lngValueCount As Long
    Dim lngValueItem As Long
    lngValueCount = xmlValueList.length
    If lngValueCount > 0 Then
        strValueSQL = " AND ( NWD.NONWORKINGTYPE = ? "
        Set adoParameter = adoCommand.CreateParameter("NonWorkingType", adInteger, adParamInput)
        adoCommand.Parameters.Append adoParameter
        adoCommand(adoCommand.Parameters.Count - 1) = xmlValueList(0).Text
        For lngValueItem = 1 To lngValueCount - 1
            Set adoParameter = adoCommand.CreateParameter("NonWorkingType", adInteger, adParamInput)
            adoCommand.Parameters.Append adoParameter
            adoCommand(adoCommand.Parameters.Count - 1) = xmlValueList(lngValueItem).Text
            strValueSQL = strValueSQL & " OR NWD.NONWORKINGTYPE = ? "
        Next
        strValueSQL = strValueSQL & " ) "
    Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "Value ID tag(s) not found"
    End If
    strSQL = strSQL & strValueSQL
    adoCommand.CommandText = strSQL
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
        adoConnection.Open strConnString
        If adoConnection.ERRORS.Count = 0 Then
            blnOpenedOk = True
        End If
        intAttempt = intAttempt + 1
    Wend
    If blnOpenedOk = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeUnableToConnect
    End If
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.RecordCount > 0 Then
        strIsWorkingDay = "1"
    Else
        strIsWorkingDay = "0"
    End If
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("NONWORKINGIND")
    xmlListElem.Text = strIsWorkingDay
    xmlOut.appendChild xmlListElem
    adoRecordset.Close
    Set ISystemDatesDO_CheckNonWorkingOccurence = xmlListElem
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    m_objContext.SetComplete
    Exit Function
CheckNonWorkingOccurenceVbErr:
    
    
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
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objSQLAssist = New SQLAssist
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
    Set m_objSQLAssist = Nothing
End Sub
