VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "OrganisationDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OrganisationDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Code template for omiga4 Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function GetXMLFromRecordSet(vrstRecordSet As ADODB.Recordset) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     24/09/99 Created
'RF     09/11/99 Updated LoadData functions to match database changes.
'                Corrected FindCurrentUnitList return xml hierarchy.
'RF     11/11/99 Updated LoadUserRoleData to match database change.
'RF     11/11/99 Use ErrAssist.ThrowError rather than RaiseError.
'                Remove obsolete, unimplemented functions.
'RF     19/11/99 Change to error numbers used in CreatePassword
'RF     19/11/99 Fix to FindCurrentUnitList (applied in GetCurUnitListSearchCriteria)
'RF     24/11/99 Correct SQL syntax: use "IS NULL" instead of "= NULL"
'RF     24/11/99 AQR SC1: Don't do a MoveFirst in an empty recordset
'RF     29/11/99 Fix to IsBankHoliday.
'RF     04/01/99 Add optional parameter to LoadDepartmentData.
'MC     17/05/00 SYS0166 - Add method UpdateChangePasswordIndicator
'IW     23/05/00 SYS0774 DISTRIBUTIONCHANNELID S/B/ CHANNELID
'MV     20/11/00 CORE000011 Added FindUnitList Function
'CL     27/11/00 CORE00005 Modifications to meet new standards
'DJP    14/12/00 Added FindUserNameList and FindUnitNameList
'DJP    19/12/00 Join OrganisationUser on FindUserList
'MV     21/12/00 SYS1729 :Removed LoadDistributionChannelData from FindCurrentUnitList
'DRC    05/04/01 SYS1925 Added GetActionOwnerDetails
'LD     11/06/01 SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'LD     12/06/01 SYS2368 SQL Server Port - Modify select statement in IOrganisationDO_GetCurrentCompetency
'SR     03/07/01 SYS2433 Modified method 'GetOmigaUserDetails'
'JR     04/09/01 Omiplus24 Modified SQL in FindUnitSummary method to include ContactTelephoneDetails table
'AD     10/10/01 Removed the MoveFirst as it was ForwardOnly cursor. SYS2718.
'JR     04/09/01 SYS2860, Comment out code ref. to Omiplus24 telephone number in FindUnitSummary SQL
'JR     03/12/01 SYS2459 re-apply Omiplus24 telephone changes
'SAB    21/09/05 CORE188 Added optional functionality to FindCurrentUserList to only
'                return those units that are linked to a user and that allow logon.
'SAB    22/09/05 CORE191 Updated isBankHoliday so that it accepts an optional date parameter
'JD     20/12/2005  CORE227 set return value for IsDateBankHoliday
'------------------------------------------------------------------------------------------
Option Explicit
Implements IOrganisationDO
Implements ObjectControl
 '=============================================
 'Variable Declaration Section
 '=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
Private Enum ConditionalOperator
    enumOR = 1
    enumAND
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'
'   AW     13/08/2002  BM029 - Added IOrganisationDO_UpdateUnitChequeNo
'
Private Function IOrganisationDO_FindDepartmentList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDepartmentListVbErr
    Const strFunctionName As String = "IOrganisationDO_FindDepartmentList"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim objSQLAssist As SQLAssist
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim strChannelID As String
    Dim strActiveOnly As String
    Dim strToDay As String
    Dim strSQL As String
    Set objXmlAssist = New XMLAssist
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    ' Check the correct keys have been passed in
    strChannelID = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//CHANNELID", False)
    strActiveOnly = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//ACTIVEONLY", False)
    'Preparing SQL Statement
    strSQL = "SELECT * FROM DEPARTMENT "
    If Len(strChannelID) <> 0 Then
        'If ChannelId exists then append to the Where clause
        strSQL = strSQL & " WHERE CHANNELID = ? "
        Set adoParameter = adoCommand.CreateParameter("ChannelID", adBSTR, adParamInput, Len(strChannelID), strChannelID)
        adoCommand.Parameters.Append adoParameter
    End If
    'Store Todays date based on Database
    strToDay = objSQLAssist.GetSystemDate
    If strActiveOnly <> "0" Then
        If Len(strChannelID) > 0 Then
            strSQL = strSQL & " AND "
        Else
            strSQL = strSQL & " WHERE "
        End If
        strSQL = strSQL & " DEPARTMENTACTIVEFROM <= " & strToDay
        strSQL = strSQL & " AND ( DEPARTMENTACTIVETO >= " & strToDay
        strSQL = strSQL & " OR  DEPARTMENTACTIVETO IS NULL ) "
    End If
    adoCommand.CommandText = strSQL
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDepartmentData()
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("DEPARTMENTLIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IOrganisationDO_FindDepartmentList = xmlListElem
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
    Set objSQLAssist = Nothing
    m_objContext.SetComplete
    Exit Function
FindDepartmentListVbErr:
    
    
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
    Set objSQLAssist = Nothing
      
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
Private Function IOrganisationDO_FindDistributionChannelList() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDistributionChannelListVbErr
    Const strFunctionName As String = "IOrganisationDO_FindDistributionChannelList"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim objSQLAssist As SQLAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim strChannelID As String
    Dim strActiveOnly As String
    Dim strToDay As String
    Dim strSQL As String
        
    Set objXmlAssist = New XMLAssist
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    strSQL = "SELECT * FROM DISTRIBUTIONCHANNEL "
    adoCommand.CommandText = strSQL
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDistributionChannelData()
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("DISTRIBUTIONCHANNELLIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IOrganisationDO_FindDistributionChannelList = xmlListElem
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
    Set objSQLAssist = Nothing
    m_objContext.SetComplete
    Exit Function
FindDistributionChannelListVbErr:
    
    
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
    Set objSQLAssist = Nothing
       
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
Private Function IOrganisationDO_FindUnitSummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'history:
'BG SYS1860 22/01/01 Added Method
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitSummaryVbErr
    Const strFunctionName As String = "IOrganisationDO_FindUnitSummary"
           
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim strUnitID As String
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    Set objXmlAssist = New XMLAssist
    ' Check the correct keys have been passed in
    strUnitID = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//UNITID")
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    Dim strSQL As String
    'JR Omiplus24, replaced with SQL code below
'    strSQL = "SELECT * FROM UNIT, UNITCONTACTDETAILS, CONTACTDETAILS WHERE " & _
'             "UNIT.UNITID=? AND UNITCONTACTDETAILS.UNITID = UNIT.UNITID AND " & _
'             "CONTACTDETAILS.CONTACTDETAILSGUID = UNITCONTACTDETAILS.CONTACTDETAILSGUID"
   
    strSQL = "SELECT * FROM UNIT, UNITCONTACTDETAILS, CONTACTDETAILS, CONTACTTELEPHONEDETAILS WHERE " & _
             "UNIT.UNITID=? AND UNITCONTACTDETAILS.UNITID = UNIT.UNITID AND " & _
             "CONTACTDETAILS.CONTACTDETAILSGUID = UNITCONTACTDETAILS.CONTACTDETAILSGUID AND " & _
             "CONTACTDETAILS.CONTACTDETAILSGUID = CONTACTTELEPHONEDETAILS.CONTACTDETAILSGUID"
    adoCommand.CommandText = strSQL
    ' Set up the parameter to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("UnitID", adBSTR, adParamInput, Len(strUnitID), strUnitID)
    adoCommand.Parameters.Append adoParameter
      
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
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnitSummaryData(usFindUnitSummary)
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlElement = xmlOut.createElement("UNITSUMMARY")
    xmlOut.appendChild xmlElement
    ' Convert recordset to XML
    Set xmlTempNode = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc, xmlElement)
    Set xmlClassDefDoc = Nothing
    Set xmlClassDefDoc = objIClassDef.LoadContactDetailsData(cdFindUnitSummary)
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlTempNode
    adoRecordset.Close
    Set IOrganisationDO_FindUnitSummary = xmlElement
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    m_objContext.SetComplete
    Exit Function
FindUnitSummaryVbErr:
    
    
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
    Set xmlElement = Nothing
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
Private Function IOrganisationDO_FindUserList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindUserListVbErr
    Const strFunctionName As String = "IOrganisationDO_FindUserList"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim objSQLAssist As SQLAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim strUnitID As String
    Dim strUserRole As String
    Dim strActiveOnly As String
    Dim strToDay As String
    Dim strSQL As String
    Set objXmlAssist = New XMLAssist
    Set objSQLAssist = New SQLAssist
    ' Check the correct keys have been passed in
    strUnitID = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//UNITID", False)
    strUserRole = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//USERROLE", False)
    strActiveOnly = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//ACTIVEONLY", False)
    If Len(strUnitID) = 0 And Len(strUserRole) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ChannelID and DepartmenID missing"
    End If
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    strSQL = "SELECT DISTINCT OMIGAUSER.*, USERFORENAME, USERSURNAME FROM USERROLE ,OMIGAUSER, ORGANISATIONUSER " & _
             " WHERE USERROLE.USERID = OMIGAUSER.USERID AND OMIGAUSER.USERID = ORGANISATIONUSER.USERID"
                    
    If Len(strUnitID) <> 0 Then
        strSQL = strSQL & " AND USERROLE.UNITID  = ? "
        Set adoParameter = adoCommand.CreateParameter("UnitID", adBSTR, adParamInput, Len(strUnitID), strUnitID)
        adoCommand.Parameters.Append adoParameter
    End If
    strToDay = objSQLAssist.GetSystemDate
    If strActiveOnly <> "0" Then
        strSQL = strSQL & " AND OMIGAUSER.OMIGAUSERACTIVEFROM <= " & strToDay
        strSQL = strSQL & " AND ( OMIGAUSER.OMIGAUSERACTIVETO >= " & strToDay
        strSQL = strSQL & " OR  OMIGAUSER.OMIGAUSERACTIVETO IS NULL ) "
        strSQL = strSQL & " AND USERROLE.USERROLEACTIVEFROM <= " & strToDay
        strSQL = strSQL & " AND (USERROLE.USERROLEACTIVETO >= " & strToDay
        strSQL = strSQL & " OR  USERROLE.USERROLEACTIVETO IS NULL ) "
    End If
    adoCommand.CommandText = strSQL
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUserData(udFindUserList)
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("OMIGAUSERLIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IOrganisationDO_FindUserList = xmlListElem
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
    Set objSQLAssist = Nothing
        
    m_objContext.SetComplete
    Exit Function
FindUserListVbErr:
    
    
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
    Set objSQLAssist = Nothing
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
Private Sub IOrganisationDO_UpdateUnitChequeNo(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be updated
'                    <REQUEST>
'                       <UNITLIST>
'                           <UNIT>
'                             <UNITID></UNITID>
'                             <UNHIGHCHEQUENUMBER></UNHIGHCHEQUENUMBER>
'                            </UNIT>
'                       </UNITLIST>
'                    </REQUEST>
'
' return:   none
'   AW     13/08/2002  BM029
'------------------------------------------------------------------------------------------
On Error GoTo UpdateUnitChequeNoVbErr
    Const strFunctionName As String = "UpdateUnitChequeNo"
    Dim objXMLDocument As FreeThreadedDOMDocument40
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objGeneralAssist As New GeneralAssist
    Dim strUnitID As String, strChequeNum As String
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
      
     ' Check the correct keys have been passed in
    Set objXmlAssist = New XMLAssist
    strUnitID = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//UNITID", False)
    strChequeNum = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//UNHIGHCHEQUENUMBER", False)
    If Len(strUnitID) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ChannelID is missing"
    End If
    If Len(strChequeNum) = 0 Or objGeneralAssist.IsDigits(strChequeNum) = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Cheque number is invalid"
    End If
    Set objIADOAssist = New ADOAssist
    Set objXMLDocument = New FreeThreadedDOMDocument40
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnitChequeNo()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objGeneralAssist = Nothing
    m_objContext.SetComplete
        
        
    Exit Sub
UpdateUnitChequeNoVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objGeneralAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
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
Private Function IOrganisationDO_FindUnitList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitListVbErr
    Const strFunctionName As String = "IOrganisationDO_FindUnitList"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim objSQLAssist As SQLAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim strChannelID As String
    Dim strDepartmentID As String
    Dim strActiveOnly As String
    Set objXmlAssist = New XMLAssist
    Set objSQLAssist = New SQLAssist
    ' Check the correct keys have been passed in
    strChannelID = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//CHANNELID", False)
    strDepartmentID = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//DEPARTMENTID", False)
    strActiveOnly = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//ACTIVEONLY", False)
    If Len(strChannelID) = 0 And Len(strDepartmentID) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ChannelID and DepartmenID missing"
    End If
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    Dim strToDay As String
    Dim strSQL As String
    strSQL = "SELECT UNIT.* FROM UNIT ,DEPARTMENT WHERE DEPARTMENT.DEPARTMENTID = UNIT.DEPARTMENTID "
    Set adoParameter = adoCommand.CreateParameter("ChannelID", adBSTR, adParamInput)
    adoCommand.Parameters.Append adoParameter
    If Len(strChannelID) > 0 Then
        strSQL = strSQL & "AND DEPARTMENT.CHANNELID = ? "
                adoParameter.Size = Len(strChannelID)
        adoCommand(0) = strChannelID
    End If
    If Len(strDepartmentID) > 0 Then
        strSQL = strSQL & "AND DEPARTMENT.DEPARTMENTID = ? "
        If Len(strChannelID) > 0 Then
            Set adoParameter = adoCommand.CreateParameter("DepartmentID", adBSTR, adParamInput)
            adoCommand.Parameters.Append adoParameter
                        adoParameter.Size = Len(strDepartmentID)
            adoCommand(1) = strDepartmentID
        Else
                        adoParameter.Size = Len(strDepartmentID)
            adoCommand(0) = strDepartmentID
        End If
    End If
    strToDay = objSQLAssist.GetSystemDate
    If strActiveOnly <> "0" Then
        strSQL = strSQL & " AND DEPARTMENT.DEPARTMENTACTIVEFROM <= " & strToDay
        strSQL = strSQL & " AND ( DEPARTMENT.DEPARTMENTACTIVETO >= " & strToDay
        strSQL = strSQL & " OR  DEPARTMENT.DEPARTMENTACTIVETO IS NULL ) "
        strSQL = strSQL & " AND UNIT.UNITACTIVEFROM <= " & strToDay
        strSQL = strSQL & " AND (UNIT.UNITACTIVETO >= " & strToDay
        strSQL = strSQL & " OR  UNIT.UNITACTIVETO IS NULL ) "
    End If
    adoCommand.CommandText = strSQL
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnitData()
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("UNITLIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IOrganisationDO_FindUnitList = xmlListElem
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
    Set objSQLAssist = Nothing
       
    m_objContext.SetComplete
    Exit Function
FindUnitListVbErr:
    
    
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
    Set objSQLAssist = Nothing
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
Private Function IOrganisationDO_FindCurrentUnitList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get Unit, UserRole, Department data for the supplied UserId where all are active.
' pass:
'   vxmlTableElement
'       xml Request data stream containing data to which identifies
'       the instance(s) of the persistant data to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
' Raise Errors:
'   if record not found, raise oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindCurrentUnitListVbErr
    Const strFunctionName As String = "FindCurrentUnitList"
          
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIADOAssist As IADOAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim xmlDeptListNode As IXMLDOMNode
    Dim xmlChannelListNode As IXMLDOMNode
    Dim xmlUnitListNode As IXMLDOMNode
    Dim xmlDeptNode As IXMLDOMNode
    Dim xmlChannelNode As IXMLDOMNode
    Dim xmlUnitNode As IXMLDOMNode
    Dim xmlUserRoleListNode As IXMLDOMNode
    Dim adoParameter As ADODB.Parameter
    Dim adoConnection As ADODB.Connection
    Dim objSQLAssist As SQLAssist
    Dim objClassDef As IOmOrgClassDef
    Dim xmlDeptDoc As FreeThreadedDOMDocument40
    Dim xmlChannelDoc As FreeThreadedDOMDocument40
    Dim xmlUnitDoc As FreeThreadedDOMDocument40
    Dim xmlUserRoleDoc As FreeThreadedDOMDocument40
                   
    Dim strSQL As String
    Dim strUserId As String
    Dim strChannelID As String
    Dim strUnitID As String
    Dim strSysDate As String
    Dim strThisDeptId As String
    Dim strThisChannelId As String
    Dim strThisUnitId As String
    Dim strLastDeptId As String
    Dim strLastChannelId As String
    Dim strLastUnitId As String
    Dim strAllowOmigaLogon As String 'CORE188
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
        
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set adoConnection = New ADODB.Connection
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set objClassDef = New omOrgClassDef
    Set objSQLAssist = New SQLAssist
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    'errors check
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
              
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
    strUnitID = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//UNITID", False, False)
    ' CORE188 - Extract optional key from XML
    strAllowOmigaLogon = m_objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                           ".//ALLOWOMIGALOGON", False)
    ' CORE188 - End
    
    strSysDate = objSQLAssist.GetSystemDate()
        
    strSQL = "SELECT *  "
    strSQL = strSQL + " FROM UNIT, USERROLE, DEPARTMENT , DISTRIBUTIONCHANNEL WHERE "
    strSQL = strSQL + " USERROLE.USERID = ? "
    If Len(strUnitID) > 0 Then
        strSQL = strSQL & " AND UNIT.UNITID = ?"
    End If
    
    ' CORE188
    If Len(strAllowOmigaLogon) > 0 Then
        If strAllowOmigaLogon = "0" Then
            strSQL = strSQL + " AND (UNIT.ALLOWOMIGALOGON = 0 OR UNIT.ALLOWOMIGALOGON IS NULL) "
        Else
            If strAllowOmigaLogon = "1" Then
                strSQL = strSQL + " AND UNIT.ALLOWOMIGALOGON = 1 "
            Else
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInValidKeyValue, "ALLOWOMIGALOGON"
            End If
        End If
    End If
    ' CORE188 - End
    ' RF 19/11/99
    strSQL = strSQL & " AND USERROLE.UNITID = UNIT.UNITID"
    strSQL = strSQL & " AND UNIT.DEPARTMENTID = DEPARTMENT.DEPARTMENTID"
    strSQL = strSQL & " AND DEPARTMENT.CHANNELID = DISTRIBUTIONCHANNEL.CHANNELID"
    strSQL = strSQL & " AND USERROLE.USERROLEACTIVEFROM <= " & strSysDate
    strSQL = strSQL & " AND (USERROLE.USERROLEACTIVETO >= " & strSysDate
    strSQL = strSQL & " OR USERROLE.USERROLEACTIVETO IS NULL)"
    strSQL = strSQL & " AND UNIT.UNITACTIVEFROM <= " & strSysDate
    strSQL = strSQL & " AND (UNIT.UNITACTIVETO >= " & strSysDate
    strSQL = strSQL & " OR UNIT.UNITACTIVETO IS NULL)"
    strSQL = strSQL & " AND DEPARTMENT.DEPARTMENTACTIVEFROM <= " & strSysDate
    strSQL = strSQL & " AND (DEPARTMENT.DEPARTMENTACTIVETO >= " & strSysDate
    strSQL = strSQL & " OR DEPARTMENT.DEPARTMENTACTIVETO IS NULL)"
    strSQL = strSQL & " ORDER BY DEPARTMENT.DEPARTMENTID, "
    strSQL = strSQL & " UNIT.UNITID, USERROLE.USERID"
    cmd.CommandText = strSQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
            
    If Len(strUnitID) > 0 Then
        Set adoParameter = cmd.CreateParameter("UnitID", adBSTR, adParamInput, Len(strUnitID), strUnitID)
        cmd.Parameters.Append adoParameter
    End If
       
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    Set xmlElem = xmlOut.createElement("DEPARTMENTLIST")
    Set xmlDeptListNode = xmlOut.appendChild(xmlElem)
    Set xmlDeptDoc = objClassDef.LoadDepartmentData(ddFindCurrentUnitList)
    Set xmlUnitDoc = objClassDef.LoadUnitData()
    Set xmlUserRoleDoc = objClassDef.LoadUserRoleData()
     
    'AD 10/10/01 Removed the MoveFirst as it was ForwardOnly cursor. SYS2718.
    ' loop through the record set
    While Not rst.EOF
        
        ' check if different department
        strThisDeptId = rst.Fields.Item("DEPARTMENTID").Value
        If strThisDeptId <> strLastDeptId Then
            strLastDeptId = strThisDeptId
            Set xmlDeptNode = m_objIDOAssist.GetXMLFromRecordSet(rst, _
                                                                 xmlDeptDoc, _
                                                                 xmlDeptListNode)
            Set xmlElem = xmlOut.createElement("UNITLIST")
            Set xmlUnitListNode = xmlDeptNode.appendChild(xmlElem)
        End If
        ' check if different unit
        strThisUnitId = rst.Fields.Item("UNITID").Value
        If strThisUnitId <> strLastUnitId Then
            strLastUnitId = strThisUnitId
            Set xmlUnitNode = m_objIDOAssist.GetXMLFromRecordSet(rst, _
                                                                 xmlUnitDoc, _
                                                                 xmlUnitListNode)
            Set xmlElem = xmlOut.createElement("USERROLELIST")
            Set xmlUserRoleListNode = xmlUnitNode.appendChild(xmlElem)
        End If
        ' handle all userroles
        m_objIDOAssist.GetXMLFromRecordSet rst, xmlUserRoleDoc, xmlUserRoleListNode
              
        rst.MoveNext
    Wend
            
    Set IOrganisationDO_FindCurrentUnitList = xmlDeptListNode
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlDeptListNode = Nothing
    Set xmlUnitListNode = Nothing
    Set xmlDeptNode = Nothing
    Set xmlUnitNode = Nothing
    Set xmlUserRoleListNode = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
    Set objSQLAssist = Nothing
    Set objClassDef = Nothing
    Set xmlDeptDoc = Nothing
    Set xmlUnitDoc = Nothing
    Set xmlUserRoleDoc = Nothing
                
    m_objContext.SetComplete
    Exit Function
FindCurrentUnitListVbErr:
    
    
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
      
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
      
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlDeptListNode = Nothing
    Set xmlUnitListNode = Nothing
    Set xmlDeptNode = Nothing
    Set xmlUnitNode = Nothing
    Set xmlUserRoleListNode = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
    Set objSQLAssist = Nothing
    Set objClassDef = Nothing
    Set xmlDeptDoc = Nothing
    Set xmlUnitDoc = Nothing
    Set xmlUserRoleDoc = Nothing
                       
    m_objErrAssist.AddToErrSource (strFunctionName)
    m_objContext.SetComplete
            
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IOrganisationDO_GetOmigaUserDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data from OMIGAUSER, ORGANISATIONUSER, INTERMEDIARYUSER,
'   INTERMEDIARYINDIVIDUAL, QUALIFICATION, USERCOMPETENCY
' pass:
'   vstrXMLRequest
'       Format:
'           <REQUEST>
'               <OMIGAUSER>
'                   <USERID></USERID>
'               </OMIGAUSER>
'           </REQUEST>
' return:
'       Format:
'           <OMIGAUSER>
'               ...
'               <ORGANISATIONUSER>  (optional)
'                   ...
'               </ORGANISATIONUSER>
'               <SURNAME></SURNAME>
'               <FORENAME></FORENAME>
'               <TITLE></TITLE>
'               <QUALIFICATIONLIST> (optional)
'                   ...
'               </QUALIFICATIONLIST>
'               <COMPETENCY>        (optional)
'                   ...
'               </COMPETENCY>
'           </OMIGAUSER>
'------------------------------------------------------------------------------------------
On Error GoTo GetOmigaUserDetailsVbErr
    Const strFunctionName As String = "GetOmigaUserDetails"
        
    Dim xmlOmigaUserNode As IXMLDOMNode
    Dim objComboDO As ComboDO
    Dim xmlOrganisationUser  As IXMLDOMNode
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlOmigaUser As IXMLDOMNode
    Dim xmlIntermediaryUser As IXMLDOMNode
    Dim xmlIntermediaryIndividual As IXMLDOMNode
    Dim xmlQualificationList As IXMLDOMNode
    Dim xmlCompetency As IXMLDOMNode
    Dim xmlUserDataDoc As FreeThreadedDOMDocument40
    Dim xmlIntUserDataDef As FreeThreadedDOMDocument40
    Dim xmlIntIndividualDataDef As FreeThreadedDOMDocument40
    Dim xmlCopyNode As IXMLDOMNode
    Dim xmlSurname As IXMLDOMElement
    Dim xmlForeName As IXMLDOMElement
    Dim xmlTitle As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
      
    Dim blnIsOrganisationUser As Boolean
    Dim blnCompetencyFound As Boolean
    Dim blnQualificationListFound As Boolean
                 
                 
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
             
    Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
                           
    '------------------------------------------------------------------------------------------
    ' get the OmigaUser record if user is active
    '------------------------------------------------------------------------------------------
        
    Set xmlOmigaUser = IOrganisationDO_GetCurrentOmigaUser(vxmlTableElement)
    'SR 03/07/01 : SYS2433 - Search in the current context
    If xmlOmigaUser.nodeName = "OMIGAUSER" Then
        Set xmlOmigaUserNode = xmlOmigaUser
    Else
        'Set xmlOmigaUserNode = m_objXmlAssist.GetNode(xmlOmigaUser, "//OMIGAUSER")
        Set xmlOmigaUserNode = m_objXmlAssist.GetNode(xmlOmigaUser, ".//OMIGAUSER")
    End If
    '------------------------------------------------------------------------------------------
    ' establish whether the user if a member of staff or an intermediary and get the
    ' required data
    '------------------------------------------------------------------------------------------
         
    Set xmlCopyNode = vxmlTableElement.cloneNode(True)
    m_objXmlAssist.ChangeNodeName xmlCopyNode, "OMIGAUSER", "ORGANISATIONUSER"
    Err.Clear
    On Error Resume Next
    Set objIClassDef = New omOrgClassDef
    Set xmlUserDataDoc = objIClassDef.LoadOrganisationUserData()
              
    Set xmlOrganisationUser = m_objIDOAssist.GetData(xmlCopyNode, xmlUserDataDoc)
        
    m_objErrAssist.SaveErr
    On Error GoTo GetOmigaUserDetailsVbErr
    If m_objErrAssist.LastErr = 0 Then
        blnIsOrganisationUser = True
    Else
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
            blnIsOrganisationUser = False
            Err.Clear
        Else
            m_objErrAssist.ReRaise
        End If
    End If
      
    Set xmlCopyNode = Nothing
          
    If blnIsOrganisationUser Then
        Set xmlSurname = m_objXmlAssist.GetNode(xmlOrganisationUser, ".//USERSURNAME")
        Set xmlForeName = m_objXmlAssist.GetNode(xmlOrganisationUser, ".//USERFORENAME")
        Set xmlTitle = m_objXmlAssist.GetNode(xmlOrganisationUser, ".//USERTITLE")
        xmlOmigaUserNode.appendChild xmlOrganisationUser
    Else
        
        '------------------------------------------------------------------------------------------
        ' user is not a member of staff so get their IntermediaryUser data and
        ' IntermediaryUser data
        '------------------------------------------------------------------------------------------
        Set xmlCopyNode = vxmlTableElement.cloneNode(True)
        m_objXmlAssist.ChangeNodeName xmlCopyNode, "ORGANISATIONUSER", "INTERMEDIARYUSER"
        m_objXmlAssist.ChangeNodeName xmlCopyNode, "USERID", "INTERMEDIARYUSERID"
                
        Set xmlIntUserDataDef = objIClassDef.LoadIntermediaryUserData()
        Set xmlIntermediaryUser = m_objIDOAssist.GetData(xmlCopyNode, xmlIntUserDataDef)
        m_objXmlAssist.ChangeNodeName xmlIntermediaryUser, "INTERMEDIARYUSER", "INTERMEDIARYINDIVIDUAL"
        Set xmlIntIndividualDataDef = objIClassDef.LoadIntermediaryIndividualData()
        Set xmlIntermediaryIndividual = m_objIDOAssist.GetData(xmlIntermediaryUser, xmlIntIndividualDataDef)
        Set xmlSurname = m_objXmlAssist.GetNode(xmlIntermediaryIndividual, ".//SURNAME")
        Set xmlForeName = m_objXmlAssist.GetNode(xmlIntermediaryIndividual, ".//FORENAME")
        Set xmlTitle = m_objXmlAssist.GetNode(xmlIntermediaryIndividual, ".//TITLE")
    End If
       
    Set xmlElement = xmlOmigaUserNode.ownerDocument.createElement("SURNAME")
    xmlOmigaUserNode.appendChild xmlElement
    xmlElement.Text = xmlSurname.Text
    Set xmlElement = xmlOmigaUserNode.ownerDocument.createElement("FORENAME")
    xmlOmigaUserNode.appendChild xmlElement
    xmlElement.Text = xmlForeName.Text
    Set xmlElement = xmlOmigaUserNode.ownerDocument.createElement("TITLE")
    xmlOmigaUserNode.appendChild xmlElement
    xmlElement.Text = xmlTitle.Text
    xmlElement.setAttribute "TEXT", xmlTitle.getAttribute("TEXT")
    '------------------------------------------------------------------------------------------
    ' get the User's current qualifications
    '------------------------------------------------------------------------------------------
    Err.Clear
    On Error Resume Next
    Set xmlQualificationList = IOrganisationDO_FindCurrentQualificationList(vxmlTableElement)
      
    m_objErrAssist.SaveErr
    On Error GoTo GetOmigaUserDetailsVbErr
    If m_objErrAssist.LastErr = 0 Then
        blnQualificationListFound = True
    Else
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
            blnQualificationListFound = False
            Err.Clear
        Else
            m_objErrAssist.ReRaise
        End If
    End If
    If blnQualificationListFound Then
        xmlOmigaUserNode.appendChild xmlQualificationList
    End If
    '------------------------------------------------------------------------------------------
    ' get the User's current competency
    '-----------------------------------------------------------------------------------------
    Err.Clear
    On Error Resume Next
     
    Set xmlCompetency = IOrganisationDO_GetCurrentCompetency(vxmlTableElement)
    m_objErrAssist.SaveErr
    On Error GoTo GetOmigaUserDetailsVbErr
    If m_objErrAssist.LastErr = 0 Then
        blnCompetencyFound = True
    Else
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) = oeRecordNotFound Then
            blnCompetencyFound = False
            Err.Clear
        Else
            m_objErrAssist.ReRaise
        End If
    End If
    If blnCompetencyFound Then
       
        xmlOmigaUserNode.appendChild xmlCompetency
    End If
    '------------------------------------------------------------------------------------------
    ' return the collected data
    '------------------------------------------------------------------------------------------
    Set IOrganisationDO_GetOmigaUserDetails = xmlOmigaUserNode
    Set objComboDO = Nothing
    Set xmlCopyNode = Nothing
    Set xmlOmigaUserNode = Nothing
    Set xmlOrganisationUser = Nothing
    Set objIClassDef = Nothing
    Set xmlOmigaUser = Nothing
    Set xmlIntermediaryUser = Nothing
    Set xmlIntermediaryIndividual = Nothing
    Set xmlQualificationList = Nothing
    Set xmlCompetency = Nothing
    Set xmlUserDataDoc = Nothing
    Set xmlIntUserDataDef = Nothing
    Set xmlIntIndividualDataDef = Nothing
    Set xmlCopyNode = Nothing
    Set xmlElement = Nothing
                  
    m_objContext.SetComplete
        
        
    Exit Function
GetOmigaUserDetailsVbErr:
    
    '<VSA>
    Set objComboDO = Nothing
    Set xmlCopyNode = Nothing
    Set xmlOmigaUserNode = Nothing
    Set xmlOrganisationUser = Nothing
    Set objIClassDef = Nothing
    Set xmlOmigaUser = Nothing
    Set xmlIntermediaryUser = Nothing
    Set xmlIntermediaryIndividual = Nothing
    Set xmlQualificationList = Nothing
    Set xmlCompetency = Nothing
    Set xmlUserDataDoc = Nothing
    Set xmlIntUserDataDef = Nothing
    Set xmlIntIndividualDataDef = Nothing
    Set xmlCopyNode = Nothing
    Set xmlElement = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
            
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IOrganisationDO_FindCurrentQualificationList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo FindCurrentQualificationListVbErr
    
    Const strFunctionName As String = "FindCurrentQualificationList"
             
    Dim cmd As ADODB.Command
    Dim objIADOAssist As IADOAssist
    Dim rst As ADODB.Recordset
    Dim xmlElem As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objSQLAssist As SQLAssist
    Dim xmlNode As IXMLDOMNode
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim adoParameter As ADODB.Parameter
    Dim strUserId As String
    Dim strSQL As String
    Dim strSysDate As String
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
              
                   
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objSQLAssist = New SQLAssist
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set adoConnection = New ADODB.Connection
    Set xmlOut = New FreeThreadedDOMDocument40
            
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
           
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
        
    strSysDate = objSQLAssist.GetSystemDate()
    strSQL = "SELECT * FROM QUALIFICATION WHERE " _
    & " USERID = ?" _
    & " AND QUALIFICATIONACTIVEFROM <= " & strSysDate _
    & " AND (QUALIFICATIONACTIVETO >= " & strSysDate _
    & " OR QUALIFICATIONACTIVETO IS NULL)"
    cmd.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set xmlElem = xmlOut.createElement("QUALIFICATIONLIST")
    Set xmlListNode = xmlOut.appendChild(xmlElem)
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadQualificationData
    m_objIDOAssist.GetXMLFromWholeRecordset rst, xmlClassDefDoc, xmlListNode
    rst.Close
      
    Set IOrganisationDO_FindCurrentQualificationList = xmlListNode
    Set xmlDoc = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
              
              
    m_objContext.SetComplete
       
       
    Exit Function
FindCurrentQualificationListVbErr:
    
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
        
    Set xmlDoc = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set adoConnection = Nothing
        
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
          
End Function
Private Function IOrganisationDO_GetCurrentCompetency(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentCompetencyVbErr
    Const strFunctionName As String = "GetCurrentCompetency"
     
    Dim cmd As ADODB.Command
    Dim objIADOAssist As IADOAssist
    Dim rst As ADODB.Recordset
    Dim xmlElem As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim objSQLAssist As SQLAssist
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlLoadUserCompetencyDataDef As FreeThreadedDOMDocument40
    Dim adoParameter As ADODB.Parameter
    Dim strUserId As String
    Dim strSQL As String
    Dim strSysDate As String
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
      
    Set xmlOut = New FreeThreadedDOMDocument40
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
              
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set adoConnection = New ADODB.Connection
    Set objSQLAssist = New SQLAssist
       
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
    strSysDate = objSQLAssist.GetSystemDate()
     
    strSQL = "SELECT * FROM USERCOMPETENCY WHERE " _
    & "USERCOMPETENCYACTIVEFROM = (" _
    & "SELECT MAX(USERCOMPETENCYACTIVEFROM) FROM USERCOMPETENCY WHERE " _
    & "COMPETENCYTYPE IN (" _
    & "SELECT COMPETENCYTYPE FROM COMPETENCY WHERE " _
    & "COMPETENCYACTIVEFROM <= " & strSysDate & " AND (" _
    & "COMPETENCYACTIVETO >= " & strSysDate & " OR " _
    & "COMPETENCYACTIVETO IS NULL) AND " _
    & "USERID = ? ))"
    cmd.CommandText = strSQL
       
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
      
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
              
    Set objIClassDef = New omOrgClassDef
    Set xmlLoadUserCompetencyDataDef = objIClassDef.LoadUserCompetencyData()
    Set IOrganisationDO_GetCurrentCompetency = m_objIDOAssist.GetXMLFromRecordSet(rst, xmlLoadUserCompetencyDataDef)
    rst.Close
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadUserCompetencyDataDef = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set adoConnection = Nothing
                 
    m_objContext.SetComplete
        
        
    Exit Function
GetCurrentCompetencyVbErr:
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
        
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadUserCompetencyDataDef = Nothing
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set adoConnection = Nothing
          
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
           
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IOrganisationDO_GetActionOwnerDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo GetActionOwnerDetailsVbErr
    Const strFunctionName As String = "GetActionOwnerDetails"
     
    Dim cmd As ADODB.Command
    Dim objIADOAssist As IADOAssist
    Dim rst As ADODB.Recordset
    Dim xmlElem As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim objSQLAssist As SQLAssist
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlLoadActionOwnerDataDef As FreeThreadedDOMDocument40
    Dim adoParameter As ADODB.Parameter
    Dim strUnitID As String
    Dim strTaskOwnerType As String
    Dim strSQL As String
    Dim strSysDate As String
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
      
    Set xmlOut = New FreeThreadedDOMDocument40
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
              
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set adoConnection = New ADODB.Connection
    Set objSQLAssist = New SQLAssist
       
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    strUnitID = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//UNITID")
    strTaskOwnerType = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//TASKOWNERTYPE")
     
    strSQL = "SELECT * FROM ACTIONOWNER WHERE " _
      & "UNITID = ? AND " _
      & "TASKOWNERTYPE = ?"
    cmd.CommandText = strSQL
       
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UnitID", adBSTR, adParamInput, Len(strUnitID), strUnitID)
    cmd.Parameters.Append adoParameter
    Set adoParameter = cmd.CreateParameter("TaskOwnerType", adBSTR, adParamInput, Len(strTaskOwnerType), strTaskOwnerType)
    cmd.Parameters.Append adoParameter
      
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
              
    Set objIClassDef = New omOrgClassDef
    Set xmlLoadActionOwnerDataDef = objIClassDef.LoadActionOwnerData()
    Set IOrganisationDO_GetActionOwnerDetails = m_objIDOAssist.GetXMLFromRecordSet(rst, xmlLoadActionOwnerDataDef)
    rst.Close
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadActionOwnerDataDef = Nothing
    Set adoParameter = Nothing
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set adoConnection = Nothing
                 
    m_objContext.SetComplete
        
        
    Exit Function
GetActionOwnerDetailsVbErr:
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
        
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadActionOwnerDataDef = Nothing
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set adoConnection = Nothing
          
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
           
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IOrganisationDO_GetWorkingHoursForToday(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetWorkingHoursForTodayVbErr
    Const strFunctionName As String = "GetWorkingHoursForToday"
       
    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlLoadWorkingHoursDataDef As FreeThreadedDOMDocument40
    Dim adoParameter As Parameter
    Dim strUserId As String
    Dim strChannelID As String
    Dim strTagName As String
    Dim strSQL As String
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set objSQLAssist = New SQLAssist
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set adoConnection = New ADODB.Connection
           
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    strTagName = ".//USERID"
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, strTagName)
        
    strTagName = ".//CHANNELID"
    strChannelID = m_objXmlAssist.GetNodeValue(vxmlTableElement, strTagName)
    strSQL = "SELECT WORKINGHOURS.*, WORKINGHOURTYPE.* " _
    & "FROM WORKINGHOURS, WORKINGHOURTYPE, OMIGAUSER " _
    & "WHERE OMIGAUSER.USERID = ? " _
    & " AND " & "WORKINGHOURTYPE.WORKINGHOURTYPE = OMIGAUSER.WORKINGHOURTYPE AND " _
    & "WORKINGHOURTYPE.WORKINGHOURTYPE = WORKINGHOURS.WORKINGHOURTYPE AND " _
    & "WORKINGHOURS.DAYOFTHEWEEK = ?"
    cmd.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
    Set adoParameter = cmd.CreateParameter("UserID", adInteger, adParamInput)
    cmd.Parameters.Append adoParameter
    cmd(1) = Weekday(Date)
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    If rst!BANKHOLIDAYINDICATOR <> 1 Then
        ' check if it's a bank holiday
        If IOrganisationDO_IsBankHoliday(strChannelID) Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeRecordNotFound, _
                "This WORKINGHOURTYPE requires a BANKHOLIDAY record"
        End If
    End If
    Set objIClassDef = New omOrgClassDef
    Set xmlLoadWorkingHoursDataDef = objIClassDef.LoadWorkingHoursData()
    Set IOrganisationDO_GetWorkingHoursForToday = m_objIDOAssist.GetXMLFromRecordSet(rst, xmlLoadWorkingHoursDataDef)
           
    rst.Close
    Set objSQLAssist = New SQLAssist
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadWorkingHoursDataDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
            
    m_objContext.SetComplete
        
        
    Exit Function
GetWorkingHoursForTodayVbErr:
  
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    Set objSQLAssist = New SQLAssist
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set objIADOAssist = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadWorkingHoursDataDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
            
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
          
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IOrganisationDO_GetCurrentPassword(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the most recent password (i.e. one with latest PASSWORD.CREATIONDATE) for a
'   specified USERID
' pass:
'   vstrXMLRequest
'       Format:
'           <PASSWORD>
'               <USERID></USERID>
'           </PASSWORD>
' return:
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentPasswordVbErr
    Const strFunctionName As String = "GetCurrentPassword"
    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim xmlLoadPasswordData As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlLoadPasswordDataDef As FreeThreadedDOMDocument40
    Dim adoParameter As Parameter
    Dim strUserId As String
    Dim strSQL As String
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
           
      
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objSQLAssist = New SQLAssist
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set adoConnection = New ADODB.Connection
             
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
                
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
            
    strSQL = "select * from PASSWORD where USERID = ? " _
    & " AND PASSWORDCREATIONDATE = " _
    & "(SELECT MAX(PASSWORDCREATIONDATE) FROM PASSWORD " _
    & " where USERID = ? )"
    cmd.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
        
    Set adoParameter = cmd.CreateParameter("UserID2", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
      
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    Set objIClassDef = New omOrgClassDef
    Set xmlLoadPasswordDataDef = objIClassDef.LoadPasswordData()
    Set IOrganisationDO_GetCurrentPassword = m_objIDOAssist.GetXMLFromRecordSet(rst, xmlLoadPasswordDataDef)
      
    rst.Close
      
    Set rst = Nothing
    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set objSQLAssist = Nothing
    Set xmlLoadPasswordData = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadPasswordDataDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
                     
    m_objContext.SetComplete
        
    Exit Function
GetCurrentPasswordVbErr:
    
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    Set rst = Nothing
    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set objSQLAssist = Nothing
    Set xmlLoadPasswordData = Nothing
    Set objIClassDef = Nothing
    Set xmlLoadPasswordDataDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
        
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
            
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub IOrganisationDO_CreatePassword(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create a new password record for the supplied UserId. The input PASSWORDVALUE
'   is encrypted before being saved.
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'                   Format:
'                   <REQUEST>
'                       <PASSWORD>
'                           <USERID></USERID>
'                           <PASSWORDVALUE></PASSWORDVALUE>
'                       </PASSWORD>
'                   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CreatePasswordVbErr
    
    Const strFunctionName As String = "CreatePassword"
            
    Dim objSQLAssist As SQLAssist
    Dim objIADOAssist As IADOAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim xmlCopyNode As IXMLDOMNode
    Dim xmlPasswordElem As IXMLDOMElement
    Dim xmlPassWordValueNode As IXMLDOMNode
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
            
    Dim blnRecFound As Boolean
    Dim strNewPassword As String
    Dim dblMaxPasswords As Double
    Dim intPasswordCount As Integer
    Dim strSQL As String
    Dim strFields As String
    Dim strValues As String
    Dim strTable As String
        
    Set objSQLAssist = New SQLAssist
    Set objIClassDef = New IOmOrgClassDef
     'Check for MTS error
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
        
    Set xmlCopyNode = vxmlRequest.cloneNode(True)
            
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set objGlobalParameterDO = _
            m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    '------------------------------------------------------------------------------------------
    ' encrypt the password and put the encryted version into the request xml
    '------------------------------------------------------------------------------------------
    Set xmlNode = m_objXmlAssist.GetNode(xmlCopyNode, ".//PASSWORDVALUE")
    strNewPassword = xmlNode.Text
    strNewPassword = Encrypt(strNewPassword)
    xmlNode.Text = strNewPassword
    '------------------------------------------------------------------------------------------
    ' check if the password already exists;
    ' also count the number of existing passwords
    '------------------------------------------------------------------------------------------
    ' remove PASSWORDVALUE from the vxmlRequest so it doesn't restrict the search
    'Check this, we are checking from the top of the doc
    Set xmlPassWordValueNode = xmlCopyNode.selectSingleNode("PASSWORDVALUE")
    xmlCopyNode.removeChild xmlPassWordValueNode
        
    Err.Clear
    On Error Resume Next
    Set xmlNode = IOrganisationDO_FindPasswordList(xmlCopyNode)
    m_objErrAssist.SaveErr
    On Error GoTo CreatePasswordVbErr
    If m_objErrAssist.LastErr = 0 Then
        blnRecFound = True
    Else
        If m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeRecordNotFound Then
            m_objErrAssist.ReRaise
        End If
    End If
      
    If blnRecFound = True Then
        
        Set xmlNodeList = xmlNode.selectNodes("//PASSWORDVALUE")
        intPasswordCount = xmlNodeList.Length
          
        For Each xmlPasswordElem In xmlNodeList
            If xmlPasswordElem.Text = strNewPassword Then
                'RF 19/11/99 Change to error number used
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 145
            End If
        Next
    Else
        intPasswordCount = 0
    End If
    '------------------------------------------------------------------------------------------
    ' check count against max passwords
    '------------------------------------------------------------------------------------------
    dblMaxPasswords = objGlobalParameterDO.GetCurrentParameterByType( _
        "MaximumPasswords", "AMOUNT")
    If intPasswordCount >= dblMaxPasswords Then
        
        '------------------------------------------------------------------------------------------
        ' delete old passwords
        '------------------------------------------------------------------------------------------
                
        Set xmlTableNode = xmlCopyNode.ownerDocument.createElement("PASSWORD")
        Set xmlElem = m_objXmlAssist.GetNode(vxmlRequest, ".//USERID")
        xmlTableNode.appendChild xmlElem.cloneNode(True)
        Set xmlElem = xmlCopyNode.ownerDocument.createElement("NUMBEROFRECORDS")
        xmlTableNode.appendChild xmlElem
        xmlElem.Text = CStr(dblMaxPasswords - intPasswordCount + 1)
        DeleteOldestPasswords xmlTableNode
                
    End If
    '------------------------------------------------------------------------------------------
    ' create the password using the updated Xml Request string
    '------------------------------------------------------------------------------------------
      
    ' Put password back into copy of request
    xmlCopyNode.appendChild xmlPassWordValueNode
       
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadPasswordData
    m_objIDOAssist.BuildSQLString xmlCopyNode, xmlClassDefDoc, strTable, _
        sftFieldValueSeparated, cdkPRIMARYKEY, cdkaALLKEYS, , strFields, strValues
    m_objIDOAssist.BuildSQLString xmlCopyNode, xmlClassDefDoc, strTable, _
        sftFieldValueSeparated, cdkOTHERS, cdkaANYKEYS, , strFields, strValues
    'append the date
    strFields = strFields & ",PASSWORDCREATIONDATE"
    strValues = strValues & "," & objSQLAssist.GetSystemDate()
    strSQL = "INSERT INTO " & strTable & _
        " (" & strFields & ") VALUES (" & strValues & ")"
    objIADOAssist.executeSQLCommand strSQL
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTableNode = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlCopyNode = Nothing
    Set xmlPasswordElem = Nothing
    Set xmlPassWordValueNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    m_objContext.SetComplete
    Exit Sub
CreatePasswordVbErr:
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlTableNode = Nothing
    Set objGlobalParameterDO = Nothing
    Set xmlCopyNode = Nothing
    Set xmlPasswordElem = Nothing
    Set xmlPassWordValueNode = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objSQLAssist = Nothing
    Set objIADOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function IOrganisationDO_FindPasswordList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindPasswordListVbErr
    
    Const strFunctionName As String = "FindPasswordList"
       
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlPasswordDataDef As FreeThreadedDOMDocument40
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objIClassDef = New omOrgClassDef
    Set xmlPasswordDataDef = objIClassDef.LoadPasswordData()
    Set IOrganisationDO_FindPasswordList = m_objIDOAssist.FindList(vxmlTableElement, xmlPasswordDataDef)
      
    m_objContext.SetComplete
        
    Set objIClassDef = Nothing
    Set xmlPasswordDataDef = Nothing
      
        
    Exit Function
FindPasswordListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlPasswordDataDef = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub DeleteOldestPasswords(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete the specified number of oldest password records for the specified user.
' pass:
'   vstrXMLRequest
'       Format:
'           <PASSWORD>
'               <USERID></USERID>
'               <NUMBEROFRECORDS></NUMBEROFRECORDS>
'           <PASSWORD>
'------------------------------------------------------------------------------------------
On Error GoTo DeleteOldestPasswordsVbErr
    Const strFunctionName As String = "DeleteOldestPasswords"
    Dim strUserId As String
    Dim strNumToDelete As String
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim cmd As ADODB.Command
    Dim adoParameter  As Parameter
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    Dim intLoop As Integer
    Set objSQLAssist = New SQLAssist
    Set cmd = New ADODB.Command
            
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
      
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
         
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
        
    strNumToDelete = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//NUMBEROFRECORDS")
      
    strSQL = "DELETE FROM PASSWORD WHERE " _
    & "USERID = " _
    & objSQLAssist.FormatString(strUserId) _
    & " AND PASSWORDCREATIONDATE = " _
    & "(SELECT MIN(PASSWORDCREATIONDATE) FROM PASSWORD WHERE " _
    & "USERID = " _
    & objSQLAssist.FormatString(strUserId) _
    & ")"
              
    For intLoop = 1 To CInt(strNumToDelete)
        objIADOAssist.executeSQLCommand (strSQL)
    Next
       
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set objIClassDef = Nothing
    Set cmd = Nothing
    Set adoParameter = Nothing
          
    m_objContext.SetComplete
    Exit Sub
DeleteOldestPasswordsVbErr:
    
    
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set objIClassDef = Nothing
    Set cmd = Nothing
    Set adoParameter = Nothing
      
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function IOrganisationDO_GetCurrentOmigaUser(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
'       xml Request data stream containing data to which identifies
'       the instance(s) of the persistant data to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentOmigaUserVbErr
    Const strFunctionName As String = "GetCurrentOmigaUser"
           
    Dim cmd As ADODB.Command
    Dim objIADOAssist As IADOAssist
    Dim rst As ADODB.Recordset
    Dim xmlElem As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objSQLAssist As SQLAssist
    Dim xmlOmigaUserData  As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim adoParameter  As Parameter
    Dim strUserId As String
    Dim strSQL As String
    Dim strSysDate As String
    Dim strClassDef As String
    Dim strXML As String
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
    Dim xmlGetCurrOmigaUserDef As FreeThreadedDOMDocument40
           
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set adoConnection = New ADODB.Connection
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
         
    strUserId = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//USERID")
     
    strSysDate = objSQLAssist.GetSystemDate()
    strSQL = "SELECT * FROM OMIGAUSER WHERE " _
    & " USERID = ? " _
    & " AND OMIGAUSERACTIVEFROM <= " & strSysDate _
    & " AND (OMIGAUSERACTIVETO >= " & strSysDate _
    & " OR OMIGAUSERACTIVETO IS NULL)"
    cmd.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
      
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    Set objIClassDef = New omOrgClassDef
    Set xmlGetCurrOmigaUserDef = objIClassDef.LoadUserData()
         
    Set IOrganisationDO_GetCurrentOmigaUser = m_objIDOAssist.GetXMLFromRecordSet(rst, xmlGetCurrOmigaUserDef)
    rst.Close
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlNode = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlOmigaUserData = Nothing
    Set objIClassDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
    Set xmlGetCurrOmigaUserDef = Nothing
      
    m_objContext.SetComplete
        
        
    Exit Function
GetCurrentOmigaUserVbErr:
    
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set rst = Nothing
    Set xmlElem = Nothing
    Set xmlListNode = Nothing
    Set xmlNode = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlOmigaUserData = Nothing
    Set objIClassDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
    Set xmlGetCurrOmigaUserDef = Nothing
      
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
        
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IOrganisationDO_IsBankHoliday( _
    ByVal vstrDistributionChannelId As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Returns true if today is a bank holiday for the supplied channel id
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo IsBankHolidayVbErr
    Const strFunctionName As String = "IsBankHoliday"
        
    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim cmd As ADODB.Command
    Dim strCriteria As String
    Dim intNumRecs As Integer
    Dim IsBankHoliday As Boolean
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    strCriteria = "DISTRIBUTIONCHANNEL.CHANNELID = " _
    & objSQLAssist.FormatString(vstrDistributionChannelId) _
    & " AND DISTRIBUTIONCHANNEL.COUNTRYNUMBER = " _
    & "BANKHOLIDAY.COUNTRYNUMBER" _
    & " AND BANKHOLIDAY.BANKHOLIDAYDATE = " _
    & objSQLAssist.GetSystemDate()
        
    intNumRecs = objIADOAssist.GetNumberOfRecords("BANKHOLIDAY,DISTRIBUTIONCHANNEL", _
        strCriteria)
    If intNumRecs > 0 Then
        IsBankHoliday = True
    Else
        IsBankHoliday = False
    End If
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
         
    m_objContext.SetComplete
              
    Exit Function
IsBankHolidayVbErr:
    
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
           
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'CORE191 - New Method
Private Function IOrganisationDO_IsDateBankHoliday( _
    ByVal vstrDistributionChannelId As String, _
    ByVal vdtmBankHoliday As Date) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Returns true if today or the optional date supplied is a bank holiday for the
'   supplied channel id.
' pass:
'------------------------------------------------------------------------------------------
On Error GoTo IsBankHolidayVbErr
    Const strFunctionName As String = "IsBankHoliday"
    
    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As SQLAssist
    Dim objConvertAssist As New ConvertAssist
    Dim eDbEngineType As DBENGINETYPE
    Dim cmd As ADODB.Command
    Dim strCriteria As String
    Dim strOracleSearchDate As String
    Dim intNumRecs As Integer
    Dim IsBankHoliday As Boolean
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    eDbEngineType = objIADOAssist.GetDBEngine
    
    strCriteria = "DISTRIBUTIONCHANNEL.CHANNELID = " _
    & objSQLAssist.FormatString(vstrDistributionChannelId) _
    & " AND DISTRIBUTIONCHANNEL.COUNTRYNUMBER = " _
    & "BANKHOLIDAY.COUNTRYNUMBER" _
    
    Select Case eDbEngineType
    Case SQLServer
        strCriteria = strCriteria & " AND DATEDIFF(dd,BANKHOLIDAY.BANKHOLIDAYDATE," & _
            objSQLAssist.FormatDate(vdtmBankHoliday) & ")=0 "
    Case Oracle
        strOracleSearchDate = "TO_CHAR(" & objSQLAssist.FormatDate(vdtmBankHoliday) & _
            ", 'YYYY-MM-DD') "
        strCriteria = strCriteria & " AND TO_CHAR(BANKHOLIDAY.BANKHOLIDAYDATE, 'YYYY-MM-DD') = " & _
            strOracleSearchDate
    End Select
        
    intNumRecs = objIADOAssist.GetNumberOfRecords("BANKHOLIDAY,DISTRIBUTIONCHANNEL", _
        strCriteria)
    If intNumRecs > 0 Then
        IsBankHoliday = True
    Else
        IsBankHoliday = False
    End If
    'JD CORE227
    IOrganisationDO_IsDateBankHoliday = IsBankHoliday
    
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
         
    m_objContext.SetComplete
              
    Exit Function
IsBankHolidayVbErr:
    
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
           
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' CORE191 - End

Private Function IOrganisationDO_ValidateUserMandateLevel(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo ValidateUserMandateLevelVbErr
    
    Const strFunctionName As String = "ValidateUserMandateLevel"
    Dim objIADOAssist As IADOAssist
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim objSQLAssist As SQLAssist
    Dim objIClassDef As IOmOrgClassDef
    Dim blnOpenedOk As Boolean
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim adoParameter As ADODB.Parameter
    Dim strUserId As String
    Dim strSQL As String
    Dim strSysDate As String
    Dim intRetries As Integer
    Dim adoConnection As ADODB.Connection
       
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set adoConnection = New ADODB.Connection
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
            
    strUserId = m_objXmlAssist.GetAttributeValue(vxmlTableElement, "REQUEST", "USERID")
    strSysDate = objSQLAssist.GetSystemDate()
    strSQL = "SELECT * FROM COMPETENCY WHERE " & _
                 "COMPETENCYACTIVEFROM <= " & strSysDate & " AND " & _
                 "(COMPETENCYACTIVETO >= " & strSysDate & " OR COMPETENCYACTIVETO IS NULL) AND " & _
                 "COMPETENCYTYPE IN (" & _
                     "SELECT COMPETENCYTYPE FROM USERCOMPETENCY WHERE " & _
                     "USERCOMPETENCYACTIVEFROM = (" & _
                        "SELECT MAX(USERCOMPETENCYACTIVEFROM) FROM USERCOMPETENCY WHERE " & _
                        "USERID =  ? " & " AND " & _
                        "COMPETENCYTYPE IN (" & _
                            "SELECT COMPETENCYTYPE FROM COMPETENCY WHERE " & _
                            "COMPETENCYACTIVEFROM <= " & strSysDate & " AND " & _
                            "(COMPETENCYACTIVETO >= " & strSysDate & " OR COMPETENCYACTIVETO IS NULL))) AND " & _
                     "USERID =  ? " & ")"
    cmd.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = cmd.CreateParameter("UserID", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
    Set adoParameter = cmd.CreateParameter("UserID2", adBSTR, adParamInput, Len(strUserId), strUserId)
    cmd.Parameters.Append adoParameter
        
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
                
    'Try and connect to database
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
                
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
         
    Set objIClassDef = New omOrgClassDef
    Set IOrganisationDO_ValidateUserMandateLevel = m_objIDOAssist.GetXMLFromRecordSet(rst, objIClassDef.LoadCompetencyData)
    Set cmd = Nothing
    Set rst = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
          
    m_objContext.SetComplete
        
            
    Exit Function
ValidateUserMandateLevelVbErr:
       
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If
        
    Set cmd = Nothing
    Set rst = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objIClassDef = Nothing
    Set adoParameter = Nothing
    Set adoConnection = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub IOrganisationDO_UpdateChangePasswordIndicator(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
'                    <REQUEST>
'                       <OMIGAUSER>
'                           <USERID></USERID>
'                           <CHANGEPASSWORDINDICATOR></CHANGEPASSWORDINDICATOR>
'                       </OMIGAUSER>
'                    </REQUEST>
'
' return:   none
'
'------------------------------------------------------------------------------------------
On Error GoTo UpdateChangePasswordIndicatorVbErr
    Const strFunctionName As String = "UpdateChangePasswordIndicator"
    Dim strChangePasswordIndicator As String
    Dim objXMLDocument As FreeThreadedDOMDocument40
    Dim objIADOAssist As IADOAssist
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
      
    Set objIADOAssist = New ADOAssist
    Set objXMLDocument = New FreeThreadedDOMDocument40
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        
    strChangePasswordIndicator = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//CHANGEPASSWORDINDICATOR")
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUserData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    m_objContext.SetComplete
        
        
    Exit Sub
UpdateChangePasswordIndicatorVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function IOrganisationDO_FindUnitNameList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindUnitNameListVbErr
    Const strFunctionName As String = "IOrganisationDO_FindUnitNameList"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strUserId As String
    Dim strSQL As String
    Dim strSQL1 As String
    Dim intUserCount  As Integer
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Check the correct keys have been passed in
    ''''''''''''''''''''''''''''''''''''''''''''
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    strSQL = "SELECT UNITID,UNITNAME FROM UNIT WHERE "
                    
    strSQL1 = CreateSQLFromList(adoCommand, vxmlTableElement, enumOR, "UNITID")
    strSQL = strSQL & strSQL1
    If Len(strSQL1) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "UserID"
    End If
    adoCommand.CommandText = strSQL
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnitNameData()
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("UNITNAMELIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IOrganisationDO_FindUnitNameList = xmlListElem
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
        
    m_objContext.SetComplete
    Exit Function
FindUnitNameListVbErr:
    
    
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
Private Function IOrganisationDO_FindUserNameList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindUserNameListVbErr
    Const strFunctionName As String = "IOrganisationDO_FindUserNameList"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlListElem As IXMLDOMElement
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlNodeList As IXMLDOMNodeList
    Dim strUserId As String
    Dim strSQL As String
    Dim strSQL1 As String
    Dim intUserCount  As Integer
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Check the correct keys have been passed in
    ''''''''''''''''''''''''''''''''''''''''''''
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    strSQL = "SELECT USERID,USERFORENAME,USERSURNAME,USERTITLE FROM ORGANISATIONUSER WHERE "
                    
    strSQL1 = CreateSQLFromList(adoCommand, vxmlTableElement, enumOR, "USERID")
    strSQL = strSQL & strSQL1
    If Len(strSQL1) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "UserID"
    End If
    adoCommand.CommandText = strSQL
    'Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand
    'Populate recordset
    adoRecordset.Open
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUserNameData()
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("OMIGAUSERNAMELIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IOrganisationDO_FindUserNameList = xmlListElem
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
        
    m_objContext.SetComplete
    Exit Function
FindUserNameListVbErr:
    
    
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
Private Function CreateSQLFromList(adoCommand As ADODB.Command, ByVal vxmlTableElement As IXMLDOMElement, operatorType As ConditionalOperator, strTag As String) As String
    Dim intItemCount As Long
    Dim intItemLength As Long
    Dim strSQL As String
    Dim strVal As String
    Dim strCondition As String
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim adoParameter As ADODB.Parameter
        
    If operatorType = enumAND Then
        strCondition = "AND "
    Else
        strCondition = "OR "
    End If
    Set objXmlNodeList = vxmlTableElement.getElementsByTagName(strTag)
    intItemCount = 0
    intItemLength = objXmlNodeList.Length
    While intItemCount < intItemLength
        Set xmlNode = objXmlNodeList.Item(intItemCount)
        strSQL = strSQL & strTag & " = ? "
        strVal = xmlNode.Text
        If intItemCount < intItemLength - 1 Then
            strSQL = strSQL & strCondition
        End If
        Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strVal), strVal)
        adoCommand.Parameters.Append adoParameter
        intItemCount = intItemCount + 1
    Wend
    CreateSQLFromList = strSQL
End Function


