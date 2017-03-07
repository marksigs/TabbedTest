VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CreditCheckDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CreditCheckDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for Credit Check Data Object
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MDC    10/03/00    Created
'MDC    18/05/00    Added DemoOptions methods and used module level ObjectContext
'                   Updated interface
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'LD     07/11/00    Explicity close database connections
'LD     07/11/00        Explicity close recordsets
'SR     08/12/00    included methods for Full Bureau Download
'SR     13/12/00    included methods for Full Bureau Download
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    05/09/2002  BMIDS00336 Credit Check & Bureau Download
'GD     05/03/2003  BM0373 Don't call setAbort() in the case of recordNotFound in all Find/Get methods
'INR    24/05/2004  BMIDS744  Get Credit Check Data for TPD changes.
'INR    12/08/2004  BMIDS744 Added GetCreditCheckStatus
'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'HM     31/08/2004  MAR13 Added GetCurrentXMLCreditCheckResults
'SR     20/10/2005  MAR24 New method FindApplicationCreditCheckDetailList
'INR    24/11/2005  MAR645 Added CreateRuleOverride
'PJO    30/11/2005  MAR741 make user id length 64
'HMA    15/12/2005  MAR550 Changed FindApplicationCreditCheckDetailList
'PE     04/01/2006  MAR945 New private sub - ValidateUser
'PE     27/01/2006  MAR1060 Increased password parameter length to 15
'------------------------------------------------------------------------------------------
'EPSOM History:
'
'Prog   Date        Description
'IK     09/08/2006  EP1032 CreateRuleOverride
'                   - use actual string length for ADO parameter length
'------------------------------------------------------------------------------------------

Option Explicit

Implements ICreditCheckDO
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


Private Function ICreditCheckDO_FindApplicationCreditCheckDetailsList( _
            ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindApplicationCreditCheckDetailsList"
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim Cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim adoStr As ADODB.Stream
    Dim Param As ADODB.Parameter
    
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    
    Dim objResponseElement As IXMLDOMElement
    Dim objResponseNode As IXMLDOMNode
    Dim objADOAssist As ADOAssist

    Set objResponseElement = xmlTempResponseDoc.createElement("RESPONSE")
    Set objResponseNode = xmlTempResponseDoc.appendChild(objResponseElement)
    
    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
 
    Set Cn = New ADODB.Connection
    Cn.ConnectionString = objADOAssist.GetConnStr

    Set cmd = New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETAPPLICATIONCREDITCHECKDETAILS"
    End With
    
    Set Param = New ADODB.Parameter
    With Param
        .Name = "@p_ApplicationNumber"
        .Type = adVarChar
        .Direction = adParamInput
        .Size = 12
        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//APPLICATIONNUMBER")
    End With
    cmd.Parameters.Append Param
    
    Set Param = New ADODB.Parameter
    With Param
        .Name = "@p_ApplicationFactFindNumber"
        .Type = adInteger
        .Direction = adParamInput
        If Len(m_objXmlAssist.GetElementText(vxmlTableElem, "//APPLICATIONFACTFINDNUMBER")) > 0 Then
            .Value = CInt(m_objXmlAssist.GetElementText(vxmlTableElem, "//APPLICATIONFACTFINDNUMBER"))
        End If
    End With
    cmd.Parameters.Append Param
    
    Set adoStr = New ADODB.Stream
    Cn.Open
    adoStr.Open

    cmd.ActiveConnection = Cn
    cmd.Properties("Output Stream") = adoStr
    cmd.Execute , , adExecuteStream

    If xmlTempResponseDoc.loadXML("<APPLICATIONCREDITCHECKLIST>" & adoStr.ReadText() & "</APPLICATIONCREDITCHECKLIST>") Then
         Set ICreditCheckDO_FindApplicationCreditCheckDetailsList = xmlTempResponseDoc.documentElement
    Else
         Set ICreditCheckDO_FindApplicationCreditCheckDetailsList = objResponseNode
    End If
    
FindListExit:
    'MAR550  Change SetAbort to SetComplete
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            Cn.Close
        End If
    End If

    Set Cn = Nothing
    Set cmd = Nothing
    Set adoStr = Nothing
    Set Param = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objResponseNode = Nothing
    Set objADOAssist = Nothing

    Exit Function
    
FindListVbErr:
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            Cn.Close
        End If
    End If

    Set Cn = Nothing
    Set cmd = Nothing
    Set adoStr = Nothing
    Set Param = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objResponseNode = Nothing
    Set objADOAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICreditCheckDO_FindApplicationCreditCheckList( _
                    ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode


On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindApplicationCreditCheckList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationCreditCheckData
    
    Set ICreditCheckDO_FindApplicationCreditCheckList = m_objIDOAssist.FindList(vxmlTableElem, xmlClassDefDoc, "SEQUENCENUMBER")

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'INR 23/11/2005 : MAR645
Private Function ICreditCheckDO_CreateRuleOverride( _
            ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode

On Error GoTo CreateRuleOverrideVbErr

    Const strFunctionName As String = "ICreditCheckDO_CreateRuleOverride"
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim Cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim adoStr As ADODB.Stream
    Dim Param As ADODB.Parameter
    Dim xmlGUIDNode As IXMLDOMNode
    Dim objSQLAssist As New SQLAssist
    
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    
    Dim objResponseElement As IXMLDOMElement
    Dim objResponseNode As IXMLDOMNode
    Dim objADOAssist As ADOAssist
    
    'EP1032
    Dim strValue As String

    Set objResponseElement = xmlTempResponseDoc.createElement("RESPONSE")
    Set objResponseNode = xmlTempResponseDoc.appendChild(objResponseElement)
    
    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")

    Set xmlGUIDNode = vxmlTableElem.selectSingleNode("CREDITCHECKGUID")
    If xmlGUIDNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    'MAR945 - Password validation is not working in Case Assessment Override screen.
    'Validate the user
    Dim strUserId As String
    Dim strUnitId As String
    Dim strPassword As String
    strUserId = m_objXmlAssist.GetElementText(vxmlTableElem, "USERID")
    strUnitId = m_objXmlAssist.GetElementText(vxmlTableElem, "UNITID")
    strPassword = m_objXmlAssist.GetElementText(vxmlTableElem, "USERPASSWORD")
    ValidateUser strUserId, strPassword, strUnitId, 1

    Set Cn = New ADODB.Connection
    Cn.ConnectionString = objADOAssist.GetConnStr

    Set cmd = New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "USP_CCCreateRuleOverride"
    End With

    Dim rCreditCheckGuid() As Byte
    rCreditCheckGuid = objSQLAssist.GuidStringToByteArray(xmlGUIDNode.Text)
    cmd.Parameters.Append cmd.CreateParameter("rawCreditCheckGuid", adVarBinary, adParamInput, UBound(rCreditCheckGuid) + 1, rCreditCheckGuid)
    
    'EP1032
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//REASONCODE")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_ReasonCode", adVarChar, adParamInput, 4, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_ReasonCode", adVarChar, adParamInput, Len(strValue), strValue)
    End If
    
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//USERID")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_UserID", adVarChar, adParamInput, 10, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_UserID", adVarChar, adParamInput, Len(strValue), strValue)
    End If
    
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEREASON")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_SMOverrideReason", adVarChar, adParamInput, 1900, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_SMOverrideReason", adVarChar, adParamInput, Len(strValue), strValue)
    End If
    
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEAPPROVALREFERENCE")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_SMOverrideApprovalReference", adVarChar, adParamInput, 10, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_SMOverrideApprovalReference", adVarChar, adParamInput, Len(strValue), strValue)
    End If
    
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEREASONCODE")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_SMOverrideReasonCode", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_SMOverrideReasonCode", adInteger, adParamInput, , CInt(strValue))
    End If
    
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//USERPASSWORD")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_UserPwd", adVarChar, adParamInput, 10, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_UserPwd", adVarChar, adParamInput, Len(strValue), strValue)
    End If
    
    strValue = m_objXmlAssist.GetElementText(vxmlTableElem, "//UNITID")
    If Len(strValue) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@p_UnitID", adVarChar, adParamInput, 10, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@p_UnitID", adVarChar, adParamInput, Len(strValue), strValue)
    End If
    
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_ReasonCode"
'        .Type = adVarChar
'        .Direction = adParamInput
'        .Size = 4
'        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//REASONCODE")
'    End With
'    cmd.Parameters.Append Param
'
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_UserID"
'        .Type = adVarChar
'        .Direction = adParamInput
'        .Size = 64      'PJO 30/11/2005 MAR741
'        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//USERID")
'    End With
'    cmd.Parameters.Append Param
'
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_SMOverrideReason"
'        .Type = adVarChar
'        .Direction = adParamInput
'        .Size = 1900
'        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEREASON")
'    End With
'    cmd.Parameters.Append Param
'
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_SMOverrideApprovalReference"
'        .Type = adVarChar
'        .Direction = adParamInput
'        .Size = 10
'        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEAPPROVALREFERENCE")
'    End With
'    cmd.Parameters.Append Param
'
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_SMOverrideReasonCode"
'        .Type = adInteger
'        .Direction = adParamInput
'        If Len(m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEREASONCODE")) > 0 Then
'            .Value = CInt(m_objXmlAssist.GetElementText(vxmlTableElem, "//SMOVERRIDEREASONCODE"))
'        End If
'    End With
'    cmd.Parameters.Append Param
'
'
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_UserPwd"
'        .Type = adVarChar
'        .Direction = adParamInput
'        .Size = 15  'PE - 27/01/2006 - MAR1060
'        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//USERPASSWORD")
'    End With
'    cmd.Parameters.Append Param
'
'    Set Param = New ADODB.Parameter
'    With Param
'        .Name = "@p_UnitID"
'        .Type = adVarChar
'        .Direction = adParamInput
'        .Size = 10
'        .Value = m_objXmlAssist.GetElementText(vxmlTableElem, "//UNITID")
'    End With
'    cmd.Parameters.Append Param
    'EP1032_ends

    Set adoStr = New ADODB.Stream
    Cn.Open
    adoStr.Open

    cmd.ActiveConnection = Cn
    cmd.Properties("Output Stream") = adoStr
    cmd.Execute , , adExecuteStream

    If xmlTempResponseDoc.loadXML("<APPLICATIONCREDITCHECKLIST>" & adoStr.ReadText() & "</APPLICATIONCREDITCHECKLIST>") Then
         Set ICreditCheckDO_CreateRuleOverride = xmlTempResponseDoc.documentElement
    Else
         Set ICreditCheckDO_CreateRuleOverride = objResponseNode
    End If
    
    m_objContext.SetComplete
CreateRuleOverrideExit:
      If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            Cn.Close
        End If
    End If

    Set Cn = Nothing
    Set cmd = Nothing
    Set adoStr = Nothing
    Set Param = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objResponseNode = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing

    Exit Function
    
CreateRuleOverrideVbErr:
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If

    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            Cn.Close
        End If
    End If

    Set Cn = Nothing
    Set cmd = Nothing
    Set adoStr = Nothing
    Set Param = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objResponseNode = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'MAR945 - Password validation is not working in Case Assessment Override screen.
Private Sub ValidateUser(ByVal vstrUserID As String, ByVal vstrPassword As String, _
    ByVal vstrUnitID As String, ByVal vstrAuditRecordType As String)

    On Error GoTo ValidateUserVbErr

    Dim strRequest As String
    Dim objOrg As Object

    Const strFunctionName As String = "ValidateUser"
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Set objOrg = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO")
        
    strRequest = _
        "<REQUEST USERID=""" & vstrUserID & """ MACHINEID=""CreditCheckDO"">" & _
            "<OMIGAUSER>" & _
            "<USERID>" & vstrUserID & "</USERID>" & _
            "<PASSWORDVALUE>" & vstrPassword & "</PASSWORDVALUE>" & _
            "<UNITID>" & vstrUnitID & "</UNITID>" & _
            "<AUDITRECORDTYPE>1</AUDITRECORDTYPE>" & _
            "</OMIGAUSER>" & _
        "</REQUEST>"

    If m_objErrAssist.CheckXMLResponse(objOrg.ValidateUserLogon(strRequest)) <> 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 111, "Not authorised"
    End If
        
    m_objContext.SetComplete
        
    If Not objOrg Is Nothing Then Set objOrg = Nothing
    
    Exit Sub

ValidateUserVbErr:
            
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
        
    If Not objOrg Is Nothing Then Set objOrg = Nothing

    ' re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ICreditCheckDO_FindFullBureauAssociationsList( _
                ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


    On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauAssociationsList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauAssociationsData
    
    Set ICreditCheckDO_FindFullBureauAssociationsList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



Private Function ICreditCheckDO_FindFullBureauCAISList( _
            ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauCAISList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauCAISData
    
    Set ICreditCheckDO_FindFullBureauCAISList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



Private Function ICreditCheckDO_FindFullBureauCAPSList( _
                ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

    
    On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauCAPSList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauCAPSData
    
    Set ICreditCheckDO_FindFullBureauCAPSList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



Private Function ICreditCheckDO_FindFullBureauCIFASList( _
            ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauCIFASList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauCIFASData
    
    Set ICreditCheckDO_FindFullBureauCIFASList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



Private Function ICreditCheckDO_FindFullBureauCorrectionLinesList( _
                ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


    On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauCorrectionLinesList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauCorrectionLinesData
    
    Set ICreditCheckDO_FindFullBureauCorrectionLinesList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description


End Function


Private Function ICreditCheckDO_FindFullBureauPublicInfoList( _
                    ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauPublicInfoList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
    Dim xmlAttrib As IXMLDOMAttribute   'BMIDS00336 MDC 11/09/2002
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauPublicInfoData
    
    'BMIDS00336 MDC 11/09/2002 - Enable sorting
    Set xmlAttrib = vobjRequestNode.Attributes.getNamedItem("SORT")
    'Set ICreditCheckDO_FindFullBureauPublicInfoList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)
    If xmlAttrib Is Nothing Then
        Set ICreditCheckDO_FindFullBureauPublicInfoList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)
    Else
        Set ICreditCheckDO_FindFullBureauPublicInfoList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc, xmlAttrib.Text)
    End If
    'BMIDS00336 MDC 11/09/2002 - End

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlAttrib = Nothing 'BMIDS00336 MDC 11/09/2002
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICreditCheckDO_FindFullBureauVotersRollList( _
                ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindFullBureauVotersRollList"
    

    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
         
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauVotersRollData
    
    Set ICreditCheckDO_FindFullBureauVotersRollList = m_objIDOAssist.FindList(vobjRequestNode, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICreditCheckDO_GetCurrentCAISSummary(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo GetCurrentCAISSummaryVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetCurrentCAISSummary"

    Dim objSQLAssist As New SQLAssist
    Dim objADOAssist As ADOAssist
    Dim objClassDefs As ICreditCheckClassDef

    Dim conn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim objXmlSchemaDoc1 As FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
    Dim objResponseNode As IXMLDOMNode
    
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strCreditCheckGuid As String
    Dim strPersonType As String
    Dim strSQL As String

    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    Set objClassDefs = New CreditCheckClassDef
    
    Set objResponseNode = vobjRequestNode.ownerDocument.createElement("RESPONSE")
    
    'Get and (if necessary) convert input parameters
    strCustomerNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERNUMBER")
    strCustomerNumber = objSQLAssist.FormatString(strCustomerNumber)
    strCustomerVersionNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERVERSIONNUMBER")
    strPersonType = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/PERSONTYPE")
    strCreditCheckGuid = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CREDITCHECKGUID")
    strCreditCheckGuid = objSQLAssist.FormatGuid(strCreditCheckGuid)
    
    'Construct SQL string
    strSQL = "SELECT TOTALVALUEACTIVECAIS3M, TOTALACTIVECAIS3M, TOTALVALUEACTIVECAIS4M, "
    strSQL = strSQL & "WORSTSTATUSACTIVECAIS4M, TOTALVALUEACTIVECAIS12M, WORSTSTATUSACTIVECAIS12M "
    strSQL = strSQL & "FROM CREDITCHECKDETAILS WHERE CREDITCHECKGUID = " & strCreditCheckGuid
    strSQL = strSQL & " AND CUSTOMERNUMBER = " & strCustomerNumber
    strSQL = strSQL & " AND CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber
    strSQL = strSQL & " AND PERSONTYPE = " & strPersonType
    
   'Get Database Connection
    conn.ConnectionString = objADOAssist.GetConnStr
    conn.Open
        
    'Open recordset
    rst.CursorLocation = adUseClient
    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    'Disconnect recordset
    Set rst.ActiveConnection = Nothing

    
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    rst.MoveFirst
   
    'Load Class Def for Credit Check Details Data
    Set objXmlSchemaDoc1 = objClassDefs.LoadCreditCheckDetailsData(ccCAIS)
            
    'Get XML from the recordset
    Set objXmlNode = m_objIDOAssist.GetXMLFromRecordSet(rst, objXmlSchemaDoc1, objResponseNode)
    rst.Close
     
    Set ICreditCheckDO_GetCurrentCAISSummary = objResponseNode
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    'BMIDS00336 MDC 05/09/2002
'    rst.Close
'    Set rst = Nothing
'    conn.Close
'    Set conn = Nothing
    ' Close the recordset if still open
    If Not rst Is Nothing Then
            If rst.State = adStateOpen Then
                    rst.Close
            End If
    End If

    ' Close the connection if still open
    If Not conn Is Nothing Then
            If conn.State = adStateOpen Then
                    conn.Close
            End If
    End If
    
    Set rst = Nothing
    Set conn = Nothing
    'BMIDS00336 MDC 05/09/2002 - End
    
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    Set objClassDefs = Nothing
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlNode = Nothing
    
    Exit Function
    
GetCurrentCAISSummaryVbErr:

    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    ' Close the connection if still open
        If Not conn Is Nothing Then
                If conn.State = adStateOpen Then
                        conn.Close
                End If
        End If

    Set rst = Nothing
    Set conn = Nothing
    
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    Set objClassDefs = Nothing
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlNode = Nothing

    If Not m_objContext Is Nothing Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function ICreditCheckDO_GetCurrentCreditCheckResults(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo GetCurrentCreditCheckResultsVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetCurrentCreditCheckResults"
    
    Dim objSQLAssist As New SQLAssist

    Dim objXmlSchemaDoc1 As FreeThreadedDOMDocument40
    
    Dim objXMLElement As IXMLDOMElement
    Dim objXmlNodeList As IXMLDOMNodeList
    
    Dim objXmlParentNode As IXMLDOMNode
    Dim objXmlDetailListNode As IXMLDOMNode
    Dim objXmlMiscListNode As IXMLDOMNode
    Dim objXmlChildNode As IXMLDOMNode
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objResponseElement As IXMLDOMElement
    Dim objResponseNode As IXMLDOMNode
    
    Dim objClassDefs As ICreditCheckClassDef
    
    Dim objADOAssist As ADOAssist
   
    Dim conn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim strApplicationNumber As String, strApplicationFactFindNumber As String
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strCustomerOrder As String
    Dim strCustomerName As String
    Dim strCreditCheckGuid As String
    Dim strSQL As String
    Dim bytCreditCheckGUID() As Byte

    
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    Set objClassDefs = New CreditCheckClassDef
    
    Set objResponseElement = objXmlOut.createElement("RESPONSE")
    Set objResponseNode = objXmlOut.appendChild(objResponseElement)
    
   'Get and format App Number and App Fact Find Number
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/APPLICATIONNUMBER")
    strApplicationNumber = objSQLAssist.FormatString(strApplicationNumber)
    strApplicationFactFindNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/APPLICATIONFACTFINDNUMBER")
    strApplicationFactFindNumber = objSQLAssist.FormatString(strApplicationFactFindNumber)
   
   'Get Database Connection
    conn.ConnectionString = objADOAssist.GetConnStr
    conn.Open
    
    rst.CursorLocation = adUseClient

    'BMIDS744 get new field, TPDOUTCOMECODE.
    'Construct SQL string
    strSQL = "SELECT USERID, DATETIME, CREDITCHECKGUID, TPDOUTCOMECODE FROM APPLICATIONCREDITCHECK "
    strSQL = strSQL & "WHERE APPLICATIONNUMBER = " & strApplicationNumber
    strSQL = strSQL & "AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
    strSQL = strSQL & "AND SEQUENCENUMBER = "
    strSQL = strSQL & "(SELECT MAX(SEQUENCENUMBER) FROM APPLICATIONCREDITCHECK "
    strSQL = strSQL & "WHERE APPLICATIONNUMBER= " & strApplicationNumber
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber & ")"

    'Open recordset
    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    Set rst.ActiveConnection = Nothing
    
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "No credit check exists for this application."
    End If
    
    rst.MoveFirst
    
    'If no CreditCheckGUID, return Credit Check failed error 276
    If IsNull(rst("CreditCheckGUID")) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 276
    End If
        
    'Load Class Def for Application Credit Check Data
    Set objXmlSchemaDoc1 = objClassDefs.LoadApplicationCreditCheckData(ccCurrent)
            
    'Get XML from the recordset
    Set objXmlParentNode = m_objIDOAssist.GetXMLFromRecordSet(rst, objXmlSchemaDoc1, objResponseNode)
        
    Set objXmlSchemaDoc1 = Nothing
    
    'Get CreditCheckGuid and convert to string for SQL query string
    bytCreditCheckGUID = rst("CreditCheckGuid")
    strCreditCheckGuid = objSQLAssist.GuidToString(bytCreditCheckGUID)
    strCreditCheckGuid = objSQLAssist.FormatGuid(strCreditCheckGuid)
        
    'Close the recordset
    rst.Close
    
    strCustomerOrder = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERORDER")
    strSQL = "select cr.customernumber, cr.customerversionnumber, cv.surname, cv.firstforename "
    strSQL = strSQL & "from customerrole cr, customerversion cv "
    strSQL = strSQL & "where cr.customernumber = cv.customernumber "
    strSQL = strSQL & "and cr.customerversionnumber = cv.customerversionnumber "
    strSQL = strSQL & "and applicationnumber = " & strApplicationNumber
    strSQL = strSQL & " and applicationfactfindnumber = " & strApplicationFactFindNumber
    strSQL = strSQL & " and customerorder = " & strCustomerOrder

    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    Set rst.ActiveConnection = Nothing
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "Unable to find customer number."
    End If
    
    'Get and convert CustomerNumber and CustomerVersionNumber
    rst.MoveFirst
    strCustomerName = rst("firstforename") & " " & rst("surname")
    strCustomerNumber = rst("customernumber") 'm_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERNUMBER")
    strCustomerNumber = objSQLAssist.FormatString(strCustomerNumber)
    strCustomerVersionNumber = rst("customerversionnumber") 'm_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERVERSIONNUMBER")
    strCustomerVersionNumber = objSQLAssist.FormatString(strCustomerVersionNumber)
    
    rst.Close

    'Construct SQL string for query
    'Retrieve all records matching CreditCheckGUID and check Customer Numbers afterwards
    strSQL = "SELECT CREDITCHECKMISCELLANEOUS.SCORECARDID1, CREDITCHECKMISCELLANEOUS.SCORE1, "
    strSQL = strSQL & "CREDITCHECKDETAILS.CUSTOMERNUMBER, CREDITCHECKDETAILS.CUSTOMERVERSIONNUMBER, "
    strSQL = strSQL & "CREDITCHECKDETAILS.PERSONTYPE, CREDITCHECKDETAILS.TOTALCCJS, "
    strSQL = strSQL & "CREDITCHECKDETAILS.TOTALVALUECCJS, CREDITCHECKDETAILS.AGEOFLASTCCJ, "
    strSQL = strSQL & "CREDITCHECKDETAILS.TOTALOUTSTANDINGCCJ, CREDITCHECKDETAILS.TOTALCAIS89, "
    strSQL = strSQL & "CREDITCHECKDETAILS.TOTALVALUECAIS89, CREDITCHECKDETAILS.AGEOFLASTCAIS89, "
    strSQL = strSQL & "CREDITCHECKDETAILS.TOTALVALUEACTIVECAISINCMORT, CREDITCHECKDETAILS.WORSTSTATUSMORTGAGE6M, "
    strSQL = strSQL & "CREDITCHECKDETAILS.MORTGAGEACCOUNTTIME, CREDITCHECKDETAILS.WORSTSTATUSOC, "
    strSQL = strSQL & "CREDITCHECKDETAILS.TOTALVALUEOC, CREDITCHECKDETAILS.VOTERSROLLPAF, "
    strSQL = strSQL & "CREDITCHECKCUSTOMERMISC.NAMESCONFIRMED "
    strSQL = strSQL & "FROM CREDITCHECKDETAILS, CREDITCHECKCUSTOMERMISC, CREDITCHECKMISCELLANEOUS "
    strSQL = strSQL & "WHERE CREDITCHECKMISCELLANEOUS.CREDITCHECKGUID = CREDITCHECKDETAILS.CREDITCHECKGUID "
    strSQL = strSQL & "AND CREDITCHECKMISCELLANEOUS.CREDITCHECKGUID = CREDITCHECKCUSTOMERMISC.CREDITCHECKGUID "
    strSQL = strSQL & "AND CREDITCHECKDETAILS.CUSTOMERNUMBER = CREDITCHECKCUSTOMERMISC.CUSTOMERNUMBER "
    strSQL = strSQL & "AND CREDITCHECKDETAILS.CUSTOMERVERSIONNUMBER = CREDITCHECKCUSTOMERMISC.CUSTOMERVERSIONNUMBER "
    strSQL = strSQL & "AND CREDITCHECKMISCELLANEOUS.CREDITCHECKGUID = " & strCreditCheckGuid
    strSQL = strSQL & " AND CREDITCHECKDETAILS.CUSTOMERNUMBER = " & strCustomerNumber
    strSQL = strSQL & " AND CREDITCHECKDETAILS.CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber
    strSQL = strSQL & " ORDER BY CREDITCHECKDETAILS.PERSONTYPE"
   
    'Open recordset
    rst.Open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    ' disconnect RecordSet
    Set rst.ActiveConnection = Nothing
    
    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "No Credit Check Summary data available for this application."
    Else
        rst.MoveFirst
        'Credit Check Miscellaneous Data
        Set objXmlParentNode = objResponseNode.selectSingleNode("APPLICATIONCREDITCHECK")
        Set objXmlSchemaDoc1 = objClassDefs.LoadCreditCheckMiscellaneousData(ccCurrent)
        Set objXmlChildNode = m_objIDOAssist.GetXMLFromRecordSet(rst, objXmlSchemaDoc1, objXmlParentNode)
        Set objXmlSchemaDoc1 = Nothing
                       
        'Credit Check Details Data
        Set objXmlSchemaDoc1 = objClassDefs.LoadCreditCheckDetailsData(ccCurrent)
        Set objXmlChildNode = m_objIDOAssist.GetXMLFromRecordSet(rst, objXmlSchemaDoc1, objXmlParentNode)
        
        Set objXmlSchemaDoc1 = Nothing
        
        'Credit Check Customer Misc Data
        Set objXmlSchemaDoc1 = objClassDefs.LoadCreditCheckCustomerMiscData(ccCurrent)
        Set objXmlChildNode = m_objIDOAssist.GetXMLFromRecordSet(rst, objXmlSchemaDoc1, objXmlParentNode)
                       
        rst.Close
        Set objXmlSchemaDoc1 = Nothing
        
        'Customer Name
        Set objXMLElement = objXmlParentNode.ownerDocument.createElement("CUSTOMERNAME")
        objXMLElement.Text = strCustomerName
        Set objXmlChildNode = objXmlParentNode.appendChild(objXMLElement)
    End If
    
    Set ICreditCheckDO_GetCurrentCreditCheckResults = objResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    'BMIDS00336 MDC 05/09/2002
'    rst.Close
'    Set rst = Nothing
'    conn.Close
'    Set conn = Nothing
    ' Close the recordset if still open
    If Not rst Is Nothing Then
            If rst.State = adStateOpen Then
                    rst.Close
            End If
    End If

    ' Close the connection if still open
    If Not conn Is Nothing Then
            If conn.State = adStateOpen Then
                    conn.Close
            End If
    End If
    
    Set rst = Nothing
    Set conn = Nothing
    'BMIDS00336 MDC 05/09/2002 - End
    
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    Set objClassDefs = Nothing
    Set objXmlParentNode = Nothing
    Set objXMLElement = Nothing
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlChildNode = Nothing
    Exit Function
    
GetCurrentCreditCheckResultsVbErr:

    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    ' Close the connection if still open
        If Not conn Is Nothing Then
                If conn.State = adStateOpen Then
                        conn.Close
                End If
        End If

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set rst = Nothing
    Set conn = Nothing
    
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    Set objClassDefs = Nothing
    Set objXmlParentNode = Nothing
    Set objXMLElement = Nothing
    Set objXmlSchemaDoc1 = Nothing
    Set objXmlChildNode = Nothing

    If Not m_objContext Is Nothing Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    End If

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Function ICreditCheckDO_FindCurrentFullBureauResultsList( _
                ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode


On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICreditCheckDO_FindCurrentFullBureauResultsList"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
                        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
    
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFBResultsData()
    Set ICreditCheckDO_FindCurrentFullBureauResultsList = m_objIDOAssist.FindList(vxmlTableElem, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICreditCheckDO_GetCurrentXMLCreditCheckResults(ByVal vobjRequestNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode

On Error GoTo GetCurrentXMLCreditCheckResultsVbErr

    Const strFunctionName   As String = "ICreditCheckDO_GetCurrentXMLCreditCheckResults"
    
    Dim objSQLAssist        As New SQLAssist
    Dim objADOAssist        As New ADOAssist

    Dim adoConn             As ADODB.Connection
    Dim adoCmd              As ADODB.Command
    Dim adoStr              As ADODB.Stream
    Dim adoParam            As ADODB.Parameter
    
    Dim objXmlOut           As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    
    Dim objResponseElement  As IXMLDOMElement
    Dim objResponseNode     As IXMLDOMNode

    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strCustomerNumber1 As String
    Dim strCustomerNumber2 As String
    Dim strCustomerVersionNumber1 As String
    Dim strCustomerVersionNumber2 As String
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    'create response object
    Set objResponseElement = objXmlOut.createElement("RESPONSE")
    Set objResponseNode = objXmlOut.appendChild(objResponseElement)
    
    'Get Context Data: App Number, App Fact Find Number, CustNo 1 & 2, CustVer 1 & 2
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/APPLICATIONFACTFINDNUMBER")
    strCustomerNumber1 = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERNUMBER1")
    strCustomerVersionNumber1 = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/CUSTOMERVERSIONNUMBER1")
    strCustomerNumber2 = m_objXmlAssist.GetElementText(vobjRequestNode, "SEARCH/CUSTOMERNUMBER2")
    strCustomerVersionNumber2 = m_objXmlAssist.GetElementText(vobjRequestNode, "SEARCH/CUSTOMERVERSIONNUMBER2")
      
    'Get Database Connection
    Set adoConn = New ADODB.Connection
    adoConn.ConnectionString = objADOAssist.GetConnStr
    'adoConn.CursorLocation = adUseClient

    Set adoCmd = New ADODB.Command
    With adoCmd
        .CommandType = adCmdStoredProc
        .CommandText = "usp_GetCreditCheckResult"
    End With
    
    'create sp parameters
    Set adoParam = adoCmd.CreateParameter(, adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCmd.Parameters.Append adoParam
    
    Set adoParam = adoCmd.CreateParameter(, adBSTR, adParamInput, Len(strApplicationFactFindNumber), strApplicationFactFindNumber)
    adoCmd.Parameters.Append adoParam
    
    Set adoParam = adoCmd.CreateParameter(, adBSTR, adParamInput, Len(strCustomerNumber1), strCustomerNumber1)
    adoCmd.Parameters.Append adoParam
    
    Set adoParam = adoCmd.CreateParameter(, adBSTR, adParamInput, Len(strCustomerVersionNumber1), strCustomerVersionNumber1)
    adoCmd.Parameters.Append adoParam
    
    'optional params
    If Len(strCustomerNumber2) > 0 Then
        Set adoParam = adoCmd.CreateParameter(, adBSTR, adParamInput, Len(strCustomerNumber2), strCustomerNumber2)
        adoCmd.Parameters.Append adoParam

        Set adoParam = adoCmd.CreateParameter(, adBSTR, adParamInput, Len(strCustomerVersionNumber2), strCustomerVersionNumber2)
        adoCmd.Parameters.Append adoParam
    End If
    
    'open connection
    Set adoStr = New ADODB.Stream
    adoConn.Open
    adoStr.Open

    'execute sp
    adoCmd.ActiveConnection = adoConn
    adoCmd.Properties("Output Stream") = adoStr
    adoCmd.Properties("xml root") = "RESPONSE" 'xml root'
    adoCmd.Execute , , adExecuteStream

    ' Raise errors if any while executing the stored procedure
    If adoConn.Errors.Count <> 0 Then
        Err.Raise adoConn.Errors.Item(0).Number, _
                  adoConn.Errors.Item(0).Source, _
                  adoConn.Errors.Item(0).Description
    End If
          
    'create response object
    Set objXmlOut = New FreeThreadedDOMDocument40
    
    'debug to check whenever the response is well formatted
    Dim strTemp As String
    Dim i As Integer
    strTemp = adoStr.ReadText()
    Debug.Print strTemp
    
    'prepare response xml
    If objXmlOut.loadXML(strTemp) Then
        If objXmlOut.documentElement.childNodes.length > 0 Then
            For i = 0 To objXmlOut.documentElement.childNodes.length - 1
                objResponseNode.appendChild objXmlOut.documentElement.childNodes(0)
            Next i
        End If
    End If
    
    'send back response
    Set ICreditCheckDO_GetCurrentXMLCreditCheckResults = objResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
TheEnd:
    
    ' Close the connection if still open
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
                adoConn.Close
        End If
    End If
    
    Set adoConn = Nothing
    Set adoCmd = Nothing
    Set adoStr = Nothing
    Set adoParam = Nothing

    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    Set objResponseElement = Nothing
    Set objResponseNode = Nothing
    Set objXmlOut = Nothing
    
    Exit Function
    
GetCurrentXMLCreditCheckResultsVbErr:

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
    End If

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

    Resume TheEnd
End Function

Private Function ICreditCheckDO_GetDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode


Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim objADOAssist As New ADOAssist
Dim objSQLAssist As New SQLAssist

Dim objXMLElement As IXMLDOMElement
Dim objXmlChildNode As IXMLDOMNode
Dim objClassDefs As ICreditCheckClassDef 'IExperianClassDef

Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim objResponseElement As IXMLDOMElement
Dim objResponseNode As IXMLDOMNode

Dim objXmlSchemaDoc1 As FreeThreadedDOMDocument40
Dim objXmlNode As IXMLDOMNode

Dim strSQL As String
Dim sValue As String

On Error GoTo GetDemoOptionsVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetDemoOptions"

    Set objClassDefs = New CreditCheckClassDef 'ExperianClassDef
    
    Set objResponseElement = objXmlOut.createElement("RESPONSE")
    Set objResponseNode = objXmlOut.appendChild(objResponseElement)
    
    'Check if using demo mode
    Set conn = New ADODB.Connection
    conn.ConnectionString = objADOAssist.GetConnStr
    conn.Open
    
    'Open recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "SELECT * FROM CreditCheckDemo", conn, adOpenStatic, adLockReadOnly, adCmdText
    
    'Disconnect recordset
    Set rst.ActiveConnection = Nothing


    If rst.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    rst.MoveFirst
   
    'Load Class Def for Credit Check Details Data
    Set objXmlSchemaDoc1 = objClassDefs.LoadDemoOptionsData()
            
    'Get XML from the recordset
    Set objXmlNode = m_objIDOAssist.GetXMLFromRecordSet(rst, objXmlSchemaDoc1, objResponseNode)
        
    rst.Close
        
    Set ICreditCheckDO_GetDemoOptions = objResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
   
    Set rst = Nothing
        conn.Close
    Set conn = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Exit Function
        
GetDemoOptionsVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

        ' Close the connection if still open
        If Not conn Is Nothing Then
                If conn.State = adStateOpen Then
                        conn.Close
                End If
        End If

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If

    Set rst = Nothing
        Set conn = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description


End Function


Private Function ICreditCheckDO_GetFullBureauDataHeaderData( _
                ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode


On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetFullBureauDataHeaderData"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
                        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
    
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauDataHeaderData()
    Set ICreditCheckDO_GetFullBureauDataHeaderData = m_objIDOAssist.GetData(vxmlTableElem, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetDataVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICreditCheckDO_GetFullBureauRequest( _
                    ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetFullBureauRequest"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement

    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As SQLAssist   'BMIDS00336 MDC 06/09/2002
    
    Dim strUnitId As String, strUserId As String, strApplicationNumber As String, strSequenceNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                                                                
    strUnitId = m_objXmlAssist.GetElementText(vXMLRequest, "UNITID")
    If Len(Trim$(strUnitId)) = 0 Then
        'Get the UnitID attribute from the REQUEST node
        strUnitId = m_objXmlAssist.GetAttributeValue(vXMLRequest.ownerDocument.documentElement, "REQUEST", "UNITID")
    End If
                                                                   
    strUserId = m_objXmlAssist.GetElementText(vXMLRequest, "USERID")
    If Len(Trim$(strUserId)) = 0 Then
        'Get the UserID attribute from the REQUEST node
        strUserId = m_objXmlAssist.GetAttributeValue(vXMLRequest.ownerDocument.documentElement, "REQUEST", "USERID")
        If Len(Trim$(strUserId)) = 0 Then
            'Get default user ID from global parameters
            strUserId = GetDefaultUserId()
        End If
    End If
                                                     
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vXMLRequest, "APPLICATIONNUMBER")
    strSequenceNumber = m_objXmlAssist.GetMandatoryElementText(vXMLRequest, "SEQUENCENUMBER")
    
    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set objSQLAssist = New SQLAssist    'BMIDS00336 MDC 06/09/2002
    
    Dim adoParameter As ADODB.Parameter
    Dim ADOConnection As New ADODB.Connection
    ADOConnection.ConnectionString = objADOAssist.GetConnStr
    Set objADOAssist = Nothing
    ADOConnection.CursorLocation = adUseClient
    ADOConnection.Open
    
    Dim adoCommand As New ADODB.Command
    Set adoCommand.ActiveConnection = ADOConnection
    adoCommand.CommandType = adCmdStoredProc
    'BMIDS00336 MDC 06/09/2002
    'adoCommand.CommandText = "sp_FullBureau.GetFullBureauRequest"
    adoCommand.CommandText = "usp_GetFullBureauRequest"
    'BMIDS00336 MDC 06/09/2002 - End
    
    'BMIDS00336 MDC 06/09/2002
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamOutput, 1000) ' TODO - LENGTH REQUIRED WHEN CONVERTED TO SQL SERVER.
    'Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamOutput) ' TODO - LENGTH REQUIRED WHEN CONVERTED TO SQL SERVER.
    'BMIDS00336 MDC 06/09/2002 - End
    adoCommand.Parameters.Append adoParameter
    
    'BMIDS00336 MDC 06/09/2002
    Set adoParameter = adoCommand.CreateParameter(, adBinary, adParamOutput, 16)
    'Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamOutput) ' TODO - LENGTH REQUIRED WHEN CONVERTED TO SQL SERVER.
    'BMIDS00336 MDC 06/09/2002 - End
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strUserId), strUserId)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strUnitId), strUnitId)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, Len(strSequenceNumber), strSequenceNumber)
    adoCommand.Parameters.Append adoParameter
    
    adoCommand.Execute
    ' Raise errors if any while executing the stored procedure
    If ADOConnection.Errors.Count <> 0 Then
        Err.Raise ADOConnection.Errors.Item(0).Number, _
                  ADOConnection.Errors.Item(0).Source, _
                  ADOConnection.Errors.Item(0).Description
    End If
       
    ' Add the output from stored proc to Respone
    Set xmlElement = xmlOut.createElement("FBREQUEST")
    xmlElement.Text = adoCommand(0)
    xmlResponseElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("CREDITCHECKGUID")
    'BMIDS00336 MDC 06/09/2002
    'xmlElement.Text = adoCommand(1)
    xmlElement.Text = objSQLAssist.ByteArrayToGuidString(adoCommand(1).Value)
    'BMIDS00336 MDC 06/09/2002 - End
    xmlResponseElem.appendChild xmlElement
        
    Set ICreditCheckDO_GetFullBureauRequest = xmlResponseElem

    m_objContext.SetComplete
    
    ' Clear all the objects
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing  'BMIDS00336 MDC 06/09/2002
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    
    Set adoParameter = Nothing
    Set ADOConnection = Nothing
    Set adoCommand = Nothing
    
    
    Exit Function

GetDataVbErr:
    
        
    ' Clear all the objects
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing  'BMIDS00336 MDC 06/09/2002
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    
    Set adoParameter = Nothing
    Set ADOConnection = Nothing
    Set adoCommand = Nothing
         
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function GetDefaultUserId() As String

On Error GoTo GetDefaultUserIdVbErr
    
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    GetDefaultUserId = objGlobalParamDO.GetCurrentParameterByType("DefaultUserId", "STRING")
    
GetDefaultUserIdVbExit:
    Set objGlobalParamDO = Nothing
    Exit Function
    
GetDefaultUserIdVbErr:
    
    'If error occurs return empty string
    GetDefaultUserId = ""
    GoTo GetDefaultUserIdVbExit
    
End Function

Private Function ICreditCheckDO_GetFullBureauStandardHeaderData( _
                    ByVal vxmlTableElem As IXMLDOMNode) As IXMLDOMNode


On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetFullBureauStandardHeaderData"
    
    
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
                        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICreditCheckClassDef
    
    Set objIClassDef = New CreditCheckClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFullBureauStandardHeaderData()
    Set ICreditCheckDO_GetFullBureauStandardHeaderData = m_objIDOAssist.GetData(vxmlTableElem, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetDataVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        'GD BM0373 START
        If m_objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeRecordNotFound Then
            m_objContext.SetAbort
        End If
        'GD BM0373 END
    Else
        m_objContext.SetComplete
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICreditCheckDO_UpdateDemoOptions(ByVal vobjRequestNode As IXMLDOMNode) As IXMLDOMNode

Dim conn As ADODB.Connection
Dim objADOAssist As New ADOAssist
Dim objSQLAssist As New SQLAssist

Dim objXMLElement As IXMLDOMElement
Dim objXmlChildNode As IXMLDOMNode
Dim objXmlOut As New FreeThreadedDOMDocument40
objXmlOut.validateOnParse = False
objXmlOut.setProperty "NewParser", True
Dim objResponseElement As IXMLDOMElement
Dim objResponseNode As IXMLDOMNode

Dim strSQL As String
Dim sValue As String

On Error GoTo UpdateDemoOptionsVbErr

    Const strFunctionName As String = "ICreditCheckDO_UpdateDemoOptions"

    Set objResponseElement = objXmlOut.createElement("RESPONSE")
    Set objResponseNode = objXmlOut.appendChild(objResponseElement)

    'Construct SQL string
    strSQL = "INSERT INTO CreditCheckDemo VALUES ("
    sValue = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "UPDATE/DEMOOPTIONS/APPLICATIONNUMBER")
    sValue = objSQLAssist.FormatString(sValue)
    strSQL = strSQL & sValue & ", "
    sValue = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "UPDATE/DEMOOPTIONS/APPLICATIONFACTFINDNUMBER")
    strSQL = strSQL & sValue & ", "
    sValue = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "UPDATE/DEMOOPTIONS/SEQUENCENUMBER")
    strSQL = strSQL & sValue & ", "
    sValue = m_objXmlAssist.GetMandatoryElementText(vobjRequestNode, "UPDATE/DEMOOPTIONS/DEMOMODEIND")
    strSQL = strSQL & sValue & ")"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = objADOAssist.GetConnStr
    conn.Open
    
    'Remove existing records
    conn.Execute "DELETE FROM CreditCheckDemo"
    
    'Add new record
    conn.Execute strSQL
    
    conn.Close
    
    Set objXMLElement = objXmlOut.createElement("SUCCESSINDICATOR")
    objXMLElement.Text = CStr(True)
    Set objXmlChildNode = objResponseNode.appendChild(objXMLElement)
    
    Set ICreditCheckDO_UpdateDemoOptions = objResponseNode
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
   
    Set conn = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Exit Function
        
UpdateDemoOptionsVbErr:

    ' Close the connection if still open
        If Not conn Is Nothing Then
                If conn.State = adStateOpen Then
                        conn.Close
                End If
        End If

    If m_objErrAssist.IsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
   
    Set conn = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Exit Function
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'BMIDS744
Private Function ICreditCheckDO_GetCreditCheckStatus(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo GetCreditCheckStatusVbErr

    Const strFunctionName As String = "ICreditCheckDO_GetCreditCheckStatus"

    Dim adoConn As ADODB.Connection
    Dim adoCmd As ADODB.Command
    Dim adoStr As ADODB.Stream
    Dim adoParam As ADODB.Parameter
    
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    
    Dim objResponseElement As IXMLDOMElement
    Dim objResponseNode As IXMLDOMNode
    Dim objADOAssist As ADOAssist
    Dim strFactFindInt As String

    Set objResponseElement = xmlTempResponseDoc.createElement("RESPONSE")
    Set objResponseNode = xmlTempResponseDoc.appendChild(objResponseElement)
    
    Set objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")


    Set adoConn = New ADODB.Connection
    adoConn.ConnectionString = objADOAssist.GetConnStr

    Set adoCmd = New ADODB.Command
    With adoCmd
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GetCreditCheckStatus"
    End With

    Set adoParam = New ADODB.Parameter
    With adoParam
        .Name = "pAPPLICATIONNUMBER"
        .Type = adVarChar
        .Direction = adParamInput
        .Size = 12
        .Value = m_objXmlAssist.GetElementText(vxmlRequestNode, "//APPLICATIONNUMBER")
    End With
    adoCmd.Parameters.Append adoParam
    
    Set adoParam = New ADODB.Parameter
    With adoParam
        .Name = "pAPPLICATIONFACTFINDNUMBER"
        .Type = adInteger
        .Direction = adParamInput
        If Len(m_objXmlAssist.GetElementText(vxmlRequestNode, "//APPLICATIONFACTFINDNUMBER")) > 0 Then
            .Value = CInt(m_objXmlAssist.GetElementText(vxmlRequestNode, "//APPLICATIONFACTFINDNUMBER"))
        End If
    End With
    adoCmd.Parameters.Append adoParam

    Set adoStr = New ADODB.Stream
    adoConn.Open
    adoStr.Open

    adoCmd.ActiveConnection = adoConn
    adoCmd.Properties("Output Stream") = adoStr
    adoCmd.Execute , , adExecuteStream

    If xmlTempResponseDoc.loadXML(adoStr.ReadText()) Then
        objResponseNode.appendChild xmlTempResponseDoc.documentElement
    End If
    
    Set ICreditCheckDO_GetCreditCheckStatus = objResponseNode

GetCreditCheckStatusExit:

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If

    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
            adoConn.Close
        End If
    End If

    Set adoConn = Nothing
    Set adoCmd = Nothing
    Set adoStr = Nothing
    Set adoParam = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objResponseNode = Nothing
    Set objADOAssist = Nothing

    Exit Function
    
GetCreditCheckStatusVbErr:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If

    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
            adoConn.Close
        End If
    End If

    Set adoConn = Nothing
    Set adoCmd = Nothing
    Set adoStr = Nothing
    Set adoParam = Nothing
    Set xmlTempResponseDoc = Nothing
    Set objResponseNode = Nothing
    Set objADOAssist = Nothing
    Set m_objXmlAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

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

