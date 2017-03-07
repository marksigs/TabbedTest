VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LockManagerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LockManagerBO.cls
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:   A business object that co-ordinates the calls to create and
'               release locks for the customer and application.
'               Note: XML is not used internally within this module, but only when interfacing
'               to other components that have an XML based interface.
'Dependencies:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AS     05/04/2005  Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private m_objContext As ObjectContext

' TODO Replace with real error number.
Private Const gstrLockManagerErrorNumber As Integer = 8009

Public Function omRequest(ByVal vstrXmlIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Executes a application/customer lock request.
'   vstrRequest
'       xml Request data stream
'       Format:
'           <REQUEST
'               USERID=""
'               UNITID=""
'               MACHINEID=""
'               CHANNELID=""
'               LOCKING=""
'               APPLICATIONNUMBER=""
'               OPERATION="LOCK|UNLOCK">
'           </REQUEST>
On Error GoTo omRequestExit
    
    Const cstrFunctionName As String = "omLockManager.LockManagerBO.omRequest"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    Debug.Print vstrXmlIn
    xmlInDoc.loadXML vstrXmlIn
    
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")

    Dim strUserId As String
    Dim strUnitID As String
    Dim strMachineID As String
    Dim strChannelID As String
    Dim strApplicationNumber As String
    Dim strCustomerNumber As String
    Dim strOperation As String
    Dim blnEnableLocking As Boolean
    Dim blnCreateNew As Boolean
    Dim strLocks As String

    strUserId = xmlGetMandatoryAttributeText(xmlRequestNode, "USERID")
    strUnitID = xmlGetMandatoryAttributeText(xmlRequestNode, "UNITID")
    strMachineID = xmlGetMandatoryAttributeText(xmlRequestNode, "MACHINEID")
    strChannelID = xmlGetMandatoryAttributeText(xmlRequestNode, "CHANNELID")
    strApplicationNumber = xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER")
    strCustomerNumber = xmlGetAttributeText(xmlRequestNode, "CUSTOMERNUMBER")
    strOperation = xmlGetMandatoryAttributeText(xmlRequestNode, "OPERATION")

    ' For MAPS this is the global parameter "Locking"
    ' For POS this is the global parameter "POSLocking"
    ' The caller (GUI etc) decides which parameter to use, and passes it in as the "ENABLELOCKING" attribute.
    blnEnableLocking = xmlGetAttributeAsBoolean(xmlRequestNode, "ENABLELOCKING", "1")
    
    ' If True then there must not already be an APPLICATIONLOCK for this application.
    ' If False then there must not already be an APPLICATIONLOCK for this application,
    ' unless it is for this user.
    blnCreateNew = xmlGetAttributeAsBoolean(xmlRequestNode, "CREATENEW", "1")
    
    If ValidateWorkingHours(strUserId, strChannelID) Then
        If blnEnableLocking Then
            If strOperation = "LOCK" Then
                If strApplicationNumber <> "" Then
                    blnSuccess = LockApplication(strApplicationNumber, strUserId, strUnitID, strMachineID, blnCreateNew, strLocks)
                ElseIf strCustomerNumber <> "" Then
                    blnSuccess = LockCustomer(strCustomerNumber, strUserId, strUnitID, strMachineID, strLocks)
                Else
                    Call errThrowError(cstrFunctionName, oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified")
                End If
            ElseIf strOperation = "UNLOCK" Then
                If strApplicationNumber <> "" Then
                    blnSuccess = UnlockApplication(strApplicationNumber)
                ElseIf strCustomerNumber <> "" Then
                    blnSuccess = UnlockCustomer(strCustomerNumber)
                Else
                    Call errThrowError(cstrFunctionName, oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified")
                End If
            Else
                Call errThrowError(cstrFunctionName, oeInvalidParameter, "OPERATION attribute in REQUEST tag must be either ""LOCK"" or ""UNLOCK""")
            End If
        Else
            blnSuccess = True
        End If
    End If
        
    omRequest = xmlOutDoc.xml
        
omRequestExit:
    
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
        
    If Err.Number = 0 Then
        ' No error.
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    Else
        ' Error.
        Err.Source = cstrFunctionName & ", " & Err.Source
        
        omRequest = CreateErrorResponse(strLocks)
        
        If Not m_objContext Is Nothing Then
            m_objContext.SetAbort
        End If
    End If
        
    If strApplicationNumber <> "" Then
        ' Ignore any errors that occur when auditing.
        On Error Resume Next
        ' Only audit if locking an application.
        ' The audit record must be created after any error handling
        ' (as calling CreateApplicationAccessAudit will reset the Err object).
        Call CreateApplicationAccessAudit(blnSuccess, strUserId, strMachineID, strApplicationNumber)
    End If
    
    Debug.Print omRequest
    
End Function

Private Function CreateErrorResponse(strLocks As String) As String
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlLocksDoc As FreeThreadedDOMDocument40
    Dim xmlErrorElem As IXMLDOMElement
    Dim xmlLocksElem As IXMLDOMElement
    
    Set xmlResponseElem = CreateErrorResponseNode(False)
    
' Must be after calling CreateErrorResponseNode, as "On Error" clears the error object.
On Error GoTo CreateErrorResponseExit
    
    If Len(strLocks) > 0 Then
        
        Set xmlLocksDoc = New FreeThreadedDOMDocument40
        xmlLocksDoc.setProperty "NewParser", True
        xmlLocksDoc.async = False
        xmlLocksDoc.loadXML strLocks
        
        Set xmlErrorElem = xmlResponseElem.selectSingleNode("ERROR")
        If Not xmlErrorElem Is Nothing Then
            Set xmlLocksElem = xmlLocksDoc.selectSingleNode("LOCKS")
            If Not xmlLocksElem Is Nothing Then
                xmlErrorElem.appendChild xmlLocksElem.cloneNode(True)
            End If
        End If
    End If
    
CreateErrorResponseExit:
    If Not xmlResponseElem Is Nothing Then
        CreateErrorResponse = xmlResponseElem.xml
    End If
    
    Set xmlLocksElem = Nothing
    Set xmlErrorElem = Nothing
    Set xmlLocksDoc = Nothing
    Set xmlResponseElem = Nothing
End Function

Private Function LockApplication(strApplicationNumber As String, strUserId As String, strUnitID As String, strMachineID As String, blnCreateNew As Boolean, ByRef strLocks As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Places an application lock and related customer locks into the database via a stored-
'   procedure.
'   Based on omApp.ApplicationManagerDO.cls:IApplicationManagerDO_LockCustomersForApplication().
'------------------------------------------------------------------------------------------
On Error GoTo LockApplicationExit

    Const cstrFunctionName As String = "omLockManager.LockManagerBO.LockApplication"
    
    Dim objADOAssist As ADOAssist
    Dim cmd As ADODB.Command
    Dim strm As ADODB.Stream

    Set objADOAssist = IIf(m_objContext Is Nothing, New ADOAssist, m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist"))

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_LockManagerLockApplication"
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.Parameters.Append cmd.CreateParameter("@applicationNumber", adVarChar, adParamInput, 12, strApplicationNumber)
    cmd.Parameters.Append cmd.CreateParameter("@userID", adVarChar, adParamInput, 10, strUserId)
    cmd.Parameters.Append cmd.CreateParameter("@unitID", adVarChar, adParamInput, 10, strUnitID)
    cmd.Parameters.Append cmd.CreateParameter("@machineID", adVarChar, adParamInput, 30, strMachineID)
    cmd.Parameters.Append cmd.CreateParameter("@createNew", adBoolean, adParamInput, , blnCreateNew)
    
    Set strm = New ADODB.Stream
    strm.Open
    cmd.Properties("Output Stream") = strm
    cmd.Properties("XML Root") = "LOCKS"

    cmd.Execute , , adExecuteStream
    strLocks = strm.ReadText
    If InStr(1, strLocks, "<LOCKS></LOCKS>") = 0 Then
        ' Assume strLocks matches "<LOCKS><CUSTOMERLOCK>...</CUSTOMERLOCK></LOCKS>",
        ' i.e., existing customer locks have been returned.
        Call errThrowError(cstrFunctionName, gstrLockManagerErrorNumber)
    Else
        strLocks = ""
    End If
            
LockApplicationExit:
    
    If Not strm Is Nothing Then
        If strm.State = adStateOpen Then
            strm.Close
        End If
    End If
    
    Set strm = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrFunctionName & ", " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        LockApplication = True
    End If
End Function

Private Function UnlockApplication(strApplicationNumber As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Removes an application lock and related customer locks via a stored-
'   procedure.
'   Based on omApp.ApplicationManagerDO.cls:IApplicationManagerDO_UnlockApplication().
'------------------------------------------------------------------------------------------
On Error GoTo UnlockApplicationExit

    Const cstrFunctionName As String = "omLockManager.LockManagerBO.UnlockApplication"
    
    Dim objADOAssist As ADOAssist
    Dim cmd As ADODB.Command

    Set objADOAssist = IIf(m_objContext Is Nothing, New ADOAssist, m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist"))

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_LockManagerUnlockApplication"
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.Parameters.Append cmd.CreateParameter("@applicationNumber", adVarChar, adParamInput, 12, strApplicationNumber)
    
    cmd.Execute
        
UnlockApplicationExit:
    
    Set objADOAssist = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrFunctionName & ", " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        UnlockApplication = True
    End If
End Function

Private Function LockCustomer(strCustomerNumber As String, strUserId As String, strUnitID As String, strMachineID As String, ByRef strLocks As String) As Boolean
On Error GoTo LockCustomerExit

    Const cstrFunctionName As String = "omLockManager.LockManagerBO.LockCustomer"
    
    Dim objADOAssist As ADOAssist
    Dim cmd As ADODB.Command
    Dim strm As ADODB.Stream
    
    Set objADOAssist = IIf(m_objContext Is Nothing, New ADOAssist, m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist"))

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_LockManagerLockCustomer"
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.Parameters.Append cmd.CreateParameter("@customerNumber", adVarChar, adParamInput, 12, strCustomerNumber)
    cmd.Parameters.Append cmd.CreateParameter("@userID", adVarChar, adParamInput, 10, strUserId)
    cmd.Parameters.Append cmd.CreateParameter("@unitID", adVarChar, adParamInput, 10, strUnitID)
    cmd.Parameters.Append cmd.CreateParameter("@machineID", adVarChar, adParamInput, 30, strMachineID)

    Set strm = New ADODB.Stream
    strm.Open
    cmd.Properties("Output Stream") = strm
    cmd.Properties("XML Root") = "LOCKS"

    cmd.Execute , , adExecuteStream
    strLocks = strm.ReadText
    If InStr(1, strLocks, "<LOCKS></LOCKS>") = 0 Then
        ' Assume strLocks matches "<LOCKS><CUSTOMERLOCK>...</CUSTOMERLOCK></LOCKS>",
        ' i.e., existing customer locks have been returned.
        Call errThrowError(cstrFunctionName, gstrLockManagerErrorNumber)
    Else
        strLocks = ""
    End If
    
LockCustomerExit:
       
    If Not strm Is Nothing Then
        If strm.State = adStateOpen Then
            strm.Close
        End If
    End If
    
    Set strm = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing

    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrFunctionName & ", " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        LockCustomer = True
    End If
End Function

Private Function UnlockCustomer(strCustomerNumber As String) As Boolean
On Error GoTo UnlockCustomerExit

    Const cstrFunctionName As String = "omLockManager.UnlockManagerBO.LockCustomer"
    
    Dim objADOAssist As ADOAssist
    Dim cmd As ADODB.Command
    
    Set objADOAssist = IIf(m_objContext Is Nothing, New ADOAssist, m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist"))

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_LockManagerUnlockCustomer"
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.Parameters.Append cmd.CreateParameter("@customerNumber", adVarChar, adParamInput, 12, strCustomerNumber)

    cmd.Execute
                    
UnlockCustomerExit:
    
    Set cmd = Nothing
    Set objADOAssist = Nothing

    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrFunctionName & ", " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        UnlockCustomer = True
    End If
End Function

Private Function ValidateWorkingHours(strUserId As String, strChannelID As String) As Boolean
On Error GoTo ValidateWorkingHoursExit

    Const cstrFunctionName As String = "omLockManager.LockManagerBO.ValidateWorkingHours"
    Dim objOrganisationBO As OrganisationBO
        
    Set objOrganisationBO = IIf(m_objContext Is Nothing, New OrganisationBO, m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationBO"))
    
    ' This will raise an error if the user is outside working hours.
    Call objOrganisationBO.ValidateWorkingHours("<REQUEST><OMIGAUSER><USERID>" & strUserId & "</USERID><CHANNELID>" & strChannelID & "</CHANNELID></OMIGAUSER></REQUEST>")

ValidateWorkingHoursExit:
    
    Set objOrganisationBO = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrFunctionName & ", " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        ValidateWorkingHours = True
    End If
End Function

Private Function CreateApplicationAccessAudit(blnSuccess As Boolean, strUserId As String, strMachineID As String, strApplicationNumber As String) As Boolean
On Error GoTo CreateApplicationAccessAuditExit

    Const cstrFunctionName As String = "omLockManager.LockManagerBO.CreateApplicationAccessAudit"
    
    ' CORE113 AS 04/05/2005 Do not use NTxBO as this results in lock time outs on
    ' the APPLICATIONACCESS table when calling objIAuditNTxBO.CreateAccessAudit.
    ' Because we're not using NTxBO (requires new), but TxBO (requires), then the audit
    ' records are created in the same transaction as the locks. Thus, if the transaction is
    ' aborted (e.g., on error), no audit records will be inserted.
    Dim objIAuditTxBO As IAuditTxBO
    Dim objIAuditDO As IAuditDO
    Dim xmlRequest As FreeThreadedDOMDocument40
    Dim strAuditRecordType As String
            
    Set objIAuditTxBO = IIf(m_objContext Is Nothing, New AuditTxBO, m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditTxBO"))
    Set objIAuditDO = IIf(m_objContext Is Nothing, New AuditDO, m_objContext.CreateInstance(gstrAUDIT_COMPONENT & ".AuditDO"))

    strAuditRecordType = objIAuditDO.GetApplicationLockValueId()

    Set xmlRequest = New FreeThreadedDOMDocument40
    xmlRequest.setProperty "NewParser", True
    
    Dim strRequest As String
    strRequest = _
        "<REQUEST>" & _
            "<ACCESSAUDIT>" & _
                "<USERID>" & strUserId & "</USERID>" & _
                "<MACHINEID>" & strMachineID & "</MACHINEID>" & _
                "<AUDITRECORDTYPE>" & strAuditRecordType & "</AUDITRECORDTYPE>" & _
                "<SUCCESSINDICATOR>" & IIf(blnSuccess, "1", "0") & "</SUCCESSINDICATOR>" & _
                "<APPLICATIONNUMBER>" & strApplicationNumber & "</APPLICATIONNUMBER>" & _
            "</ACCESSAUDIT>" & _
        "</REQUEST>"
    Call xmlRequest.loadXML(strRequest)
    
    Call errCheckXMLResponse(objIAuditTxBO.CreateAccessAudit(xmlRequest.documentElement).xml, True)    ' Raise an error on unsuccessful response.

CreateApplicationAccessAuditExit:
    
    Set objIAuditDO = Nothing
    Set objIAuditTxBO = Nothing
    Set xmlRequest = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, cstrFunctionName & ", " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        CreateApplicationAccessAudit = True
    End If
End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

' From LockApplication
'    If Not rst Is Nothing Then
'        If rst.State = adStateOpen Then
'            If Not rst.EOF Then
'                ' Stored procedure returns recordset to indicate there are existing locks.
'                Dim strTable As String
'                Dim strLockedCustomerNumber As String
'                Dim strLockingUserID As String
'                Dim strLockingUnitID As String
'                strTable = rst.Fields.Item(0).Value
'
'                ' The first field in the recordset is the name of the table with the existing lock.
'                If strTable = "APPLICATIONLOCK" Then
'                    ' The application is already locked.
'                    strLockingUserID = rst.Fields.Item(3).Value
'                    strLockingUnitID = rst.Fields.Item(2).Value
'                    Call errThrowError(cstrFunctionName, 7014, , strLockingUserID, strLockingUnitID)
'                ElseIf strTable = "CUSTOMERLOCK" Then
'                    ' One or more of the customers on this application are already locked,
'                    ' and they are not locked to this application.
'                    Dim strLockedCustomers As String
'                    strLockedCustomers = "" & vbCrLf & vbCrLf
'                    While Not rst.EOF
'                        strLockedCustomerNumber = rst.Fields.Item(1).Value
'                        strLockingUserID = rst.Fields.Item(3).Value
'                        strLockingUnitID = rst.Fields.Item(2).Value
'                        strLockedCustomers = strLockedCustomers & "Customer Number: " & strLockedCustomerNumber & " by " & strLockingUserID & " (" & strLockingUnitID & ")" & vbCrLf
'                        rst.MoveNext
'                    Wend
'                    strLockedCustomers = strLockedCustomers & vbCrLf
'                    Call errThrowError(cstrFunctionName, 7013, , vbCrLf, strLockedCustomers)
'                ElseIf strTable = "CUSTOMERLOCKAPPLICATIONLOCK" Then
'                    ' One or more of the customers on the application are locked to another application.
'                    Call errThrowError(cstrFunctionName, 190)
'                Else
'                    Call errThrowError(cstrFunctionName, oeUnspecifiedError, "Invalid table (" & strTable & ") returned in recordset from usp_LockManagerLockApplication")
'                End If
'            End If
'        End If
'    End If
   
' From LockCustomer
'    Set rst = cmd.Execute
'    If Not rst Is Nothing Then
'        If rst.State = adStateOpen Then
'            If Not rst.EOF Then
'                ' Stored procedure returns recordset to indicate there is an existing lock.
'                Dim strLockingUserID As String
'                Dim strLockingUnitID As String
'                strLockingUserID = rst.Fields.Item(3).Value
'                strLockingUnitID = rst.Fields.Item(2).Value
'                ' TODO - require proper error message number.
'                Call errThrowError(cstrFunctionName, 20000, , strCustomerNumber, strLockingUserID, strLockingUnitID)
'            End If
'        End If
'    End If
        

