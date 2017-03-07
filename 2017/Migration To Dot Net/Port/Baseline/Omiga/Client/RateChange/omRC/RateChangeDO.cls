VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "RateChangeDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------------------------------
'Workfile:      RateChangeDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Rate Change Data Object.
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV/MC  20/08/2001  Created
'GHun   14/03/2002  SYS4276 - Base Rate Applied Date is not set after a rate change batch run
'DRC    08/04/2002  SYS4362 - in FindBaseRateDataforLoanComponent re-initialise temporary response elem
'-------------------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IRateChangeDO

Private gobjContext As ObjectContext
'BMIDS622 gobjBO is never used
'Private gobjBO As Object


'-------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    09/01/2003  Added method FindFailedApplications
'GHun   20/02/2003  BM0367 - Added method GetDatabaseServerDate
'GHun   18/08/2003  BMIDS622 restructure and clean up, added method PrepareRateChange
'GHun   05/11/2003  BMIDS622 Amend PrepareRateChange to increase the database timeout
'-------------------------------------------------------------------------------------------------------

Private Sub IRateChangeDO_CreateRateChangeAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_CreateRateChangeAudit
' Author:       Mark Coleman
' Date:         06/09/2001
' Purpose:      Create a single RateChangeBatchAudit record
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------
    
On Error GoTo CreateRateChangeAuditExit
    
Const cstrFunctionName As String = "IRateChangeDO_CreateRateChangeAudit"
Dim xmlRCBatchAuditRequestNode As IXMLDOMNode
    
    'Checking the Mandatory Node
    If vxmlRequestNode.nodeName = "RATECHANGEBATCHAUDIT" Then
        Set xmlRCBatchAuditRequestNode = vxmlRequestNode
    Else
        Set xmlRCBatchAuditRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "RATECHANGEBATCHAUDIT")
    End If
    xmlCheckMandatoryAttribute xmlRCBatchAuditRequestNode, "BATCHAUDITGUID"
    
    adoCreateFromNode xmlRCBatchAuditRequestNode, "RATECHANGEBATCHAUDIT"
    gobjContext.SetComplete
    
CreateRateChangeAuditExit:
    Set xmlRCBatchAuditRequestNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Sub IRateChangeDO_FindApplications(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_FindApplications
' Author:       Mark Coleman
' Date:         06/09/01
' Purpose:      Find all applications affected by a Rate Change
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------

'BMIDS622 No longer used, but cannot be deleted because OOSS will just replace it
'On Error GoTo FindApplicationsExit
'
'Dim xmlRateChangeNode As IXMLDOMElement
'
'Const strFunctionName As String = "IRateChangeDO_FindApplications"
'
''    Set xmlRateChangeNode = vxmlRequestNode.selectSingleNode(".//RATECHANGE")
'    Set xmlRateChangeNode = vxmlRequestNode.selectSingleNode(".//BATCH")
'    adoGetAsXML xmlRateChangeNode, vxmlResponseNode, "FINDRATECHANGEAPPLICATIONS"
'
'    'Checking the Record Existence
'    If vxmlResponseNode.hasChildNodes = False Then
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
'    gobjContext.SetComplete
'
'FindApplicationsExit:
'    Set xmlRateChangeNode = Nothing
'
'    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IRateChangeDO_FindBaseRateDataforLoanComponent(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_FindBaseRateDataforLoanComponent
' Author:       Dave Crossley
' Date:         01/03/02
' Purpose:      Find all Info for BaseRate change for an LoanComponent
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------

On Error GoTo FindBaseRateDataforLoanComponentExit
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRateChangeNode As IXMLDOMElement
Dim strTemp As String
Dim xmlTempNode As IXMLDOMNode
Dim xmlLCNodeList As IXMLDOMNodeList
Dim xmlTempResponseNode As IXMLDOMElement
                                
Const cstrFunctionName As String = "IRateChangeDO_FindBaseRateDataforLoanComponent"
    
    'build up the ratechangenode
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlRateChangeNode = xmlDoc.createElement("RATECHANGE")

    xmlCopyAttribute vxmlRequestNode, xmlRateChangeNode, "APPLICATIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlRateChangeNode, "APPLICATIONFACTFINDNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlRateChangeNode, "MORTGAGESUBQUOTENUMBER"
    
    ' go through the  loan components for this application
    Set xmlLCNodeList = vxmlResponseNode.selectNodes(".//LOANCOMPONENT")
    For Each xmlTempNode In xmlLCNodeList
        
        ' get the loan component sequence number from the loan component node
        strTemp = xmlGetAttributeText(xmlTempNode, _
                                         "LOANCOMPONENTSEQUENCENUMBER")
        
        xmlRateChangeNode.setAttribute "LOANCOMPONENTSEQUENCENUMBER", strTemp
        ' AQR SYS4362 - Re initialise
        Set xmlTempResponseNode = xmlDoc.createElement("TEMP")
        adoGetAsXML xmlRateChangeNode, xmlTempResponseNode, "FINDBASERATEDATA"
        
        'Checking the Record Existence
        If xmlTempResponseNode.hasChildNodes = False Then
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
        'now need to transfer attribs across to response
    
        xmlCopyAttribute xmlTempResponseNode.firstChild, xmlTempNode, "BASERATESET"
        xmlCopyAttribute xmlTempResponseNode.firstChild, xmlTempNode, "LOANAMOUNT"
        xmlCopyAttribute xmlTempResponseNode.firstChild, xmlTempNode, "LTV"
        xmlCopyAttribute xmlTempResponseNode.firstChild, xmlTempNode, "MORTGAGEPRODUCTCODE"
        xmlCopyAttribute xmlTempResponseNode.firstChild, xmlTempNode, "STARTDATE"
    
    Next
    gobjContext.SetComplete
    
FindBaseRateDataforLoanComponentExit:
    Set xmlRateChangeNode = Nothing
    Set xmlDoc = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempResponseNode = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Sub IRateChangeDO_FindBatchApplicationList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_FindBatchApplicationList
' Author:       Mark Coleman
' Date:         20/08/01
' Purpose:      Find a list of applications for a Rate Change Batch Run
'               returning the status of each item in the batch
' Input parameters:
'               vxmlRequestNode - XML Request of the form
'               <REQUEST>
'                   <BATCH BATCHNUMBER='A' BATCHRUNNUMBER='B'/>
'               </REQUEST>
'
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               No return value
' Notes:
'               Uses view FINDRATECHANGEAPPLICATIONLIST
'------------------------------------------------------------------------

On Error GoTo IRateChangeDO_FindBatchApplicationList_Exit

Const cstrFunctionName As String = "IRateChangeDO_FindBatchApplicationList"

Dim xmlRequest As IXMLDOMNode

    If vxmlRequestNode.nodeName = "BATCH" Then
        Set xmlRequest = vxmlRequestNode
    Else
        Set xmlRequest = vxmlRequestNode.selectSingleNode("BATCH")
    End If
    
    adoGetAsXML xmlRequest, vxmlResponseNode, "FINDRATECHANGEAPPLICATIONLIST"
    gobjContext.SetComplete

IRateChangeDO_FindBatchApplicationList_Exit:
    Set xmlRequest = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'BM0230 MDC 09/01/2003
Private Sub IRateChangeDO_FindFailedApplications(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_FindFailedApplications
' Author:       Mark Coleman
' Date:         09/01/2003
' Purpose:      Find a list of applications that failed in a previous Rate Change
'               batch run.
' Input parameters:
'               vxmlRequestNode - XML Request of the form
'               <REQUEST>
'                   <BATCH BATCHNUMBER='A' BATCHRUNNUMBER='B'/>
'               </REQUEST>
'
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               No return value
' Notes:
'               Uses stored procedure USP_FINDFAILEDAPPLICATIONS to retrieve
'               XML directly from the database in the form required.
'------------------------------------------------------------------------

'BMIDS622 No longer used, but cannot be deleted because OOSS will just replace it
'On Error GoTo FindFailedApplicationsExit
'
'Dim strFunctionName As String
'strFunctionName = "IRateChangeDO_FindFailedApplications"
'
'Dim adoConn As ADODB.Connection
'Dim adoCommand As ADODB.Command
'Dim adoStream As ADODB.Stream
'Dim adoParameter As ADODB.Parameter
'Dim xmlOut As FreeThreadedDOMDocument40
'
'Dim strParamValue As String
'Dim strXML As String
'
'    'Initialise objects
'    Set adoConn = New ADODB.Connection
'    Set adoCommand = New ADODB.Command
'    Set adoStream = New ADODB.Stream
'
'    'Create parameters
'    strParamValue = xmlGetMandatoryAttributeText(vxmlRequestNode, "BATCHNUMBER")
'    Set adoParameter = adoCommand.CreateParameter("@szBatchNumber", adBSTR, adParamInput, Len(strParamValue), strParamValue)
'    adoCommand.Parameters.Append adoParameter
'
'    strParamValue = xmlGetMandatoryAttributeText(vxmlRequestNode, "BATCHRUNNUMBER")
'    Set adoParameter = adoCommand.CreateParameter("@nBatchRunNumber", adInteger, adParamInput, , CSafeInt(strParamValue))
'    adoCommand.Parameters.Append adoParameter
'
'    'Open connection
'    With adoConn
'        .ConnectionString = adoGetDbConnectString
'        .open
'    End With
'
'    ' Use stream object to retrieve XML from the database
'    adoStream.open
'    With adoCommand
'        .CommandText = "USP_FINDFAILEDAPPLICATIONS"
'        .CommandType = adCmdStoredProc
'        .ActiveConnection = adoConn
'        .Properties("Output Stream") = adoStream
'        .Properties("XML Root") = "RATECHANGEBATCHAUDITLIST"
'        .Execute , , adExecuteStream
'    End With
'
'    strXML = adoStream.ReadText()
'
'    'Close connection
'    adoConn.Close
'    adoStream.Close
'
'    'Set up return xml
'    Set xmlOut = xmlLoad(strXML, strFunctionName)
'    vxmlResponseNode.appendChild xmlOut.documentElement
'
'    gobjContext.SetComplete
'
'FindFailedApplicationsExit:
'    Set adoConn = Nothing
'    Set adoCommand = Nothing
'    Set adoStream = Nothing
'    Set adoParameter = Nothing
'    Set xmlOut = Nothing
'
'    errCheckError strFunctionName, TypeName(Me)

End Sub
'BM0230 MDC 09/01/2003 - End

Private Sub IRateChangeDO_FindLoanComponentsForApplication(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_FindLoanComponentsForApplication
' Author:       Mark Coleman
' Date:         06/09/2001
' Purpose:      Find all Loan Component Sequence Numbers for a given application
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------

Const cstrFunctionName As String = "IRateChangeDO_FindLoanComponentsForApplication"
On Error GoTo FindLoanComponentsForApplicationExit

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlElement As IXMLDOMElement

Dim objCommand As ADODB.Command
Dim objRs As ADODB.Recordset

Dim strAppNo As String
Dim strAppFFNo As String
Dim strMSQNo As String
Dim strSQL As String

    strAppNo = xmlGetMandatoryAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
    strAppFFNo = xmlGetMandatoryAttributeText(vxmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    strMSQNo = xmlGetMandatoryAttributeText(vxmlRequestNode, "MORTGAGESUBQUOTENUMBER")
    
    Set objCommand = New ADODB.Command
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strSQL = "select loancomponentsequencenumber from loancomponent"
    strSQL = strSQL & " where applicationnumber = '" & strAppNo & "' "
    strSQL = strSQL & " and applicationfactfindnumber = " & strAppFFNo
    strSQL = strSQL & " and mortgagesubquotenumber = " & strMSQNo
    
    'Execute the Command SQL
    objCommand.CommandType = adCmdText
    objCommand.CommandText = strSQL
    Debug.Print strSQL
    Set objRs = executeGetRecordSet(objCommand)

    'Convert Recordset to XML Response
    If Not objRs Is Nothing Then
        If Not (objRs.BOF Or objRs.EOF) Then
            objRs.MoveFirst
            Do Until objRs.EOF
                Set xmlElement = xmlDoc.createElement("LOANCOMPONENT")
                xmlElement.setAttribute "LOANCOMPONENTSEQUENCENUMBER", objRs.Fields("LOANCOMPONENTSEQUENCENUMBER")
                vxmlResponseNode.appendChild xmlElement
                objRs.MoveNext
            Loop
        Else
            errThrowError cstrFunctionName, oeRecordNotFound
        End If
    Else
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
    gobjContext.SetComplete
    
FindLoanComponentsForApplicationExit:
    Set objCommand = Nothing
    Set objRs = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IRateChangeDO_FindRateChangeAuditList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_FindRateChangeAuditList
' Author:       Mark Coleman
' Date:         06/09/2001
' Purpose:      Find RateChangeBatchAudit records for this batch
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------
    
On Error GoTo FindRateChangeAuditListExit
    
Const cstrFunctionName As String = "IRateChangeDO_FindRateChangeAuditList"

    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "RATECHANGEBATCHAUDIT", , "APPLICATIONNUMBER, APPLICATIONFACTFINDNUMBER"
    
    'Checking the Record Existence
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    gobjContext.SetComplete
    
FindRateChangeAuditListExit:
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

''BM0367
'Private Sub IRateChangeDO_GetDatabaseServerDate(ByVal vxmlResponseNode As IXMLDOMNode)
'' Procedure:    IRateChangeDO_GetDatabaseServerDate
'' Author:       GHun
'' Date:         20/02/2003
'' Purpose:      Get the current date an time from the database
'' Input parameters:
''               vxmlResponseNode - XML Response to be appended to
'' Output parameters:
''               N/A
''------------------------------------------------------------------------
'
'On Error GoTo GetDatabaseServerDateExit
'
'    Const cstrFunctionName  As String = "IRateChangeDO_GetDatabaseServerDate"
'    Dim objCommand          As ADODB.Command
'    Dim objConn             As ADODB.Connection
'    Dim objRecordSet        As ADODB.Recordset
'    Dim strSQL              As String
'    Dim strDate             As String
'    Dim xmlDoc              As FreeThreadedDOMDocument40
'    Dim xmlServer           As IXMLDOMElement
'    Dim objSQLAssist        As omBase.SQLAssist
'
'On Error GoTo GetDatabaseServerDateExit
'
'    Set objCommand = New ADODB.Command
'    Set objConn = New ADODB.Connection
'    Set objRecordSet = New ADODB.Recordset
'
'    Set objSQLAssist = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".SQLAssist")
'    strSQL = "SELECT " & objSQLAssist.GetSystemDate
'
'    objCommand.CommandType = adCmdText
'    objCommand.CommandText = strSQL
'
'    objConn.ConnectionString = adoGetDbConnectString
'    objConn.Open
'
'    objCommand.ActiveConnection = objConn
'
'    objRecordSet.CursorLocation = adUseClient
'    objRecordSet.CursorType = adOpenStatic
'    objRecordSet.LockType = adLockReadOnly
'
'    objRecordSet.Open objCommand
'
'    ' disconnect RecordSet
'    Set objRecordSet.ActiveConnection = Nothing
'    objConn.Close
'
'    If Not objRecordSet.EOF Then
'        objRecordSet.MoveFirst
'        If IsNull(objRecordSet.Fields.Item(0)) = False Then
'            strDate = objRecordSet.Fields.Item(0).Value
'        End If
'    End If
'
'    If Len(strDate) = 0 Then
'        errThrowError cstrFunctionName, oeRecordNotFound
'    Else
'        Set xmlDoc = New FreeThreadedDOMDocument40
'        Set xmlServer = xmlDoc.createElement("SERVER")
'        xmlServer.setAttribute "DATE", strDate
'        vxmlResponseNode.appendChild xmlServer
'    End If
'
'GetDatabaseServerDateExit:
'
'    Set objRecordSet = Nothing
'    Set objCommand = Nothing
'    Set objConn = Nothing
'    Set xmlDoc = Nothing
'    Set xmlServer = Nothing
'    Set objSQLAssist = Nothing
'
'    errCheckError cstrFunctionName, TypeName(Me)
'End Sub
''BM0367 End

Private Sub IRateChangeDO_SetBaseRateAppliedDate(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeDO_SetBaseRateAppliedDate
' Author:       Mark Coleman
' Date:         06/09/2001
' Purpose:      Set the BaseRateAppliedDate field upon completion of the batch run
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------
    
On Error GoTo SetBaseRateAppliedDateExit

Const cstrFunctionName As String = "IRateChangeDO_SetBaseRateAppliedDate"
Dim objCommand As ADODB.Command
Dim strSQL As String
Dim lngDays As Long
Dim conn                As ADODB.Connection    'SYS4276 GHun 14/03/2002
Dim strGetSystemDate    As String              'SYS4276 GHun 14/03/2002
Dim objSQLAssist        As omBase.SQLAssist    'SYS4276 GHun 14/03/2002
    
    'Get Global Parameter RateChangeLookAheadDays
    lngDays = GetGlobalParamAmount("RateChangeLookAheadDays")
    
    Set objCommand = New ADODB.Command
    
    'SYS4276 GHun 14/03/2002 Use of SYSDATE was Oracle specific
    'strSQL = "UPDATE BASERATE SET BASERATEAPPLIEDDATE = SYSDATE"
    'strSQL = strSQL & " WHERE BASERATESTARTDATE <= SYSDATE"
    Set objSQLAssist = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".SQLAssist")
    strGetSystemDate = objSQLAssist.GetSystemDate
    strSQL = "UPDATE BASERATE SET BASERATEAPPLIEDDATE = " & strGetSystemDate
    strSQL = strSQL & " WHERE BASERATESTARTDATE <= " & strGetSystemDate
    'SYS4276
    
    If lngDays > 0 Then
        strSQL = strSQL & " + " & CStr(lngDays)
    End If
    strSQL = strSQL & " AND BASERATEAPPLIEDDATE IS NULL"
        
    'Execute the Command SQL
    objCommand.CommandType = adCmdText
    objCommand.CommandText = strSQL
    Debug.Print strSQL
    
    'SYS4276 GHun 14/03/2002 Set a connection for the Command object to use
    Set conn = New ADODB.Connection
    conn.ConnectionString = adoGetDbConnectString
    conn.open
    objCommand.ActiveConnection = conn
    'SYS4276
    
    objCommand.Execute
    conn.Close 'SYS4276
    
    gobjContext.SetComplete

SetBaseRateAppliedDateExit:
    Set objCommand = Nothing
    Set conn = Nothing          'SYS4276
    Set objSQLAssist = Nothing  'SYS4276

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

'BMIDS622 Sub is never used, but cannot be deleted as OOSS will just replace it
Private Sub IRateChangeDO_FindBaseRaqteDataforApplications(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
''------------------------------------------------------------------------
'' Procedure:    IRateChangeDO_FindBaseRaqteDataforApplications
'' Author:       Mark Coleman
'' Date:         06/09/01
'' Purpose:      Find all applications affected by a Rate Change
'' Input parameters:
''               vxmlRequestNode - XML Request
''               vxmlResponseNode - XML Response to be appended to
'' Output parameters:
''               N/A
''------------------------------------------------------------------------
'
'On Error GoTo FindBaseRaqteDataforApplicationsExit
'
'Dim xmlRateChangeNode As IXMLDOMElement
'
'Const strFunctionName As String = "IRateChangeDO_FindBaseRaqteDataforApplications"
'
''    Set xmlRateChangeNode = vxmlRequestNode.selectSingleNode(".//RATECHANGE")
'    Set xmlRateChangeNode = vxmlRequestNode.selectSingleNode(".//BATCH")
'    adoGetAsXML xmlRateChangeNode, vxmlResponseNode, "FINDRATECHANGEAPPLICATIONS"
'
'    'Checking the Record Existence
'    If vxmlResponseNode.hasChildNodes = False Then
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
'    gobjContext.SetComplete
'
'FindBaseRaqteDataforApplicationsExit:
'    Set xmlRateChangeNode = Nothing
'
'    errCheckError strFunctionName, TypeName(Me)
'
End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'BMIDS622 GHun 18/08/2003
Private Sub IRateChangeDO_PrepareRateChange(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName      As String = "PrepareRateChange"
    Dim sBatchNumber            As String
    Dim lBatchRunNumber         As Long
    Dim lFailedBatchRunNumber   As Long
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoStream               As ADODB.Stream
    Dim xmlBatchSchedule        As IXMLDOMElement
    Dim xmlDoc                  As FreeThreadedDOMDocument40

On Error GoTo PrepareRateChangeExit
    
    Set xmlBatchSchedule = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    sBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSchedule, "BATCHNUMBER")
    lBatchRunNumber = xmlGetMandatoryAttributeAsLong(xmlBatchSchedule, "BATCHRUNNUMBER")
    'FailedBatchRunNumber defaults to -1 if it does not exist
    lFailedBatchRunNumber = xmlGetAttributeAsLong(xmlBatchSchedule, "FAILEDBATCHRUNNUMBER", "-1")

    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.open
        
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    adoStream.open
    
    'Create the parameters and call the stored proc
    With adoCommand
        Set .ActiveConnection = adoConnection
        'BMIDS622 Change the timeout to 1 hour (same as the MQL default) as the
        ' default 30 seconds is normally far too short for this proc to complete in
        .CommandTimeout = 3600
        'BMIDS622 End
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "BATCH"
        .CommandType = adCmdStoredProc
        .CommandText = "USP_PREPARERATECHANGE"
        .Parameters.Append .CreateParameter("@p_BatchNumber", adBSTR, adParamInput, 10, sBatchNumber)
        .Parameters.Append .CreateParameter("@p_BatchRunNumber", adInteger, adParamInput, , lBatchRunNumber)
        .Parameters.Append .CreateParameter("@p_FailedBatchRunNumber", adInteger, adParamInput, , lFailedBatchRunNumber)
        .Execute , , adExecuteStream
    End With
            
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
        
    If Not (xmlDoc.loadXML(adoStream.ReadText)) Then
        ' raise application error to be interpreted by calling object
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
       
    'Close the database connection
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close

    vxmlResponseNode.appendChild xmlDoc.documentElement

PrepareRateChangeExit:
    gobjContext.SetComplete
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    Set adoStream = Nothing
    Set xmlBatchSchedule = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BMIDS622 End
