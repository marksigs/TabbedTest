VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MortgageLenderDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageLenderDO.cls
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
' Prog  Date        Description
' MCS   19/11/99    Created
' APS   03/03/00    Change object reference so we get the IOmOrgClassDef interface of
'                   the omOrgClasssDef class
' CL    27/11/00    CORE000005
' CL    05/12/00    CORE000005 - Reinstate FindLenderNameAndIncomeMultiples on default
'                                interface for backward compatibility
' PSC   12/01/01    SYS1753 Added GetMainLenderDirectoryGuid()
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'DRC    12/06/01    Added GetLenderDetails method
'DRC    28/09/01    SYS2696  - change GetLenderDetails method
'AS     13/11/03    CORE1 Removed GENERIC_SQL.
'SAB    20/09/05    Added GetApplicationMortgageLender()
'SR     23/01/08    CORE00000452 Apply pre-migration convertor tool
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'BMIDS Change History:
'---------------------
' Prog      Date        AQR         Description
' GD        17/04/03    BM0198      Adjusted IMortgageLenderDO_GetLenderDetails so that Get can be done on Org Id or Lender code
' GD        20/05/03    BM0198      Adjusted IMortgageLenderDO_GetLenderDetails to check for elements, not attributes
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'Baseline specific History:
'
' Prog      Date        AQR         Description
' MV        20/11/2003  CORE6       Added IMortgageLenderDO_GetPanelLenderDetails()
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'BBG Change History:
'---------------------
' Prog      Date        AQR         Description
'
' AW        14/09/04    BBG1416     Added GetLenderTelephoneDetails
' AW        14/09/04    BBG1436     Added GetLenderNameDirectoryDetails
'
'-----------------------------------------------------------------------------------------------


Option Explicit
Implements IMortgageLenderDO
Implements ObjectControl
 '=============================================
 'Variable Declaration Section
 '=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
Private m_objSQLAssist As SQLAssist
Private Function IMortgageLenderDO_FindAllowableIncomeFactors( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
On Error GoTo FindAllowableIncomeFactorsVbErr
    
    Const strFunctionName As String = "FindAllowableIncomeFactors"
       
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
        
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadAllowableIncomeFactorsData
      
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
           
    Set IMortgageLenderDO_FindAllowableIncomeFactors = m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
        
    m_objContext.SetComplete
        
            
    Exit Function
FindAllowableIncomeFactorsVbErr:
            
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
            
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objContext.SetAbort
          
    ' re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' AW        14/09/04    BBG1416
Private Function IMortgageLenderDO_GetLenderTelephoneDetails(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Gets the data for a  lender given a directory guid
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
'GD BM0198 Added ability for Org id or lendercode to be used for the Get
'As per Coventry implementation
On Error GoTo GetLenderTelephoneDetailsVbErr
    Const strFunctionName As String = "IMortgageLenderDO_GetLenderTelephoneDetails"
                            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlReturnElement As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    Dim strDirectoryGuid As String
    Dim strSQL As String
    Dim objSQLAssist As SQLAssist
    Set objIADOAssist = New ADOAssist
    Set objSQLAssist = New SQLAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    strDirectoryGuid = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//DIRECTORYGUID", False, False)

    'Search using Directory guid
    strSQL = "SELECT * FROM CONTACTTELEPHONEDETAILS WHERE CONTACTTELEPHONEDETAILS.CONTACTDETAILSGUID = "
    strSQL = strSQL + "( SELECT CONTACTDETAILSGUID FROM NAMEANDADDRESSDIRECTORY "
    strSQL = strSQL + "WHERE NAMEANDADDRESSDIRECTORY.DIRECTORYGUID =  ?  )"
    
    adoCommand.CommandText = strSQL
    
    'Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("strDirectoryGuid", adBinary, _
                                                adParamInput, Len(strDirectoryGuid), _
                                                objSQLAssist.GuidStringToByteArray(strDirectoryGuid))

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
    Set xmlClassDefDoc = objIClassDef.LoadContactTelephoneData()
    
    While Not adoRecordset.EOF
        ' convert record set to xml
        Set xmlReturnElement = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
        Set xmlTempElement = xmlResponseElem.appendChild(xmlReturnElement)
        adoRecordset.MoveNext
    Wend

    Set IMortgageLenderDO_GetLenderTelephoneDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
GetLenderTelephoneDetailsExit:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlTempElement = Nothing
    Set xmlReturnElement = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function
GetLenderTelephoneDetailsVbErr:
      
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
      
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlTempElement = Nothing
    Set xmlReturnElement = Nothing
    
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

Private Function IMortgageLenderDO_GetPanelLenderDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Gets the data for a  lender given a DirectoryGUID
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetPanelLenderDetailsVbErr
    Const strFunctionName As String = "IMortgageLenderDO_GetPanelLenderDetails"
    ' <VSA>
    ' Visual Studio Analyser Support
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984
    ' </VSA>
                            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As New SQLAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    Dim strOrganisationID As String
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    Dim strDirectoryGuid As String
    strDirectoryGuid = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//DIRECTORYGUID")
        
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT MORTGAGELENDER.* FROM MORTGAGELENDER, MORTGAGELENDERDIRECTORY" & _
                        " WHERE MORTGAGELENDERDIRECTORY.DIRECTORYGUID = ? " & _
                        " AND MORTGAGELENDER.ORGANISATIONID = MORTGAGELENDERDIRECTORY.ORGANISATIONID"
    adoCommand.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Dim rbyteDirectoryGuid() As Byte
            
            
            ' Convert the GUID string to a binary
           rbyteDirectoryGuid() = objSQLAssist.GuidStringToByteArray(strDirectoryGuid)
    Set adoParameter = adoCommand.CreateParameter("rawDirectoryGuid", adVarBinary, adParamInput, UBound(rbyteDirectoryGuid) + 1, rbyteDirectoryGuid)
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
    Set xmlClassDefDoc = objIClassDef.LoadMortgageLenderData()
    Set xmlElement = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
    Set IMortgageLenderDO_GetPanelLenderDetails = xmlElement
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    m_objContext.SetComplete
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
    ' <VSA>
    ' Visual Studio Analyser Support
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984
    ' </VSA>
    Exit Function
GetPanelLenderDetailsVbErr:
    
    ' <VSA>
    ' Visual Studio Analyser Support
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984
    ' </VSA>
      
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
      
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
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
Public Function FindLenderNameAndIncomeMultiples(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindLenderNameAndIncomeMultiplesVbErr
    Const strFunctionName As String = "FindLenderNameAndIncomeMultiples"
    Dim xmlRequest As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Set xmlRequest = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlResponse = IMortgageLenderDO_FindLenderNameAndIncomeMultiples(xmlRequest.documentElement)
    FindLenderNameAndIncomeMultiples = xmlResponse.xml
     
    m_objContext.SetComplete
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Exit Function
FindLenderNameAndIncomeMultiplesVbErr:
    
    
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
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
Private Function IMortgageLenderDO_FindLenderNameAndIncomeMultiples(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindLenderNameAndIncomeMultiplesVbErr
    
    Const strFunctionName As String = "FindLenderNameAndIncomeMultiples"
       
    Dim objIClassDef As IOmOrgClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
        
    Set objIClassDef = New omOrgClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageLenderData(LenderNameAndIncomeMultiples)
      
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
       
    ' APS 03/03/00 - Change object reference so we get the IOmOrgClassDef interface of
    ' the omOrgClasssDef class
    Set IMortgageLenderDO_FindLenderNameAndIncomeMultiples = m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
        
    m_objContext.SetComplete
        
            
    Exit Function
FindLenderNameAndIncomeMultiplesVbErr:
            
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
            
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objContext.SetAbort
          
    ' re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IMortgageLenderDO_FindLenderNames(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindLenderNamesVbErr
    Const strFunctionName As String = "IMortgageLenderDO_FindLenderNames"
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
    Set objXmlAssist = New XMLAssist
    Set objSQLAssist = New SQLAssist
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    Dim strToDay As String
    Dim strSQL As String
    strActiveOnly = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//ACTIVEONLY", False)
    strSQL = "SELECT LENDERCODE ,LENDERNAME ,ORGANISATIONID FROM MORTGAGELENDER "
    'Store Todays date based on Database
    strToDay = objSQLAssist.GetSystemDate
    If strActiveOnly <> "0" Then
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " LENDERSTARTDATE <= " & strToDay
        strSQL = strSQL & " AND ( LENDERENDDATE >= " & strToDay
        strSQL = strSQL & " OR  LENDERENDDATE IS NULL ) "
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
    Set xmlClassDefDoc = objIClassDef.LoadLenderNamesData
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlListElem = xmlOut.createElement("MORTGAGELENDERLIST")
    xmlOut.appendChild xmlListElem
    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    'Close recordset Object
    adoRecordset.Close
    Set IMortgageLenderDO_FindLenderNames = xmlListElem
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
FindLenderNamesVbErr:
    
    
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

' AW        14/09/04    BBG1416
Private Function IMortgageLenderDO_GetLenderNameDirectoryDetails(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Gets the data for a  lender given a directory guid
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
'GD BM0198 Added ability for Org id or lendercode to be used for the Get
'As per Coventry implementation
On Error GoTo IMortgageLenderDO_GetLenderNameDirectoryDetailsVbErr
    Const strFunctionName As String = "IMortgageLenderDO_GetLenderNameDirectoryDetails"
                            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlReturnElement As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    Dim strOrganisationID As String
    Dim strSQL As String
    Dim objSQLAssist As SQLAssist
    Set objIADOAssist = New ADOAssist
    Set objSQLAssist = New SQLAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    strOrganisationID = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//ORGANISATIONID", False, False)

    'Search using organisationId
    strSQL = "SELECT * FROM MORTGAGELENDERDIRECTORY, NAMEANDADDRESSDIRECTORY, CONTACTDETAILS, "
    strSQL = strSQL + "ADDRESS "
    
    strSQL = strSQL + "WHERE (MORTGAGELENDERDIRECTORY.ORGANISATIONID = ? ) "
    strSQL = strSQL + "AND ( MORTGAGELENDERDIRECTORY.DIRECTORYGUID = NAMEANDADDRESSDIRECTORY.DIRECTORYGUID ) "
    strSQL = strSQL + "AND ( NAMEANDADDRESSDIRECTORY.CONTACTDETAILSGUID = CONTACTDETAILS.CONTACTDETAILSGUID ) "
    strSQL = strSQL + "AND ( NAMEANDADDRESSDIRECTORY.ADDRESSGUID = ADDRESS.ADDRESSGUID ) "
    
    adoCommand.CommandText = strSQL
    
    'Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("strOrganisationId", adBinary, _
                                                adParamInput, Len(strOrganisationID), _
                                                objSQLAssist.GuidStringToByteArray(strOrganisationID))

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
    Set xmlClassDefDoc = objIClassDef.LoadLenderDirectoryData()
    
    While Not adoRecordset.EOF
        ' convert record set to xml
        Set xmlReturnElement = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
        Set xmlTempElement = xmlResponseElem.appendChild(xmlReturnElement)
        adoRecordset.MoveNext
    Wend

    Set IMortgageLenderDO_GetLenderNameDirectoryDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
IMortgageLenderDO_GetLenderNameDirectoryDetailsExit:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlTempElement = Nothing
    Set xmlReturnElement = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function
IMortgageLenderDO_GetLenderNameDirectoryDetailsVbErr:
      
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
      
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlTempElement = Nothing
    Set xmlReturnElement = Nothing
    
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
Private Function IMortgageLenderDO_GetLenderDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Gets the data for a  lender given a lender organisationID or Lender Code
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
'GD BM0198 Added ability for Org id or lendercode to be used for the Get
'As per Coventry implementation
On Error GoTo GetLenderDetailsVbErr
    Const strFunctionName As String = "IMortgageLenderDO_GetLenderDetails"
                            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    Dim strOrganisationID As String
    Dim strSQL As String
    Dim strLenderCode As String     'SG 17/10/01 SYS2755
    Dim objSQLAssist As SQLAssist   'SG 29/10/01 SYS2755
    Set objIADOAssist = New ADOAssist
    Set objSQLAssist = New SQLAssist   'SG 29/10/01 SYS2755
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    'SG 17/10/01 SYS2755 START
    'Use OrganisationID OR LenderCode to retrieve details.
    'If BOTH are passed in, use OrganisationID.
    'Get values out of XML
    'strOrganisationID = m_objXmlAssist.GetAttributeValue(vxmlRequest, _
                        "MORTGAGELENDER", "ORGANISATIONID", False)
    'BM0198 START
    'strOrganisationID = m_objXmlAssist.GetAttributeValue(vxmlRequest, _
                        "MORTGAGELENDER", "ORGANISATIONID", False)
    'strLenderCode = m_objXmlAssist.GetAttributeValue(vxmlRequest, _
                        "MORTGAGELENDER", "LENDERCODE", False)
    strOrganisationID = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//ORGANISATIONID", False, False)
    strLenderCode = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//LENDERCODE", False, False)
    'BM0198 END
    'Determine search criteria...
    If Len(strOrganisationID) = 0 Then
        'No OrganisationID present, check LenderCode...
        If Len(strLenderCode) = 0 Then
            'Neither parameter passed in - can't perform the search
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                        "ORGANISATIONID & LENDERCODE tags missing"
        Else
            'Search using LenderCode...
            strSQL = "SELECT * FROM MORTGAGELENDER WHERE LENDERCODE =  ? "
            adoCommand.CommandText = strSQL
            'Set up the parameters to be substituted into the SQL
            Set adoParameter = adoCommand.CreateParameter("strLenderCode", adBSTR, _
                                adParamInput, Len(strLenderCode), strLenderCode)
            adoCommand.Parameters.Append adoParameter
        End If
    Else
        'Search using OrganisationID...
        strSQL = "SELECT * FROM MORTGAGELENDER WHERE ORGANISATIONID =  ? "
        adoCommand.CommandText = strSQL
        'Set up the parameters to be substituted into the SQL
        'SG 29/10/01 Introduce this IF statement
        Set adoParameter = adoCommand.CreateParameter("strOrganisationID", adBinary, _
                                                adParamInput, Len(strOrganisationID), _
                                                objSQLAssist.GuidStringToByteArray(strOrganisationID))
        'end SG 29/10/01
        adoCommand.Parameters.Append adoParameter
    End If
    'SG 17/10/01 SYS2755 END
         
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
    Set xmlClassDefDoc = objIClassDef.LoadMortgageLenderData()
    Set xmlElement = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
    Set IMortgageLenderDO_GetLenderDetails = xmlElement
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    m_objContext.SetComplete
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
    Set objSQLAssist = Nothing
    Exit Function
GetLenderDetailsVbErr:
    
      
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
      
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
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
Private Function IMortgageLenderDO_GetMainLenderDirectoryGuid(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Gets the directory guid for the main lender given a lender code
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMainLenderDirectoryGuidVbErr
    Const strFunctionName As String = "IMortgageLenderDO_GetMainLenderDirectoryGuid"
                            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IOmOrgClassDef
    Dim objIADOAssist As IADOAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRootElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    Set objIADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    Dim strLenderCode As String
    strLenderCode = m_objXmlAssist.GetNodeValue(vxmlTableElement, ".//LENDERCODE")
        
    Dim strSQL As String
    strSQL = "SELECT MORTGAGELENDERDIRECTORY.* FROM MORTGAGELENDER, MORTGAGELENDERDIRECTORY" & _
    " WHERE MORTGAGELENDER.LENDERCODE = ? " & _
    " AND MORTGAGELENDER.ORGANISATIONID = MORTGAGELENDERDIRECTORY.ORGANISATIONID " & _
    " AND MORTGAGELENDERDIRECTORY.MAINMORTGAGELENDERIND = 1"
    adoCommand.CommandText = strSQL
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ApplicationNumber", adBSTR, adParamInput, Len(strLenderCode), strLenderCode)
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
    Set xmlClassDefDoc = objIClassDef.LoadMortgageLenderDirectoryData()
    Set xmlElement = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
    Set xmlElement = xmlElement.selectSingleNode(".//DIRECTORYGUID")
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.async = False
    Set xmlRootElement = xmlOut.createElement("MORTGAGELENDERDIRECTORY")
    xmlOut.appendChild xmlRootElement
    xmlRootElement.appendChild xmlElement
    Set IMortgageLenderDO_GetMainLenderDirectoryGuid = xmlRootElement
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    m_objContext.SetComplete
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
    Exit Function
GetMainLenderDirectoryGuidVbErr:
    
      
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
      
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
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

'*********************************************************************************
' Function Name : IMortgageLenderDO_GetApplicationMortgageLender
' Author        : Mallesh Cheekoti
' Date Created  : 14/06/2004
' Arguments     : This member function takes XMLDOMElement input object
'                 which pass ApplicationNumber and ApplicationFactFindNumber data
' description   : BMIDS763/CC075 - FEE CHANGES
' return        : XML Node containing the MortgageLender Details
'
'*********************************************************************************
Private Function IMortgageLenderDO_GetApplicationMortgageLender(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
        
On Error GoTo ErrorGetApplicationMortgageLender
    Const strFunctionName As String = "IMortgageLenderDO_GetApplicationMortgageLender"
    
    Dim xmlClassDefDoc          As FreeThreadedDOMDocument40
    Dim objIClassDef            As IOmOrgClassDef
    Dim objIADOAssist           As IADOAssist
    Dim adoConnection           As Connection
    Dim adoCommand              As Command
    Dim adoRecordset            As Recordset
    Dim adoParameter            As Parameter
    Dim xmlOut                  As FreeThreadedDOMDocument40
    Dim xmlRootElement          As IXMLDOMElement
    Dim xmlElement              As IXMLDOMElement
    Dim intRetries              As Integer
    Dim intMaxAttempts          As Integer
    Dim intAttempt              As Integer
    Dim blnOpenedOk             As Boolean
    Dim strConnString           As String
    Dim strOrganisationID       As String
    Dim strSQL                  As String
    Dim strLenderCode           As String
    Dim objSQLAssist            As SQLAssist
    Dim strAppNumber            As String
    Dim strAppFactFindNumber    As String
    
    Set objIADOAssist = New ADOAssist
    Set objSQLAssist = New SQLAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    
    strAppNumber = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//APPLICATIONNUMBER", False, False)
    strAppFactFindNumber = m_objXmlAssist.GetNodeValue(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER", False, False)
    
    'Determine search criteria...
    If Len(strAppNumber) = 0 And Len(strAppFactFindNumber) = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                        "APPLICATIONNUMBER and APPLICATIONFACTFINDNUMBER data missing from Request"
    End If
            
    strSQL = "USP_GetApplicationMortgageLender"
    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdStoredProc
    
    'Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, _
                        adParamInput, Len(strAppNumber), strAppNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("APPLICATIONFACTFINDNUMBER", adBSTR, _
                        adParamInput, Len(strAppFactFindNumber), strAppFactFindNumber)
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
    Set xmlClassDefDoc = objIClassDef.LoadMortgageLenderData()
    
    Set xmlElement = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
    'Return value
    Set IMortgageLenderDO_GetApplicationMortgageLender = xmlElement
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    Set xmlClassDefDoc = Nothing

    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
    Set objSQLAssist = Nothing
    
    Exit Function

ErrorGetApplicationMortgageLender:
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
      
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlRootElement = Nothing
    Set xmlElement = Nothing
    
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
'CORE155 - End

Public Function AddDerivedData(ByVal vstrXMLData As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vstrXMLData         base XML data stream
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:
'   AddDerivedData      base XML data stream plus any derived values
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr
    Dim strFunctionName As String
    strFunctionName = "AddDerivedData"
    Dim objErrAssist As New ErrAssist
    AddDerivedData = vstrXMLData
    Set objErrAssist = Nothing
    Exit Function
            
    Set objErrAssist = Nothing
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984
    Exit Function
AddDerivedDataVBErr:
    Set objErrAssist = Nothing
    AddDerivedData = ""
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
End Sub
