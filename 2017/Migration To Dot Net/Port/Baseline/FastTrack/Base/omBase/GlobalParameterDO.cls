VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "GlobalParameterDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      GlobalParameterDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Code template for omiga4 Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function GetXMLFromRecordSet(vrstRecordSet As ADODB.Recordset) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:  ADOAssist
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' IK    20/07/1999  Created
' MCS   02/09/1999  GetMaximumPasswords added
' RF    29/09/1999  Added GetCurrentParameterByType
' RF    22/12/1999  Fix GetCurrentParameter.
'                   Improve performance of GetCurrentParameterByType.
' RF    08/05/2000  Improved error handling.
' MS    23/06/2000  Removed additional un-needed timing output
' APS   27/07/2000  Stress testing errors in retrieving from SPM
' MC    07/08/2000  SYS1409 Amend isolation mode for SPM to LockMethod as advised following load testing
' PSC   11/08/2000  SYS1430 Back out SYS1409
' LD    07/11/2000  Explicity close recordsets
' LD    07/11/2000  Explicity destroy command objects
' APS   29/11/2000  CORE000022
' LD    19/06/2001  SYS2386 All projects to use guidassist.bas rather than guidassist.cls
' DM    15/10/2001  SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
' AS    13/11/2003  CORE1 Removed GENERIC_SQL.
' TK    22/11/2004  BBG1821 - Performance related fixes.
' AS    22/03/2006  CORE175: BMIDS858 SQL queries for getting global parameters should be improved
' PSC   26/06/2006  CORE281 - Amend GetCurrentParameterByType to return correct value if column is empty
' SR    22/01/2008  CORE00000452 - run pre-migration convertor tool
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
 
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
Private Const cstrTABLE_NAME = "GLOBALPARAMETER"
 
 '=============================================
 'Variable Declaration Section
 '=============================================
 '=============================================
 'Constant Declaration Section
 '=============================================
'------------------------------------------------------------------------------------------
'BMIDS History:
' Prog  Date        Description
' SR    21/11/02    BMIDS01050 - Modified method 'GetSearchString'
' MDC   10/03/2003  BM0493 Performance Enhancements
'------------------------------------------------------------------------------------------
Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
' XML Request Format:
'   <GLOBALPARAMETER>
'       <NAME>string</NAME>
'       <GLOBALPARAMETERSTARTDATE>date</GLOBALPARAMETERSTARTDATE>
'       <DESCRIPTION>string</DESCRIPTION>
'       <AMOUNT>double</AMOUNT>
'       <MAXIMUMAMOUNT>double</MAXIMUMAMOUNT>
'       <PERCENTAGE>double</PERCENTAGE>
'       <BOOLEAN>boolean</BOOLEAN>
'       <STRING>string</STRING>
'   </GLOBALPARAMETER>
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    Const strFunctionName As String = "Create"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlDataElem As IXMLDOMElement
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Set objXmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    If objXmlDoc.getElementsByTagName(cstrTABLE_NAME).Length <> 1 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingPrimaryTag
    End If
    Set objXmlDataElem = objXmlDoc.getElementsByTagName(cstrTABLE_NAME).Item(0)
    'GetKeyString will raise an error if primary key(s) not found
    GetKeyString vstrXMLRequest
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Dim objSQLAssist As New SQLAssist
' TODO ************************************************************************
'   Build the SQL insert string for this row.
'   Extra processing will be required if sequence numbers are involved.
'
' e.g.
'       strCustomerNumber = xmlNodeList.Item(intRow - 1).childNodes(0).Text
'       e.g
'       strSQL = "INSERT INTO ADDRESS "
'       strSQL = strSQL & "(ADDRESSGUID, BUILDINGORHOUSENAME, BUILDINGORHOUSENUMBER, MAILSORTCODE) VALUES ("
'       strSQL = strSQL & objSQLAssist.FormatGuid(CreateGUID())
'       strSQL = strSQL & ", 'MSG Towers'"
'       strSQL = strSQL & ", '66'"
'       strSQL = strSQL & ", 123456)"
' *****************************************************************************
    Dim strSQL As String        ' SQL insert statement
    Dim strSQLCols As String    ' SQL insert statement, column list
    Dim strSQLValues As String  ' SQL insert statement, value list
    Dim intChildNodes As Integer, intLoop As Integer
    Dim blnHasData As Boolean
    Dim eDbDataType As DBDATATYPE
        
    blnHasData = False
    intChildNodes = objXmlDataElem.childNodes.Length
    For intLoop = 0 To (intChildNodes - 1)
        eDbDataType = GetDbType(objXmlDataElem.childNodes.Item(intLoop).nodeName)
        If eDbDataType <> dbdtNotStored Then
            If blnHasData Then
                strSQLCols = strSQLCols & ", "
                strSQLValues = strSQLValues & ", "
            End If
            blnHasData = True
            strSQLCols = strSQLCols & objXmlDataElem.childNodes.Item(intLoop).nodeName
            Select Case eDbDataType
                Case dbdtGuid
                    strSQLValues = strSQLValues & objSQLAssist.FormatGuid(objXmlDataElem.childNodes.Item(intLoop).Text)
                Case dbdtString
                    strSQLValues = strSQLValues & objSQLAssist.FormatString(objXmlDataElem.childNodes.Item(intLoop).Text)
                Case dbdtInt, dbdtBoolean, dbdtComboId, dbdtCurrency, dbdtDouble
                    strSQLValues = strSQLValues & objXmlDataElem.childNodes.Item(intLoop).Text
                Case dbdtDate
                    strSQLValues = strSQLValues & objSQLAssist.FormatDateString(objXmlDataElem.childNodes.Item(intLoop).Text)
            End Select
        End If
    Next
    If blnHasData = False Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoDataForCreate
    End If
    strSQL = "INSERT INTO " & cstrTABLE_NAME & " (" & strSQLCols & ") VALUES (" & strSQLValues & ")"
    objADOAssist.ExecuteSQLCommand (strSQL)
    Set objXmlDataElem = Nothing
    Set objXmlDoc = Nothing
    Set objADOAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
CreateVbErr:
    Set objXmlDataElem = Nothing
    Set objXmlDoc = Nothing
    Set objADOAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    Const strFunctionName As String = "Update"
'    Dim objErrAssist As New ErrAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlDataElem As IXMLDOMElement
'    Dim objXmlAssist As New XMLAssist
    Set objXmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    If objXmlDoc.getElementsByTagName(cstrTABLE_NAME).Length <> 1 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingPrimaryTag
    End If
    Set objXmlDataElem = objXmlDoc.getElementsByTagName(cstrTABLE_NAME).Item(0)
    ' build the key string
    Dim strKeys As String
    strKeys = GetKeyString(vstrXMLRequest)
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Dim objSQLAssist As New SQLAssist
    ' build the full SQL string
    Dim strSQL As String        ' SQL insert statement
    Dim strSQLSets As String    ' SQL insert statement, column list
    Dim intChildNodes As Integer, intLoop As Integer
    Dim blnHasData As Boolean
    Dim eDbDataType As DBDATATYPE
    Dim strElementText As String
        
    blnHasData = False
    intChildNodes = objXmlDataElem.childNodes.Length
    For intLoop = 0 To (intChildNodes - 1)
        ' do not update primary keys
        If objXmlDataElem.childNodes.Item(intLoop).nodeName <> "NAME" And _
           objXmlDataElem.childNodes.Item(intLoop).nodeName <> "GLOBALPARAMETERSTARTDATE" Then
            eDbDataType = GetDbType(objXmlDataElem.childNodes.Item(intLoop).nodeName)
            If eDbDataType <> dbdtNotStored Then
                If blnHasData Then
                    strSQLSets = strSQLSets & ", "
                End If
                blnHasData = True
                strSQLSets = strSQLSets & objXmlDataElem.childNodes.Item(intLoop).nodeName & " = "
                Select Case eDbDataType
                    Case dbdtGuid
                        strSQLSets = strSQLSets & objSQLAssist.FormatGuid(objXmlDataElem.childNodes.Item(intLoop).Text)
                    Case dbdtString
                        strSQLSets = strSQLSets & objSQLAssist.FormatString(objXmlDataElem.childNodes.Item(intLoop).Text)
                    Case dbdtInt, dbdtDouble, dbdtCurrency, dbdtComboId
                        strElementText = objXmlDataElem.childNodes.Item(intLoop).Text
                        If Len(strElementText) = 0 Then
                            strSQLSets = strSQLSets & "Null"
                        Else
                            strSQLSets = strSQLSets & objXmlDataElem.childNodes.Item(intLoop).Text
                        End If
                End Select
            End If
        End If
    Next
    If blnHasData = False Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingPrimaryTag
    End If
    strSQL = "UPDATE " & cstrTABLE_NAME & " SET " & strSQLSets & " WHERE " & strKeys
        
    ' do the update
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    objADOAssist.ExecuteSQLCommand (strSQL)
    Set objXmlDataElem = Nothing
    Set objXmlDoc = Nothing
    Set objADOAssist = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
UpdateVbErr:
    
    Set objXmlDataElem = Nothing
    Set objXmlDoc = Nothing
    Set objADOAssist = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub Delete(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance to be deleted
' Raise Errors:
'       omiga4RecordNotFound
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    Const strFunctionName As String = "Delete"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
    ' build the key string
    Dim strKeys As String
    strKeys = GetKeyString(vstrXMLRequest)
    ' check the record exists on the database
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Call objADOAssist.CheckSingleRecordExists(cstrTABLE_NAME, strKeys)
    ' build the full SQL string
    Dim strSQL As String
    strSQL = "delete from " & cstrTABLE_NAME & " where " & strKeys
    ' do the delete
    objADOAssist.ExecuteSQLCommand strSQL
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteVbErr:
    Set objADOAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    Const strFunctionName As String = "GetData"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
    Dim cmd As New ADODB.Command
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    cmd.ActiveConnection = objADOAssist.GetConnStr
    ' build the key string
    Dim strKeys As String
    strKeys = GetKeyString(vstrXMLRequest)
    ' build the full select statement
    Dim strSQL As String
    strSQL = "select * from " & cstrTABLE_NAME & " where " & strKeys
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        GetData = Me.GetXMLFromRecordSet(rstThisRecSet)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
        Set cmd = Nothing
'    Set objErrAssist = Nothing
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetDataVbErr:
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If
    Set rstThisRecSet = Nothing
        Set cmd = Nothing
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        End If
    End If
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetXMLFromRecordSet(ByVal vrstRecordSet As ADODB.Recordset) As String
' header ----------------------------------------------------------------------------------
' description:
'   Create XML data stream from field elements in ADO Recordset.
'   Add any values derived from field elements to XML.
' pass:
'   vrstRecordSet       ADO Recordset positioned at the required record
' return:
'   GetXMLFromRecordSet string containing XML data stream
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetXMLFromRecordSetVbErr
    Const strFunctionName As String = "GetXMLFromRecordSet"
    Dim objSQLAssist As SQLAssist
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim objXmlDocsertNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
                        
    Set objSQLAssist = New SQLAssist
                        
    Dim objComboDO As ComboDO
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(App.Title & ".ComboDO")
    End If
    
    Set objXmlOut = New FreeThreadedDOMDocument40
    Set objXmlElem = objXmlOut.createElement(cstrTABLE_NAME)
    Set objXmlDocsertNode = objXmlOut.appendChild(objXmlElem)
    
    Dim intField As Integer
    Dim strElementName As String
    Dim strFormatted As String, strFormatMask As String, strComboGroup As String
    Dim eDbType As DBDATATYPE
    
    For intField = 0 To (vrstRecordSet.Fields.Count - 1)
        
        If Not IsNull(vrstRecordSet.Fields.Item(intField).Value) Then
            
            strElementName = vrstRecordSet.Fields.Item(intField).Name
            
            eDbType = GetDbType(strElementName)
            
            If eDbType <> dbdtNotStored Then
            
                Set objXmlElem = objXmlOut.createElement(strElementName)
                
                Select Case eDbType
                    Case dbdtGuid
                        objXmlElem.Text = objSQLAssist.GuidToString(vrstRecordSet.Fields.Item(intField).Value)
                    Case dbdtCurrency
                        strFormatted = Format$(vrstRecordSet.Fields.Item(intField).Value, "0.0000000")
                        objXmlElem.Text = Left$(strFormatted, InStr(1, strFormatted, ".") + 2)
                    Case dbdtDouble
                        If GetFormatMask(strElementName, strFormatMask) Then
                            objXmlElem.setAttribute "RAW", vrstRecordSet.Fields.Item(intField).Value
                            objXmlElem.Text = Format$(vrstRecordSet.Fields.Item(intField).Value, strFormatMask)
                        Else
                            objXmlElem.Text = vrstRecordSet.Fields.Item(intField).Value
                        End If
                    Case dbdtComboId
                        If GetComboGroup(strElementName, strComboGroup) Then
                            objXmlElem.setAttribute "TEXT", objComboDO.GetComboText(strComboGroup, vrstRecordSet.Fields.Item(intField).Value)
                            objXmlElem.Text = vrstRecordSet.Fields.Item(intField).Value
                        End If
                    Case dbdtDate
                        objXmlElem.Text = objSQLAssist.DateToString(vrstRecordSet.Fields.Item(intField).Value)
                    Case Else
                        objXmlElem.Text = vrstRecordSet.Fields.Item(intField).Value
                End Select
                Set objXmlNode = objXmlDocsertNode.appendChild(objXmlElem)
            End If
        End If
                
    Next
    '   add derived values to generated XML
    GetXMLFromRecordSet = Me.AddDerivedData(objXmlOut.xml)
    Set objComboDO = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlDocsertNode = Nothing
    Set objXmlOut = Nothing
    Set objSQLAssist = Nothing
    Exit Function
GetXMLFromRecordSetVbErr:
    
    Set objComboDO = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlDocsertNode = Nothing
    Set objXmlOut = Nothing
    Set objSQLAssist = Nothing
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    '   re-raise error for calling function to handle as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
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
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr
    Const strFunctionName As String = "AddDerivedData"
    AddDerivedData = vstrXMLData
    Exit Function
AddDerivedDataVBErr:
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function GetKeyString(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'       Check all keys are present in the XML, and return them with their
'       associated values in a format ready for including in a SQL execution
'       string.
'       Raise an error if not all keys have values specified.
' pass:
'       vstrXMLRequest              xml data
' return:
'       A string in the format "TableName.KeyFieldA = ValueX and
'       TableName.KeyFieldB = ValueY and TableName.KeyFieldC = ValueZ"
' Raise Errors:
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo GetKeyStringVBErr:
    
    Const strFunctionName As String = "GetKeyString"
'    Dim objErrAssist As New ErrAssist
    Dim objXmlDoc As New FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
    Set objXmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Dim objSQLAssist As New SQLAssist
    Dim strKeys As String
    If objXmlDoc.getElementsByTagName("NAME").Length = 0 Or _
       objXmlDoc.getElementsByTagName("GLOBALPARAMETERSTARTDATE").Length = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InvalidKeyString
    End If
    If Len(objXmlDoc.getElementsByTagName("NAME").Item(0).Text) = 0 Or _
       Len(objXmlDoc.getElementsByTagName("GLOBALPARAMETERSTARTDATE").Item(0).Text) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InvalidKeyString
    End If
    strKeys = cstrTABLE_NAME & ".NAME = " & _
              objSQLAssist.FormatString(objXmlDoc.getElementsByTagName("NAME").Item(0).Text) & _
              " AND " & cstrTABLE_NAME & ".GLOBALPARAMETERSTARTDATE = " & _
              objSQLAssist.FormatDate(objXmlDoc.getElementsByTagName("GLOBALPARAMETERSTARTDATE").Item(0).Text)
    GetKeyString = strKeys
    Set objXmlDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Exit Function
GetKeyStringVBErr:
    Set objXmlDoc = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function GetPartialKeys(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'       partial key as defined for GetList & DeleteAll is:
'       Name
'       Check all keys are present in the XML, and return them with their
'       associated values in a format ready for including in a SQL execution
'       string.
'       Raise an error if not all keys have values specified.
' pass:
'       vstrXMLRequest              xml data
' return:
'       A string in the format "TableName.KeyFieldA = ValueX and
'       TableName.KeyFieldB = ValueY and TableName.KeyFieldC = ValueZ"
' Raise Errors:
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo GetPartialKeysVBErr
    
    Const strFunctionName As String = "GetPartialKeys"
'    Dim objErrAssist As New ErrAssist
    Dim objXmlDoc As New FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
    Set objXmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Dim objSQLAssist As New SQLAssist
    Dim strKeys As String
    If objXmlDoc.getElementsByTagName("NAME").Length = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InvalidKeyString
    End If
    If Len(objXmlDoc.getElementsByTagName("NAME").Item(0).Text) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InvalidKeyString
    End If
    strKeys = cstrTABLE_NAME & ".NAME = " & _
              objSQLAssist.FormatString(objXmlDoc.getElementsByTagName("NAME").Item(0).Text)
    GetPartialKeys = strKeys
    Set objXmlDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Exit Function
GetPartialKeysVBErr:
    Set objXmlDoc = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Sub DeleteAll(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    Const strFunctionName As String = "DeleteAll"
'    Dim objErrAssist As New ErrAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim strSQL As String
    Dim strKeys As String
    '   will raise error if keys not present
    strKeys = Me.GetPartialKeys(vstrXMLRequest)
    strSQL = "delete from " & cstrTABLE_NAME & " where " & strKeys
    ' do the delete
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    objADOAssist.ExecuteSQLCommand (strSQL)
    Set objADOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteAllVbErr:
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function FindList(ByVal vstrXMLRequest As String) As String
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
'History:
'
'Prog  Date     Description
'MV    30/11/00 Deleted Function FindCurrentParameterList
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    Const strFunctionName As String = "FindList"
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlElem As IXMLDOMElement
    Dim cmd As New ADODB.Command
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    cmd.ActiveConnection = objADOAssist.GetConnStr
    Dim strSQL As String
    Dim strKeys As String
    '   will raise error if keys not present
    strKeys = Me.GetPartialKeys(vstrXMLRequest)
    ' build the full select statement
    strSQL = "select * from " & cstrTABLE_NAME & " where " & strKeys
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    ' build xml doc made up of each record, delimited by table name
    ' within TABLENAMELIST
    Set objXmlElem = objXmlOut.createElement(cstrTABLE_NAME & "LIST")
    objXmlOut.appendChild objXmlElem
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rstThisRecSet.MoveFirst
    While Not rstThisRecSet.EOF
        
        Dim strXML As String
        strXML = Me.GetXMLFromRecordSet(rstThisRecSet)
        Set objXmlDoc = m_objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        objXmlElem.appendChild objXmlDoc.documentElement
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    FindList = objXmlOut.xml
    Set rstThisRecSet = Nothing
        Set cmd = Nothing
    Set objXmlElem = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
FindListVbErr:
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objXmlElem = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        End If
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetDbType(ByVal vstrElementName As String) As DBDATATYPE
' header ----------------------------------------------------------------------------------
' description:
'   return a value from enumerated list (DBDATATYPE in SQLAssist) which indicates
'   the data 'type' for the value associated with a particular (xml) element
' pass:
'   vstrElementName the xml element name
' return:
'   GetDbType       enumerated value
' Raise Errors:     n/a
'------------------------------------------------------------------------------------------
        
    Select Case vstrElementName
        Case "NAME", "DESCRIPTION", "STRING"
            GetDbType = dbdtString
        Case "AMOUNT", "MAXIMUMAMOUNT", "PERCENTAGE"
            GetDbType = dbdtDouble
        Case "BOOLEAN"
            GetDbType = dbdtBoolean
        Case "GLOBALPARAMETERSTARTDATE"
            GetDbType = dbdtDate
        Case Else
            GetDbType = dbdtNotStored
    End Select
End Function
Private Function GetFormatMask(ByVal vstrElementName As String, ByRef rstrMask As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   return a string to be used by Format function to convert a double value
'   to a string
' pass:
'   vstrElementName the xml element name
' return:
'   rstrMask        the format mask
'   GetFormatMask   True if format mask specified for element
' Raise Errors:     n/a
'------------------------------------------------------------------------------------------
        
    Select Case vstrElementName
        Case Else
            GetFormatMask = False
    End Select
End Function
Private Function GetComboGroup(ByVal vstrElementName As String, ByRef rstrGroup As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   return the name of a Combo Group to be used to look up a Combo text value
'   for the combo index value associated with an xml element name
' pass:
'   vstrElementName the xml element name
' return:
'   rstrGroup       the combo group name
'   GetComboGroup   True if a combo group name is specified for element
' Raise Errors:     n/a
'------------------------------------------------------------------------------------------
        
    Select Case vstrElementName
        Case Else
            GetComboGroup = False
    End Select
End Function
Public Function GetCurrentParameter(ByVal vstrParamName As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object, selected the current instance of the data as determined
'   by GLOBALPARAMETERSTARTDATE
' pass:
'   strParamName
'       name of the parameter to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
'------------------------------------------------------------------------------------------
' RF 22/12/99 Fix and improve performance.
' APS 29/11/2000 : CORE000022 Added SPM functionality
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterVbErr
    Const strFunctionName As String = "GetCurrentParameter"
    
    Dim strGlobalParameterXML As String

'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984
        Dim objADOAssist As IADOAssist
        Dim cmd As ADODB.Command
        Dim rstThisRecSet As ADODB.Recordset
        Dim adoConn As ADODB.Connection
        
        If m_objContext Is Nothing Then
            Set objADOAssist = New ADOAssist
        Else
            Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
        End If
    
        ' AS 22/03/2006 CORE175: BMIDS858 GHun Retrieve GlobalParameter by calling a stored proc
        Set adoConn = New ADODB.Connection
        adoConn.ConnectionString = objADOAssist.GetConnStr
        adoConn.CursorLocation = adUseClient
        adoConn.Open
        
        Set cmd = New ADODB.Command
        
        With cmd
            Set .ActiveConnection = adoConn
            .CommandText = "USP_GetGlobalParameter"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("@p_Name", adBSTR, adParamInput, Len(vstrParamName), vstrParamName)
            Set rstThisRecSet = .Execute
            Set .ActiveConnection = Nothing
        End With
        
        If rstThisRecSet.EOF Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
        End If
        
        ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
        ' rstThisRecSet.MoveFirst
            
        strGlobalParameterXML = Me.GetXMLFromRecordSet(rstThisRecSet)
        rstThisRecSet.Close
        adoConn.Close
        ' AS 22/03/2006 CORE175 End
        
        Set rstThisRecSet = Nothing
        Set cmd = Nothing
        Set objADOAssist = Nothing
        Set adoConn = Nothing
        
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984
        
    GetCurrentParameter = strGlobalParameterXML
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetCurrentParameterVbErr:
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    Set adoConn = Nothing
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' AS 22/03/2006 CORE175 No longer used
'Private Function ReadCurrentParameter( _
'    ByVal vstrParamName As String, ByVal vstrFieldName As String) As Variant
'' header ----------------------------------------------------------------------------------
'' description:
'' pass:
'' return:
''------------------------------------------------------------------------------------------
'On Error GoTo ReadCurrentParameterVbErr
'    Dim strFunctionName As String
'    strFunctionName = "ReadCurrentParameter"
''    Dim objErrAssist As New ErrAssist
'    Dim cmd As New ADODB.Command
'    Dim rst As New ADODB.Recordset
'    Dim objADOAssist As ADOAssist
''    Dim objContext As ObjectContext
''    Set objContext = GetObjectContext()
'    If m_objContext Is Nothing Then
'        Set objADOAssist = New ADOAssist
'    Else
'        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
'    End If
'    cmd.ActiveConnection = objADOAssist.GetConnStr
'    cmd.CommandText = GetSearchString(vstrParamName, vstrFieldName)
'    Set rst = cmd.Execute(1, , adCmdText)
'    If rst.EOF Then
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
'    End If
'    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
'    ' rst.MoveFirst
'    ReadCurrentParameter = rst.Fields(0)
'        rst.Close
'    Set rst = Nothing
'    Set cmd = Nothing
''    Set objErrAssist = Nothing
'    Set objADOAssist = Nothing
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
''    Set objContext = Nothing
'    Exit Function
'ReadCurrentParameterVbErr:
'
'    ' Close the recordset if still open
'        If Not rst Is Nothing Then
'                If rst.State = adStateOpen Then
'                        rst.Close
'                End If
'        End If
'    Set rst = Nothing
'    Set cmd = Nothing
'    Set objADOAssist = Nothing
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
''    Set objContext = Nothing
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'    m_objErrAssist.AddToErrSource strFunctionName
''    Set objErrAssist = Nothing
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function
' AS 22/03/2006 CORE175 End
'FIXIT: Declare 'GetCurrentParameterByType' with an early-bound data type                  FixIT90210ae-R1672-R1B8ZE
Public Function GetCurrentParameterByType( _
    ByVal vstrParamName As String, _
    ByVal vstrFieldName As String) _
    As Variant
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterByTypeVbErr
    
    Const strFunctionName As String = "GetCurrentParameterByType"
    
'FIXIT: Declare 'varParamValue' with an early-bound data type                              FixIT90210ae-R1672-R1B8ZE
    Dim varParamValue As Variant
    Dim xmlDoc As FreeThreadedDOMDocument
    Dim xmlNode As IXMLDOMNode
    
    ' APS 29/11/2000 : CORE000022
    If (Len(vstrFieldName) = 0) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingParameter, "GlobalParameter Field Name"
    End If
    
' APS 27/07/00 - Stress testing errors in retrieving from SPM, Compiler directive
' has been added to allow omBase to be built without using SPM for Globals
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984

        ' AS 22/03/2006 CORE175
        'varParamValue = ReadCurrentParameter(vstrParamName, vstrFieldName)
        Set xmlDoc = New FreeThreadedDOMDocument
        xmlDoc.loadXML GetCurrentParameter(vstrParamName)
        Set xmlNode = xmlDoc.selectSingleNode(".//" & UCase$(vstrFieldName))
        If xmlNode Is Nothing Then
            varParamValue = Empty       ' PSC 26/06/2006 CORE281
        Else
            varParamValue = xmlNode.Text
        End If
        ' AS 22/03/2006 CORE175 End
        
'FIXIT: '#If' not upgraded reliably to Visual Basic .NET                                   FixIT90210ae-R2789-H1984

    GetCurrentParameterByType = varParamValue
    
    Exit Function
    
GetCurrentParameterByTypeVbErr:
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'FIXIT: Declare 'rvarValue' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
Private Function GetGlobalFromSpam(ByVal vstrParamName As String, _
                                   ByVal vstrFieldName As String, _
                                   ByRef rvarValue As Variant) _
                                   As Boolean
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetGlobalFromSpamVbErr
    Const strFunctionName As String = "GetGlobalFromSpam"
'    Dim objErrAssist As New ErrAssist
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim bExists As Boolean
    Dim strPropertyName As String
    ' APS 29/11/2000 : CORE000023
    If Len(vstrFieldName) > 0 Then
        strPropertyName = vstrParamName & "." & vstrFieldName
    Else
        strPropertyName = vstrParamName
    End If
    Dim lIsolationMode As Long
    'MC 07/08/2000. Amend isolation mode to LockMethod as advised following load testing
    'PSC 11/08/2000  SYS1430 Back out SYS1409
    lIsolationMode = LockSetGet
    Dim lReleaseMode As Long
    lReleaseMode = Process
        
    Set spmMgr = CreateObject(gstrMSGSPM)
        
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("GlobalParams", lIsolationMode, lReleaseMode, bExists)
            
    Set spmProperty = spmGroup.CreateProperty(strPropertyName, bExists)
            
    If bExists Then
        
        If Not IsEmpty(spmProperty.Value) Then
            rvarValue = spmProperty.Value
            GetGlobalFromSpam = True
        End If
    Else
        spmProperty.Value = Empty
        GetGlobalFromSpam = False
    End If
                
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Set objErrAssist = Nothing
    Exit Function
                    
GetGlobalFromSpamVbErr:
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'FIXIT: Declare 'vvarValue' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
Private Sub PutGlobalInSpam(ByVal vstrParamName As String, _
                            ByVal vstrFieldName As String, _
                            ByVal vvarValue As Variant)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim bExists As Boolean
    Dim strPropertyName As String
    ' APS 29/11/2000 : CORE000022
    If Len(vstrFieldName) > 0 Then
        strPropertyName = vstrParamName & "." & vstrFieldName
    Else
        strPropertyName = vstrParamName
    End If
    Dim lIsolationMode As Long
    'MC 07/08/2000. Amend isolation mode to LockMethod as advised following load testing
    'PSC 11/08/2000  SYS1430 Back out SYS1409
    lIsolationMode = LockSetGet
    Dim lReleaseMode As Long
    lReleaseMode = Process
        
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("GlobalParams", lIsolationMode, lReleaseMode, bExists)
            
    Set spmProperty = spmGroup.CreateProperty(strPropertyName, bExists)
    spmProperty.Value = vvarValue
                
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
                    
End Sub
'AS 22/03/2006 CORE175 No longer used
'Private Function GetSearchString( _
'    ByVal vstrParamName As String, ByVal vstrFieldName As String) As String
'' header ----------------------------------------------------------------------------------
'' description:
''   Get search string for GetCurrentParameter and GetCurrentParameterByType.
'' pass:
'' return:
''------------------------------------------------------------------------------------------
'    Dim strFunctionName As String
'    strFunctionName = "GetSearchString"
'    Dim objADOAssist As New ADOAssist 'SR 21/11/2002 BMIDS01050
'    Dim objSQLAssist As New SQLAssist
'    Dim strSQL As String
'    If Len(vstrFieldName) = 0 Then
'        strSQL = "select *"
'    Else
'        strSQL = "select " & vstrFieldName
'    End If
'    'SR 21/11/2002 BMIDS01050 :[JR SYSMCP1221] - Remove UPPER for SQLServer.
'    strSQL = strSQL & " from " & cstrTABLE_NAME
'    If objADOAssist.GetDBEngine = SQLServer Then
'        strSQL = strSQL & " WHERE NAME = " & objSQLAssist.FormatString(vstrParamName) & ""
'    ElseIf objADOAssist.GetDBEngine = Oracle Then
'        strSQL = strSQL & " WHERE UPPER(NAME) = UPPER(" & objSQLAssist.FormatString(vstrParamName) & ")"
'    End If
'     strSQL = strSQL & " AND GLOBALPARAMETERSTARTDATE <= " & objSQLAssist.GetSystemDate() & _
'        " ORDER BY GLOBALPARAMETERSTARTDATE ASC"
'    'SR 21/11/2002 BMIDS01050 - End
'
'    GetSearchString = strSQL
'    Set objSQLAssist = Nothing
'    Set objADOAssist = Nothing
'End Function
'AS 22/03/2006 CORE175 End
'BM0493 MDC 10/04/2003
Public Function GetCurrentParameterListEx(ByVal vstrRequestXML As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the current global parameter values for a list of parameter names.
' pass:
'   strParamName
'       name of the parameter to be retrieved
' return:
'   string containing XML data stream representation of data retrieved
' NB:
'   This method contains SQL Server specific code to retrieve XML data!
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterListExExit
    Const strFunctionName As String = "GetCurrentParameterListEx"
        
    Dim cmd As ADODB.Command
    Dim adoStr As ADODB.Stream
    Dim objADOAssist As IADOAssist
    'Dim objErrAssist As ErrAssist
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlName As IXMLDOMNode
    Dim xmlNameList As IXMLDOMNodeList
    Dim strSQL As String
    Dim strSQLWhere As String
    Dim intCount As Integer
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
'    Set objErrAssist = New ErrAssist
    Set xmlDoc = New FreeThreadedDOMDocument40
    If Not xmlDoc.loadXML(vstrRequestXML) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeXMLParserError
    End If
    strSQL = "SELECT NAME, MAXIMUMAMOUNT, PERCENTAGE, AMOUNT, BOOLEAN, STRING FROM GLOBALPARAMETER"
    Set xmlNameList = xmlDoc.selectNodes(".//GLOBALPARAMETER/NAME")
    If xmlNameList.Length > 0 Then
        strSQLWhere = strSQLWhere & " WHERE NAME IN ("
        For Each xmlName In xmlNameList
            If intCount > 0 Then
                strSQLWhere = strSQLWhere & ", "
            End If
            strSQLWhere = strSQLWhere & "'" & xmlName.Text & "'"
            intCount = intCount + 1
        Next
        strSQLWhere = strSQLWhere & ") AND"
    Else
        strSQLWhere = " WHERE"
    End If
    strSQLWhere = strSQLWhere & " GLOBALPARAMETERSTARTDATE = (SELECT MAX(GLOBALPARAMETERSTARTDATE)"
    strSQLWhere = strSQLWhere & " from GLOBALPARAMETER GP2 WHERE GP2.NAME = GLOBALPARAMETER.NAME "
    strSQLWhere = strSQLWhere & " AND GLOBALPARAMETERSTARTDATE <= GETDATE() GROUP BY NAME) "
    strSQL = strSQL & strSQLWhere & "for xml auto, elements"
    Set adoStr = New ADODB.Stream
    adoStr.Open
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = objADOAssist.GetConnStr
        .CommandText = strSQL
        .CommandType = adCmdText
        .Properties("Output Stream") = adoStr
        .Properties("XML Root") = "GLOBALPARAMETERLIST"
        .Execute , , adExecuteStream
    End With
        
    GetCurrentParameterListEx = adoStr.ReadText
       
GetCurrentParameterListExExit:
    Set cmd = Nothing
    Set objADOAssist = Nothing
    Set adoStr = Nothing
'    Set objErrAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlName = Nothing
    Set xmlNameList = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Function
End Function
'BM0493 MDC 10/04/2003 - End
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing
End Sub
