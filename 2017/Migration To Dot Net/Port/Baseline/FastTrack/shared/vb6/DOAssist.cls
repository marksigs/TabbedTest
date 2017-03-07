VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "DOAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      DOAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MCS   26/08/99    Initial Creation
' MCS   24/09/99    Modification following code review
' AY    13/10/99    FindList changed to allow an "order by" setting
'                   Improved error message for incorrect combination of parameters
'                   to BuildSQLString
' RF    20/10/99    Improved error handling in GetDbType
' MCS   19/11/99    Findlist/BuildSQLString modified. To use the specific fields needed
'                   to be returned in the SQL search
' PSC   15/12/99    Make the xml node obptional on GetXMLFromRecordsetEx and
'                   GetXMLFromTableSchema. Remove node from GetDataEx and FindListEx
' PSC   17/12/99    Put all Ex functions onto new interface and remove Ex from name
' PSC   20/12/99    Amend to use new IADOAssist interface
' JLD   07/12/00    Changed GetNextSequenceNumberEx to use getElementsByTagName()
'                   instead of selectNodes() to guarantee the node being found anywhere in the xml.
' RF    25/01/00    GetNextSequenceNumberEx put in interface as GetNextSequenceNumber.
' PSC   01/02/00    Don't clone nodes on return
' SR    03/02/00    New method - GetComponentData
' SR    14/03/00    New parameters to method - FindListMultiple
' MH    05/04/00    Added more diag information to GetXMLFromTableSchema
' MS    23/06/00    Removed additional un-needed timing output
' LD    07/11/00    Explicity close recordsets
' LD    07/11/00    Explicity destroy command objects
' LD    29/11/00    Explicity close recordsets
' PSC   23/01/01    SYS1871 - Amend GetNextSequenceNumber to close the recordset once and also
'                   to close it in the error handler
' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
' AS    13/11/03        CORE1 Removed GENERIC_SQL.
' SDS    20/01/2004  LIVE00009653 - NOLOCK is used only for SQLServer sql statements
' SR     24/01/2008  CORE00000452 - apply pre-migration convertor tool
'------------------------------------------------------------------------------------------

Option Explicit
Implements IDOAssist

' TK 23/08/2004 BBG1211 Start
Implements ObjectControl

Private m_objContext As ObjectContext

' n.b. m_objComboDO is not initialised in ObjectControl_Activate() - only where used
Private m_objComboDO As ComboDO
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
' TK 23/08/2004 BBG1211 End

#Const PROFILING = 1
'Public Enum SQL_FORMAT_TYPE
'    ' fields and values in 2 separate strings,
'    ' e.g. "FieldA,FieldB" and "Value1,Value2"
'    sftFieldValueSeparated
'
'    ' field=value pairs separated by commas,
'    ' e.g. "FieldA = Value1, FieldB = Value2"
'    sftCommaSeparated
'
'    ' field=value pairs separated by " AND ",
'    ' e.g. "FieldA = Value1 AND FieldB = Value2"
'    sftAndSeparated
'End Enum
'
'Public Enum CLASS_DEF_KEY
'    cdkPRIMARYKEY
'    cdkOTHERS
'End Enum
'
'Public Enum CLASS_DEF_KEY_AMOUNT
'    cdkaALLKEYS
'    cdkaANYKEYS
'End Enum
'Private Const cstrVALUE As String = " ) VALUES ( "
'Private Const cstrWHERE As String = " WHERE "
' value used to check whether an optional parameter passed ByRef has been supplied
Private Const cstrNotSpecified = "Not specified"
'------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
' SR    02/01/2003  BM0209 - Modified methods 'FindListEx' and 'IDOAssist_FindList'
' SR    08/01/2003  BM0209 - Modified methods 'FindListEx' and 'IDOAssist_FindList' (Added
'                   NOLOCK hint to SQl executed, whereever required.
' MDC   14/01/2003  BM0249 - Handle SQLNOLOCK in GetXMLForTableSchema
' RF    12/05/2003  BM0536 - Handle SQLNOLOCK in IDOAssist_FindListMultiple
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
' Prog  Date        Description
' TK    23/08/2004  BBG1211 - omApp.ApplicationManagerBO.ImportAccountsIntoApplication (extracted from BMIDS788)
'------------------------------------------------------------------------------------------
Public Function GetData( _
    ByVal vstrXMLRequest As String, ByVal vstrClassDef As String) As String
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
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    Dim strFunctionName As String
    strFunctionName = "GetData"
        
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
                
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, , , , strFieldValuePair
        
    ' build the full select statement
    strSQL = "select * from " & strTable & " where " & strFieldValuePair
'    #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName & _
'        "[ Before Dim cmd As New ADODB.Command ]", True
'    #End If
        
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    WriteSQLToFile strSQL, strFunctionName
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
            
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'    #End If
            
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        GetData = Me.GetXMLFromRecordSet(rstThisRecSet, vstrClassDef)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
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
Public Sub BuildSQLString( _
    ByVal vstrXMLRequest As String, _
    ByVal vstrXMLClassDef As String, _
    ByRef rstrTable As String, _
    Optional ByVal vsftSQLFormatType As SQL_FORMAT_TYPE = sftAndSeparated, _
    Optional ByVal vcdkKey As CLASS_DEF_KEY = cdkPRIMARYKEY, _
    Optional ByVal vcdkKeyValue As CLASS_DEF_KEY_AMOUNT = cdkaALLKEYS, _
    Optional ByRef rstrFieldValuePair As String = cstrNotSpecified, _
    Optional ByRef rstrFields As String = cstrNotSpecified, _
    Optional ByRef rstrValues As String = cstrNotSpecified, _
    Optional ByRef rstrSelect As String _
    )
' header ----------------------------------------------------------------------------------
' description:
'       Check all keys are present in the XML, and return them with their
'       associated values in a format ready for including in a SQL execution
'       string.
'       Raise an error dependant on the value CLASS_DEF_KEY_AMOUNT
'       if not all keys have values specified.
' pass:
'   vstrXMLRequest
'       xml request data from which to build SQL string
'   vstrXMLClassDef
'       xml ClassDefinition describing the table, fields and formats
'       to be used to parse the vstrXMLRequest against
'   rstrTable
'       Base table name
'   vsftSQLFormatType
'       Enumerated SQL_FORMAT_TYPE used to determine the format of the
'       rstrFieldValuePair or rstrFields And rstrValues that are returned
'       sftFieldValueSeparated:
'           As used in create, e.g "insert into TABLE ( FieldA,FieldB )
'                                   values ( Value1,Value2 )"
'       sftCommaSeparated:
'           e.g. "update TABLE set FieldA = Value1,FieldB = Value2 where"
'       sftAndSeparated:
'           e.g. "select * from TABLE where FieldA = Value1 AND FieldB = Value2"
'       DEFAULT: sftAndSeparated
'   vcdkKey
'       Enumerated value CLASS_DEF_KEY. This and vcdkKeyValue (below) are
'       used together to Build SQL string and raise appropriate errors. E.g.:
'           GetData     cdkaALLKEYS cdkPRIMARYKEY
'           Delete      cdkaALLKEYS cdkPRIMARYKEY
'           DeleteAll   cdkaANYKEYS cdkPRIMARYKEY
'           FindList    cdkaANYKEYS cdkPRIMARYKEY
'           Update      cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
'           Create      cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
'       cdkPRIMARYKEY:
'           All primary key fields for given table
'       cdkOTHERS
'           Everything else that isn't a primary key
'       DEFAULT: cdkPRIMARYKEY
'   vcdkKeyValue
'       Enumerated value CLASS_DEF_KEY_AMOUNT - used as above with vcdkKey
'       to provide error handling depending on which combination is specified.
'       cdkaALLKEYS:
'           ALL key values expected
'       cdkaANYKEYS:
'           Any of the key values expected
'       DEFAULT: cdkaALLKEYS
'   rstrFieldValuePair
'       String created or appended to, which contains SQL formatted values
'       based on the settings of SQL_FORMAT_TYPE.
'       Used with sftAndSeparated and sftCommaSeparated; raises error if either of
'       these is specified and no parameter passed.
'       Used by GetData, Delete, DeleteAll, FindList.
'   rstrFields, rstrValues
'       String created or appended to, which contains SQL formatted values
'       based on the settings of SQL_FORMAT_TYPE.
'       Used with sftFieldValueSeparated; raises error if this is specified and no
'       parameter is passed.
'       Used by Create.
'   rstrSelect
'       Optional string used in Find, when used with a subset of the XML class/table
'       Definition. Will return a string of Table.Fields for the provided XML
' return:   n/a
' Raise Errors:
'       omiga4InvalidKeyString
'       omiga4InValidKeyValueSpecified
'       omiga4Invalidparameter
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo BuildSQLStringVBErr:
    
    Dim strFunctionName As String
    strFunctionName = "BuildSQLString"
        
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlClassDefDoc As New FreeThreadedDOMDocument40
    Dim objSQLAssist As New SQLAssist
    Dim xmlTypeElem As IXMLDOMElement
    Dim objElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Set objXmlDoc = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlClassDefDoc = m_objXmlAssist.load( _
        vstrXMLClassDef, TypeName(Me), strFunctionName)
    Set objElem = objXmlClassDefDoc.getElementsByTagName("TABLENAME").Item(0)
    If objElem Is Nothing Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingXMLTableName
    End If
    rstrTable = objElem.firstChild.Text
    If Len(rstrTable) = 0 Then
        '"Missing Table Description in XMLClass Definition"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingTableDesc
    End If
    If objXmlDoc.getElementsByTagName(rstrTable).Length < 1 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, "Expected " & rstrTable & " tag"
    End If
    Dim strPairSeparator As String
    Select Case vsftSQLFormatType ' add error checking
       Case sftFieldValueSeparated
            
            If StrComp(rstrFields, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "Fields must be specified for this format type"
            End If
            If StrComp(rstrValues, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "Values must be specified for this format type"
            End If
        Case sftAndSeparated
            If StrComp(rstrFieldValuePair, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "FieldValuePair must be specified for this format type"
            End If
            strPairSeparator = " AND "
        Case sftCommaSeparated
            If StrComp(rstrFieldValuePair, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "FieldValuePair must be specified for " & sftFieldValueSeparated
            End If
            strPairSeparator = " , "
    End Select
        
    Select Case vcdkKey
        Case cdkPRIMARYKEY
            Set xmlNodeList = objXmlClassDefDoc.getElementsByTagName("PRIMARYKEY")
        Case cdkOTHERS
            Set xmlNodeList = objXmlClassDefDoc.getElementsByTagName("OTHERS")
        Case Else
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InValidKey, _
            "Of " & vcdkKey & " for " & rstrTable
    End Select
    Dim intLoop As Integer
    Dim strValue As String
    Dim strFieldName As String
    Dim strDataTypeValue As String
    'Loop round all the Elements
    For intLoop = 0 To (xmlNodeList.Length - 1)
              
        Set objElem = xmlNodeList.Item(intLoop)
                
        strFieldName = "" 'reset variables
        strValue = ""       'reset variables
        strDataTypeValue = ""   'reset variables
        'Get the field Description from the XMLClass Description
        strFieldName = objElem.firstChild.Text
                
        If Len(strFieldName) = 0 Then 'check the element
            'No description in provided XML
            '"Missing Key Description in XMLClass Definition"
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingKeyDesc
        End If
        'Combine the field elements to be used in the Select statement
        'MCS 19/11/99
        If Len(rstrSelect) > 0 Then
            rstrSelect = rstrSelect & "," & rstrTable & "." & strFieldName
        Else
            rstrSelect = rstrSelect & rstrTable & "." & strFieldName
        End If
        If objXmlDoc.getElementsByTagName(strFieldName).Length <> 0 Then
            ' we have a field to process
            strValue = objXmlDoc.getElementsByTagName(strFieldName).Item(0).Text
            If (Trim$(strValue) = "") And (vcdkKey = cdkPRIMARYKEY) Then
                ' Primary key value cannot be null
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInValidKeyValue, _
                    "Of " & vcdkKeyValue & " for " & strValue
            End If
            Select Case vcdkKey
                Case cdkPRIMARYKEY
                    Set xmlTypeElem = objXmlClassDefDoc.getElementsByTagName( _
                        "PRIMARYKEY").Item(intLoop)
                Case cdkOTHERS
                    Set xmlTypeElem = objXmlClassDefDoc.getElementsByTagName( _
                        "OTHERS").Item(intLoop)
                Case Else
                    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InValidKey, _
                   "Of " & vcdkKey & " for " & strValue
            End Select
            'Get the Type Description from the XMLClass Description
            strDataTypeValue = m_objXmlAssist.GetTagValue(xmlTypeElem, "TYPE")
            If Len(strDataTypeValue) = 0 Then
                '"Missing Type Description Value in XMLClass Definition"
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingTypeDesc
            End If
            ' format the data according to its data type
            If strValue = "" Then
                strValue = "Null"
            Else
                strValue = objSQLAssist.GetFormattedValue( _
                    strValue, GetDbType(strDataTypeValue))
            End If
                
            Select Case vsftSQLFormatType
                Case sftFieldValueSeparated
                    If Len(rstrFields) > 0 Then
                        rstrFields = rstrFields & ", "
                        rstrValues = rstrValues & ", "
                    End If
                    rstrFields = rstrFields & strFieldName
                    rstrValues = rstrValues & strValue
                        
               Case sftAndSeparated, sftCommaSeparated
                    
                    If Len(rstrFieldValuePair) > 0 Then
                        rstrFieldValuePair = rstrFieldValuePair & _
                                            strPairSeparator
                    End If
                    rstrFieldValuePair = rstrFieldValuePair & _
                                        strFieldName & "=" & strValue
                Case Else
                ' fixme
                'Raise error
            End Select
        Else
            'No Element in XML passed in
            Select Case vcdkKeyValue
                Case cdkaALLKEYS
                    'Only an error if we want all of them
                    '"Not ALL the Elements for the Keys specified have been supplied"
                    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                        omiga4MissingElement, "Table name: " & rstrTable
                Case cdkaANYKEYS
                    'do nothing no error (skip this field)
                Case Else
                    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                        omiga4InValidKeyValue, _
                        "Of " & vcdkKeyValue & " for " & strValue
                End Select
        End If
    Next
        
    Set xmlTypeElem = Nothing
    Set objXmlDoc = Nothing
    Set objXmlClassDefDoc = Nothing
    Set objSQLAssist = Nothing
    Set objElem = Nothing
    Set xmlNodeList = Nothing
    Exit Sub
BuildSQLStringVBErr:
    Set xmlTypeElem = Nothing
    Set objXmlDoc = Nothing
    Set objXmlClassDefDoc = Nothing
    Set objSQLAssist = Nothing
    Set objElem = Nothing
    Set xmlNodeList = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Sub Delete(ByVal vstrXMLRequest As String, ByVal vstrClassDef As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance to be deleted
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors:
'       omiga4RecordNotFound
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    Dim strFunctionName As String
    strFunctionName = "Delete"
            
    Dim strSQL As String
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, , , , strFieldValuePair
            
    strSQL = "delete from " & strTable & " where " & strFieldValuePair
    objADOAssist.ExecuteSQLCommand strSQL
        
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteVbErr:
    Set objADOAssist = Nothing
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub DeleteAll(ByVal vstrXMLRequest As String, ByVal vstrClassDef As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    Dim strFunctionName As String
    strFunctionName = "DeleteAll"

    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair
                    
    strSQL = "delete from " & strTable & " where " & strFieldValuePair
    objADOAssist.ExecuteSQLCommand (strSQL)
    Set objADOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteAllVbErr:
        
    If Err.Number = omiga4NoRowsAffected Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoRowsAffectedByDeleteAll
    End If
    Set objADOAssist = Nothing
        
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function FindList(ByVal vstrXMLRequest As String, _
                         ByVal vstrClassDef As String, _
                         Optional ByVal vstrOrderByField As String = cstrNotSpecified) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest      xml Request data stream containing data to which identifies
'                       the instance(s) of the persistant data to be retrieved
'   vstrClassDef        xml Class Def to parse vstrXMLRequest against
'   vstrOrderByField    the field name to order the search result by (optional)
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    Dim strFunctionName As String
    strFunctionName = "FindList"

    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    Dim strXML As String
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    Dim objADOAssist As ADOAssist
        
    Dim strSelect As String
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair, , , strSelect
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, , cdkOTHERS, cdkaANYKEYS, strFieldValuePair, , , strSelect
        
    If Len(strSelect) = 0 Then
        'should never get here but just incase
        strSelect = "*"
    End If
    If Len(strFieldValuePair) = 0 Then 'Find All
        'MCS 19/11/99 - only slect the required fields
        strSQL = "select " & strSelect & " from " & strTable
    Else
        strSQL = "select " & strSelect & " from " & strTable & " where " & strFieldValuePair
    End If
    'AY 13/10/1999 - if an order by field name has been passed in, add to the SQL string
    If StrComp(vstrOrderByField, cstrNotSpecified, vbTextCompare) <> 0 Then
        strSQL = strSQL & " order by " & vstrOrderByField
    End If
'   #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName _
'        & "[Before Dim cmd As New ADODB.Command]", True
'    #End If
    
    Dim cmd As New ADODB.Command
        
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    WriteSQLToFile strSQL, strFunctionName
    Set rstThisRecSet = cmd.Execute(, , adCmdText)
                
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'        Set objLogAssist2 = Nothing
'    #End If
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
        
    Set objXmlElem = objXmlOut.createElement(strTable & "LIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlElem)
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' rstThisRecSet.MoveFirst
    ' loop through the record set
    While Not rstThisRecSet.EOF
        strXML = Me.GetXMLFromRecordSet(rstThisRecSet, vstrClassDef)
        Set objXmlDoc = m_objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlListNode.appendChild(objXmlDoc.documentElement)
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    FindList = objXmlOut.xml

    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlListNode = Nothing
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
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlListNode = Nothing

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
'FIXIT: Declare 'vstrOrderByField' with an early-bound data type                           FixIT90210ae-R1672-R1B8ZE
Public Function FindListMultiple(ByVal vstrXMLRequest As String, _
                                 ByVal vstrClassDef As String, _
                                 Optional ByVal vstrOrderByField As Variant) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
'
'   Note that this method differs from FindList in that it can take multiple search criteria 'blocks'
'   For example, if searching for GROUPCONNECTION records, the passed XML could include more than
'   one 'blocks' of GROUPCONNECTION XML
'
' pass:
'   vstrXMLRequest      xml Request data stream containing data to which identifies
'                       the instance(s) of the persistant data to be retrieved
'   vstrClassDef        xml Class Def to parse vstrXMLRequest against
'   vstrOrderByField    the field name to order the search result by (optional)
' return:
'   FindListMultiple        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindListMultipleVbErr
    Dim strFunctionName As String
    strFunctionName = "FindListMultiple"

    Dim strFieldValuePair As String
    Dim strSQL As String
    Dim strXML As String
    Dim objXmlIn              As FreeThreadedDOMDocument40, _
        objXMLClassDefinition As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objADOAssist As ADOAssist
        
    Dim strSelect As String, _
        strTable  As String, _
        strWhere  As String
    Dim rst As ADODB.Recordset
    Dim cmd As New ADODB.Command
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    strSelect = ""
    strTable = ""
    strWhere = ""
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXMLClassDefinition = m_objXmlAssist.load(vstrClassDef, TypeName(Me), strFunctionName)
    ' Get table name
    Set objXmlNodeList = objXMLClassDefinition.getElementsByTagName("TABLENAME")
    If objXmlNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTable = objXmlNodeList.Item(0).firstChild.Text
    End If
    ' Loop through each search criteria block in turn, adding the criteria to tbe WHERE clause
    Set objXmlNodeList = objXmlIn.getElementsByTagName(strTable)
    For Each objXmlNode In objXmlNodeList
        strFieldValuePair = ""
        ' Get the part of the WHERE clause corresponding to this criteria block
        BuildSQLString objXmlNode.xml, vstrClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair, , , strSelect
        BuildSQLString objXmlNode.xml, vstrClassDef, strTable, , cdkOTHERS, cdkaANYKEYS, strFieldValuePair, , , strSelect
        strFieldValuePair = "(" & strFieldValuePair & ")"
        If (Len(strWhere) > 0) Then strWhere = strWhere & " OR "
        strWhere = strWhere & strFieldValuePair
    Next objXmlNode
        
    If Len(strSelect) = 0 Then
        'should never get here but just incase
        strSelect = "*"
    End If
    strSQL = "select " & strSelect & _
             " from " & strTable & _
             IIf(Len(strWhere) > 0, " where " & strWhere, "") & _
             IIf(Not IsMissing(vstrOrderByField), " order by " & CStr(vstrOrderByField), "")
    '
    ' THE REST OF THIS ROUTINE IS THE SAME AS FOR FINDLIST
    '
'    #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName _
'        & "[cmd.ActiveConnection = ...]", True
'    #End If
        
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    WriteSQLToFile strSQL, strFunctionName
    Set rst = cmd.Execute(, , adCmdText)
                
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'        Set objLogAssist2 = Nothing
'    #End If
    
    If rst.EOF Then
        ' raise application error to be interpreted by calling object
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
        
    Set objXmlElem = objXmlOut.createElement(strTable & "LIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlElem)
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rst.MoveFirst
    ' loop through the record set
    While Not rst.EOF
        strXML = Me.GetXMLFromRecordSet(rst, vstrClassDef)
        Set objXmlDoc = m_objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlListNode.appendChild(objXmlDoc.documentElement)
        rst.MoveNext
    Wend
        rst.Close
    FindListMultiple = objXmlOut.xml
    Set rst = Nothing
    Set cmd = Nothing
    Set objXmlIn = Nothing
    Set objXMLClassDefinition = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlListNode = Nothing
    Set objXmlNodeList = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
            
    Exit Function
FindListMultipleVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
    Set rst = Nothing
    Set cmd = Nothing
    Set objXmlIn = Nothing
    Set objXMLClassDefinition = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlListNode = Nothing
    Set objXmlNodeList = Nothing
        
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
Public Sub Create(ByVal vstrXMLRequest As String, ByVal vstrClassDef As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    Dim strFunctionName As String
    strFunctionName = "Create"
        
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As New SQLAssist
    Dim strSQL As String        ' SQL insert statement
    Dim strFields As String
    Dim strValues As String
    Dim strTable As String
        
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, sftFieldValueSeparated _
        , , , , strFields, strValues
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, sftFieldValueSeparated _
        , cdkOTHERS, cdkaANYKEYS, , strFields, strValues
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
            
    strSQL = "INSERT INTO " & strTable & " (" & strFields & ") values ( " & strValues & ")"
    objADOAssist.ExecuteSQLCommand strSQL
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
CreateVbErr:
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
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

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub Update(ByVal vstrXMLRequest As String, ByVal vstrClassDef As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    Dim strFunctionName As String
    strFunctionName = "Update"

    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As New SQLAssist
     
    Dim strSQL As String        ' SQL insert statement
    Dim strAndSeparated  As String
    Dim strCommaSeparated As String
    Dim strTable As String
           
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, sftAndSeparated, , , strAndSeparated
        
    BuildSQLString vstrXMLRequest, vstrClassDef, strTable, sftCommaSeparated, _
        cdkOTHERS, cdkaANYKEYS, strCommaSeparated
    strSQL = "UPDATE " & strTable & " SET " & strCommaSeparated & " where " & strAndSeparated
    objADOAssist.ExecuteSQLCommand strSQL
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
UpdateVbErr:
    
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
        
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
End Sub
Public Function GetXMLFromRecordSet(ByVal vrstRecordSet As ADODB.Recordset, ByVal vstrXMLClassDef As String) As String
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
    Dim strFunctionName As String
    strFunctionName = "GetXMLFromRecordSet"
    Dim objSQLAssist As New SQLAssist
    Dim objXmlOut    As New FreeThreadedDOMDocument40
    Dim objXmlClassDefElem As IXMLDOMElement
    Dim objXmlClassDefDoc As New FreeThreadedDOMDocument40
    Dim objXmlTableNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim fld As ADODB.Field
    'Dim objComboDO As ComboDO
    Dim strTableName As String
    Dim ERRNUMBER As Long

    'TK 23/08/2004 BBG1211 End
'    If m_objContext Is Nothing Then
'        Set objComboDO = New ComboDO
'    Else
'        Set objComboDO = m_objContext.CreateInstance(App.Title & ".ComboDO")
'    End If
    'TK 23/08/2004 BBG1211 End
    Set objXmlClassDefDoc = m_objXmlAssist.load(vstrXMLClassDef, TypeName(Me), strFunctionName)
    'get the number of elements for this class def
    Set objXmlClassDefElem = objXmlClassDefDoc.getElementsByTagName("TABLENAME").Item(0)
    If objXmlClassDefElem Is Nothing Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingXMLTableName
    End If
    strTableName = objXmlClassDefElem.firstChild.Text
    If Len(strTableName) = 0 Then
        '"Missing Table Description in XMLClass Definition"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingTableDesc
    End If
        
    Set objXmlTableNode = objXmlOut.appendChild(objXmlOut.createElement(strTableName))
    Dim intField As Integer
    Dim strElementName As String
    Dim strFormatted As String, strFormatMask As String, strComboGroup As String
    Dim eDbType As DBDATATYPE
    Dim objElemNode As IXMLDOMNode
    Dim bExists As Boolean
    If objXmlClassDefElem.childNodes.Length < 1 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoFieldsFound, _
        "For " & strTableName
    End If
    For intField = 1 To (objXmlClassDefElem.childNodes.Length - 1)
        'Get the field out
        Set objElemNode = objXmlClassDefElem.childNodes.Item(intField)
        If objElemNode Is Nothing Then
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoFieldItemFound, _
            "For " & strTableName
        End If
        'This is the field (PRIMARYKEY or OTHER as defined in ClassDef)
        strElementName = objElemNode.childNodes.Item(0).Text
        If Len(strElementName) = 0 Then
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoFieldItemName, _
            "For " & strElementName
        End If
        bExists = False
        Dim strDataTypeValue As String
        'Get the Data type out, this is the TYPE as defined under (PRIMARYKEY or OTHER in ClassDef)
        strDataTypeValue = m_objXmlAssist.GetTagValue(objElemNode, "TYPE")
            
            eDbType = GetDbType(strDataTypeValue) 'This is the element name for the type we want
                      
            If eDbType <> dbdtNotStored Then
                
                Set objXmlElem = objXmlOut.createElement(strElementName)
                If Not IsNull(vrstRecordSet.Fields.Item(strElementName).Value) Then
                    
                    ' set element text; also set element attribute where applicable
                    Select Case eDbType
                    Case dbdtGuid
                        objXmlElem.Text = objSQLAssist.GuidToString(vrstRecordSet.Fields.Item(strElementName).Value)
                    Case dbdtCurrency
                        strFormatted = Format(vrstRecordSet.Fields.Item(strElementName), "0.0000000")
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
                        objXmlElem.Text = Left(strFormatted, InStr(1, strFormatted, ".") + 2)
                    Case dbdtDouble
                        'Get the  format mask out
                        strFormatMask = m_objXmlAssist.GetTagValue(objElemNode, "FORMATMASK", bExists)
                        If Not bExists Then
                            'No FORMAT MASK defined so just use the default mask
'                            objXmlElem.Text = vrstRecordSet.Fields.Item(strElementName).Value
                            objXmlElem.setAttribute "RAW", vrstRecordSet.Fields.Item(strElementName).Value
                            objXmlElem.Text = Format(vrstRecordSet.Fields.Item(strElementName).Value, "0.00")
                        Else 'This is the FORMAT MASK  value as defined under (TYPE in ClassDef)
                            'Add the defined FORMATMASK
                            objXmlElem.setAttribute "RAW", vrstRecordSet.Fields.Item(strElementName).Value
                            objXmlElem.Text = Format(vrstRecordSet.Fields.Item(strElementName).Value, strFormatMask)
                        End If
                    Case dbdtComboId
                        'Get the  combo entry out
                        strComboGroup = m_objXmlAssist.GetTagValue(objElemNode, "COMBO", bExists)
                        'trap no records found
                        If Not bExists Then
                            'No combo entry
                            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoComboTagValue, _
                            "For " & strElementName
                        Else 'This is the COMBO value as defined under (COMBO in ClassDef)
                            'Add the defined COMBO value
                            Dim strComboText As String
                            'TK 23/08/2004 BBG1211 Start
                                If m_objComboDO Is Nothing Then
                                    If m_objContext Is Nothing Then
                                        Set m_objComboDO = New ComboDO
                                    Else
                                        Set m_objComboDO = m_objContext.CreateInstance(App.Title & ".ComboDO")
                                    End If
                                End If
                                'TK 23/08/2004 BBG1211 End
                            strComboText = m_objComboDO.GetComboText(strComboGroup, vrstRecordSet.Fields.Item(strElementName).Value)
                            If ERRNUMBER = 0 Then
                                objXmlElem.setAttribute "TEXT", strComboText
                                objXmlElem.Text = vrstRecordSet.Fields.Item(strElementName).Value
                            End If
                            ERRNUMBER = 0 'reset the error number
                        End If
                    Case dbdtDate
                        objXmlElem.Text = objSQLAssist.DateToString(vrstRecordSet.Fields.Item(strElementName).Value)
                    Case dbdtDateTime
                        objXmlElem.Text = objSQLAssist.DateTimeToString(vrstRecordSet.Fields.Item(strElementName).Value)
                    Case dbdtNotStored
                        ' skip this field (shouldn't get here anyway)
                    Case Else
                        objXmlElem.Text = vrstRecordSet.Fields.Item(strElementName).Value
                    End Select
                End If
                ' append the element whether or not the field has a null value
                Set objXmlNode = objXmlTableNode.appendChild(objXmlElem)
            Else
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InValidDataTypeValue, _
                    "Of " & strDataTypeValue & " for " & strElementName
            End If      ' end If eDbType <> dbdtNotStored
    Next
    '   add derived values to generated XML in base Class
    GetXMLFromRecordSet = objXmlOut.xml
    'Set objComboDO = Nothing
    Set objXmlClassDefElem = Nothing
    Set objXmlClassDefDoc = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlTableNode = Nothing
    Set objXmlOut = Nothing
    Set objElemNode = Nothing
    Set fld = Nothing
    Exit Function
GetXMLFromRecordSetVbErr:
    
    If Err.Number = omiga4RecordNotFound Then
        ERRNUMBER = Err.Number
        Resume Next
    End If
       
    'Set objComboDO = Nothing
    Set objXmlClassDefElem = Nothing
    Set objXmlClassDefDoc = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlTableNode = Nothing
    Set objXmlOut = Nothing
    Set objElemNode = Nothing
    Set fld = Nothing
        
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

    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetDbType(ByVal vstrDataTypeValue As String) As DBDATATYPE
' header ----------------------------------------------------------------------------------
' description:
'   Get the datatype (DBDATATYPE, an enumerated list defined in SQLAssist) associated
'   with a string description of the data type.
' pass:
'   vstrDataTypeValue
'       String description of the required datatype, e.g. "dbdtString"
' return:
'   enumerated value
'------------------------------------------------------------------------------------------
On Error GoTo GetDbTypeVbErr
    Dim strFunctionName As String
    strFunctionName = "GetDbType"

    Select Case vstrDataTypeValue
        Case "dbdtString"
            GetDbType = dbdtString
        Case "dbdtInt"
            GetDbType = dbdtInt
        Case "dbdtDouble"
            GetDbType = dbdtDouble
        Case "dbdtDate"
            GetDbType = dbdtDate
        Case "dbdtGuid"
            GetDbType = dbdtGuid
        Case "dbdtCurrency"
            GetDbType = dbdtCurrency
        Case "dbdtComboId"
            GetDbType = dbdtComboId
        Case "dbdtBoolean"
            GetDbType = dbdtBoolean
        Case "dbdtLong"
            GetDbType = dbdtLong
        Case "dbdtDateTime"
            GetDbType = dbdtDateTime
        Case "dbdtNotStored"
            GetDbType = dbdtNotStored
        Case Else
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                omiga4InvalidParameter, "Invalid datatype value: " & vstrDataTypeValue
    End Select
        
    Exit Function
GetDbTypeVbErr:
    ' RF 20/10/99 This was causing a type mismatch error
    'GetDbType = ""
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub GetKeyString( _
    ByVal vstrXMLRequest As String, _
    ByVal vstrXMLClassDef As String, _
    ByVal strKey As String, _
    ByVal strType As String, _
    ByRef strFieldValuePair As String, _
    ByRef strTable As String _
    )
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
    
    Dim strFunctionName As String
    strFunctionName = "GetKeyString"
    '------------------------------------------------------------------------------------------
    ' This function has been replaced by BuildSQLString
    '------------------------------------------------------------------------------------------
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
GetKeyStringVBErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetNextSequenceNumber( _
                    ByVal vstrXMLRequest As String, _
                    ByVal vstrClassDefinition As String, _
                    ByVal vstrSequenceField As String) As Long
' header ----------------------------------------------------------------------------------
' description:  Determines the next sequence number for a field given a set of primary key
'               values to search on.
'
'               The primary key is determined from the class definition parameter and the values
'               for that key from the XML request parameter. A SQL statement is then built
'               to find the maximum value for the specified sequence field on the appropriate table
'               given the specified key values.
'
' pass:         vstrXMLRequest          (in) String
'               vstrClassDefinition     (in) String
'               vstrSequenceField       (in) String
' return:       Long
' History:
' DM  14/11/02 BMIDS00935 Added nolock hint
' SDS 20/01/04 LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
'------------------------------------------------------------------------------------------
On Error GoTo GetNextSequenceNumberVbErr
    
    Dim strFunctionName As String
        strFunctionName = "GetNextSequenceNumber"
    Dim objSQLAssist As New SQLAssist, _
        objADOAssist As New ADOAssist
    Dim xmlIn              As FreeThreadedDOMDocument40, _
        xmlClassDefinition As FreeThreadedDOMDocument40
    Dim xmlKeyNodeList   As IXMLDOMNodeList, _
        xmlValueNodeList As IXMLDOMNodeList
    Dim xmlKeyNode As IXMLDOMNode
    Dim rst As ADODB.Recordset, _
        cmd As New ADODB.Command
    Dim strTableName As String, _
        strValue As String, _
        strType As String
    Dim strSQL As String, _
        strWhere As String
    Dim blnEndOfKey As Boolean
    Dim strSqlNoLock As String ' DM 14/11/02 BMIDS00935
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlClassDefinition = m_objXmlAssist.load(vstrClassDefinition, TypeName(Me), strFunctionName)
    ' Get table name
    Set xmlValueNodeList = xmlClassDefinition.getElementsByTagName("TABLENAME")
    If xmlValueNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTableName = xmlValueNodeList.Item(0).firstChild.Text
    End If
    ' Get the primary key nodes from the class definition
    Set xmlKeyNodeList = xmlClassDefinition.getElementsByTagName("PRIMARYKEY")
    If xmlKeyNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "No keys found"
    End If
    '
    ' Build WHERE clause from primary key values
    '
    strWhere = ""
    blnEndOfKey = False
    For Each xmlKeyNode In xmlKeyNodeList
        If xmlKeyNode.firstChild.Text = vstrSequenceField Then
            ' Field in primary key matching sequence field has been found; so stop building WHERE clause here
            blnEndOfKey = True
            strSqlNoLock = m_objXmlAssist.GetTagValue(xmlKeyNode, "SQLNOLOCK") ' DM 14/11/02 BMIDS00935
        End If
        If Not blnEndOfKey Then
            ' Get type of field
            strType = m_objXmlAssist.GetTagValue(xmlKeyNode, "TYPE")
            ' Get node in the data XML corresponding to this key
            Set xmlValueNodeList = xmlIn.getElementsByTagName(xmlKeyNode.firstChild.Text)
            If xmlValueNodeList.Length <> 0 Then
                If xmlKeyNodeList.Length = 0 Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Incomplete key specified"
                End If
                strValue = xmlValueNodeList.Item(0).Text
                ' AND value to WHERE clause
                If strWhere <> "" Then
                    strWhere = strWhere & " AND "
                End If
                strWhere = strWhere & "(" & xmlKeyNode.firstChild.Text & _
                     " = " & objSQLAssist.GetFormattedValue(strValue, GetDbType(strType)) & ")"
            End If
        End If
    Next xmlKeyNode
    '
    ' Execute the SQL
    '
    strSQL = "SELECT MAX(" & vstrSequenceField & ") FROM " & strTableName
    ' DM 14/11/02 BMIDS00935 BEGIN
    If strSqlNoLock = "TRUE" And objADOAssist.GetDBEngine = SQLServer Then ' SDS LIVE00009653 / 20/01/2004
        strSQL = strSQL & " WITH (NOLOCK)"
    End If
    ' DM 14/11/02 BMIDS00935 END
    If strWhere <> "" Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    WriteSQLToFile strSQL, strFunctionName
    Set rst = cmd.Execute(1, , adCmdText)
    If IsNull(rst.Fields(0).Value) Then
        ' No records exist yet, so start sequence at 1
        GetNextSequenceNumber = 1
    Else
        GetNextSequenceNumber = rst.Fields(0).Value + 1
    End If
    rst.Close
    Set rst = Nothing
GetNextSequenceNumberExit:
    Set xmlIn = Nothing
    Set xmlClassDefinition = Nothing
    Set xmlKeyNode = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNodeList = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
GetNextSequenceNumberVbErr:
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
    Set xmlIn = Nothing
    Set xmlClassDefinition = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNodeList = Nothing
    Set xmlKeyNode = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function IDOAssist_GetNextSequenceNumber( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    ByVal vxmlClassDefinition As FreeThreadedDOMDocument40, _
                    ByVal vstrTableName As String, _
                    ByVal vstrSequenceField As String) As Long
' header ----------------------------------------------------------------------------------
' description:  Determines the next sequence number for a field given a set of primary key
'               values to search on.
'
'               The primary key is determined from the class definition parameter and the values
'               for that key from the XML request parameter. A SQL statement is then built
'               to find the maximum value for the specified sequence field on the appropriate table
'               given the specified key values.
'
' pass:         vstrXMLRequest          (in) String
'               vstrClassDefinition     (in) String
'               vstrSequenceField       (in) String
' return:       Long
' History:
' DM 14/11/02  BMIDS00935 Added nolock hint
' SDS 20/01/04 LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
'------------------------------------------------------------------------------------------
On Error GoTo IDOAssist_GetNextSequenceNumberVbErr
    
    Dim strFunctionName As String
        strFunctionName = "IDOAssist_GetNextSequenceNumber"
    Dim objSQLAssist As New SQLAssist, _
        objADOAssist As New ADOAssist
    Dim xmlKeyNodeList   As IXMLDOMNodeList, _
        xmlValueNodeList As IXMLDOMNodeList
    Dim xmlKeyNode As IXMLDOMNode
    Dim rst As ADODB.Recordset, _
        cmd As New ADODB.Command
    Dim strValue As String, _
        strType As String
    Dim strSQL As String, _
        strWhere As String
    Dim blnEndOfKey As Boolean
    Dim strSqlNoLock As String ' DM 14/11/02 BMIDS00935
    ' Get the primary key nodes from the class definition
    Set xmlKeyNodeList = vxmlClassDefinition.documentElement.selectNodes("PRIMARYKEY")
    If xmlKeyNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "No keys found"
    End If
    '
    ' Build WHERE clause from primary key values
    '
    strWhere = ""
    blnEndOfKey = False
    For Each xmlKeyNode In xmlKeyNodeList
        If xmlKeyNode.firstChild.Text = vstrSequenceField Then
            ' Field in primary key matching sequence field has been found; so stop building WHERE clause here
            blnEndOfKey = True
            strSqlNoLock = m_objXmlAssist.GetTagValue(xmlKeyNode, "SQLNOLOCK") ' DM 14/11/02 BMIDS00935
        End If
        If Not blnEndOfKey Then
            ' Get type of field
            strType = m_objXmlAssist.GetTagValue(xmlKeyNode, "TYPE")
            ' Get node in the data XML corresponding to this key
            Set xmlValueNodeList = vxmlRequest.getElementsByTagName(xmlKeyNode.firstChild.Text)
            If xmlValueNodeList.Length <> 0 Then
                If xmlKeyNodeList.Length = 0 Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Incomplete key specified"
                End If
                strValue = xmlValueNodeList.Item(0).Text
                ' AND value to WHERE clause
                If strWhere <> "" Then
                    strWhere = strWhere & " AND "
                End If
                strWhere = strWhere & "(" & xmlKeyNode.firstChild.Text & _
                     " = " & objSQLAssist.GetFormattedValue(strValue, GetDbType(strType)) & ")"
            End If
        End If
    Next xmlKeyNode
    '
    ' Execute the SQL
    '
    strSQL = "SELECT MAX(" & vstrSequenceField & ") FROM " & vstrTableName
    ' DM 14/11/02 BMIDS00935 BEGIN
    If strSqlNoLock = "TRUE" And objADOAssist.GetDBEngine = SQLServer Then ' SDS LIVE00009653 / 20/01/2004
        strSQL = strSQL & " WITH (NOLOCK)"
    End If
    ' DM 14/11/02 BMIDS00935 END
    If strWhere <> "" Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    WriteSQLToFile strSQL, strFunctionName
    Set rst = cmd.Execute(1, , adCmdText)
    If IsNull(rst.Fields(0).Value) Then
        ' No records exist yet, so start sequence at 1
        IDOAssist_GetNextSequenceNumber = 1
    Else
        IDOAssist_GetNextSequenceNumber = rst.Fields(0).Value + 1
    End If
    rst.Close
    Set rst = Nothing
IDOAssist_GetNextSequenceNumberExit:

    Set xmlKeyNode = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNodeList = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Exit Function
IDOAssist_GetNextSequenceNumberVbErr:
    
    ' PSC 23/01/01 SYS1871 - Start
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
    ' PSC 23/01/01 SYS1871 - End
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set xmlKeyNodeList = Nothing
    Set xmlValueNodeList = Nothing
    Set xmlKeyNode = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Sub GenerateSequenceNumber( _
                    ByVal rxmlDoc As FreeThreadedDOMDocument40, _
                    ByVal vstrClassDefinition As String, _
                    ByVal vstrSequenceField As String)
' header ----------------------------------------------------------------------------------
' description:  Determines the next sequence number for a field given a set of primary key
'               values to search on.
'
'               The primary key is determined from the class definition parameter and the values
'               for that key from the XML request parameter. A SQL statement is then built
'               to find the maximum value for the specified sequence field on the appropriate table
'               given the specified key values.
'
' pass:         vstrXMLRequest          (in) String
'               vstrClassDefinition     (in) String
'               vstrSequenceField       (in) String
' return:       Long
'------------------------------------------------------------------------------------------
On Error GoTo GetNextSequenceNumberVbErr
    
    Dim strFunctionName As String
        strFunctionName = "GetNextSequenceNumber"
    Dim xmlClassDefinition As FreeThreadedDOMDocument40
    Dim xmlNodeList        As IXMLDOMNodeList, _
        xmlTableElement    As IXMLDOMElement, _
        xmlSequenceNode    As IXMLDOMNode
    Dim strTableName As String
    Set xmlClassDefinition = m_objXmlAssist.load(vstrClassDefinition, TypeName(Me), strFunctionName)
    ' Get table name
    Set xmlNodeList = xmlClassDefinition.getElementsByTagName("TABLENAME")
    If xmlNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTableName = xmlNodeList.Item(0).firstChild.Text
    End If
    ' Find table root node in the XML
    Set xmlNodeList = rxmlDoc.getElementsByTagName(strTableName)
    If xmlNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, strTableName & " tag not found"
    Else
        Set xmlTableElement = xmlNodeList.Item(0)
    End If
    ' Generate the sequence number element if it does not already exist
    Set xmlNodeList = xmlTableElement.getElementsByTagName(vstrSequenceField)
    If xmlNodeList.Length = 0 Then
        ' No sequence number tag exists yet, so create one now
        Set xmlSequenceNode = xmlTableElement.appendChild(rxmlDoc.createElement(vstrSequenceField))
    Else
        Set xmlSequenceNode = xmlNodeList.Item(0)
    End If
    ' Insert new sequence number into XML
    xmlSequenceNode.Text = GetNextSequenceNumber(rxmlDoc.xml, vstrClassDefinition, vstrSequenceField)
GetNextSequenceNumberExit:
    Set xmlClassDefinition = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElement = Nothing
    Set xmlSequenceNode = Nothing
    Exit Sub
GetNextSequenceNumberVbErr:
    Set xmlClassDefinition = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElement = Nothing
    Set xmlSequenceNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Sub GenerateSequenceNumberEx( _
                    ByVal rxmlElement As IXMLDOMElement, _
                    ByVal vxmlClassDefinition As FreeThreadedDOMDocument40, _
                    ByVal vstrSequenceField As String)
' header ----------------------------------------------------------------------------------
' description:  Determines the next sequence number for a field given a set of primary key
'               values to search on.
'
'               The primary key is determined from the class definition parameter and the values
'               for that key from the XML request parameter. A SQL statement is then built
'               to find the maximum value for the specified sequence field on the appropriate table
'               given the specified key values.
'
' pass:         vstrXMLRequest          (in) String
'               vstrClassDefinition     (in) String
'               vstrSequenceField       (in) String
' return:       Long
'------------------------------------------------------------------------------------------
    On Error GoTo GenerateSequenceNumberExVbErr
    Dim strFunctionName As String
        strFunctionName = "GenerateSequenceNumberEx"
    Dim xmlNode As IXMLDOMNode
    Dim xmlTableNode       As IXMLDOMNode, _
        xmlSequenceNode    As IXMLDOMNode
    Dim strTableName As String
    ' Get table name
    Set xmlNode = vxmlClassDefinition.selectSingleNode("TABLENAME")
    If xmlNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTableName = xmlNode.firstChild.Text
    End If
    If rxmlElement.nodeName = strTableName Then
        Set xmlTableNode = rxmlElement
    Else
        Set xmlTableNode = rxmlElement.selectSingleNode(strTableName)
    End If
    If xmlTableNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, strTableName & " tag not found"
    End If
    ' Generate the sequence number element if it does not already exist
    Set xmlSequenceNode = xmlTableNode.selectSingleNode(vstrSequenceField)
    If xmlSequenceNode Is Nothing Then
        ' No sequence number tag exists yet, so create one now
        Set xmlSequenceNode = xmlTableNode.appendChild(rxmlElement.ownerDocument.createElement(vstrSequenceField))
    End If
    ' Insert new sequence number into XML
    xmlSequenceNode.Text = IDOAssist_GetNextSequenceNumber(rxmlElement, vxmlClassDefinition, strTableName, vstrSequenceField)
GenerateSequenceNumberExExit:
    Set xmlNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlSequenceNode = Nothing
    Exit Sub
GenerateSequenceNumberExVbErr:
    Set xmlNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlSequenceNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Private Sub WriteSQLToFile(ByVal vstrXML As String, _
                           ByVal vstrFunctionName As String)
' Header ----------------------------------------------------------------------------------
' Description:
'   Writes the passed SQL to a file on the C drive.
'
'   Note the conditional compilation constants around the code. These are present so that the
'   file is only created when specifically required.
'
' Pass:
'   vstrXML
'       The XML to write to the file.
' Returns:
'       N/A
' Raise Errors:
'------------------------------------------------------------------------------------------
End Sub
Public Function GetXMLFromRecordsetEx(ByVal vrstRecordSet As ADODB.Recordset, _
                                      ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                      Optional ByVal vxmlInNode As IXMLDOMNode = Nothing) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create XML data stream from field elements in ADO Recordset.
'   Add any values derived from field elements to XML.
' pass:
'   vrstRecordSet       ADO Recordset positioned at the required record
'   vxmlXMLClassDef     Class definition of recordset in the form of an XML document
'   vxmlInNode          Node to attach XML to
' return:               The node added
'
'------------------------------------------------------------------------------------------
On Error GoTo GetXMLForRecordsetVbErr
    Dim strFunctionName As String
    strFunctionName = "GetXMLForRecordset"
    Dim xmlPrimaryTableNode As IXMLDOMNode
    Dim xmlPrimaryDataNode  As IXMLDOMNode, _
        xmlDataNode         As IXMLDOMNode, _
        xmlViewDataNode     As IXMLDOMNode
    Dim xmlChildNode As IXMLDOMNode
    Dim blnFromView As Boolean
    Set xmlDataNode = Nothing
    Set xmlDataNode = GetXMLForTableSchema(vrstRecordSet, vxmlClassDef.documentElement, vxmlInNode)
    '
    ' Ascertain whether the data was retrieved via a view or not
    '
    blnFromView = True
    For Each xmlChildNode In vxmlClassDef.firstChild.childNodes
        If (xmlChildNode.baseName <> "") And (xmlChildNode.baseName <> "TABLENAME") Then
            ' Found a child node that does not represent a sub-table for a view - therefore the class definition
            ' must represent a physical table rather than a view
            blnFromView = False
        End If
    Next xmlChildNode
    '
    ' If the data was retrieved via a view then make sure that the 'primary' table data in the view replaces
    ' the view node in the data XML schema
    '
    If blnFromView Then
        ' The first table specified in the table class definition will be used as the 'root' data node in the returned
        ' XML. Any data for other tables in the schema will be created as child nodes of this primary data node
        Set xmlPrimaryTableNode = vxmlClassDef.documentElement.selectSingleNode("TABLENAME/PRIMARYKEY").parentNode
        Set xmlViewDataNode = xmlDataNode
        If Not (xmlPrimaryTableNode Is Nothing) Then
            Set xmlPrimaryDataNode = xmlViewDataNode.selectSingleNode(xmlPrimaryTableNode.firstChild.Text)
            If Not (xmlPrimaryDataNode Is Nothing) Then
                ' Move the primary table data up one level in the XML to replace the node named after the view used in the
                ' SQL query
                ' Move all children of the data (i.e. view) node to be children of the primary data node instead
                For Each xmlChildNode In xmlViewDataNode.childNodes
                    If Not (xmlChildNode Is xmlPrimaryDataNode) Then
                        xmlPrimaryDataNode.appendChild xmlChildNode
                    End If
                Next xmlChildNode
                If vxmlInNode Is Nothing Then
                    Set xmlDataNode = xmlPrimaryDataNode
                Else
                    ' Move the primary data node up one level in the XML
                    Set xmlDataNode = xmlDataNode.parentNode.appendChild(xmlPrimaryDataNode)
                    ' Remove the now empty view data node from the return XML
                    xmlDataNode.parentNode.removeChild xmlViewDataNode
                End If
            End If
        End If
    End If
    Set GetXMLFromRecordsetEx = xmlDataNode
    Set xmlPrimaryTableNode = Nothing
    Set xmlPrimaryDataNode = Nothing
    Set xmlChildNode = Nothing
    Set xmlViewDataNode = Nothing
    Set xmlDataNode = Nothing
    Exit Function
GetXMLForRecordsetVbErr:
    Set xmlPrimaryTableNode = Nothing
    Set xmlPrimaryDataNode = Nothing
    Set xmlChildNode = Nothing
    Set xmlViewDataNode = Nothing
    Set xmlDataNode = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetXMLForTableSchema(ByVal vrstRecordSet As ADODB.Recordset, _
                                      ByVal vxmlClassDef As IXMLDOMElement, _
                                      Optional ByVal vxmlInNode As IXMLDOMNode = Nothing) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create XML data stream from field elements in ADO Recordset.
'   Add any values derived from field elements to XML.
' pass:
'   vrstRecordSet       ADO Recordset positioned at the required record
'   vxmlXMLClassDef     Class definition of recordset in the form of an XML document
'   vxmlInNode          Node to attach XML to
' return:               The node added
'
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
' MDC   14/01/2003  BM0249 - Ignore SQLNOLOCK element in classdefs
'------------------------------------------------------------------------------------------
' BBG History
' TK    23/08/2004  BBG1211 (Extracted from BMIDS788) Only create ComboDO instance if actually needed
'------------------------------------------------------------------------------------------
On Error GoTo GetXMLForTableSchemaVbErr
    Dim strFunctionName As String
    strFunctionName = "GetXMLForTableSchema"
    Dim objSQLAssist As New SQLAssist
'    Dim objComboDO As ComboDO
    Dim xmlTableNode     As IXMLDOMNode, _
        xmlFieldDataElem As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim fld As ADODB.Field
    Dim strTableName As String, _
        strValue As String
    Dim ERRNUMBER As Long
    Dim blnSkipField As Boolean
    'TK 23/08/2004 BBG1211 Start
'    If m_objContext Is Nothing Then
'        Set objComboDO = New ComboDO
'    Else
'        Set objComboDO = m_objContext.CreateInstance(App.Title & ".ComboDO")
'    End If
    'TK 23/08/2004 BBG1211 End
    'BM0249 MDC 14/01/2003
    Dim xmlNoLockNode As IXMLDOMNode
    For Each xmlNoLockNode In vxmlClassDef.selectNodes("//SQLNOLOCK")
        xmlNoLockNode.parentNode.removeChild xmlNoLockNode
    Next
    'BM0249 MDC 14/01/2003 - End
    strTableName = vxmlClassDef.firstChild.Text
    If Len(strTableName) = 0 Then
        '"Missing Table Description in XMLClass Definition"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingTableDesc
    End If
        
    ' If a node hasn't been sent in then create a new document
    If vxmlInNode Is Nothing Then
        Set xmlDoc = New FreeThreadedDOMDocument40
        Set xmlTableNode = xmlDoc.appendChild(xmlDoc.createElement(strTableName))
    Else
        Set xmlDoc = vxmlInNode.ownerDocument
        Set xmlTableNode = vxmlInNode.appendChild(xmlDoc.createElement(strTableName))
    End If
    Dim strElementName As String
    Dim strFormatted As String, strFormatMask As String, strComboGroup As String
    Dim eDbType As DBDATATYPE
    Dim objElemNode As IXMLDOMNode
    Dim bExists As Boolean
    If vxmlClassDef.childNodes.Length < 1 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoFieldsFound, _
        "For " & strTableName
    End If
    blnSkipField = True ' This flag is used to skip the first (i.e. tablename) node in each table definition
    For Each objElemNode In vxmlClassDef.childNodes
        If Not blnSkipField Then
            'This is the field (PRIMARYKEY or OTHER as defined in ClassDef)
            strElementName = objElemNode.firstChild.Text
            If Len(strElementName) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoFieldItemName, _
                "For " & strElementName
            End If
            If objElemNode.nodeName = "TABLENAME" Then
                '
                ' Node represents a 'sub-table' part of the schema
                '
                GetXMLForTableSchema vrstRecordSet, objElemNode, xmlTableNode
            Else
                '
                ' Add the value for this field to the resulting XML
                '
                bExists = False
                Dim strDataTypeValue As String
                'Get the Data type out, this is the TYPE as defined under (PRIMARYKEY or OTHER in ClassDef)
                strDataTypeValue = objElemNode.childNodes(1).Text 'objXmlAssist.GetTagValue(objElemNode, "TYPE")
                    
                eDbType = GetDbType(strDataTypeValue) 'This is the element name for the type we want
                If eDbType <> dbdtNotStored Then
                    Set xmlFieldDataElem = xmlDoc.createElement(strElementName)
                    Set fld = vrstRecordSet.Fields.Item(strElementName)
                    If Not IsNull(fld.Value) Then
                        strValue = CStr(fld.Value)
                        ' set element text; also set element attribute where applicable
                        Select Case eDbType
                        Case dbdtGuid
                            xmlFieldDataElem.Text = objSQLAssist.GuidToString(fld.Value)
                        Case dbdtCurrency
                            strFormatted = Format$(strValue, "0.0000000")
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
                            xmlFieldDataElem.Text = Left(strFormatted, InStr(1, strFormatted, ".") + 2)
                        Case dbdtDouble
                            'Get the  format mask out
                            strFormatMask = m_objXmlAssist.GetTagValue(objElemNode, "FORMATMASK", bExists)
                            If Not bExists Then
                                'No FORMAT MASK defined so just use the default mask
                                xmlFieldDataElem.Text = strValue
                            Else 'This is the FORMAT MASK  value as defined under (TYPE in ClassDef)
                                'Add the defined FORMATMASK
                                xmlFieldDataElem.setAttribute "RAW", strValue
                                xmlFieldDataElem.Text = Format$(strValue, strFormatMask)
                            End If
                        Case dbdtComboId
                            'Get the  combo entry out
                            strComboGroup = m_objXmlAssist.GetTagValue(objElemNode, "COMBO", bExists)
                            'trap no records found
                            If Not bExists Then
                                'No combo entry
                                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoComboTagValue, _
                                "For " & strElementName
                            Else 'This is the COMBO value as defined under (COMBO in ClassDef)
                                'Add the defined COMBO value
                                Dim strComboText As String
                                'TK 23/08/2004 BBG1211 Start
                                If m_objComboDO Is Nothing Then
                                    If m_objContext Is Nothing Then
                                        Set m_objComboDO = New ComboDO
                                    Else
                                        Set m_objComboDO = m_objContext.CreateInstance(App.Title & ".ComboDO")
                                    End If
                                End If
                                'TK 23/08/2004 BBG1211 End
                                strComboText = m_objComboDO.GetComboText(strComboGroup, strValue)
                                If ERRNUMBER = 0 Then
                                    xmlFieldDataElem.setAttribute "TEXT", strComboText
                                    xmlFieldDataElem.Text = strValue
                                End If
                                ERRNUMBER = 0 'reset the error number
                            End If
                        Case dbdtDate
                            xmlFieldDataElem.Text = objSQLAssist.DateToString(strValue)
                        Case dbdtDateTime
                            xmlFieldDataElem.Text = objSQLAssist.DateTimeToString(strValue)
                        Case dbdtNotStored
                            ' skip this field (shouldn't get here anyway)
                        Case Else
                            xmlFieldDataElem.Text = strValue
                        End Select
                    End If
                        
                    ' append the element whether or not the field has a null value
                    xmlTableNode.appendChild xmlFieldDataElem
                Else
                    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                        omiga4InValidDataTypeValue, _
                        "Of " & strDataTypeValue & " for " & strElementName
                End If      ' end If eDbType <> dbdtNotStored
            End If
        End If
        blnSkipField = False
    Next
    Set GetXMLForTableSchema = xmlTableNode
     
'    Set objComboDO = Nothing
    Set xmlFieldDataElem = Nothing
    Set xmlTableNode = Nothing
    Set xmlDoc = Nothing
    Set fld = Nothing
    Set xmlNoLockNode = Nothing 'BM0249 MDC 14/01/2003
    Set objElemNode = Nothing
    Exit Function
GetXMLForTableSchemaVbErr:
    
    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
       
    'Add table and field name to the description
     If eDbType <> dbdtNotStored Then
        Err.Description = Err.Description & vbCrLf & "Table=""" & strTableName & """ Field=""" & strElementName & """"
    Else
        Err.Description = Err.Description & vbCrLf & "Table=""" & strTableName & """"
    End If
'    Set objComboDO = Nothing
    Set xmlFieldDataElem = Nothing
    Set xmlTableNode = Nothing
    Set xmlDoc = Nothing
    Set fld = Nothing
    Set xmlNoLockNode = Nothing 'BM0249 MDC 14/01/2003
    Set objElemNode = Nothing

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

    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub BuildSQLStringEx( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
    ByRef rstrTable As String, _
    Optional ByVal vsftSQLFormatType As SQL_FORMAT_TYPE = sftAndSeparated, _
    Optional ByVal vcdkKey As CLASS_DEF_KEY = cdkPRIMARYKEY, _
    Optional ByVal vcdkKeyValue As CLASS_DEF_KEY_AMOUNT = cdkaALLKEYS, _
    Optional ByRef rstrFieldValuePair As String = cstrNotSpecified, _
    Optional ByRef rstrFields As String = cstrNotSpecified, _
    Optional ByRef rstrValues As String = cstrNotSpecified, _
    Optional ByRef rstrSelect As String, _
    Optional ByVal vstrItemName As String = "")
' header ----------------------------------------------------------------------------------
' description:
'       Check all keys are present in the XML, and return them with their
'       associated values in a format ready for including in a SQL execution
'       string.
'       Raise an error dependant on the value CLASS_DEF_KEY_AMOUNT
'       if not all keys have values specified.
' pass:
'   vxmlRequest
'       xml request element data from which to build SQL string
'   vxmlClassDef
'       xml ClassDefinition describing the table, fields and formats
'       to be used to parse the vxmlRequest against
'   rstrTable
'       Base table name
'   vsftSQLFormatType
'       Enumerated SQL_FORMAT_TYPE used to determine the format of the
'       rstrFieldValuePair or rstrFields And rstrValues that are returned
'       sftFieldValueSeparated:
'           As used in create, e.g "insert into TABLE ( FieldA,FieldB )
'                                   values ( Value1,Value2 )"
'       sftCommaSeparated:
'           e.g. "update TABLE set FieldA = Value1,FieldB = Value2 where"
'       sftAndSeparated:
'           e.g. "select * from TABLE where FieldA = Value1 AND FieldB = Value2"
'       DEFAULT: sftAndSeparated
'   vcdkKey
'       Enumerated value CLASS_DEF_KEY. This and vcdkKeyValue (below) are
'       used together to Build SQL string and raise appropriate errors. E.g.:
'           GetData     cdkaALLKEYS cdkPRIMARYKEY
'           Delete      cdkaALLKEYS cdkPRIMARYKEY
'           DeleteAll   cdkaANYKEYS cdkPRIMARYKEY
'           FindList    cdkaANYKEYS cdkPRIMARYKEY
'           Update      cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
'           Create      cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
'       cdkPRIMARYKEY:
'           All primary key fields for given table
'       cdkOTHERS
'           Everything else that isn't a primary key
'       DEFAULT: cdkPRIMARYKEY
'   vcdkKeyValue
'       Enumerated value CLASS_DEF_KEY_AMOUNT - used as above with vcdkKey
'       to provide error handling depending on which combination is specified.
'       cdkaALLKEYS:
'           ALL key values expected
'       cdkaANYKEYS:
'           Any of the key values expected
'       DEFAULT: cdkaALLKEYS
'   rstrFieldValuePair
'       String created or appended to, which contains SQL formatted values
'       based on the settings of SQL_FORMAT_TYPE.
'       Used with sftAndSeparated and sftCommaSeparated; raises error if either of
'       these is specified and no parameter passed.
'       Used by GetData, Delete, DeleteAll, FindList.
'   rstrFields, rstrValues
'       String created or appended to, which contains SQL formatted values
'       based on the settings of SQL_FORMAT_TYPE.
'       Used with sftFieldValueSeparated; raises error if this is specified and no
'       parameter is passed.
'       Used by Create.
'   rstrSelect
'       Optional string used in Find, when used with a subset of the XML class/table
'       Definition. Will return a string of Table.Fields for the provided XML
' return:   n/a
' Raise Errors:
'       omiga4InvalidKeyString
'       omiga4InValidKeyValueSpecified
'       omiga4Invalidparameter
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo BuildSQLStringExVBErr:
    
    Dim strFunctionName As String
    strFunctionName = "BuildSQLStringEx"

    Dim objSQLAssist As New SQLAssist
    Dim xmlElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlTableElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlPrimaryTable As IXMLDOMNode
    Dim strItemName As String
    Set xmlElem = vxmlClassDef.getElementsByTagName("TABLENAME").Item(0)
    If xmlElem Is Nothing Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingXMLTableName
    End If
    If rstrTable = "" Then
        rstrTable = xmlElem.firstChild.Text
    End If
    ' The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
    strItemName = vstrItemName
    If strItemName = "" Then
        strItemName = rstrTable
    End If
    If Len(rstrTable) = 0 Then
        '"Missing Table Description in XMLClass Definition"
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingTableDesc
    End If
    ' The first table specified in the table class definition will be used to ascertain the key of the table/view
    ' being retrieved on
    Set xmlPrimaryTable = vxmlClassDef.documentElement.selectSingleNode(".//PRIMARYKEY").parentNode
    If vxmlRequest.nodeName = strItemName Then
        Set xmlTableElement = vxmlRequest
    Else
        Set xmlTableElement = vxmlRequest.selectSingleNode(strItemName)
    End If
    If xmlTableElement Is Nothing Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, "Expected " & strItemName & " tag"
    End If
    Dim strPairSeparator As String
    Select Case vsftSQLFormatType ' add error checking
       Case sftFieldValueSeparated
            
            If StrComp(rstrFields, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "Fields must be specified for this format type"
            End If
            If StrComp(rstrValues, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "Values must be specified for this format type"
            End If
        Case sftAndSeparated
            If StrComp(rstrFieldValuePair, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "FieldValuePair must be specified for this format type"
            End If
            strPairSeparator = " AND "
        Case sftCommaSeparated
            If StrComp(rstrFieldValuePair, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                    omiga4InvalidParameter, _
                    "FieldValuePair must be specified for " & sftFieldValueSeparated
            End If
            strPairSeparator = " , "
    End Select
        
    Select Case vcdkKey
        Case cdkPRIMARYKEY
            Set xmlNodeList = vxmlClassDef.documentElement.selectNodes(".//PRIMARYKEY")
        Case cdkOTHERS
            Set xmlNodeList = vxmlClassDef.documentElement.selectNodes(".//OTHERS")
        Case Else
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InValidKey, _
            "Of " & vcdkKey & " for " & rstrTable
    End Select
    Dim intLoop As Integer
    Dim strValue As String
    Dim strFieldName As String
    Dim strDataTypeValue As String
    'Loop round all the Elements
    For intLoop = 0 To (xmlNodeList.Length - 1)
              
        Set xmlElem = xmlNodeList.Item(intLoop)
                
        strFieldName = "" 'reset variables
        strValue = ""       'reset variables
        strDataTypeValue = ""   'reset variables
        'Get the field Description from the XMLClass Description
        strFieldName = xmlElem.firstChild.Text
                
        If Len(strFieldName) = 0 Then 'check the element
            'No description in provided XML
            '"Missing Key Description in XMLClass Definition"
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingKeyDesc
        End If
        'Combine the field elements to be used in the Select statement
        'MCS 19/11/99
        If Len(rstrSelect) > 0 Then
            rstrSelect = rstrSelect & "," & rstrTable & "." & strFieldName
        Else
            rstrSelect = rstrSelect & rstrTable & "." & strFieldName
        End If
        Set xmlNode = xmlTableElement.selectSingleNode(strFieldName)
        If Not (xmlNode Is Nothing) Then
            ' we have a field to process
            strValue = xmlNode.Text
            If (Trim$(strValue) = "") And (vcdkKey = cdkPRIMARYKEY) And (xmlElem.parentNode Is xmlPrimaryTable) Then
                ' Primary key value cannot be null
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInValidKeyValue, _
                    " NULL value for " & strFieldName
            End If
            'Get the Type Description from the XMLClass Description
            strDataTypeValue = m_objXmlAssist.GetTagValue(xmlElem, "TYPE")
            If Len(strDataTypeValue) = 0 Then
                '"Missing Type Description Value in XMLClass Definition"
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingTypeDesc
            End If
            ' format the data according to its data type
            If strValue = "" Then
                strValue = "Null"
            Else
                strValue = objSQLAssist.GetFormattedValue(strValue, GetDbType(strDataTypeValue))
            End If
                
            Select Case vsftSQLFormatType
                Case sftFieldValueSeparated
                    If Len(rstrFields) > 0 Then
                        rstrFields = rstrFields & ", "
                        rstrValues = rstrValues & ", "
                    End If
                    rstrFields = rstrFields & strFieldName
                    rstrValues = rstrValues & strValue
                        
               Case sftAndSeparated, sftCommaSeparated
                    
                    If Len(rstrFieldValuePair) > 0 Then
                        rstrFieldValuePair = rstrFieldValuePair & _
                                            strPairSeparator
                    End If
                    rstrFieldValuePair = rstrFieldValuePair & _
                                        strFieldName & "=" & strValue
                Case Else
                ' fixme
                'Raise error
            End Select
        Else
            'No Element in XML passed in
            Select Case vcdkKeyValue
                Case cdkaALLKEYS
                    If (xmlElem.parentNode Is xmlPrimaryTable) Then
                        'Only an error if we want all of them
                        '"Not ALL the Elements for the Keys specified have been supplied"
                        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                            omiga4MissingElement, "Table name: " & rstrTable
                    End If
                Case cdkaANYKEYS
                    'do nothing no error (skip this field)
                Case Else
                    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                        omiga4InValidKeyValue, _
                        "Of " & vcdkKeyValue & " for " & strValue
                End Select
        End If
    Next
    rstrTable = vxmlClassDef.documentElement.firstChild.Text

    Set objSQLAssist = Nothing
    Set xmlNode = Nothing
    Set xmlPrimaryTable = Nothing
    Set xmlElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElement = Nothing
    Exit Sub
BuildSQLStringExVBErr:
    Set objSQLAssist = Nothing
    Set xmlNode = Nothing
    Set xmlPrimaryTable = Nothing
    Set xmlElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElement = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Public Function GetDataEx(ByVal vxmlRequest As IXMLDOMElement, _
                          ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                          Optional ByVal vstrTableName As String = "", _
                          Optional ByVal vstrItemName As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' return:
'                   xml Node containing data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo GetDataExVbErr
    Dim strFunctionName As String
    strFunctionName = "GetDataEx"
        
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
                
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    Dim strItemName As String
    strTable = vstrTableName
    ' The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
    strItemName = vstrItemName
    If strItemName = "" Then
        strItemName = strTable
    End If
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, _
        rstrFieldValuePair:=strFieldValuePair, _
        vstrItemName:=strItemName
    ' build the full select statement
    strSQL = "select * from " & strTable & " where " & strFieldValuePair
'    #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName & _
'        "[ Before Dim cmd As New ADODB.Command ]", True
'    #End If
        
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    WriteSQLToFile strSQL, strFunctionName
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
            
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'    #End If
            
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        Set GetDataEx = Me.GetXMLFromRecordsetEx(rstThisRecSet, vxmlClassDef)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetDataExVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
        
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
Public Sub DeleteEx(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance to be deleted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'       omiga4RecordNotFound
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo DeleteExVbErr
    Dim strFunctionName As String
    strFunctionName = "DeleteEx"
            
    Dim strSQL As String
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, , , , strFieldValuePair
            
    strSQL = "delete from " & strTable & " where " & strFieldValuePair
    objADOAssist.ExecuteSQLCommand strSQL
        
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteExVbErr:
    Set objADOAssist = Nothing
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub DeleteAllEx(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllExVbErr
    Dim strFunctionName As String
    strFunctionName = "DeleteAllEx"
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair
                    
    strSQL = "delete from " & strTable & " where " & strFieldValuePair
    objADOAssist.ExecuteSQLCommand (strSQL)
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteAllExVbErr:
        
    If Err.Number = omiga4NoRowsAffected Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoRowsAffectedByDeleteAll
    End If
    Set objADOAssist = Nothing
        
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function FindListEx(ByVal vxmlRequest As IXMLDOMElement, _
                           ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                           Optional ByVal vstrOrderByField As String = cstrNotSpecified) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlRequest         xml Request data stream containing data to which identifies
'                       the instance(s) of the persistant data to be retrieved
'   vxmlClassDef        xml Class Def to parse vxmlRequest against
'   vstrOrderByField    the field name to order the search result by (optional)
' return:
'                       xml Node containing data retrieved
' History:
' SDS 20/01/04 LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
'------------------------------------------------------------------------------------------
On Error GoTo FindListExVbErr
    Dim strFunctionName As String
    strFunctionName = "FindListEx"

    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    Dim xmlListElem As IXMLDOMElement
    Dim objADOAssist As ADOAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
        
    Dim strSelect As String
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    ' SR 20/12/02 : BM0209 - Add 'With (NOLOCK)' to SQL statement, if required
    Dim strSqlNoLock As String
    If Not vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK") Is Nothing Then
        If objADOAssist.GetDBEngine = SQLServer Then ' SDS LIVE00009653 / 20/01/2004
            strSqlNoLock = " with (nolock)"
        End If
        vxmlClassDef.documentElement.removeChild vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK")
    End If
    ' SR 20/12/02 : BM0209 - End
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair, , , strSelect
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, , cdkOTHERS, cdkaANYKEYS, strFieldValuePair, , , strSelect
        
    If Len(strSelect) = 0 Then
        'should never get here but just incase
        strSelect = "*"
    End If
    If Len(strFieldValuePair) = 0 Then 'Find All
        'MCS 19/11/99 - only slect the required fields
        strSQL = "select " & strSelect & " from " & strTable & strSqlNoLock  ' SR 20/12/02 : BM0209 - add nolock hint
    Else
        strSQL = "select " & strSelect & " from " & strTable & strSqlNoLock & " where " & strFieldValuePair ' SR 20/12/02 : BM0209 - add nolock hint
    End If
    'AY 13/10/1999 - if an order by field name has been passed in, add to the SQL string
    If StrComp(vstrOrderByField, cstrNotSpecified, vbTextCompare) <> 0 Then
        strSQL = strSQL & " order by " & vstrOrderByField
    End If
'   #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName _
'        & "[Before Dim cmd As New ADODB.Command]", True
'    #End If
    
    Dim cmd As New ADODB.Command
        
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    WriteSQLToFile strSQL, strFunctionName
    Set rstThisRecSet = cmd.Execute(, , adCmdText)
                
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'        Set objLogAssist2 = Nothing
'    #End If
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
        
    Set xmlListElem = xmlOut.createElement(strTable & "LIST")
    xmlOut.appendChild xmlListElem
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rstThisRecSet.MoveFirst
    ' loop through the record set
    While Not rstThisRecSet.EOF
        Me.GetXMLFromRecordsetEx rstThisRecSet, vxmlClassDef, xmlListElem
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set FindListEx = xmlListElem
    Set objADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlListElem = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
            
    Exit Function
FindListExVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlListElem = Nothing

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
'FIXIT: Declare 'vstrOrderByField' with an early-bound data type                           FixIT90210ae-R1672-R1B8ZE
Public Function FindListMultipleEx(ByVal vxmlRequest As IXMLDOMElement, _
                                   ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                   Optional ByVal vstrOrderByField As Variant, _
                                   Optional ByVal vstrItemName As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
'
'   Note that this method differs from FindList in that it can take multiple search criteria 'blocks'
'   For example, if searching for GROUPCONNECTION records, the passed XML could include more than
'   one 'blocks' of GROUPCONNECTION XML
'
' pass:
'   vxmlRequest         xml Request data stream containing data to which identifies
'                       the instance(s) of the persistant data to be retrieved
'   vxmlClassDef        xml Class Def to parse vxmlRequest against
'   vstrOrderByField    the field name to order the search result by (optional)
' return:
'                       xml Node containing data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo FindListMultipleExVbErr
    Dim strFunctionName As String
    strFunctionName = "FindListMultipleEx"

    Dim sgl As Single
    Dim strFieldValuePair As String
    Dim strSQL As String
    Dim strXML As String
    Dim strItemName As String
    Dim intNodeListCount As Integer
    Dim xmlListElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim objADOAssist As ADOAssist
        
    Dim strSelect As String, _
        strTable  As String, _
        strWhere  As String, _
        strTempSelect As String
    Dim i As Integer
    Dim rst As ADODB.Recordset
    Dim cmd As New ADODB.Command
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    strSelect = ""
    strTable = ""
    strWhere = ""
    strTempSelect = ""
    ' Get table name
    Set xmlNodeList = vxmlClassDef.getElementsByTagName("TABLENAME")
    If xmlNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTable = xmlNodeList.Item(0).firstChild.Text
    End If
    Set xmlNodeList = Nothing
    ' The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
    strItemName = vstrItemName
    If strItemName = "" Then
        strItemName = strTable
    End If
    ' Loop through each search criteria block in turn, adding the criteria to tbe WHERE clause
    If vxmlRequest.nodeName = strItemName Then
        Set xmlNode = vxmlRequest
        intNodeListCount = 0
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strItemName)
        intNodeListCount = xmlNodeList.Length
    End If
    i = 0
    Do While Not (xmlNode Is Nothing) Or (i < intNodeListCount)
        If xmlNode Is Nothing Then
            ' This will happen every time except when the xmlNode is set prior to the While..Loop
            Set xmlNode = xmlNodeList.Item(i)
        End If
        strFieldValuePair = ""
        ' Get the part of the WHERE clause corresponding to this criteria block
        BuildSQLStringEx xmlNode, vxmlClassDef, strTable, _
            vcdkKeyValue:=cdkaANYKEYS, _
            rstrFieldValuePair:=strFieldValuePair, _
            rstrSelect:=strTempSelect, _
            vstrItemName:=strItemName
        BuildSQLStringEx xmlNode, vxmlClassDef, strTable, _
            vcdkKey:=cdkOTHERS, _
            vcdkKeyValue:=cdkaANYKEYS, _
            rstrFieldValuePair:=strFieldValuePair, _
            rstrSelect:=strTempSelect, _
            vstrItemName:=strItemName
        strFieldValuePair = "(" & strFieldValuePair & ")"
        If (Len(strWhere) > 0) Then strWhere = strWhere & " OR "
        strWhere = strWhere & strFieldValuePair
        If Len(strSelect) = 0 Then
            strSelect = strTempSelect
        End If
        Set xmlNode = Nothing
        i = i + 1
    Loop
        
    If Len(strSelect) = 0 Then
        'should never get here but just incase
        strSelect = "*"
    Else
        strSelect = "distinct " & strSelect
    End If
    strSQL = "select " & strSelect & _
             " from " & strTable & _
             IIf(Len(strWhere) > 0, " where " & strWhere, "") & _
             IIf(Not IsMissing(vstrOrderByField), " order by " & CStr(vstrOrderByField), "")
    '
    ' THE REST OF THIS ROUTINE IS THE SAME AS FOR FINDLIST
    '
'    #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName _
'        & "[cmd.ActiveConnection = ...]", True
'    #End If
        
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    WriteSQLToFile strSQL, strFunctionName
    Set rst = cmd.Execute(, , adCmdText)
                
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'        Set objLogAssist2 = Nothing
'    #End If
    
    If rst.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
        
    Set xmlListElem = xmlOut.createElement(strItemName & "LIST")
    xmlOut.appendChild xmlListElem
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rst.MoveFirst
    ' loop through the record set
    While Not rst.EOF
        Me.GetXMLFromRecordsetEx rst, vxmlClassDef, xmlListElem
        rst.MoveNext
    Wend
    rst.Close
    Set FindListMultipleEx = xmlListElem
    Set rst = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlListElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
            
    Exit Function
FindListMultipleExVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
    Set rst = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlListElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
        
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
Public Sub CreateEx(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlRequest     xml Request data stream containing data to be persisted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateExVbErr
    Dim strFunctionName As String
    strFunctionName = "CreateEx"
        
    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As New SQLAssist
    Dim strSQL As String        ' SQL insert statement
    Dim strFields As String
    Dim strValues As String
    Dim strTable As String
        
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, sftFieldValueSeparated _
        , , , , strFields, strValues
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, sftFieldValueSeparated _
        , cdkOTHERS, cdkaANYKEYS, , strFields, strValues
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
            
    strSQL = "INSERT INTO " & strTable & " (" & strFields & ") values ( " & strValues & ")"
    objADOAssist.ExecuteSQLCommand strSQL
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
CreateExVbErr:
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub UpdateEx(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlRequest     xml Request data stream containing data to be persisted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateExVbErr
    Dim strFunctionName As String
    strFunctionName = "UpdateEx"

    Dim objADOAssist As ADOAssist
    Dim objSQLAssist As New SQLAssist

    Dim strSQL As String        ' SQL insert statement
    Dim strAndSeparated  As String
    Dim strCommaSeparated As String
    Dim strTable As String
           
    If m_objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, sftAndSeparated, , , strAndSeparated
        
    BuildSQLStringEx vxmlRequest, vxmlClassDef, strTable, sftCommaSeparated, _
        cdkOTHERS, cdkaANYKEYS, strCommaSeparated
    strSQL = "UPDATE " & strTable & " SET " & strCommaSeparated & " where " & strAndSeparated
    objADOAssist.ExecuteSQLCommand strSQL
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
UpdateExVbErr:
    
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
        
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
End Sub
Private Sub IDOAssist_BuildSQLString( _
        ByVal vxmlRequest As IXMLDOMElement, _
        ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
        rstrTable As String, _
        Optional ByVal vsftSQLFormatType As SQL_FORMAT_TYPE = sftAndSeparated, _
        Optional ByVal vcdkKey As CLASS_DEF_KEY = cdkPRIMARYKEY, _
        Optional ByVal vcdkKeyValue As CLASS_DEF_KEY_AMOUNT = cdkaALLKEYS, _
        Optional rstrFieldValuePair As String = "Not specified", _
        Optional rstrFields As String = "Not specified", _
        Optional rstrValues As String = "Not specified", _
        Optional rstrSelect As String, _
        Optional ByVal vstrItemName As String = "")
' header ----------------------------------------------------------------------------------
' description:
'       Check all keys are present in the XML, and return them with their
'       associated values in a format ready for including in a SQL execution
'       string.
'       Raise an error dependant on the value CLASS_DEF_KEY_AMOUNT
'       if not all keys have values specified.
' pass:
'   vxmlRequest
'       xml request element data from which to build SQL string
'   vxmlClassDef
'       xml ClassDefinition describing the table, fields and formats
'       to be used to parse the vxmlRequest against
'   rstrTable
'       Base table name
'   vsftSQLFormatType
'       Enumerated SQL_FORMAT_TYPE used to determine the format of the
'       rstrFieldValuePair or rstrFields And rstrValues that are returned
'       sftFieldValueSeparated:
'           As used in create, e.g "insert into TABLE ( FieldA,FieldB )
'                                   values ( Value1,Value2 )"
'       sftCommaSeparated:
'           e.g. "update TABLE set FieldA = Value1,FieldB = Value2 where"
'       sftAndSeparated:
'           e.g. "select * from TABLE where FieldA = Value1 AND FieldB = Value2"
'       DEFAULT: sftAndSeparated
'   vcdkKey
'       Enumerated value CLASS_DEF_KEY. This and vcdkKeyValue (below) are
'       used together to Build SQL string and raise appropriate errors. E.g.:
'           GetData     cdkaALLKEYS cdkPRIMARYKEY
'           Delete      cdkaALLKEYS cdkPRIMARYKEY
'           DeleteAll   cdkaANYKEYS cdkPRIMARYKEY
'           FindList    cdkaANYKEYS cdkPRIMARYKEY
'           Update      cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
'           Create      cdkaALLKEYS cdkPRIMARYKEY Plus cdkaANYKEYS cdkOTHERS
'       cdkPRIMARYKEY:
'           All primary key fields for given table
'       cdkOTHERS
'           Everything else that isn't a primary key
'       DEFAULT: cdkPRIMARYKEY
'   vcdkKeyValue
'       Enumerated value CLASS_DEF_KEY_AMOUNT - used as above with vcdkKey
'       to provide error handling depending on which combination is specified.
'       cdkaALLKEYS:
'           ALL key values expected
'       cdkaANYKEYS:
'           Any of the key values expected
'       DEFAULT: cdkaALLKEYS
'   rstrFieldValuePair
'       String created or appended to, which contains SQL formatted values
'       based on the settings of SQL_FORMAT_TYPE.
'       Used with sftAndSeparated and sftCommaSeparated; raises error if either of
'       these is specified and no parameter passed.
'       Used by GetData, Delete, DeleteAll, FindList.
'   rstrFields, rstrValues
'       String created or appended to, which contains SQL formatted values
'       based on the settings of SQL_FORMAT_TYPE.
'       Used with sftFieldValueSeparated; raises error if this is specified and no
'       parameter is passed.
'       Used by Create.
'   rstrSelect
'       Optional string used in Find, when used with a subset of the XML class/table
'       Definition. Will return a string of Table.Fields for the provided XML
' return:   n/a
'------------------------------------------------------------------------------------------
On Error GoTo BuildSQLStringVBErr:
    
    Dim strFunctionName As String
    strFunctionName = "BuildSQLString"

    Dim objSQLAssist As New SQLAssist
    Dim xmlElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlTableElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlPrimaryTable As IXMLDOMNode
    Dim strItemName As String
    Set xmlElem = vxmlClassDef.getElementsByTagName("TABLENAME").Item(0)
    If xmlElem Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingXMLTableName
    End If
    If rstrTable = "" Then
        rstrTable = xmlElem.firstChild.Text
    End If
    ' The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
    strItemName = vstrItemName
    If strItemName = "" Then
        strItemName = rstrTable
    End If
    If Len(rstrTable) = 0 Then
        '"Missing Table Description in XMLClass Definition"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableDesc
    End If
    ' The first table specified in the table class definition will be used to ascertain the key of the table/view
    ' being retrieved on
    Set xmlPrimaryTable = vxmlClassDef.documentElement.selectSingleNode(".//PRIMARYKEY").parentNode
    If vxmlRequest.nodeName = strItemName Then
        Set xmlTableElement = vxmlRequest
    Else
        Set xmlTableElement = vxmlRequest.selectSingleNode(strItemName)
    End If
    If xmlTableElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "Expected " & strItemName & " tag"
    End If
    Dim strPairSeparator As String
    Select Case vsftSQLFormatType ' add error checking
       Case sftFieldValueSeparated
            
            If StrComp(rstrFields, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, _
                    "Fields must be specified for this format type"
            End If
            If StrComp(rstrValues, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, _
                    "Values must be specified for this format type"
            End If
        Case sftAndSeparated
            If StrComp(rstrFieldValuePair, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, _
                    "FieldValuePair must be specified for this format type"
            End If
            strPairSeparator = " AND "
        Case sftCommaSeparated
            If StrComp(rstrFieldValuePair, cstrNotSpecified, vbTextCompare) = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInvalidParameter, _
                    "FieldValuePair must be specified for " & sftFieldValueSeparated
            End If
            strPairSeparator = " , "
    End Select
        
    Select Case vcdkKey
        Case cdkPRIMARYKEY
            Set xmlNodeList = vxmlClassDef.documentElement.selectNodes(".//PRIMARYKEY")
        Case cdkOTHERS
            Set xmlNodeList = vxmlClassDef.documentElement.selectNodes(".//OTHERS")
        Case Else
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInValidKey, _
            "Of " & vcdkKey & " for " & rstrTable
    End Select
    Dim intLoop As Integer
    Dim strValue As String
    Dim strFieldName As String
    Dim strDataTypeValue As String
    'Loop round all the Elements
    For intLoop = 0 To (xmlNodeList.Length - 1)
              
        Set xmlElem = xmlNodeList.Item(intLoop)
                
        strFieldName = "" 'reset variables
        strValue = ""       'reset variables
        strDataTypeValue = ""   'reset variables
        'Get the field Description from the XMLClass Description
        strFieldName = xmlElem.firstChild.Text
                
        If Len(strFieldName) = 0 Then 'check the element
            'No description in provided XML
            '"Missing Key Description in XMLClass Definition"
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingKeyDesc
        End If
        'Combine the field elements to be used in the Select statement
        'MCS 19/11/99
        If Len(rstrSelect) > 0 Then
            rstrSelect = rstrSelect & "," & rstrTable & "." & strFieldName
        Else
            rstrSelect = rstrSelect & rstrTable & "." & strFieldName
        End If
        Set xmlNode = xmlTableElement.selectSingleNode(strFieldName)
        If Not (xmlNode Is Nothing) Then
            ' we have a field to process
            strValue = xmlNode.Text
            If (Trim$(strValue) = "") And (vcdkKey = cdkPRIMARYKEY) And (xmlElem.parentNode Is xmlPrimaryTable) Then
                ' Primary key value cannot be null
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                    oeInValidKeyValue, _
                    " NULL value for " & strFieldName
            End If
            'Get the Type Description from the XMLClass Description
            strDataTypeValue = m_objXmlAssist.GetTagValue(xmlElem, "TYPE")
            If Len(strDataTypeValue) = 0 Then
                '"Missing Type Description Value in XMLClass Definition"
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTypeDesc
            End If
            ' format the data according to its data type
            If strValue = "" Then
                strValue = "Null"
            Else
                strValue = objSQLAssist.GetFormattedValue(strValue, GetDbType(strDataTypeValue))
            End If
                
            Select Case vsftSQLFormatType
                Case sftFieldValueSeparated
                    If Len(rstrFields) > 0 Then
                        rstrFields = rstrFields & ", "
                        rstrValues = rstrValues & ", "
                    End If
                    rstrFields = rstrFields & strFieldName
                    rstrValues = rstrValues & strValue
                        
               Case sftAndSeparated, sftCommaSeparated
                    
                    If Len(rstrFieldValuePair) > 0 Then
                        rstrFieldValuePair = rstrFieldValuePair & _
                                            strPairSeparator
                    End If
                    rstrFieldValuePair = rstrFieldValuePair & _
                                        strFieldName & "=" & strValue
                Case Else
                ' fixme
                'Raise error
            End Select
        Else
            'No Element in XML passed in
            Select Case vcdkKeyValue
                Case cdkaALLKEYS
                    If (xmlElem.parentNode Is xmlPrimaryTable) Then
                        'Only an error if we want all of them
                        '"Not ALL the Elements for the Keys specified have been supplied"
                        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                            omiga4MissingElement, "Table name: " & rstrTable
                    End If
                Case cdkaANYKEYS
                    'do nothing no error (skip this field)
                Case Else
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                        oeInValidKeyValue, _
                        "Of " & vcdkKeyValue & " for " & strValue
                End Select
        End If
    Next
    rstrTable = vxmlClassDef.documentElement.firstChild.Text

    Set objSQLAssist = Nothing
    Set xmlNode = Nothing
    Set xmlPrimaryTable = Nothing
    Set xmlElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElement = Nothing
    Exit Sub
BuildSQLStringVBErr:
    Set objSQLAssist = Nothing
    Set xmlNode = Nothing
    Set xmlPrimaryTable = Nothing
    Set xmlElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElement = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Private Sub IDOAssist_Create(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlRequest     xml Request data stream containing data to be persisted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    Dim strFunctionName As String
    strFunctionName = "Create"

    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As New SQLAssist
    Dim strSQL As String        ' SQL insert statement
    Dim strFields As String
    Dim strValues As String
    Dim strTable As String
        
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, sftFieldValueSeparated _
        , , , , strFields, strValues
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, sftFieldValueSeparated _
        , cdkOTHERS, cdkaANYKEYS, , strFields, strValues
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
            
    strSQL = "INSERT INTO " & strTable & " (" & strFields & ") values ( " & strValues & ")"
    objIADOAssist.ExecuteSQLCommand strSQL
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
CreateVbErr:
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
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

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub IDOAssist_Delete(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance to be deleted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    Dim strFunctionName As String
    strFunctionName = "Delete"
            
    Dim strSQL As String
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim objIADOAssist As IADOAssist
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, , , , strFieldValuePair
            
    strSQL = "delete from " & strTable & " where " & strFieldValuePair
    objIADOAssist.ExecuteSQLCommand strSQL
        
    Set objIADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteVbErr:
    Set objIADOAssist = Nothing
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub IDOAssist_DeleteAll(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    Dim strFunctionName As String
    strFunctionName = "DeleteAll"

    Dim objIADOAssist As IADOAssist
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair
                    
    strSQL = "delete from " & strTable & " where " & strFieldValuePair
    objIADOAssist.ExecuteSQLCommand (strSQL)
    Set objIADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteAllVbErr:
        
    If m_objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNoRowsAffectedByDeleteAll
    End If
    Set objIADOAssist = Nothing
        
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function IDOAssist_FindList(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40, Optional ByVal vstrOrderByField As String = "Not specified") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlRequest         xml Request data stream containing data to which identifies
'                       the instance(s) of the persistant data to be retrieved
'   vxmlClassDef        xml Class Def to parse vxmlRequest against
'   vstrOrderByField    the field name to order the search result by (optional)
' return:
'                       xml Node containing data retrieved
' History:
' SDS 20/01/04 LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    Dim strFunctionName As String
    strFunctionName = "FindList"

    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    Dim xmlListElem As IXMLDOMElement
    Dim objIADOAssist As IADOAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
        
    Dim strSelect As String
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    ' SR 20/12/02 : BM0209 -  Add 'With (NOLOCK)' to SQL statement, if required
    Dim strSqlNoLock As String
    If Not vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK") Is Nothing Then
        If objIADOAssist.GetDBEngine = SQLServer Then ' SDS LIVE00009653 / 20/01/2004
            strSqlNoLock = " with (nolock)"
        End If
        vxmlClassDef.documentElement.removeChild vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK")
    End If
    ' SR 20/12/02 : BM0209 - End
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, , , cdkaANYKEYS, strFieldValuePair, , , strSelect
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, , cdkOTHERS, cdkaANYKEYS, strFieldValuePair, , , strSelect
        
    If Len(strSelect) = 0 Then
        'should never get here but just incase
        strSelect = "*"
    End If
    If Len(strFieldValuePair) = 0 Then 'Find All
        'MCS 19/11/99 - only slect the required fields
        strSQL = "select " & strSelect & " from " & strTable & strSqlNoLock  ' SR 20/12/02 : BM0209 - add nolock hint
    Else
        strSQL = "select " & strSelect & " from " & strTable & strSqlNoLock & " where " & strFieldValuePair  ' SR 20/12/02 : BM0209 - add nolock hint
    End If
    'AY 13/10/1999 - if an order by field name has been passed in, add to the SQL string
    If StrComp(vstrOrderByField, cstrNotSpecified, vbTextCompare) <> 0 Then
        strSQL = strSQL & " order by " & vstrOrderByField
    End If
'   #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName _
'        & "[Before Dim cmd As New ADODB.Command]", True
'    #End If
    
    Dim cmd As New ADODB.Command
        
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    WriteSQLToFile strSQL, strFunctionName
    Set rstThisRecSet = cmd.Execute(, , adCmdText)
                
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'        Set objLogAssist2 = Nothing
'    #End If
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
        
    Set xmlListElem = xmlOut.createElement(strTable & "LIST")
    xmlOut.appendChild xmlListElem
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rstThisRecSet.MoveFirst
    ' loop through the record set
    While Not rstThisRecSet.EOF
        IDOAssist_GetXMLFromRecordSet rstThisRecSet, vxmlClassDef, xmlListElem
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set IDOAssist_FindList = xmlListElem
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
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
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
        
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
'FIXIT: Declare 'vstrOrderByField' with an early-bound data type                           FixIT90210ae-R1672-R1B8ZE
Private Function IDOAssist_FindListMultiple(ByVal vxmlRequest As IXMLDOMElement, _
                                            ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                            Optional ByVal vstrOrderByField As Variant, _
                                            Optional ByVal vstrItemName As String = "", _
                                            Optional ByVal vstrAdditionalCodition As String = "", _
                                            Optional ByVal vstrAndOrToCodition As String = "AND") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
'
'   Note that this method differs from FindList in that it can take multiple search criteria 'blocks'
'   For example, if searching for GROUPCONNECTION records, the passed XML could include more than
'   one 'blocks' of GROUPCONNECTION XML
'
' pass:
'   vxmlRequest         - xml Request data stream containing data to which identifies
'                       the instance(s) of the persistant data to be retrieved
'   vxmlClassDef        - xml Class Def to parse vxmlRequest against
'   vstrOrderByField    - the field name to order the search result by (optional)
'   vstrItemName        - name of primary table (optional)
'   vstrAdditionalCodition - Additional Condition to be added to the one built with the values
'                            in the Request
'   vstrAndOrToCodition  - The word used (AND/OR) to add the additional condition to the
'                          one built with Request node
' return:
'                       xml Node containing data retrieved
' history
' RF    12/05/2003  BM0536 - Handle SQLNOLOCK: Add 'With (NOLOCK)' to SQL statement, if required
' SDS   20/01/2004  LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
'------------------------------------------------------------------------------------------
On Error GoTo FindListMultipleVbErr
    Dim strFunctionName As String
    strFunctionName = "FindListMultiple"

    Dim strFieldValuePair As String
    Dim strSQL As String
    Dim strItemName As String
    Dim intNodeListCount As Integer
    Dim xmlListElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim objIADOAssist As IADOAssist
        
    Dim strSelect As String, _
        strTable  As String, _
        strWhere  As String, _
        strTempSelect As String
    Dim i As Integer
    Dim rst As ADODB.Recordset
    Dim cmd As New ADODB.Command
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    strSelect = ""
    strTable = ""
    strWhere = ""
    strTempSelect = ""
    ' Get table name
    Set xmlNodeList = vxmlClassDef.getElementsByTagName("TABLENAME")
    If xmlNodeList.Length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTable = xmlNodeList.Item(0).firstChild.Text
    End If
    Set xmlNodeList = Nothing
    ' RF 12/05/2003 BM0536 Start
    Dim strSqlNoLock As String
    If Not vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK") Is Nothing Then
        If objIADOAssist.GetDBEngine = SQLServer Then ' SDS LIVE00009653 / 20/01/2004
            strSqlNoLock = " with (nolock)"
        End If
        vxmlClassDef.documentElement.removeChild vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK")
    End If
    ' RF 12/05/2003 BM0536 End
    ' The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
    strItemName = vstrItemName
    If strItemName = "" Then
        strItemName = strTable
    End If
    ' Loop through each search criteria block in turn, adding the criteria to tbe WHERE clause
    If vxmlRequest.nodeName = strItemName Then
        Set xmlNode = vxmlRequest
        intNodeListCount = 0
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strItemName)
        intNodeListCount = xmlNodeList.Length
    End If
    i = 0
    Do While Not (xmlNode Is Nothing) Or (i < intNodeListCount)
        If xmlNode Is Nothing Then
            ' This will happen every time except when the xmlNode is set prior to the While..Loop
            Set xmlNode = xmlNodeList.Item(i)
        End If
        strFieldValuePair = ""
        ' Get the part of the WHERE clause corresponding to this criteria block
        IDOAssist_BuildSQLString xmlNode, vxmlClassDef, strTable, _
            vcdkKeyValue:=cdkaANYKEYS, _
            rstrFieldValuePair:=strFieldValuePair, _
            rstrSelect:=strTempSelect, _
            vstrItemName:=strItemName
        IDOAssist_BuildSQLString xmlNode, vxmlClassDef, strTable, _
            vcdkKey:=cdkOTHERS, _
            vcdkKeyValue:=cdkaANYKEYS, _
            rstrFieldValuePair:=strFieldValuePair, _
            rstrSelect:=strTempSelect, _
            vstrItemName:=strItemName
        strFieldValuePair = "(" & strFieldValuePair & ")"
        If (Len(strWhere) > 0) Then strWhere = strWhere & " OR "
        strWhere = strWhere & strFieldValuePair
        If Len(strSelect) = 0 Then
            strSelect = strTempSelect
        End If
        Set xmlNode = Nothing
        i = i + 1
    Loop
    ' SR 14/03/00
    ' Add the Additional Condition, if required
    If vstrAdditionalCodition <> "" Then
        If Len(strWhere) <> 0 Then
            strWhere = strWhere & " " & vstrAndOrToCodition & " (" & vstrAdditionalCodition & ")"
        Else
            strWhere = strWhere & vstrAdditionalCodition
        End If
    End If
    If Len(strSelect) = 0 Then
        'should never get here but just incase
        strSelect = "*"
    End If
    ' RF 12/05/2003 BM0536 Start
    'strSQL = "select " & strSelect & _
    '         " from " & strTable & _
    '         IIf(Len(strWhere) > 0, " where " & strWhere, "") & _
    '         IIf(Not IsMissing(vstrOrderByField), " order by " & CStr(vstrOrderByField), "")
    strSQL = "select " & strSelect & _
             " from " & strTable & strSqlNoLock & _
             IIf(Len(strWhere) > 0, " where " & strWhere, "") & _
             IIf(Not IsMissing(vstrOrderByField), " order by " & CStr(vstrOrderByField), "")
    ' RF 12/05/2003 BM0536 End
    '
    ' THE REST OF THIS ROUTINE IS THE SAME AS FOR FINDLIST
    '
'    #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName _
'        & "[cmd.ActiveConnection = ...]", True
'    #End If
        
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    WriteSQLToFile strSQL, strFunctionName
    Set rst = cmd.Execute(, , adCmdText)
                
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'        Set objLogAssist2 = Nothing
'    #End If
    
    If rst.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
        
    Set xmlListElem = xmlOut.createElement(strItemName & "LIST")
    xmlOut.appendChild xmlListElem
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rst.MoveFirst
    ' loop through the record set
    While Not rst.EOF
        IDOAssist_GetXMLFromRecordSet rst, vxmlClassDef, xmlListElem
        rst.MoveNext
    Wend
        rst.Close
    Set IDOAssist_FindListMultiple = xmlListElem
    Set rst = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlListElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
            
    Exit Function
FindListMultipleVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
    Set rst = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlListElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
        
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
Private Function IDOAssist_GenerateSequenceNumber(ByVal rxmlElement As IXMLDOMElement, ByVal vxmlClassDefinition As FreeThreadedDOMDocument40, ByVal vstrSequenceField As String) As Long
' header ----------------------------------------------------------------------------------
' description:  Determines the next sequence number for a field given a set of primary key
'               values to search on.
'
'               The primary key is determined from the class definition parameter and the values
'               for that key from the XML request parameter. A SQL statement is then built
'               to find the maximum value for the specified sequence field on the appropriate table
'               given the specified key values.
'
' pass:         vstrXMLRequest          (in) String
'               vstrClassDefinition     (in) String
'               vstrSequenceField       (in) String
' return:       Long
'------------------------------------------------------------------------------------------
On Error GoTo GenerateSequenceNumberVbErr
    
    Dim strFunctionName As String
        strFunctionName = "GenerateSequenceNumber"
    Dim xmlNode As IXMLDOMNode
    Dim xmlTableNode    As IXMLDOMNode, _
        xmlSequenceNode As IXMLDOMNode
    Dim lngNextNumber As Long
    Dim strTableName As String
    ' Get table name
    Set xmlNode = vxmlClassDefinition.selectSingleNode("TABLENAME")
    If xmlNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingTableName, "TABLENAME tag not found"
    Else
        strTableName = xmlNode.firstChild.Text
    End If
    If rxmlElement.nodeName = strTableName Then
        Set xmlTableNode = rxmlElement
    Else
        Set xmlTableNode = rxmlElement.selectSingleNode(strTableName)
    End If
    If xmlTableNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, strTableName & " tag not found"
    End If
    ' Generate the sequence number element if it does not already exist
    Set xmlSequenceNode = xmlTableNode.selectSingleNode(vstrSequenceField)
    If xmlSequenceNode Is Nothing Then
        ' No sequence number tag exists yet, so create one now
        Set xmlSequenceNode = xmlTableNode.appendChild(rxmlElement.ownerDocument.createElement(vstrSequenceField))
    End If
    ' Insert new sequence number into XML
    lngNextNumber = IDOAssist_GetNextSequenceNumber(rxmlElement, vxmlClassDefinition, strTableName, vstrSequenceField)
    xmlSequenceNode.Text = CStr(lngNextNumber)
    IDOAssist_GenerateSequenceNumber = lngNextNumber
GenerateSequenceNumberExit:
    Set xmlTableNode = Nothing
    Set xmlSequenceNode = Nothing
    Exit Function
GenerateSequenceNumberVbErr:
    Set xmlTableNode = Nothing
    Set xmlSequenceNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function IDOAssist_GetData(ByVal vxmlRequest As IXMLDOMElement, _
                                   ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                   Optional ByVal vstrTableName As String = "", _
                                   Optional ByVal vstrItemName As String = "") As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' return:
'                   xml Node containing data retrieved
' History:
' DM 14/11/02 BMIDS00935 Added nolock hint
' SDS 20/01/04 LIVE00009653 - Code amended to make 'NOLOCK' to be used only for SQLServer sql statements
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    Dim strFunctionName As String
    strFunctionName = "GetData"
        
    Dim objIADOAssist As IADOAssist
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    Dim strSqlNoLock As String ' DM 14/11/02 BMIDS00935
    Dim strFieldValuePair As String
    Dim strTable As String
    Dim strSQL As String
    Dim strItemName As String
    strTable = vstrTableName
    ' The items (i.e. XML request blocks) to search for are defined either by the paramter or the table name
    strItemName = vstrItemName
    If strItemName = "" Then
        strItemName = strTable
    End If
    ' DM 14/11/02 BMIDS00935 BEGIN
    If Not vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK") Is Nothing Then
        If objIADOAssist.GetDBEngine = SQLServer Then ' SDS LIVE00009653 / 20/01/2004
            strSqlNoLock = " with (nolock)"
        End If
        vxmlClassDef.documentElement.removeChild vxmlClassDef.selectSingleNode("TABLENAME/SQLNOLOCK")
    End If
    ' DM 14/11/02 BMIDS00935 END
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, _
        rstrFieldValuePair:=strFieldValuePair, _
        vstrItemName:=strItemName
    ' build the full select statement
    ' DM 14/11/02 BMIDS00935 BEGIN
    'strSQL = "select * from " & strTable & " where " & strFieldValuePair
    strSQL = "select * from " & strTable & strSqlNoLock & " where " & strFieldValuePair
    ' DM 14/11/02 BMIDS00935 END
'    #If PROFILING Then
'        Dim objLogAssist2 As New LogAssist
'        objLogAssist2.StartTimer TypeName(Me) & "." & strFunctionName & _
'        "[ Before Dim cmd As New ADODB.Command ]", True
'    #End If
        
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Dim rstThisRecSet As New ADODB.Recordset
    WriteSQLToFile strSQL, strFunctionName
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
            
'    #If PROFILING Then
'        objLogAssist2.StopTimer
'    #End If
            
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        Set IDOAssist_GetData = IDOAssist_GetXMLFromRecordSet(rstThisRecSet, vxmlClassDef)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objIADOAssist = Nothing
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
    Set objIADOAssist = Nothing
        
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
Private Function IDOAssist_GetXMLFromRecordSet(ByVal vrstRecordSet As ADODB.Recordset, _
                                               ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                               Optional ByVal vxmlInNode As IXMLDOMNode = Nothing) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create XML data stream from field elements in ADO Recordset.
'   Add any values derived from field elements to XML.
' pass:
'   vrstRecordSet       ADO Recordset positioned at the required record
'   vxmlXMLClassDef     Class definition of recordset in the form of an XML document
'   vxmlInNode          Node to attach XML to
' return:               The node added
'
'------------------------------------------------------------------------------------------
On Error GoTo GetXMLForRecordsetVbErr
    Dim strFunctionName As String
    strFunctionName = "GetXMLForRecordset"
    Set IDOAssist_GetXMLFromRecordSet = GetXMLForTableSchema(vrstRecordSet, vxmlClassDef.documentElement, vxmlInNode)
    Exit Function
GetXMLForRecordsetVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub IDOAssist_GetXMLFromWholeRecordset(ByVal vrstRecordSet As ADODB.Recordset, _
                                               ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                               ByVal vxmlInNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Create XML data stream from all records in an ADO Recordset.
'   Add any values derived from field elements to XML.
' pass:
'   vrstRecordSet       ADO Recordset
'   vxmlXMLClassDef     Class definition of recordset in the form of an XML document
'   vxmlInNode          Node to attach XML to
' return:               n/a
'
'------------------------------------------------------------------------------------------
On Error GoTo GetXMLFromWholeRecordsetVbErr
    Dim strFunctionName As String
    strFunctionName = "GetXMLFromWholeRecordset"
    If vrstRecordSet Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "Invalid recordset"
    End If
    If vxmlClassDef Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "Invalid class definition"
    End If
    ' DM 15/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' Checked that none of the code that calls this method passes in a record set that is not at the
    ' begining.
    ' vrstRecordSet.MoveFirst
    While Not vrstRecordSet.EOF
        IDOAssist_GetXMLFromRecordSet vrstRecordSet, vxmlClassDef, vxmlInNode
        vrstRecordSet.MoveNext
    Wend
    Exit Sub
GetXMLFromWholeRecordsetVbErr:
    
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub IDOAssist_Update(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlRequest     xml Request data stream containing data to be persisted
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    Dim strFunctionName As String
    strFunctionName = "Update"

    Dim objIADOAssist As IADOAssist
    Dim objSQLAssist As New SQLAssist
     
    Dim strSQL As String        ' SQL insert statement
    Dim strAndSeparated  As String
    Dim strCommaSeparated As String
    Dim strTable As String
           
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
    End If
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, sftAndSeparated, , , strAndSeparated
        
    IDOAssist_BuildSQLString vxmlRequest, vxmlClassDef, strTable, sftCommaSeparated, _
        cdkOTHERS, cdkaANYKEYS, strCommaSeparated
    strSQL = "UPDATE " & strTable & " SET " & strCommaSeparated & " where " & strAndSeparated
    objIADOAssist.ExecuteSQLCommand strSQL
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
UpdateVbErr:
    
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
        
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
End Sub
Private Function IDOAssist_GetComponentData(ByVal vxmlRequest As IXMLDOMElement, _
                                    ByVal xmlClassDef As FreeThreadedDOMDocument40) As IXMLDOMNode
'-----------------------------------------------------------------------------------------
'Description :
'
'-----------------------------------------------------------------------------------------
On Error GoTo GetComponentDataVbErr
    Dim strFunctionName As String
    strFunctionName = "GetComponentData"
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlCriteria As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlClassTableNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlClonedClassDef As FreeThreadedDOMDocument40 ' Clone of original class def
    Dim xmlNewClassDef As New FreeThreadedDOMDocument40   ' Class Def to be passed to GetData method
    Dim strTableName As String
    Dim blnAddAllNodes As Boolean
    Dim strFieldName As String
    '   Create the top level nodes for new request
    Set xmlNewRequest = xmlOut.createElement("REQUEST")
    xmlOut.appendChild xmlNewRequest
    strTableName = xmlClassDef.firstChild.firstChild.Text
    Set xmlTableNode = xmlOut.createElement(strTableName)
    xmlNewRequest.appendChild xmlTableNode
    Set xmlClassTableNode = xmlNewClassDef.createElement("TABLENAME")
    xmlClassTableNode.Text = strTableName
    xmlNewClassDef.appendChild xmlClassTableNode
    '   Build all the PrimaryKey node in NewRequest, with the values mentioned in request
    Set xmlClonedClassDef = New FreeThreadedDOMDocument40
    Set xmlClonedClassDef.documentElement = xmlClassDef.documentElement.cloneNode(True)
    blnAddAllNodes = IIf(m_objXmlAssist.GetAttributeValue(vxmlRequest, "OUTPUT", "ALLFIELDS") = "1", _
                                        True, False)
    Set xmlCriteria = vxmlRequest.selectSingleNode(".//CRITERIA")
    Set xmlNodeList = xmlClonedClassDef.selectNodes("//PRIMARYKEY")
    For Each xmlNode In xmlNodeList
        strFieldName = xmlNode.firstChild.Text
        ' Add an element to the NewRequest
        Set xmlTempNode = xmlOut.createElement(strFieldName)
        xmlTempNode.Text = m_objXmlAssist.GetTagValue(xmlCriteria, strFieldName)
        xmlTableNode.appendChild xmlTempNode
        ' Add this element to New Class Definition, only if the AllNodes attrib is set to False
        If Not blnAddAllNodes Then
            xmlClassTableNode.appendChild xmlNode
        End If
    Next xmlNode
    ' Add the other fields to be queried to new class definition, only if the AllNodes
    ' attribute is set to False
    If Not blnAddAllNodes Then
        Set xmlNodeList = vxmlRequest.selectNodes(".//FIELD")
        For Each xmlNode In xmlNodeList
            strFieldName = xmlNode.firstChild.Text
            xmlClassTableNode.appendChild _
                    m_objXmlAssist.GetNodeFromClassDefByNodeValue(xmlClonedClassDef, strFieldName)
        Next xmlNode
    End If
    Set IDOAssist_GetComponentData = IDOAssist_GetData(xmlNewRequest, xmlNewClassDef)
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlCriteria = Nothing
    Set xmlTempNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlClassTableNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlClonedClassDef = Nothing
    Set xmlNewClassDef = Nothing
    Exit Function
     
GetComponentDataVbErr:
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlCriteria = Nothing
    Set xmlTempNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlClassTableNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlClonedClassDef = Nothing
    Set xmlNewClassDef = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' TK 23/08/2004 BBG1211 Start
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    ' n.b. m_objComboDO not initialised here - only where used
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objComboDO = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
' TK 23/08/2004 BBG1211 End
