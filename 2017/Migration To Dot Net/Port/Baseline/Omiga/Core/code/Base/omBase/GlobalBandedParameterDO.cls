VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "GlobalBandedParameterDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      GlobalBandedParameterDO.cls
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
'Prog  Date     Description
'MCS    17/08/99 Created
'LD     07/11/00 Explicity close recordsets
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog  Date     Description
'TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------
' Table format:
'   e.g.
'VALUATIONFEESET                         NOT NULL FLOAT(63)
'STARTDATE                      NOT NULL DATE
'TYPEOFVALUATION                NOT NULL NUMBER(5)
'MAXIMUMVALUE                   NOT NULL NUMBER(10)
'LOCATION                       NOT NULL NUMBER(5)
'AMOUNT                         Number(10)
' *****************************************************************************
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
 
 '=============================================
 'Variable Declaration Section
 '=============================================
 '=============================================
 'Constant Declaration Section
 '=============================================
Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
' XML Request Format:
' TODO ************************************************************************
'   put in actual request details
'   e.g.
'   <CustomerTelephoneNumberList>
'       <CustomerTelephoneNumber>
'           <CustomerNumber>value</CustomerNumber>
'           <CustomerVersionNumber>value</CustomerVersionNumber>
'           <TelephoneNumber>value</TelephoneNumber>
'           <Usage>value</Usage>
'           <ContactTime>value</ContactTime>
'       </CustomerTelephoneNumber>
'       <CustomerTelephoneNumber>
'       ...
'       </CustomerTelephoneNumber>
'   </CustomerTelephoneNumberList>
' *****************************************************************************
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    Dim strFunctionName As String
    strFunctionName = "Create"
'    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
                
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
    objDOAssist.Create vstrXMLRequest, LoadData
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
CreateVbErr:
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
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
    Dim strFunctionName As String
    strFunctionName = "Delete"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objDOAssist As DOAssist
'    Dim objErrAssist As New ErrAssist
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
            
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
                
    objDOAssist.Delete vstrXMLRequest, LoadData
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteVbErr:
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
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
    Dim strFunctionName As String
    strFunctionName = "DeleteAll"
'    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
        
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
                    
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
                
    objDOAssist.DeleteAll vstrXMLRequest, LoadData
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteAllVbErr:
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
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
    Dim strFunctionName As String
    strFunctionName = "GetData"
'    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
            
    GetData = objDOAssist.GetData(vstrXMLRequest, LoadData)
    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetDataVbErr:
        
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetCurrentParameter(ByVal vstrParameterName As String, _
                                    ByVal vstrQualifier) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrParameterName  Name of the parameter
'   vstrQualifier      Qualifier to select appropriate band
'
' return:
'   GetCurrentParameter     string containing XML data stream representation of
'                           data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentParameterVbErr
    Dim strFunctionName As String
    strFunctionName = "GetCurrentParameter"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
'    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
    Dim objDOAssist As DOAssist
    Dim rstThisRecSet As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim objADOAssist As ADOAssist
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
                        
    Dim strSQL As String
    Dim strClassDef As String
    strClassDef = LoadData
    ' Build up the SQL string
    strSQL = "SELECT * FROM GLOBALBANDEDPARAMETER GBPARAM1 WHERE"
    strSQL = strSQL & " NAME = " & objSQLAssist.FormatString(vstrParameterName)
    strSQL = strSQL & " AND GBPARAM1.GBPARAMSTARTDATE ="
    strSQL = strSQL & " (SELECT MAX(GBPARAMSTARTDATE) FROM GLOBALBANDEDPARAMETER WHERE"
    strSQL = strSQL & " NAME = " & objSQLAssist.FormatString(vstrParameterName)
    strSQL = strSQL & " AND GBPARAMSTARTDATE <= " & objSQLAssist.GetSystemDate & ")"
    strSQL = strSQL & " AND GBPARAM1.HIGHBAND = "
    strSQL = strSQL & " (SELECT MIN(HIGHBAND) FROM GLOBALBANDEDPARAMETER WHERE"
    strSQL = strSQL & " NAME = " & objSQLAssist.FormatString(vstrParameterName)
    strSQL = strSQL & " AND HIGHBAND >= " & vstrQualifier
    strSQL = strSQL & " AND GBPARAMSTARTDATE = GBPARAM1.GBPARAMSTARTDATE)"
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
    If Not rstThisRecSet.EOF Then
        GetCurrentParameter = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strClassDef)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    rstThisRecSet.Close
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set objDOAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
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
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set objDOAssist = Nothing
    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objADOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
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
On Error GoTo FindListVbErr
    Dim strFunctionName As String
    strFunctionName = "FindList"
'    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
     
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(App.Title & ".DOAssist")
    End If
     
    FindList = objDOAssist.FindList(vstrXMLRequest, LoadData)
    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
'    Set objErrAssist = Nothing
    Set objDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
FindListVbErr:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    Dim strFunctionName As String
    strFunctionName = "Update"
'    Dim objErrAssist As New ErrAssist
'    Dim objDOAssist As DOAssist
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
'    If objContext Is Nothing Then
'        Set objDOAssist = New DOAssist
'    Else
'        Set objDOAssist = objContext.CreateInstance(App.Title & ".DOAssist")
'    End If
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    ' TODO ************************************************************************
    '   Un comment these lines to implement this function
    ' *****************************************************************************
       
    'Update = objDOAssist.Update(vstrXMLRequest, LoadData)
'    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
UpdateVbErr:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
'    Set objDOAssist = Nothing
'    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function LoadData()
'------------------------------------------------------------------------------------------
' NAME         NOT NULL VARCHAR2(30)
' STARTDATE    NOT NULL DATE
' HIGHBAND     NOT NULL NUMBER
' DESCRIPTION  VARCHAR2(255)
' AMOUNT       NUMBER
' PERCENTAGE   NUMBER
' MAXIMUM      NUMBER
' BOOLEAN      NUMBER(1)
' STRING       VARCHAR2(30)
'------------------------------------------------------------------------------------------
Dim XMLstr As String
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "GLOBALBANDEDPARAMETER"
    strXML = strXML & "<PRIMARYKEY>NAME<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>GBPARAMSTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>HIGHBAND<TYPE>dbdtDouble</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PERCENTAGE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUM<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BOOLEAN<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STRING<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
LoadData = strXML
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
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr
    Dim strFunctionName As String
    strFunctionName = "AddDerivedData"
'    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when this function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
    AddDerivedData = vstrXMLData
       
    Exit Function
AddDerivedDataVBErr:
    
    AddDerivedData = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
