VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CustomerRoleDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CustomerRoleDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for CustomerRole Data Object
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     07/02/00    Created
'SR     16/02/00    Added Order by clause in FindList
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'LD     07/11/00    Explicity close recordsets
'AD     03/10/01    Remove MoveFirst as it's a ForwardOnly Cursor. SYS2718
'HMA    13/09/04    BMIDS864  Remove ApplicationNumber and ApplicationFactFindNumber from the key in Verification table.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ICustomerRoleDO
Implements ObjectControl
 
'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objIDOAssist As IDOAssist

Private Const cstrTABLE_NAME = "CUSTOMERROLE"

Private Sub ICustomerRoleDO_Create(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "ICustomerRoleDO_Create"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

Private Sub ICustomerRoleDO_Delete(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "ICustomerRoleDO_Delete"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIADOAssist As IADOAssist
    'Dim objSqlAssist As New SQLAssist
    Dim objSqlAssist As SQLAssist
    Set objSqlAssist = New SQLAssist
    
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    ' delete the  related records from the table 'CUSTOMERRELATIONSHIP', if any
    Dim strCondition As String, strApplicationNumber As String, strAFFNumber As String
    Dim strCustomerNumber As String, strCustomrVersionNumber As String
    Dim strSQL As String
    
    strApplicationNumber = vxmlTableElement.selectSingleNode("APPLICATIONNUMBER").Text
    strAFFNumber = vxmlTableElement.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
    strCustomerNumber = vxmlTableElement.selectSingleNode("CUSTOMERNUMBER").Text
    strCustomrVersionNumber = vxmlTableElement.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
    
    strCondition = " APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplicationNumber) _
                   & " And APPLICATIONFACTFINDNUMBER = " & strAFFNumber _
                   & " And ((OWNERCUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustomerNumber) _
                   & " AND OWNERCUSTOMERVERSIONNUMBER = " & strCustomrVersionNumber & ")" _
                   & " OR (RELTOCUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustomerNumber) _
                   & " AND RELTOCUSTOMERVERSIONNUMBER = " & strCustomrVersionNumber & "))"
    
    strSQL = " DELETE FROM CUSTOMERRELATIONSHIP WHERE " & strCondition
    
    Err.Clear
    On Error Resume Next
    objIADOAssist.executeSQLCommand strSQL
    m_objErrAssist.SaveErr
    
    On Error GoTo DeleteVbErr
    If m_objErrAssist.LastErr <> 0 And _
        m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeNoRowsAffected Then
        m_objErrAssist.ReRaise
    End If
                 
    ' delete the verification records from the table 'VERFIFCATION', if any
    ' BMIDS864  Remove Application Number and Application Fact Find Number
    strCondition = "CUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustomerNumber) _
                   & " And CUSTOMERVERSIONNUMBER = " & strCustomrVersionNumber
    
    strSQL = " DELETE FROM VERIFICATION WHERE " & strCondition
    
    Err.Clear
    On Error Resume Next
    objIADOAssist.executeSQLCommand strSQL
    m_objErrAssist.SaveErr
    
    On Error GoTo DeleteVbErr
    If m_objErrAssist.LastErr <> 0 And _
        m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeNoRowsAffected Then
        m_objErrAssist.ReRaise
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objSqlAssist = Nothing
    Set objIADOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objSqlAssist = Nothing
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
End Sub

Private Function ICustomerRoleDO_FindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "ICustomerRoleDO_FindList"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim strOrderBy As String
         
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()
    
    strOrderBy = " CUSTOMERROLETYPE ASC, CUSTOMERORDER ASC "
    Set ICustomerRoleDO_FindList = m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc, strOrderBy)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

Private Function ICustomerRoleDO_GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ICustomerRoleDO_GetData"
      
                
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()
    Set ICustomerRoleDO_GetData = m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

Private Sub ICustomerRoleDO_Update(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "ICustomerRoleDO_Update"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
       
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

Private Function ICustomerRoleDO_FindOmigaBusinessForCustomer( _
                            ByVal vstrXMLRequest As String) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description: Get omiga business for a particular customer and customer version
'              combination
' pass:
'   vstrXMLRequest  string containing the xml Request containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        XML node representation of  data retrieved
'                   Response format:
'                        <OMIGABUSINESSLIST>
'                          <OMIGABUSINESS>
'                              <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                              <TYPEOFAPPLICATION></TYPEOFAPPLICATION>
'                              <BUSINESSTYPEINDICATOR></BUSINESSTYPEINDICATOR>
'                              <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                              <DATECREATED></DATECREATED>
'                              <PACKAGENUMBER></PACKAGENUMBER>
'                              <CORRESPONDENCESALUTATION></CORRESPONDENCESALUTATION>
'                              <AMOUNT></AMOUNT>
'                              <STATUS></STATUS>
'                          </OMIGABUSINESS>
'                        <OMIGABUSINESSLIST>
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindOmigaBusinessForCustomerVbErr

    Const strFunctionName As String = "ICustomerRoleDO_FindOmigaBusinessForCustomer"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlBusinessNode As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objIApplicationClassDef As IApplicationClassDef
    
    Dim strOrderBy As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    vstrXMLRequest = Replace(vstrXMLRequest, "CUSTOMERROLE", "FINDOMIGABUSINESSFORCUSTOMER")
    vstrXMLRequest = Replace(vstrXMLRequest, "/CUSTOMERROLE", "/FINDOMIGABUSINESSFORCUSTOMER")
    
    Set xmlDoc = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
     
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadFindOmigaBusinessForCustomerData()
    
    strOrderBy = " DATECREATED DESC, APPLICATIONNUMBER ASC," _
                 & " DATECREATED DESC, APPLICATIONNUMBER ASC "
    
    Set xmlNode = _
            objIDOAssist.FindListMultiple(xmlDoc.documentElement, xmlClassDefDoc, _
                                          strOrderBy)
    
    Set xmlNodeList = xmlNode.selectNodes(".//" & "FINDOMIGABUSINESSFORCUSTOMER")
    
    ' For each node add 'BusinessTypeIndicator' as 'A' and create a node setvalue for Amount
    Dim strAmount As String
    For Each xmlBusinessNode In xmlNodeList
        Set xmlElement = xmlDoc.createElement("BUSINESSTYPEINDICATOR")
        xmlElement.Text = "A"
        xmlBusinessNode.appendChild xmlElement
        
        Set xmlElement = xmlDoc.createElement("AMOUNT")
        strAmount = objXmlAssist.GetTagValue(xmlBusinessNode, "TOTALLOANAMOUNT")
        If strAmount = "" Then
            strAmount = objXmlAssist.GetTagValue(xmlBusinessNode, "AMOUNTREQUESTED")
        End If
        xmlElement.Text = strAmount
        xmlBusinessNode.appendChild xmlElement
    Next xmlBusinessNode
    
    Set ICustomerRoleDO_FindOmigaBusinessForCustomer = xmlNode

    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElement = Nothing
    Set xmlBusinessNode = Nothing
    Set objIApplicationClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

FindOmigaBusinessForCustomerVbErr:
    
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set objIApplicationClassDef = Nothing

    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICustomerRoleDO_FindMainApplicants( _
                                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Find the CustomerRole data for the first 2 applicants for an application
' pass:
'   vXMLRequest  xml Request node containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
'                   Format:
'                       <APPLICATIONNUMBER>
'                       </APPLICATIONNUMBER>
'                       <APPLICATIONFACTFINDNUMBER>
'                       </APPLICATIONFACTFINDNUMBER>
' return:
'   FindMainApplicants
'       string containing XML data stream representation of data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindMainApplicantsVbErr

    Const strFunctionName As String = "ICustomerRoleDO_FindMainApplicants"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim cmd As New ADODB.Command

    Dim objADOAssist As ADOAssist

    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    cmd.ActiveConnection = objADOAssist.GetConnStr

    Dim strSQL As String
    Dim strSearchCriteria As String

    ' build the where clause of the SQL expression
    strSearchCriteria = GetFindMainApplicantsSearchString(vxmlRequest)

    strSQL = "select * from " & cstrTABLE_NAME & " where " & strSearchCriteria

    ' build the full select statement
    cmd.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlNode As IXMLDOMNode

    Dim strXML As String

    Dim objXmlElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode

    Dim objIDOAssist As IDOAssist
    Dim objIClassDef As IApplicationClassDef
    
    Set objXmlElem = objXmlOut.createElement(cstrTABLE_NAME & "LIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlElem)

    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If

    ' AD 03/10/01 Remove the MoveFirst as it's a ForwardOnly Cursor. SYS2718.

    While Not rstThisRecSet.EOF
        ' convert record set to xml
        Set objIDOAssist = New DOAssist
        Set objIClassDef = New ApplicationClassDef
        objIDOAssist.GetXMLFromRecordset rstThisRecSet, objIClassDef.LoadCustomerRoleData, objXmlListNode
        
        rstThisRecSet.MoveNext
    Wend
        rstThisRecSet.Close

    Set ICustomerRoleDO_FindMainApplicants = objXmlOut.documentElement

    Set cmd = Nothing
        Set objXmlElem = Nothing
    Set objXmlListNode = Nothing
    Set objXmlAssist = Nothing
    Set rstThisRecSet = Nothing
    Set objErrAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

FindMainApplicantsVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set cmd = Nothing
    Set rstThisRecSet = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function GetFindMainApplicantsSearchString( _
                ByVal vxmlRequest As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the search string from the XML string passed. Used by FindMainApplicants.
' pass:
'   vXMLRequest      Format:
'                       <APPLICATIONNUMBER>
'                       </APPLICATIONNUMBER>
'                       <APPLICATIONFACTFINDNUMBER>
'                       </APPLICATIONFACTFINDNUMBER>
' return:
'   GetFindMainApplicantsSearchString
'       string containing SQL data stream representation of a WHERE clause
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetFindMainApplicantsSearchStringVbErr

    Const strFunctionName As String = "GetFindMainApplicantsSearchString"

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objSqlAssist As New SQLAssist
    
    Dim strSearchCriteria As String
    Dim strAppNo As String
    Dim strAppFactFindNumber As String

    strAppNo = objXmlAssist.GetTagValue(vxmlRequest, _
                                                    "APPLICATIONNUMBER")

    strAppFactFindNumber = objXmlAssist.GetTagValue(vxmlRequest, _
                                                    "APPLICATIONFACTFINDNUMBER")

    strSearchCriteria = _
        "APPLICATIONNUMBER = " & objSqlAssist.FormatString(strAppNo) & _
        " and APPLICATIONFACTFINDNUMBER = " & strAppFactFindNumber & _
        " and CUSTOMERROLETYPE = 1 and CUSTOMERORDER <= 2"

    GetFindMainApplicantsSearchString = strSearchCriteria

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSqlAssist = Nothing

    Exit Function

GetFindMainApplicantsSearchStringVbErr:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSqlAssist = Nothing

    GetFindMainApplicantsSearchString = ""

    ' re-raise the error back up the chain
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    
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
    Set m_objIDOAssist = Nothing

End Sub

